import time
import os
import math
from datetime import date

from django.shortcuts import render, redirect, get_object_or_404
from django.http import FileResponse, Http404, JsonResponse
from django.conf import settings

from .models import UploadedReport
from .forms import UploadForm

from treatement import process_file


def _period_label(report):
    start = report.date_rapport
    end = report.date_fin or start
    if report.period_type == UploadedReport.PERIOD_DAY:
        return start.strftime('%d/%m/%Y')
    return f"Du {start.strftime('%d/%m/%Y')} au {end.strftime('%d/%m/%Y')}"


def home(request):
    if request.user.is_superuser:
        recent_reports = UploadedReport.objects.filter(processed=True)[:5]
    else:
        recent_reports = UploadedReport.objects.filter(processed=True, user=request.user)[:5]
    return render(request, 'reports/home.html', {'recent_reports': recent_reports})


def upload(request):
    if request.method == 'POST':
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            report = form.save(commit=False)
            report.original_filename = request.FILES['file'].name
            report.user = request.user  # ← associe l'utilisateur
            report.save()
            return redirect('process_report', pk=report.pk)
    else:
        form = UploadForm()
    return render(request, 'reports/upload.html', {'form': form})


def process_report(request, pk):
    report = get_object_or_404(UploadedReport, pk=pk)

    if report.processed:
        return redirect('results', pk=report.pk)

    if request.method == 'GET':
        return render(request, 'reports/processing.html', {'report': report})

    # POST → traitement réel
    start = time.time()
    input_path = report.file.path
    date_debut_str = report.date_rapport.strftime('%Y-%m-%d')
    date_fin_str = report.date_fin.strftime('%Y-%m-%d') if report.date_fin else None

    results_dir = os.path.join(settings.MEDIA_ROOT, 'results')
    os.makedirs(results_dir, exist_ok=True)
    output_path = os.path.join(results_dir, f'{report.pk}_detailed.xlsx')

    try:
        df_export, df_dedup, df_synthese = process_file(
            input_path, date_debut_str, date_fin=date_fin_str,
        )
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)

    df_export.to_excel(output_path, index=False)
    synthesis_path = output_path.replace('.xlsx', '_Synthese.xlsx')
    df_synthese.to_excel(synthesis_path, index=False)

    import pandas as pd
    df_original = pd.read_excel(input_path)
    report.total_rows = len(df_original)
    report.filtered_rows = len(df_export)

    try:
        total_row = df_synthese[df_synthese['Escalade'] == 'TOTAL']
        report.total_incidents = int(total_row['Inc count'].values[0]) if len(total_row) > 0 else 0
    except (IndexError, KeyError):
        report.total_incidents = 0

    unresolved = 0
    for _, row in df_synthese.iterrows():
        status = str(row.get('Status', ''))
        if 'Non resolu' in status:
            try:
                unresolved += int(status.split()[0])
            except (ValueError, IndexError):
                pass
    report.unresolved_count = unresolved

    def _parse_duration(x):
        try:
            parts = str(x).split(':')
            if len(parts) == 3:
                return int(float(parts[0])) * 3600 + int(float(parts[1])) * 60 + int(float(parts[2]))
        except (ValueError, AttributeError):
            pass
        return 0

    report.total_duration_sec = float(
        df_export['Duration'].astype(str).apply(_parse_duration).sum()
        if 'Duration' in df_export.columns and len(df_export) > 0 else 0
    )
    report.processing_time_sec = round(time.time() - start, 2)

    import json
    import numpy as np

    class _NpEncoder(json.JSONEncoder):
        def default(self, obj):
            if isinstance(obj, np.integer): return int(obj)
            if isinstance(obj, np.floating): return float(obj)
            return super().default(obj)

    report.synthesis_json = json.loads(
        json.dumps(df_synthese.to_dict('records'), cls=_NpEncoder)
    )

    site_col = next((c for c in ('Site Name', 'Site name', 'SITE NAME') if c in df_dedup.columns), None)
    if site_col and len(df_dedup) > 0:
        top_sites = (
            df_dedup[site_col].dropna().astype(str)
            .value_counts().head(10).reset_index()
        )
        top_sites.columns = ['name', 'count']
        report.top_sites_json = json.loads(json.dumps(top_sites.to_dict('records'), cls=_NpEncoder))
    else:
        report.top_sites_json = []

    report.detailed_file.name = os.path.relpath(output_path, settings.MEDIA_ROOT)
    if os.path.exists(synthesis_path):
        report.synthesis_file.name = os.path.relpath(synthesis_path, settings.MEDIA_ROOT)

    report.processed = True
    report.save()

    return JsonResponse({
        'done': True,
        'redirect': f'/results/{report.pk}/',
        'total_incidents': report.total_incidents,
        'processing_time': report.processing_time_sec,
    })


def process_status(request, pk):
    report = get_object_or_404(UploadedReport, pk=pk)
    if report.processed:
        return JsonResponse({'done': True, 'redirect': f'/results/{report.pk}/'})
    return JsonResponse({'done': False})


def delete_report(request, pk):
    report = get_object_or_404(UploadedReport, pk=pk)

    # Seul le propriétaire ou l'admin peut supprimer
    if report.user != request.user and not request.user.is_superuser:
        return JsonResponse({'error': 'Non autorisé'}, status=403)

    if request.method == 'POST':
        files_to_delete = []
        if report.file:
            files_to_delete.append(report.file.path)
        if report.detailed_file:
            try: files_to_delete.append(report.detailed_file.path)
            except Exception: pass
        if report.synthesis_file:
            try: files_to_delete.append(report.synthesis_file.path)
            except Exception: pass

        for filepath in files_to_delete:
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
            except Exception:
                pass

        report.delete()
        return JsonResponse({'success': True, 'redirect': '/history/'})

    return JsonResponse({'error': 'Méthode non autorisée'}, status=405)


def results(request, pk):
    report = get_object_or_404(UploadedReport, pk=pk, processed=True)
    # Vérifie que l'utilisateur est propriétaire ou admin
    if report.user != request.user and not request.user.is_superuser:
        return redirect('history')
    return render(request, 'reports/results.html', {'report': report, 'period_label': _period_label(report)})


def download_file(request, pk, file_type):
    report = get_object_or_404(UploadedReport, pk=pk, processed=True)
    if report.user != request.user and not request.user.is_superuser:
        raise Http404("Non autorisé")
    safe_label = _period_label(report).replace('/', '-').replace(' ', '_')
    if file_type == 'detailed' and report.detailed_file:
        file_field = report.detailed_file
        filename = f"Rapport_Detail_{safe_label}.xlsx"
    elif file_type == 'synthesis' and report.synthesis_file:
        file_field = report.synthesis_file
        filename = f"Rapport_Synthese_{safe_label}.xlsx"
    else:
        raise Http404("File not found")
    return FileResponse(open(file_field.path, 'rb'), as_attachment=True, filename=filename)


def history(request):
    from datetime import timedelta

    if request.user.is_superuser:
        all_reports = UploadedReport.objects.filter(processed=True).order_by('-date_rapport')
    else:
        all_reports = UploadedReport.objects.filter(processed=True, user=request.user).order_by('-date_rapport')

    period_filter = request.GET.get('period', 'all')

    if period_filter == 'day':
        filtered_reports = list(all_reports.filter(date_fin__isnull=True))
    elif period_filter == 'week':
        filtered_reports = [
            r for r in all_reports
            if r.date_fin and 5 <= (r.date_fin - r.date_rapport).days <= 10
        ]
    elif period_filter == 'month':
        filtered_reports = [
            r for r in all_reports
            if r.date_fin and 20 <= (r.date_fin - r.date_rapport).days <= 45
        ]
    elif period_filter == 'year':
        filtered_reports = [
            r for r in all_reports
            if r.date_fin and (r.date_fin - r.date_rapport).days > 45
        ]
    else:
        filtered_reports = list(all_reports)

    total_reports    = len(filtered_reports)
    total_incidents  = sum(r.total_incidents for r in filtered_reports)
    total_unresolved = sum(r.unresolved_count for r in filtered_reports if r.unresolved_count)
    total_resolved   = sum(1 for r in filtered_reports if r.unresolved_count == 0)

    return render(request, 'reports/history.html', {
        'reports':          filtered_reports,
        'period_filter':    period_filter,
        'total_reports':    total_reports,
        'total_incidents':  total_incidents,
        'total_unresolved': total_unresolved,
        'total_resolved':   total_resolved,
    })


# ── SVG Donut helper ────────────────────────────────────────────────────────
DONUT_COLORS = [
    '#003087', '#FFC72C', '#e53e3e', '#2196F3', '#FF9800',
    '#4CAF50', '#9C27B0', '#00BCD4', '#8BC34A', '#FF5722',
    '#607D8B', '#795548', '#009688', '#0047cc', '#F44336',
]

def _make_donut_svg(data, total_h):
    if not data:
        return ''
    total = sum(d['outage_h'] for d in data)
    if total == 0:
        return ''
    cx = cy = 100
    R  = 88
    Ri = 48
    paths = []
    angle = -math.pi / 2
    for i, d in enumerate(data):
        sweep = (d['outage_h'] / total) * 2 * math.pi
        ea    = angle + sweep
        large = 1 if sweep > math.pi else 0
        x1o = cx + R  * math.cos(angle); y1o = cy + R  * math.sin(angle)
        x2o = cx + R  * math.cos(ea);    y2o = cy + R  * math.sin(ea)
        x1i = cx + Ri * math.cos(ea);    y1i = cy + Ri * math.sin(ea)
        x2i = cx + Ri * math.cos(angle); y2i = cy + Ri * math.sin(angle)
        color = DONUT_COLORS[i % len(DONUT_COLORS)]
        mid   = angle + sweep / 2
        p  = f'<path d="M{x1o:.2f},{y1o:.2f} A{R},{R} 0 {large},1 {x2o:.2f},{y2o:.2f} '
        p += f'L{x1i:.2f},{y1i:.2f} A{Ri},{Ri} 0 {large},0 {x2i:.2f},{y2i:.2f} Z" '
        p += f'fill="{color}" stroke="#fff" stroke-width="2.5">'
        p += f'<title>{d["name"]}: {d["outage_h"]}h ({d["pct"]}%)</title></path>'
        if d['pct'] >= 5:
            lx = cx + (R + Ri) / 2 * math.cos(mid)
            ly = cy + (R + Ri) / 2 * math.sin(mid)
            p += (f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="middle" '
                  f'dominant-baseline="middle" fill="white" '
                  f'font-size="10" font-weight="bold" font-family="Arial,sans-serif">'
                  f'{d["pct"]}%</text>')
        paths.append(p)
        angle = ea

    center  = (f'<text x="{cx}" y="{cy-7}" text-anchor="middle" fill="#003087" '
               f'font-size="15" font-weight="bold" font-family="Arial,sans-serif">{total_h}h</text>')
    center += (f'<text x="{cx}" y="{cy+9}" text-anchor="middle" fill="#888" '
               f'font-size="9" font-family="Arial,sans-serif">TOTAL</text>')
    return f'<svg width="200" height="200" viewBox="0 0 200 200">{"".join(paths)}{center}</svg>'

def export_pdf(request, pk):
    import datetime
    from django.http import HttpResponse
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from io import BytesIO

    report = get_object_or_404(UploadedReport, pk=pk, processed=True)
    if report.user != request.user and not request.user.is_superuser:
        raise Http404("Non autorisé")

    YAS_BLUE   = colors.HexColor('#003087')
    YAS_YELLOW = colors.HexColor('#FFC72C')
    LIGHT_BLUE = colors.HexColor('#e8f0ff')
    LIGHT_GRAY = colors.HexColor('#f8faff')

    buffer = BytesIO()
    W = A4[0] - 40*mm  # largeur utile

    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm,
        topMargin=15*mm, bottomMargin=20*mm,
    )
    styles = getSampleStyleSheet()
    elements = []
    now = datetime.datetime.now().strftime('%d/%m/%Y à %H:%M')

    def p(text, size=9, color='#333333', bold=False, align='LEFT'):
        fn = 'Helvetica-Bold' if bold else 'Helvetica'
        return Paragraph(
            f'<font size="{size}" color="{color}">{text}</font>',
            ParagraphStyle('_', fontName=fn, alignment={'LEFT':0,'CENTER':1,'RIGHT':2}[align])
        )

   # ── HEADER ──────────────────────────────────────────────────
    ht = Table([[
        p('<b>YAS</b>', size=24, color='#003087'),
        Table([
            [p('● RAPPORT D\'INCIDENTS RÉSEAU', size=8, color='#FFC72C')],
            [p('<b>Rapport Automatique</b>', size=14, color='#003087')],
            [p(f'📅 {_period_label(report)}', size=9, color="#888888")],
        ], colWidths=[W*0.55]),
    ]], colWidths=[W*0.35, W*0.65])
    ht.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
        ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor('#e8edf5')),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING',    (0,0), (-1,-1), 8),
        ('LINEBELOW',  (0,0), (-1,-1), 3, YAS_YELLOW),
    ]))
    elements.append(ht)
    elements.append(Spacer(1, 5*mm))

    # ── META ────────────────────────────────────────────────────
    fname = report.original_filename
    if len(fname) > 40: fname = fname[:37] + '...'
    mt = Table([[
        Table([[p('FICHIER SOURCE', size=7, color='#aaaaaa')],[p(f'<b>{fname}</b>', size=9, color='#003087')]], colWidths=[W*0.33]),
        Table([[p('GÉNÉRÉ PAR', size=7, color='#aaaaaa')],[p(f'<b>{request.user.get_full_name() or request.user.username}</b>', size=9, color='#003087')]], colWidths=[W*0.25]),
        Table([[p('DATE', size=7, color='#aaaaaa')],[p(f'<b>{now}</b>', size=9, color='#003087')]], colWidths=[W*0.28]),
        Table([[p('DURÉE', size=7, color='#aaaaaa')],[p(f'<b>{report.processing_time_sec}s</b>', size=9, color='#003087')]], colWidths=[W*0.14]),
    ]], colWidths=[W*0.33, W*0.25, W*0.28, W*0.14])
    mt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), LIGHT_GRAY),
        ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor('#e8edf5')),
        ('INNERGRID',  (0,0), (-1,-1), 0.5, colors.HexColor('#e8edf5')),
        ('PADDING',    (0,0), (-1,-1), 6),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
    ]))
    elements.append(mt)
    elements.append(Spacer(1, 5*mm))

    # ── KPIs ────────────────────────────────────────────────────
    elements.append(p('RÉSUMÉ DU TRAITEMENT', size=9, color='#003087', bold=True))
    elements.append(HRFlowable(width='100%', thickness=2, color=YAS_YELLOW, spaceAfter=4))
    elements.append(Spacer(1, 2*mm))

    unresolved_color = '#e53e3e' if report.unresolved_count > 0 else '#22c55e'
    kpi_w = W / 4
    kt = Table([[
        Table([[p(str(report.total_rows), size=20, color='#003087', bold=True)],[p('LIGNES FICHIER', size=7, color='#888888')]]),
        Table([[p(str(report.filtered_rows), size=20, color='#003087', bold=True)],[p('ALARMES FILTRÉES', size=7, color='#888888')]]),
        Table([[p(str(report.total_incidents), size=20, color='#FFC72C', bold=True)],[p('INCIDENTS', size=7, color='#ffffff')]]),
        Table([[p(str(report.unresolved_count), size=20, color=unresolved_color, bold=True)],[p('NON RÉSOLUS', size=7, color='#888888')]]),
    ]], colWidths=[kpi_w]*4)
    kt.setStyle(TableStyle([
        ('BACKGROUND',  (0,0), (1,0), LIGHT_GRAY),
        ('BACKGROUND',  (2,0), (2,0), YAS_BLUE),
        ('BACKGROUND',  (3,0), (3,0), LIGHT_GRAY),
        ('BOX',         (0,0), (-1,-1), 0.5, colors.HexColor('#e8edf5')),
        ('INNERGRID',   (0,0), (-1,-1), 0.5, colors.HexColor('#e8edf5')),
        ('ALIGN',       (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING',     (0,0), (-1,-1), 10),
    ]))
    elements.append(kt)
    elements.append(Spacer(1, 5*mm))

    # ── SYNTHÈSE ─────────────────────────────────────────────────
    if report.synthesis_json:
        elements.append(p('SYNTHÈSE PAR ESCALADE', size=9, color='#003087', bold=True))
        elements.append(HRFlowable(width='100%', thickness=2, color=YAS_YELLOW, spaceAfter=4))
        elements.append(Spacer(1, 2*mm))

        # Largeurs colonnes optimisées
        cw = [W*0.30, W*0.12, W*0.16, W*0.16, W*0.16, W*0.10]

        def cell(txt, size=9, color='#333333', bold=False, align='CENTER'):
            fn = 'Helvetica-Bold' if bold else 'Helvetica'
            return Paragraph(
                f'<font size="{size}" color="{color}">{txt}</font>',
                ParagraphStyle('_', fontName=fn, alignment={'LEFT':0,'CENTER':1,'RIGHT':2}[align])
            )

        header_row = [
            cell('Escalade',  size=8, color='#ffffff', bold=True, align='LEFT'),
            cell('Incidents', size=8, color='#ffffff', bold=True),
            cell('Durée',     size=8, color='#ffffff', bold=True),
            cell('MTTR',      size=8, color='#ffffff', bold=True),
            cell('Outage',    size=8, color='#ffffff', bold=True),
            cell('Statut',    size=8, color='#ffffff', bold=True),
        ]

        data = [header_row]
        total_idx = None

        for i, row in enumerate(report.synthesis_json):
            esc = row.get('Escalade', '')
            inc = row.get('Inc count', 0)
            is_total = (esc == 'TOTAL')
            if is_total:
                total_idx = i + 1

            status = str(row.get('Status', ''))
            if 'Non resolu' in status:
                status_txt = f'⚠ {status}'
                s_color = '#7a5a00'
            elif status == 'Résolu':
                status_txt = '✓ Résolu'
                s_color = '#166534'
            elif status == 'N/A':
                status_txt = 'N/A'
                s_color = '#aaaaaa'
            else:
                status_txt = status
                s_color = '#333333'

            esc_color = '#003087' if is_total else '#111111'
            data.append([
                cell(esc,                    size=9, color=esc_color, bold=is_total, align='LEFT'),
                cell(str(inc),               size=9, color=esc_color, bold=is_total),
                cell(row.get('DUREE',''),    size=8, color='#555555'),
                cell(row.get('MTTR',''),     size=8, color='#555555'),
                cell(row.get('OUTAGE',''),   size=9, color='#003087', bold=is_total),
                cell(status_txt,             size=7, color=s_color),
            ])

        st = Table(data, colWidths=cw, repeatRows=1)
        ts = TableStyle([
            ('BACKGROUND',  (0,0), (-1,0),  YAS_BLUE),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
            ('GRID',        (0,0), (-1,-1), 0.3, colors.HexColor('#e8edf5')),
            ('ALIGN',       (1,0), (-1,-1), 'CENTER'),
            ('ALIGN',       (0,0), (0,-1),  'LEFT'),
            ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING',     (0,0), (-1,-1), 5),
            ('LEFTPADDING', (0,0), (0,-1),  8),
        ])
        if total_idx:
            ts.add('BACKGROUND', (0, total_idx), (-1, total_idx), LIGHT_BLUE)
            ts.add('LINEABOVE',  (0, total_idx), (-1, total_idx), 1.5, YAS_BLUE)
        st.setStyle(ts)
        elements.append(st)
        elements.append(Spacer(1, 6*mm))

    # ── FOOTER ──────────────────────────────────────────────────
    ft = Table([[
        p('YAS Togo — Rapport généré automatiquement par le système NOC', size=8, color='#ffffff'),
        p(f'ISOC • Confidentiel • {now}', size=8, color='#ffffff', align='RIGHT'),
    ]], colWidths=[W*0.60, W*0.40])
    ft.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), YAS_BLUE),
        ('PADDING',    (0,0), (-1,-1), 8),
    ]))
    elements.append(ft)

    doc.build(elements)
    buffer.seek(0)

    filename = f"Rapport_{report.original_filename.replace('.xlsx','')}.pdf"
    response = HttpResponse(buffer.read(), content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response

def notifications(request):
    """Retourne les incidents non résolus pour l'utilisateur."""
    if request.user.is_superuser:
        reports = UploadedReport.objects.filter(processed=True, unresolved_count__gt=0).order_by('-date_rapport')
    else:
        reports = UploadedReport.objects.filter(processed=True, user=request.user, unresolved_count__gt=0).order_by('-date_rapport')
    
    data = [{
        'pk':       str(r.pk),
        'filename': r.original_filename[:40],
        'date':     r.date_rapport.strftime('%d/%m/%Y'),
        'count':    r.unresolved_count,
    } for r in reports[:10]]
    
    return JsonResponse({
        'total': reports.count(),
        'items': data,
    })

def statistiques(request):
    from collections import defaultdict
    from datetime import timedelta

    if request.user.is_superuser:
        reports = UploadedReport.objects.filter(processed=True).order_by('-date_rapport')
    else:
        reports = UploadedReport.objects.filter(processed=True, user=request.user).order_by('-date_rapport')

    period_filter = request.GET.get('period', 'all')
    today = date.today()

    if period_filter == 'day':
        reports = reports.filter(date_rapport=today)
    elif period_filter == 'week':
        reports = reports.filter(date_rapport__gte=today - timedelta(days=7))
    elif period_filter == 'month':
        reports = reports.filter(date_rapport__gte=today - timedelta(days=30))
    elif period_filter == 'year':
        reports = reports.filter(date_rapport__year=today.year)

    def parse_hms(s):
        try:
            parts = str(s).split(':')
            if len(parts) == 3:
                return int(float(parts[0])) * 3600 + int(float(parts[1])) * 60 + int(float(parts[2]))
        except Exception:
            pass
        return 0

    escalade_data = defaultdict(lambda: {'count': 0, 'outage_sec': 0, 'duree_sec': 0})
    for r in reports:
        if not r.synthesis_json:
            continue
        for row in r.synthesis_json:
            esc = row.get('Escalade', '')
            if esc == 'TOTAL' or not esc:
                continue
            escalade_data[esc]['count']      += row.get('Inc count', 0)
            escalade_data[esc]['outage_sec'] += parse_hms(row.get('OUTAGE', '0:00:00'))
            escalade_data[esc]['duree_sec']  += parse_hms(row.get('DUREE',  '0:00:00'))

    escalades_sorted = sorted(escalade_data.items(), key=lambda x: x[1]['count'], reverse=True)
    max_esc = escalades_sorted[0][1]['count'] if escalades_sorted else 1
    escalades_chart = [
        {'name': k, 'count': v['count'], 'pct': round(v['count'] / max_esc * 100), 'outage': round(v['outage_sec'] / 3600, 1)}
        for k, v in escalades_sorted if v['count'] > 0
    ]

    site_data = defaultdict(int)
    for r in reports:
        if r.top_sites_json:
            for s in r.top_sites_json:
                site_data[s['name']] += s['count']
    sites_top10 = sorted(site_data.items(), key=lambda x: x[1], reverse=True)[:10]
    max_site = sites_top10[0][1] if sites_top10 else 1
    sites_chart = [{'name': k, 'count': v, 'pct': round(v / max_site * 100)} for k, v in sites_top10]

    total_outage_sec = sum(v['outage_sec'] for v in escalade_data.values())
    outage_chart = []
    for k, v in escalades_sorted:
        if v['outage_sec'] > 0:
            pct = round(v['outage_sec'] / total_outage_sec * 100) if total_outage_sec else 0
            outage_chart.append({'name': k, 'outage_h': round(v['outage_sec'] / 3600, 1), 'pct': pct})
    total_outage_h = round(total_outage_sec / 3600, 1)

    site_duration = defaultdict(float)
    for r in reports:
        if not r.detailed_file:
            continue
        file_name = r.detailed_file.name or ''
        if not (('results/' in file_name or 'results\\' in file_name) and file_name.endswith('_detailed.xlsx')):
            continue
        try:
            import pandas as pd
            df = pd.read_excel(r.detailed_file.path)
            site_col = next((c for c in df.columns if c.strip().lower() == 'site name'), None)
            if site_col and 'Duration' in df.columns:
                for _, row in df.iterrows():
                    site = str(row[site_col]).strip()
                    dur  = parse_hms(row['Duration'])
                    if site and site != 'nan':
                        site_duration[site] += dur
        except Exception:
            continue

    degraded_top10 = sorted(site_duration.items(), key=lambda x: x[1], reverse=True)[:10]
    max_deg = degraded_top10[0][1] if degraded_top10 else 1
    degraded_chart = [
        {'name': k, 'duration_h': round(v / 3600, 1), 'pct': round(v / max_deg * 100)}
        for k, v in degraded_top10
    ]

    donut_svg = _make_donut_svg(outage_chart, total_outage_h)
    outage_chart_colored = [
        {**d, 'color': DONUT_COLORS[i % len(DONUT_COLORS)]}
        for i, d in enumerate(outage_chart)
    ]

    return render(request, 'reports/statistiques.html', {
        'period_filter':   period_filter,
        'escalades_chart': escalades_chart,
        'sites_chart':     sites_chart,
        'outage_chart':    outage_chart_colored,
        'degraded_chart':  degraded_chart,
        'total_outage_h':  total_outage_h,
        'total_reports':   reports.count(),
        'donut_svg':       donut_svg,
    })