import json
import time
import os
import math
from datetime import date

from django.shortcuts import render, redirect, get_object_or_404
from django.utils.safestring import mark_safe
from django.http import FileResponse, Http404, JsonResponse
from django.conf import settings
from django.contrib import messages

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
        all_reports = UploadedReport.objects.filter(processed=True).order_by('-uploaded_at')
    else:
        all_reports = UploadedReport.objects.filter(processed=True, user=request.user).order_by('-uploaded_at')

    recent_reports = all_reports[:5]

    # KPIs globaux
    total_reports    = all_reports.count()
    total_incidents  = sum(r.total_incidents  for r in all_reports)
    total_unresolved = sum(r.unresolved_count for r in all_reports if r.unresolved_count)
    total_outage_h   = round(sum(r.total_duration_sec for r in all_reports) / 3600, 1)

    # Mini sparkline — incidents des 7 derniers rapports (du plus ancien au plus récent)
    spark_reports = list(reversed(list(all_reports[:7])))
    spark_labels  = [r.date_rapport.strftime('%d/%m') for r in spark_reports]
    spark_values  = [r.total_incidents for r in spark_reports]

    # Dernier rapport traité
    last_report = all_reports.first()

    return render(request, 'reports/home.html', {
        'recent_reports':   recent_reports,
        'total_reports':    total_reports,
        'total_incidents':  total_incidents,
        'total_unresolved': total_unresolved,
        'total_outage_h':   total_outage_h,
        'spark_labels':     spark_labels,
        'spark_values':     spark_values,
        'last_report':      last_report,
    })


def upload(request):
    if request.method == 'POST':
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            report = form.save(commit=False)
            report.original_filename = request.FILES['file'].name
            report.user = request.user
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


# ── Palette couleurs ──────────────────────────────────────────────────────────
DONUT_COLORS = [
    '#003087', '#e05a2b', '#FF9800', '#2196F3', '#FFC72C',
    '#4CAF50', '#9C27B0', '#00BCD4', '#8BC34A', '#FF5722',
    '#607D8B', '#795548', '#009688', '#0047cc', '#F44336',
]
DONUT_DARK = [
    '#001245', '#8b3318', '#b36a00', '#0d5ca8', '#a07800',
    '#2e7d32', '#6a1b9a', '#00838f', '#558b2f', '#bf360c',
    '#37474f', '#4e342e', '#00695c', '#002fa0', '#b71c1c',
]


# ── SVG Camembert 3D Premium ──────────────────────────────────────────────────
def _make_donut_svg(data, total_h):
    if not data:
        return ''
    total = sum(d['outage_h'] for d in data)
    if total == 0:
        return ''

    CX, CY  = 330, 160
    RX, RY  = 235, 92
    DEPTH   = 60
    W       = 720
    n       = len(data)
    COLS    = 3
    LEG_ROWS = math.ceil(n / COLS)
    H_BASE = CY + RY + DEPTH + 30 + LEG_ROWS * 52 + 20

    def pt(a, r=1.0):
        return (CX + r * RX * math.cos(a), CY + r * RY * math.sin(a))

    # ── Defs ────────────────────────────────────────────────────────────────
    defs = '<defs>'
    for i, d in enumerate(data):
        c  = DONUT_COLORS[i % len(DONUT_COLORS)]
        dk = DONUT_DARK[i  % len(DONUT_DARK)]
        defs += (
            f'<linearGradient id="pg{i}" x1="0" y1="0" x2="0.3" y2="1">'
            f'<stop offset="0%" stop-color="{c}"/>'
            f'<stop offset="100%" stop-color="{dk}"/>'
            f'</linearGradient>'
        )
    defs += (
        '<filter id="pshadow" x="-20%" y="-20%" width="140%" height="160%">'
        '<feDropShadow dx="0" dy="8" stdDeviation="10" flood-color="#00000035"/>'
        '</filter>'
        '</defs>'
    )

    # ── Tranches ────────────────────────────────────────────────────────────
    slices = []
    angle  = -math.pi / 2
    for i, d in enumerate(data):
        sweep = (d['outage_h'] / total) * 2 * math.pi
        slices.append({
            'a1':    angle,
            'a2':    angle + sweep,
            'mid':   angle + sweep / 2,
            'color': DONUT_COLORS[i % len(DONUT_COLORS)],
            'dark':  DONUT_DARK[i  % len(DONUT_DARK)],
            'grad':  f'url(#pg{i})',
            'pct':   d['pct'],
            'h':     d['outage_h'],
            'name':  d['name'],
        })
        angle += sweep

    # ── Parois latérales ────────────────────────────────────────────────────
    sides = ''
    for s in reversed(slices):
        a1, a2 = s['a1'], s['a2']
        vs = max(a1, 0)
        ve = min(a2, math.pi)
        if vs < ve:
            N = 32
            tp, bp = [], []
            for k in range(N + 1):
                a = vs + (ve - vs) * k / N
                x = CX + RX * math.cos(a)
                y = CY + RY * math.sin(a)
                tp.append(f'{x:.2f},{y:.2f}')
                bp.append(f'{x:.2f},{y + DEPTH:.2f}')
            bp.reverse()
            sides += (
                f'<path d="M{" L".join(tp + bp)}Z" '
                f'fill="{s["dark"]}" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>'
            )
        for a in [a1, a2]:
            if 0 <= a <= math.pi:
                ox, oy = pt(a)
                sides += (
                    f'<path d="M{CX:.2f},{CY:.2f} L{ox:.2f},{oy:.2f} '
                    f'L{ox:.2f},{oy+DEPTH:.2f} L{CX:.2f},{CY+DEPTH:.2f} Z" '
                    f'fill="{s["dark"]}" opacity="0.5" stroke="rgba(255,255,255,0.12)" stroke-width="0.8"/>'
                )

    bot_ellipse = (
        f'<ellipse cx="{CX}" cy="{CY+DEPTH}" rx="{RX}" ry="{RY}" '
        f'fill="none" stroke="rgba(0,0,0,0.08)" stroke-width="1.5"/>'
    )

    # ── Faces supérieures ───────────────────────────────────────────────────
    tops = ''
    for s in slices:
        a1, a2 = s['a1'], s['a2']
        x1, y1 = pt(a1)
        x2, y2 = pt(a2)
        large  = 1 if (a2 - a1) > math.pi else 0
        tops += (
            f'<path d="M{CX:.2f},{CY:.2f} L{x1:.2f},{y1:.2f} '
            f'A{RX},{RY} 0 {large},1 {x2:.2f},{y2:.2f} Z" '
            f'fill="{s["grad"]}" stroke="rgba(255,255,255,0.3)" stroke-width="1.5"/>'
        )

    # ── Reflets ─────────────────────────────────────────────────────────────
    highlights = ''
    for s in slices:
        hs = max(s['a1'], -math.pi / 2)
        he = min(s['a2'], -math.pi / 2 + 1.0)
        if hs < he:
            hx1, hy1 = pt(hs, 0.97)
            hx2, hy2 = pt(he, 0.97)
            highlights += (
                f'<path d="M{hx1:.2f},{hy1:.2f} A{RX*0.97:.1f},{RY*0.97:.1f} 0 0,1 {hx2:.2f},{hy2:.2f}" '
                f'fill="none" stroke="rgba(255,255,255,0.4)" stroke-width="2" stroke-linecap="round"/>'
            )

    # ── Labels avec anti-collision ───────────────────────────────────────────
    inner_threshold = 11  # pct >= 11 -> label interne

    BOX_H = 34
    BOX_PAD_TOP = 18   # rect top = ly - BOX_PAD_TOP
    # bas du rectangle (ly = ligne de base du texte du haut)
    RECT_BOTTOM_FROM_LY = -BOX_PAD_TOP + BOX_H  # ly + RECT_BOTTOM_FROM_LY = bas du rect
    VERT_GAP = 12      # espace vertical entre deux boîtes empilées
    BOX_W = 92

    ext_labels = []
    for s in slices:
        if s['pct'] < inner_threshold:
            lx_raw, ly_raw = pt(s['mid'], 1.45)
            ext_labels.append({
                's': s,
                'lx': lx_raw,
                'ly': ly_raw,
                'anchor': 'start' if math.cos(s['mid']) >= 0 else 'end',
            })

    def _pack_external_vertical(side_list):
        """Empile verticalement : chaque boîte sous la précédente (pas de chevauchement)."""
        side_list.sort(key=lambda e: e['ly'])
        prev_bottom = -1e9
        for e in side_list:
            top = e['ly'] - BOX_PAD_TOP
            min_top = prev_bottom + VERT_GAP
            if top < min_top:
                e['ly'] = min_top + BOX_PAD_TOP
            prev_bottom = max(prev_bottom, e['ly'] + RECT_BOTTOM_FROM_LY)

    left = [e for e in ext_labels if math.cos(e['s']['mid']) < 0]
    right = [e for e in ext_labels if math.cos(e['s']['mid']) >= 0]
    _pack_external_vertical(left)
    _pack_external_vertical(right)

    # Colonnes d'ancrage fixes : évite le chevauchement horizontal (stagger sur lx).
    # Toutes les étiquettes d'un même côté partagent le même lx ; seul ly varie.
    LEFT_COL_X, RIGHT_COL_X = 52, W - 52
    for e in left:
        e['lx'] = LEFT_COL_X
    for e in right:
        e['lx'] = RIGHT_COL_X

    ext_labels = left + right

    # Reserve de place sous les etiquettes externes pour ne pas recouvrir la legende
    base_leg_top = CY + RY + DEPTH + 32
    max_ext_bottom = max((e['ly'] + RECT_BOTTOM_FROM_LY for e in ext_labels), default=0)
    legend_push = max(0.0, max_ext_bottom - base_leg_top + 18)
    leg_top = base_leg_top + legend_push
    H = H_BASE + legend_push

    labels = ''

    # Labels internes — grandes tranches
    for s in slices:
        if s['pct'] >= inner_threshold:
            lx, ly = pt(s['mid'], 0.60)
            labels += (
                f'<rect x="{lx-34:.1f}" y="{ly-18:.1f}" width="68" height="37" rx="8" '
                f'fill="rgba(0,0,0,0.42)"/>'
                f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="middle" '
                f'font-family="Arial,sans-serif" font-size="17" font-weight="800" fill="white">'
                f'{s["pct"]}%</text>'
                f'<text x="{lx:.1f}" y="{ly+15:.1f}" text-anchor="middle" '
                f'font-family="Arial,sans-serif" font-size="12.5" font-weight="700" fill="rgba(255,255,255,0.97)">'
                f'{s["h"]}h</text>'
            )

    # Labels externes — petites tranches avec fond blanc
    for e in ext_labels:
        s      = e['s']
        lx     = e['lx']
        ly     = e['ly']
        anchor = e['anchor']
        ox, oy = pt(s['mid'], 1.04)
        mx, my = pt(s['mid'], 1.20)
        tx     = lx + (6 if anchor == 'start' else -6)
        box_x  = tx - BOX_W if anchor == 'end' else tx

        labels += (
            f'<rect x="{box_x:.1f}" y="{ly-BOX_PAD_TOP:.1f}" width="{BOX_W}" height="{BOX_H}" rx="7" '
            f'fill="rgba(255,255,255,0.97)" stroke="{s["color"]}" stroke-width="1.6"/>'
            f'<polyline points="{ox:.1f},{oy:.1f} {mx:.1f},{my:.1f} {lx:.1f},{ly:.1f}" '
            f'fill="none" stroke="{s["color"]}" stroke-width="1.9" opacity="0.95"/>'
            f'<circle cx="{ox:.1f}" cy="{oy:.1f}" r="3.3" fill="{s["color"]}"/>'
            f'<text x="{tx:.1f}" y="{ly-2:.1f}" text-anchor="{anchor}" '
            f'font-family="Arial,sans-serif" font-size="14" font-weight="900" fill="{s["color"]}">'
            f'{s["pct"]}%</text>'
            f'<text x="{tx:.1f}" y="{ly+11:.1f}" text-anchor="{anchor}" '
            f'font-family="Arial,sans-serif" font-size="12" font-weight="700" fill="#1f2937">'
            f'{s["h"]}h</text>'
        )

    # ── Badge total ─────────────────────────────────────────────────────────
    badge = (
        f'<rect x="{CX-44}" y="{CY-20}" width="88" height="38" rx="11" '
        f'fill="white" opacity="0.96" stroke="rgba(0,48,135,0.18)" stroke-width="1"/>'
        f'<text x="{CX}" y="{CY+1}" text-anchor="middle" '
        f'font-family="Arial,sans-serif" font-size="14" font-weight="900" fill="#003087">'
        f'{total_h}h</text>'
        f'<text x="{CX}" y="{CY+15}" text-anchor="middle" '
        f'font-family="Arial,sans-serif" font-size="8.5" font-weight="700" fill="#6b7280" letter-spacing="1.4">TOTAL</text>'
    )

    # ── Légende 3 colonnes ───────────────────────────────────────────────────
    COL_W   = int((W - 40) / COLS)
    legend  = ''
    for i, s in enumerate(slices):
        col = i % COLS
        row = i // COLS
        x   = 20 + col * COL_W
        y   = leg_top + row * 50
        legend += (
            f'<rect x="{x}" y="{y}" width="{COL_W-10}" height="42" rx="9" '
            f'fill="{s["color"]}" opacity="0.08"/>'
            f'<circle cx="{x+12}" cy="{y+12}" r="6" fill="{s["color"]}"/>'
        )
        name = s['name'][:16] + ('…' if len(s['name']) > 16 else '')
        legend += (
            f'<rect x="{x}" y="{y}" width="{COL_W-10}" height="44" rx="9" '
            f'fill="{s["color"]}" opacity="0.13" stroke="{s["color"]}" stroke-opacity="0.2"/>'
            f'<circle cx="{x+14}" cy="{y+13}" r="7.5" fill="{s["color"]}"/>'
            f'<text x="{x+26}" y="{y+17}" '
            f'font-family="Arial,sans-serif" font-size="13.5" font-weight="800" fill="{s["dark"]}">'
            f'{name}</text>'
            f'<text x="{x+14}" y="{y+34}" '
            f'font-family="Arial,sans-serif" font-size="12.5" font-weight="700" fill="#374151">'
            f'{s["h"]}h · {s["pct"]}%</text>'
        )

    return (
        f'<svg width="100%" viewBox="0 0 {W} {H}" '
        f'xmlns="http://www.w3.org/2000/svg" overflow="visible">'
        f'{defs}'
        f'<g filter="url(#pshadow)">{sides}{bot_ellipse}{tops}</g>'
        f'{highlights}{labels}{badge}{legend}'
        f'</svg>'
    )


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
    W = A4[0] - 40*mm

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

    if report.synthesis_json:
        elements.append(p('SYNTHÈSE PAR ESCALADE', size=9, color='#003087', bold=True))
        elements.append(HRFlowable(width='100%', thickness=2, color=YAS_YELLOW, spaceAfter=4))
        elements.append(Spacer(1, 2*mm))

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

        data_rows = [header_row]
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
            data_rows.append([
                cell(esc,                    size=9, color=esc_color, bold=is_total, align='LEFT'),
                cell(str(inc),               size=9, color=esc_color, bold=is_total),
                cell(row.get('DUREE',''),    size=8, color='#555555'),
                cell(row.get('MTTR',''),     size=8, color='#555555'),
                cell(row.get('OUTAGE',''),   size=9, color='#003087', bold=is_total),
                cell(status_txt,             size=7, color=s_color),
            ])

        st = Table(data_rows, colWidths=cw, repeatRows=1)
        ts = TableStyle([
            ('BACKGROUND',     (0,0), (-1,0),  YAS_BLUE),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
            ('GRID',           (0,0), (-1,-1), 0.3, colors.HexColor('#e8edf5')),
            ('ALIGN',          (1,0), (-1,-1), 'CENTER'),
            ('ALIGN',          (0,0), (0,-1),  'LEFT'),
            ('VALIGN',         (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING',        (0,0), (-1,-1), 5),
            ('LEFTPADDING',    (0,0), (0,-1),  8),
        ])
        if total_idx:
            ts.add('BACKGROUND', (0, total_idx), (-1, total_idx), LIGHT_BLUE)
            ts.add('LINEABOVE',  (0, total_idx), (-1, total_idx), 1.5, YAS_BLUE)
        st.setStyle(ts)
        elements.append(st)
        elements.append(Spacer(1, 6*mm))

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


def register_view(request):
    if request.user.is_authenticated:
        return redirect('/')

    if request.method == 'POST':
        from django.contrib.auth.models import User
        username   = request.POST.get('username', '').strip()
        first_name = request.POST.get('first_name', '').strip()
        last_name  = request.POST.get('last_name', '').strip()
        email      = request.POST.get('email', '').strip()
        password1  = request.POST.get('password1', '')
        password2  = request.POST.get('password2', '')

        if not username or not password1:
            messages.error(request, 'Identifiant et mot de passe obligatoires.')
        elif User.objects.filter(username=username).exists():
            messages.error(request, 'Cet identifiant est déjà utilisé.')
        elif password1 != password2:
            messages.error(request, 'Les mots de passe ne correspondent pas.')
        elif len(password1) < 8:
            messages.error(request, 'Le mot de passe doit contenir au moins 8 caractères.')
        else:
            User.objects.create_user(
                username=username, password=password1,
                first_name=first_name, last_name=last_name, email=email,
            )
            messages.success(request, 'Compte créé avec succès ! Connectez-vous.')
            return redirect('accounts:login')

    return render(request, 'accounts/register.html')

def comparer(request):
    """
    Comparaison de deux rapports côte à côte.
    GET  /comparer/              → page de sélection
    GET  /comparer/?r1=pk&r2=pk  → comparaison
    """
    if request.user.is_superuser:
        all_reports = UploadedReport.objects.filter(processed=True).order_by('-uploaded_at')
    else:
        all_reports = UploadedReport.objects.filter(processed=True, user=request.user).order_by('-uploaded_at')

    pk1 = request.GET.get('r1')
    pk2 = request.GET.get('r2')

    r1 = r2 = None
    diff = None

    if pk1 and pk2:
        r1 = get_object_or_404(all_reports, pk=pk1)
        r2 = get_object_or_404(all_reports, pk=pk2)

        def parse_hms(s):
            try:
                parts = str(s).split(':')
                if len(parts) == 3:
                    return int(float(parts[0])) * 3600 + int(float(parts[1])) * 60 + int(float(parts[2]))
            except Exception:
                pass
            return 0

        # ── Données KPI ──────────────────────────────────────────────────
        def kpi_diff(v1, v2):
            if v1 == 0 and v2 == 0:
                return 0, 'neutral'
            if v1 == 0:
                return '+∞', 'worse'
            delta = v2 - v1
            pct   = round(delta / v1 * 100, 1)
            trend = 'better' if delta < 0 else ('worse' if delta > 0 else 'neutral')
            return (f'+{pct}%' if pct > 0 else f'{pct}%'), trend

        inc_delta,  inc_trend  = kpi_diff(r1.total_incidents,  r2.total_incidents)
        unr_delta,  unr_trend  = kpi_diff(r1.unresolved_count, r2.unresolved_count)
        rows_delta, rows_trend = kpi_diff(r1.total_rows,       r2.total_rows)
        filt_delta, filt_trend = kpi_diff(r1.filtered_rows,    r2.filtered_rows)

        # ── Comparaison synthèse par escalade ────────────────────────────
        def synth_map(report):
            m = {}
            for row in (report.synthesis_json or []):
                esc = row.get('Escalade', '')
                if esc and esc != 'TOTAL':
                    m[esc] = {
                        'count':      row.get('Inc count', 0),
                        'outage_sec': parse_hms(row.get('OUTAGE', '0:00:00')),
                        'status':     row.get('Status', ''),
                    }
            return m

        s1, s2    = synth_map(r1), synth_map(r2)
        all_escs  = sorted(set(list(s1.keys()) + list(s2.keys())))

        esc_rows = []
        for esc in all_escs:
            v1 = s1.get(esc, {'count': 0, 'outage_sec': 0, 'status': '—'})
            v2 = s2.get(esc, {'count': 0, 'outage_sec': 0, 'status': '—'})
            delta = v2['count'] - v1['count']
            esc_rows.append({
                'name':       esc,
                'count1':     v1['count'],
                'count2':     v2['count'],
                'outage1_h':  round(v1['outage_sec'] / 3600, 1),
                'outage2_h':  round(v2['outage_sec'] / 3600, 1),
                'status1':    v1['status'],
                'status2':    v2['status'],
                'delta':      delta,
                'trend':      'better' if delta < 0 else ('worse' if delta > 0 else 'neutral'),
            })

        # ── Top sites communs ────────────────────────────────────────────
        def sites_map(report):
            return {s['name']: s['count'] for s in (report.top_sites_json or [])}

        ts1, ts2    = sites_map(r1), sites_map(r2)
        all_sites   = sorted(set(list(ts1.keys()) + list(ts2.keys())),
                             key=lambda s: max(ts1.get(s, 0), ts2.get(s, 0)), reverse=True)[:10]

        site_rows = [{
            'name':   s,
            'count1': ts1.get(s, 0),
            'count2': ts2.get(s, 0),
            'delta':  ts2.get(s, 0) - ts1.get(s, 0),
        } for s in all_sites]

        diff = {
            'inc_delta':  inc_delta,  'inc_trend':  inc_trend,
            'unr_delta':  unr_delta,  'unr_trend':  unr_trend,
            'rows_delta': rows_delta, 'rows_trend': rows_trend,
            'filt_delta': filt_delta, 'filt_trend': filt_trend,
            'esc_rows':   esc_rows,
            'site_rows':  site_rows,
        }

    return render(request, 'reports/comparer.html', {
        'all_reports': all_reports,
        'r1':   r1,
        'r2':   r2,
        'pk1':  pk1,
        'pk2':  pk2,
        'diff': diff,
    })

def export_statistiques(request):
    """
    Exporte les statistiques actuelles en Excel (4 onglets).
    Accepte les mêmes paramètres GET que statistiques() :
      ?period=latest|day|week|month|year|all
      ?report=<pk>
    """
    import openpyxl
    import openpyxl.chart.label
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, GradientFill
    )
    from openpyxl.utils import get_column_letter
    from django.http import HttpResponse
    from collections import defaultdict
    from datetime import timedelta
    import io

    # ── Récupère les rapports (même logique que statistiques()) ────────────
    if request.user.is_superuser:
        base_qs = UploadedReport.objects.filter(processed=True).order_by('-uploaded_at')
    else:
        base_qs = UploadedReport.objects.filter(processed=True, user=request.user).order_by('-uploaded_at')

    report_pk     = request.GET.get('report')
    period_filter = request.GET.get('period', 'latest')

    if report_pk:
        reports = base_qs.filter(pk=report_pk)
        period_label = 'Rapport sélectionné'
    elif period_filter == 'latest' or period_filter not in ('day', 'week', 'month', 'year', 'all'):
        first = base_qs.first()
        reports = base_qs.filter(pk=first.pk) if first else base_qs.none()
        period_label = 'Dernier rapport'
    else:
        today = date.today()
        labels_map = {'day': "Aujourd'hui", 'week': '7 jours', 'month': '30 jours', 'year': 'Année', 'all': 'Tout'}
        period_label = labels_map.get(period_filter, period_filter)
        if period_filter == 'day':
            reports = base_qs.filter(uploaded_at__date=today)
        elif period_filter == 'week':
            reports = base_qs.filter(uploaded_at__date__gte=today - timedelta(days=7))
        elif period_filter == 'month':
            reports = base_qs.filter(uploaded_at__date__gte=today - timedelta(days=30))
        elif period_filter == 'year':
            reports = base_qs.filter(uploaded_at__year=today.year)
        else:
            reports = base_qs

    def parse_hms(s):
        try:
            parts = str(s).split(':')
            if len(parts) == 3:
                return int(float(parts[0])) * 3600 + int(float(parts[1])) * 60 + int(float(parts[2]))
        except Exception:
            pass
        return 0

    # ── Calcule les données ─────────────────────────────────────────────────
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
            escalade_data[esc]['duree_sec']  += parse_hms(row.get('DUREE', '0:00:00'))

    escalades_sorted = sorted(escalade_data.items(), key=lambda x: x[1]['count'], reverse=True)
    total_outage_sec = sum(v['outage_sec'] for v in escalade_data.values())

    site_data = defaultdict(int)
    for r in reports:
        for s in (r.top_sites_json or []):
            site_data[s['name']] += s['count']
    sites_top10 = sorted(site_data.items(), key=lambda x: x[1], reverse=True)[:10]

    outage_data = [
        (k, round(v['outage_sec']/3600, 1),
         round(v['outage_sec']/total_outage_sec*100) if total_outage_sec else 0)
        for k, v in escalades_sorted if v['outage_sec'] > 0
    ]

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

    # ── Styles ──────────────────────────────────────────────────────────────
    YAS_BLUE   = '003087'
    YAS_YELLOW = 'FFC72C'
    LIGHT_BLUE = 'E8F0FF'
    LIGHT_GRAY = 'F8FAFF'
    WHITE      = 'FFFFFF'

    hdr_font  = Font(name='Calibri', bold=True, color=WHITE, size=11)
    hdr_fill  = PatternFill('solid', fgColor=YAS_BLUE)
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    sub_font  = Font(name='Calibri', bold=True, color=YAS_BLUE, size=10)
    sub_fill  = PatternFill('solid', fgColor=LIGHT_BLUE)

    cell_font  = Font(name='Calibri', size=10)
    cell_align = Alignment(horizontal='left', vertical='center')
    num_align  = Alignment(horizontal='center', vertical='center')

    alt_fill = PatternFill('solid', fgColor=LIGHT_GRAY)

    thin = Side(style='thin', color='E8EDF5')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header_row(ws, row_num, ncols):
        for col in range(1, ncols + 1):
            cell = ws.cell(row=row_num, column=col)
            cell.font   = hdr_font
            cell.fill   = hdr_fill
            cell.alignment = hdr_align
            cell.border = border

    def style_data_row(ws, row_num, ncols, alt=False):
        for col in range(1, ncols + 1):
            cell = ws.cell(row=row_num, column=col)
            cell.font   = Font(name='Calibri', size=10)
            cell.fill   = alt_fill if alt else PatternFill('solid', fgColor=WHITE)
            cell.alignment = num_align if col > 1 else cell_align
            cell.border = border

    def add_title_block(ws, title, period):
        ws.merge_cells('A1:F1')
        t = ws['A1']
        t.value     = f'YAS NOC — {title}'
        t.font      = Font(name='Calibri', bold=True, size=14, color=YAS_BLUE)
        t.alignment = Alignment(horizontal='left', vertical='center')
        t.fill      = PatternFill('solid', fgColor=LIGHT_BLUE)
        ws.row_dimensions[1].height = 28

        ws.merge_cells('A2:F2')
        p = ws['A2']
        p.value     = f'Période : {period}  |  Généré le {date.today().strftime("%d/%m/%Y")}'
        p.font      = Font(name='Calibri', size=10, color='888888')
        p.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[2].height = 18

        ws.row_dimensions[3].height = 6  # espace

    # ── Workbook ─────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()

    # ══ Onglet 1 : Classement Escalades ══════════════════════════════════════
    ws1 = wb.active
    ws1.title = '📊 Escalades'
    ws1.sheet_view.showGridLines = False
    ws1.column_dimensions['A'].width = 30
    ws1.column_dimensions['B'].width = 14
    ws1.column_dimensions['C'].width = 14
    ws1.column_dimensions['D'].width = 14

    add_title_block(ws1, 'Classement des Escalades', period_label)

    headers = ['Escalade', 'Incidents', 'Outage (h)', 'Durée (h)']
    for col, h in enumerate(headers, 1):
        ws1.cell(row=4, column=col).value = h
    style_header_row(ws1, 4, len(headers))
    ws1.row_dimensions[4].height = 22

    for i, (esc, v) in enumerate(escalades_sorted):
        row = 5 + i
        ws1.cell(row=row, column=1).value = esc
        ws1.cell(row=row, column=2).value = v['count']
        ws1.cell(row=row, column=3).value = round(v['outage_sec'] / 3600, 1)
        ws1.cell(row=row, column=4).value = round(v['duree_sec']  / 3600, 1)
        style_data_row(ws1, row, len(headers), alt=(i % 2 == 1))
        ws1.row_dimensions[row].height = 20

    # Ligne total
    if escalades_sorted:
        tr = 5 + len(escalades_sorted)
        ws1.cell(tr, 1).value = 'TOTAL'
        ws1.cell(tr, 2).value = sum(v['count'] for _, v in escalades_sorted)
        ws1.cell(tr, 3).value = round(total_outage_sec / 3600, 1)
        ws1.cell(tr, 4).value = round(sum(v['duree_sec'] for _, v in escalades_sorted) / 3600, 1)
        for col in range(1, 5):
            c = ws1.cell(tr, col)
            c.font   = Font(name='Calibri', bold=True, size=10, color=WHITE)
            c.fill   = PatternFill('solid', fgColor=YAS_BLUE)
            c.alignment = num_align if col > 1 else cell_align
            c.border = border
        ws1.row_dimensions[tr].height = 22

    # ══ Onglet 2 : Récurrence des Sites ══════════════════════════════════════
    ws2 = wb.create_sheet('📡 Sites')
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions['A'].width = 32
    ws2.column_dimensions['B'].width = 16

    add_title_block(ws2, 'Récurrence des Sites (Top 10)', period_label)

    for col, h in enumerate(['Site', 'Occurrences'], 1):
        ws2.cell(row=4, column=col).value = h
    style_header_row(ws2, 4, 2)
    ws2.row_dimensions[4].height = 22

    for i, (site, cnt) in enumerate(sites_top10):
        row = 5 + i
        ws2.cell(row, 1).value = site
        ws2.cell(row, 2).value = cnt
        style_data_row(ws2, row, 2, alt=(i % 2 == 1))
        ws2.row_dimensions[row].height = 20

    # ══ Onglet 3 : Poids & Outage / Métier ═══════════════════════════════════
    from openpyxl.chart import PieChart3D, Reference
    from openpyxl.chart.series import DataPoint

    ws3 = wb.create_sheet('🥧 Outage Métier')
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions['A'].width = 30
    ws3.column_dimensions['B'].width = 16
    ws3.column_dimensions['C'].width = 12

    add_title_block(ws3, 'Poids & Outage par Métier', period_label)

    for col, h in enumerate(['Métier / Escalade', 'Outage (h)', '% Total'], 1):
        ws3.cell(row=4, column=col).value = h
    style_header_row(ws3, 4, 3)
    ws3.row_dimensions[4].height = 22

    for i, (name, h, pct) in enumerate(outage_data):
        row = 5 + i
        ws3.cell(row, 1).value = name
        ws3.cell(row, 2).value = h
        ws3.cell(row, 3).value = f'{pct}%'
        style_data_row(ws3, row, 3, alt=(i % 2 == 1))
        ws3.row_dimensions[row].height = 20

    if outage_data:
        tr = 5 + len(outage_data)
        ws3.cell(tr, 1).value = 'TOTAL'
        ws3.cell(tr, 2).value = round(total_outage_sec / 3600, 1)
        ws3.cell(tr, 3).value = '100%'
        for col in range(1, 4):
            c = ws3.cell(tr, col)
            c.font      = Font(name='Calibri', bold=True, size=10, color=WHITE)
            c.fill      = PatternFill('solid', fgColor=YAS_BLUE)
            c.alignment = num_align if col > 1 else cell_align
            c.border    = border
        ws3.row_dimensions[tr].height = 22

    # ── Camembert Excel ──────────────────────────────────────────────────────
    if outage_data:
        pie = PieChart3D()
        pie.style = 26 
        pie.title    = 'Outage par Métier'
        pie.style    = 10
        pie.width    = 16
        pie.height   = 14

        # Données : colonne B (valeurs outage_h), lignes 5 à 5+n-1
        n_rows = len(outage_data)
        data_ref   = Reference(ws3, min_col=2, min_row=4, max_row=4 + n_rows)  # inclut header
        labels_ref = Reference(ws3, min_col=1, min_row=5, max_row=4 + n_rows)

        pie.add_data(data_ref, titles_from_data=True)
        pie.set_categories(labels_ref)

        # Couleurs des tranches — palette YAS
        SLICE_COLORS = [
            '003087', 'E05A2B', 'FF9800', '2196F3', 'FFC72C',
            '4CAF50', '9C27B0', '00BCD4', '8BC34A', 'FF5722',
        ]
        series = pie.series[0]
        for idx in range(n_rows):
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = SLICE_COLORS[idx % len(SLICE_COLORS)]
            series.dPt.append(pt)

        # Affiche les labels avec % et nom de catégorie
        series.dLbls = openpyxl.chart.label.DataLabelList()
        series.dLbls.showCatName = True
        series.dLbls.showPercent = True
        series.dLbls.showVal     = False
        series.dLbls.showLegendKey = False
        series.dLbls.separator = '\n'

        # Place le graphique à droite du tableau (colonne E)
        ws3.add_chart(pie, 'E4')

    # ══ Onglet 4 : Sites les Plus Dégradés ═══════════════════════════════════
    ws4 = wb.create_sheet('🔴 Sites Dégradés')
    ws4.sheet_view.showGridLines = False
    ws4.column_dimensions['A'].width = 32
    ws4.column_dimensions['B'].width = 16

    add_title_block(ws4, 'Sites les Plus Dégradés (Top 10)', period_label)

    for col, h in enumerate(['Site', 'Durée totale (h)'], 1):
        ws4.cell(row=4, column=col).value = h
    style_header_row(ws4, 4, 2)
    ws4.row_dimensions[4].height = 22

    for i, (site, sec) in enumerate(degraded_top10):
        row = 5 + i
        ws4.cell(row, 1).value = site
        ws4.cell(row, 2).value = round(sec / 3600, 1)
        style_data_row(ws4, row, 2, alt=(i % 2 == 1))
        ws4.row_dimensions[row].height = 20

    # ── Export ──────────────────────────────────────────────────────────────
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"Statistiques_YAS_{date.today().strftime('%Y%m%d')}_{period_label.replace(' ', '_')}.xlsx"
    response = HttpResponse(
        buffer.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response

def statistiques(request):
    """
    Modes :
      1. ?report=<pk>  -> stats du rapport specifique
      2. ?period=...   -> stats agregees de la periode
      3. (defaut)      -> stats du dernier rapport traite
    """
    from collections import defaultdict
    from datetime import timedelta

    if request.user.is_superuser:
        base_qs = UploadedReport.objects.filter(processed=True).order_by('-uploaded_at')
    else:
        base_qs = UploadedReport.objects.filter(processed=True, user=request.user).order_by('-uploaded_at')

    report_pk     = request.GET.get('report')
    single_report = None
    period_filter = request.GET.get('period', 'latest')

    if report_pk:
        single_report = get_object_or_404(base_qs, pk=report_pk)
        reports = base_qs.filter(pk=report_pk)
        period_filter = 'report'
    elif period_filter == 'latest' or period_filter not in ('day', 'week', 'month', 'year', 'all'):
        single_report = base_qs.first()
        if single_report:
            reports = base_qs.filter(pk=single_report.pk)
        else:
            reports = base_qs.none()
        period_filter = 'latest'
    else:
        today = date.today()
        if period_filter == 'day':
            reports = base_qs.filter(uploaded_at__date=today)
        elif period_filter == 'week':
            reports = base_qs.filter(uploaded_at__date__gte=today - timedelta(days=7))
        elif period_filter == 'month':
            reports = base_qs.filter(uploaded_at__date__gte=today - timedelta(days=30))
        elif period_filter == 'year':
            reports = base_qs.filter(uploaded_at__year=today.year)
        else:
            reports = base_qs

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
        {
            'name':   k,
            'count':  v['count'],
            'pct':    round(v['count'] / max_esc * 100),
            'outage': round(v['outage_sec'] / 3600, 1),
        }
        for k, v in escalades_sorted if v['count'] > 0
    ]

    site_data = defaultdict(int)
    for r in reports:
        if r.top_sites_json:
            for s in r.top_sites_json:
                site_data[s['name']] += s['count']
    sites_top10 = sorted(site_data.items(), key=lambda x: x[1], reverse=True)[:10]
    max_site = sites_top10[0][1] if sites_top10 else 1
    sites_chart = [
        {'name': k, 'count': v, 'pct': round(v / max_site * 100)}
        for k, v in sites_top10
    ]

    total_outage_sec = sum(v['outage_sec'] for v in escalade_data.values())
    outage_chart = []
    for k, v in escalades_sorted:
        if v['outage_sec'] > 0:
            pct = round(v['outage_sec'] / total_outage_sec * 100) if total_outage_sec else 0
            outage_chart.append({
                'name':     k,
                'outage_h': round(v['outage_sec'] / 3600, 1),
                'pct':      pct,
            })
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

    total_reports = reports.count()
    evolution_reports = list(reports.order_by('uploaded_at'))
    evolution_labels = [r.uploaded_at.strftime('%d/%m') for r in evolution_reports]
    evolution_incidents = [r.total_incidents for r in evolution_reports]
    evolution_outage = [round(r.total_duration_sec / 3600, 1) for r in evolution_reports]

    return render(request, 'reports/statistiques.html', {
        'period_filter':   period_filter,
        'single_report':   single_report,
        'escalades_chart': escalades_chart,
        'sites_chart':     sites_chart,
        'outage_chart':    outage_chart_colored,
        'degraded_chart':  degraded_chart,
        'total_outage_h':  total_outage_h,
        'total_reports':   total_reports,
        'donut_svg':       donut_svg,
        'show_evolution_chart': len(evolution_labels) > 1,
        'evolution_labels':    mark_safe(json.dumps(evolution_labels)),
        'evolution_incidents': mark_safe(json.dumps(evolution_incidents)),
        'evolution_outage':    mark_safe(json.dumps(evolution_outage)),
        'evolution_labels':    evolution_labels,
        'evolution_incidents': evolution_incidents,
        'evolution_outage':    evolution_outage,
    })