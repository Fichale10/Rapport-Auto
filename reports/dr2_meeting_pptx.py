"""
Génération PPTX native (python-pptx) des 2 supports de réunion DR2 :
  - `generate_gdi_daily`   : « presentation a automatiser GDI.pptx »   (point quotidien / mois en cours)
  - `generate_reunion_hebdo` : « PRESENTATION REUNION[1] vendredi NEW (9).pptx » (réunion hebdomadaire,
    1 diapositive « Détail DR2 » par jour de la période sélectionnée)

Données : `Dr2ViolationRecord` (calcul auto 2G/3G, voir `dr2_availability.py`) pour tout ce qui est DR2,
API ticketing live (`analytics.py`) pour le DR1 (indisponibilités >= 1h sur les 30 derniers jours).
Design : calqué sur les 2 fichiers de référence (logo YAS, bandeau titre bleu/rouge sur fond blanc).
"""
from io import BytesIO
from datetime import date, timedelta
from collections import Counter

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

from .pptx_report import (
    _blank, _rect, _txt, _table,
    SW, SH, MARGIN, CONTENT_TOP, CONTENT_H,
    C_BLUE, C_WHITE, C_DTEXT, C_LGRAY, C_BLUE3, C_YELL,
    C_RED_BG, C_RED_FG, C_YELL_BG, C_YELL_FG, C_GREEN_BG, C_GREEN_FG,
)
from .dr2_availability import DR2_REGION_TARGETS, DR2_ESCALADE_ORDER

C_RED_T = RGBColor(0xC0, 0x00, 0x00)

_JOURS_FR = ['LUNDI', 'MARDI', 'MERCREDI', 'JEUDI', 'VENDREDI', 'SAMEDI', 'DIMANCHE']
_MOIS_FR = ['', 'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
            'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE']


# ═══════════════════════════════════════════════════════════════════════════
# BAS NIVEAU — bandeau titre / logo / diapos fixes
# ═══════════════════════════════════════════════════════════════════════════

def _logo(slide, l=None, t=Inches(0.16), h=Inches(0.78)):
    from .gdi_core import yas_logo_bytes
    data = yas_logo_bytes()
    if not data:
        return
    if l is None:
        l = SW - Inches(1.55)
    try:
        slide.shapes.add_picture(BytesIO(data), l, t, height=h)
    except Exception:
        pass


def _header(slide, tag, subtitle='', extra_right=''):
    _rect(slide, 0, 0, SW, SH, C_WHITE)
    _txt(slide, tag, MARGIN, Inches(0.16), Inches(9.5), Inches(0.45),
         size=20, bold=True, color=C_BLUE)
    if subtitle:
        _txt(slide, subtitle, MARGIN, Inches(0.64), Inches(9.5), Inches(0.4),
             size=14, bold=True, color=C_RED_T)
    if extra_right:
        _txt(slide, extra_right, Inches(8.6), Inches(0.66), Inches(3.0), Inches(0.35),
             size=12, bold=True, color=C_RED_T, align=PP_ALIGN.RIGHT)
    _logo(slide)


def _blob(slide, points_frac, color):
    """Forme organique (freeform) approximant le blob bleu des diapos de garde
    /fin du modèle de référence — points en fractions (x, y) de la diapositive."""
    pts = [(int(round(x * SW)), int(round(y * SH))) for x, y in points_frac]
    fb = slide.shapes.build_freeform(start_x=pts[0][0], start_y=pts[0][1], scale=1)
    fb.add_line_segments(pts[1:], close=True)
    shape = fb.convert_to_shape()
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


# Points du blob (haut-gauche, bord droit/bas ondulé) — calqués sur la diapo
# de fin du modèle de référence (fichier « presentation a automatiser GDI.pptx »).
_MERCI_BLOB_PTS = [
    (0.065, 0.00), (0.38, 0.00), (0.58, 0.015), (0.74, 0.05), (0.84, 0.11),
    (0.87, 0.17), (0.80, 0.23), (0.66, 0.29), (0.52, 0.35), (0.40, 0.41),
    (0.28, 0.47), (0.19, 0.51), (0.11, 0.545), (0.065, 0.555),
]


def _cover(prs):
    """Diapositive de garde — identique à la diapo 1 (layout « Pause ») des 2
    fichiers de référence : uniquement le logo YAS centré sur fond jaune."""
    sl = _blank(prs)
    _rect(sl, 0, 0, SW, SH, C_YELL)
    _logo(sl, l=Inches(4.9), t=Inches(2.35), h=Inches(2.8))
    return sl


def _closing(prs):
    """Diapositive de fin — reproduit la diapo « MERCI » (layout « End1 ») :
    blob bleu en haut à gauche, texte MERCI, mention YAS Togo, logo en bas à droite."""
    sl = _blank(prs)
    _rect(sl, 0, 0, SW, SH, C_YELL)
    _blob(sl, _MERCI_BLOB_PTS, C_BLUE)
    _txt(sl, 'MERCI', Inches(1.0), Inches(1.95), Inches(4.6), Inches(1.1),
         size=54, bold=True, italic=True, color=RGBColor(0x7E, 0xA6, 0xD9))
    _txt(sl, 'YAS Togo / DT / DOC / iSOC / GDI', Inches(0.65), Inches(6.25),
         Inches(5.8), Inches(0.4), size=13, italic=True, color=C_BLUE)
    _logo(sl, l=Inches(10.5), t=Inches(4.85), h=Inches(1.2))
    return sl


def _kpi_text(slide, text, l, t, w=Inches(3), h=Inches(0.5), size=16, color=C_RED_T, align=PP_ALIGN.LEFT):
    _txt(slide, text, l, t, w, h, size=size, bold=True, color=color, align=align)


# ═══════════════════════════════════════════════════════════════════════════
# DONNÉES
# ═══════════════════════════════════════════════════════════════════════════

def _dr2_qs(debut, fin):
    from .models import Dr2ViolationRecord
    return Dr2ViolationRecord.objects.filter(date__gte=debut, date__lte=fin).order_by('date', 'site_name')


def _dr2_dataset(debut, fin):
    """Réutilise `_build_dr2_from_rows`/`_dr2_records_to_rows` (views.py) — même
    logique que la page /reporting/dr2-daily/ (import différé pour éviter le
    cycle d'import views.py <-> ce module)."""
    from .views import _build_dr2_from_rows, _dr2_records_to_rows
    rows = _dr2_records_to_rows(_dr2_qs(debut, fin))
    return _build_dr2_from_rows(rows, debut, fin)


def _duree_str(rec):
    if not rec.alarm_time:
        return ''
    if not rec.cancel_time:
        return 'EN COURS'
    sec = int((rec.cancel_time - rec.alarm_time).total_seconds())
    if sec < 0:
        return 'EN COURS'
    return f'{sec // 3600}:{(sec % 3600) // 60:02d}:{sec % 60:02d}'


_DETAIL_HEADERS = ['N°', 'Ticket', 'Site parent', 'Site Name', 'Site ID', 'Alarm Time',
                    'Durée', 'Catégorie', 'Cause', 'Pt bloquant', 'Cancel Time', 'DR2']
_DETAIL_COL_W = [0.5, 1.6, 1.4, 1.7, 1.0, 1.5, 1.0, 1.4, 2.2, 1.6, 1.5, 0.6]


def _detail_rows(qs):
    rows = []
    for i, rec in enumerate(qs, 1):
        rows.append([
            i,
            rec.numero_ticket or '',
            rec.site_parent or '\xa0',
            rec.site_name,
            rec.site_id or '',
            rec.alarm_time.strftime('%d-%m-%Y %H:%M') if rec.alarm_time else '',
            _duree_str(rec),
            rec.categorie or '—',
            (rec.cause or '')[:38],
            rec.point_bloquant or 'N/A',
            rec.cancel_time.strftime('%d-%m-%Y %H:%M') if rec.cancel_time else '\xa0',
            'OUI',
        ])
    return rows


def _escalade_breakdown(qs, total=None):
    cnt = Counter(rec.categorie or '—' for rec in qs)
    total = total or sum(cnt.values()) or 1
    ordered = [e for e in DR2_ESCALADE_ORDER if e in cnt] + [e for e in cnt if e not in DR2_ESCALADE_ORDER]
    return [(e, cnt[e], round(cnt[e] / total * 100)) for e in ordered]


def _causes_breakdown(qs, n=10):
    cnt = Counter((rec.cause or '—').strip() for rec in qs if (rec.cause or '').strip())
    total = sum(cnt.values()) or 1
    return [(c, k, round(k / total * 100)) for c, k in cnt.most_common(n)]


def _points_bloquants(qs):
    cnt = Counter((rec.point_bloquant or '').strip() for rec in qs
                  if (rec.point_bloquant or '').strip() and (rec.point_bloquant or '').strip().upper() != 'N/A')
    return cnt.most_common(15)


def _top_sites(qs, n=15):
    cnt = Counter(rec.site_name for rec in qs)
    return cnt.most_common(n)


def _base_breakdown(qs):
    from .models import Site
    names = {rec.site_name for rec in qs}
    base_map = dict(Site.objects.filter(site_name__in=names).values_list('site_name', 'base'))
    cnt = Counter(base_map.get(rec.site_name) or 'INCONNU' for rec in qs)
    return cnt.most_common(20)


def _daily_counts(qs, debut, fin):
    cnt = Counter(rec.date for rec in qs)
    days = []
    d = debut
    while d <= fin:
        days.append((d, cnt.get(d, 0)))
        d += timedelta(days=1)
    return days


def _monthly_averages():
    """Moyenne DR2/jour par mois traité (dénominateur = nb de jours réellement
    traités ce mois-là via `Dr2ProcessedDate`, pas le nb de jours calendaires)."""
    from .models import Dr2ViolationRecord, Dr2ProcessedDate
    months = {}
    for rec in Dr2ViolationRecord.objects.all():
        key = (rec.date.year, rec.date.month)
        months.setdefault(key, [0, 0])
        months[key][0] += 1
    for pd_ in Dr2ProcessedDate.objects.all():
        key = (pd_.date.year, pd_.date.month)
        if key in months:
            months[key][1] += 1
    out = []
    for (y, m), (total, ndays) in sorted(months.items()):
        avg = round(total / ndays, 2) if ndays else 0
        out.append({'label': f'{_MOIS_FR[m]} {y}', 'total': total, 'moyenne': avg})
    return out


def _dr1_violations(fin):
    """Sites en violation DR1 (indisponible >= 1h, plus de 2 fois sur les 30
    derniers jours) — via l'API ticketing live (dégrade proprement si injoignable)."""
    from .analytics import fetch_api_dataframe, normalize_dataframe
    debut30 = fin - timedelta(days=29)
    try:
        raw = fetch_api_dataframe(debut30.isoformat(), fin.isoformat(), network='mobile')
        df = normalize_dataframe(raw)
    except Exception:
        return None  # API indisponible
    df = df[df['duration_sec'].fillna(0) >= 3600]
    if df.empty:
        return []
    grp = df.groupby(['site', 'region']).agg(
        cnt=('site', 'size'),
        cause=('cause', lambda s: Counter(s).most_common(1)[0][0] if len(s) else ''),
    ).reset_index()
    grp = grp[grp['cnt'] > 2].sort_values('cnt', ascending=False)
    return grp.to_dict('records')


# ═══════════════════════════════════════════════════════════════════════════
# GRAPHIQUES NATIFS
# ═══════════════════════════════════════════════════════════════════════════

def _line_chart(slide, categories, values, l, t, w, h, title=''):
    data = CategoryChartData()
    data.categories = [c.strftime('%d/%m') if hasattr(c, 'strftime') else str(c) for c in categories]
    data.add_series('DR2', values)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, l, t, w, h, data)
    chart = gframe.chart
    chart.has_legend = False
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
    plot = chart.plots[0]
    plot.series[0].format.line.color.rgb = C_BLUE
    plot.series[0].format.line.width = Pt(2.25)
    return chart


def _bar_chart(slide, categories, values, l, t, w, h, horizontal=True, color=C_BLUE):
    data = CategoryChartData()
    data.categories = [str(c) for c in categories]
    data.add_series('Nb', values)
    ctype = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gframe = slide.shapes.add_chart(ctype, l, t, w, h, data)
    chart = gframe.chart
    chart.has_legend = False
    plot = chart.plots[0]
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = color
    return chart


def _pie_chart(slide, categories, values, l, t, w, h):
    data = CategoryChartData()
    data.categories = [str(c) for c in categories]
    data.add_series('DR2', values)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.PIE, l, t, w, h, data)
    chart = gframe.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    return chart


# ═══════════════════════════════════════════════════════════════════════════
# DIAPOSITIVES PARTAGÉES (utilisées par les 2 supports)
# ═══════════════════════════════════════════════════════════════════════════

def _slide_dr2_trend(prs, d, kpi_label, period_label):
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'DR2 TREND')
    days = _daily_counts(_dr2_qs(d['debut'], d['fin']), d['debut'], d['fin'])
    _line_chart(sl, [x[0] for x in days], [x[1] for x in days],
                MARGIN, CONTENT_TOP, SW - Inches(3.2), CONTENT_H - Inches(0.1),
                title=period_label)
    _kpi_text(sl, f"{kpi_label} = {d['total_dr2']}", Inches(10.3), CONTENT_TOP + Inches(0.3), Inches(2.7))
    _kpi_text(sl, f"Moyenne = {d['moyenne']}".replace('.', ','),
              Inches(10.3), CONTENT_TOP + Inches(0.9), Inches(2.7), size=13)
    return sl


def _slide_apercu_global(prs, d):
    sl = _blank(prs)
    _header(sl, 'RAPPORT GESTION DES INCIDENTS', 'Aperçu global et comparatif des tendances DR2')
    color_map = {'red': (C_RED_BG, C_RED_FG), 'yellow': (C_YELL_BG, C_YELL_FG), 'green': (C_GREEN_BG, C_GREEN_FG)}
    rows = [[r['region'], r['tget'], r['dr2'], f"{r['pct_reg']}%", f"{r['pct_tget']}%"]
            for r in d['region_rows']]
    fmts = {}
    for i, r in enumerate(d['region_rows']):
        fmts[(i, 4)] = color_map[r['color']]
    rows.append(['TOTAL', d['total_tget'], d['total_dr2'], '100%',
                  f"{round(d['total_dr2'] / d['total_tget'] * 100) if d['total_tget'] else 0}%"])
    fmts[(len(rows) - 1, 0)] = (C_BLUE, C_WHITE)
    _table(sl, ['RÉGION', 'CIBLE', 'DR2', '% GLOBAL', '% CIBLE'], rows,
           col_widths=[3, 2, 2, 2, 2], cell_fmts=fmts, font_size=11)
    return sl


def _slide_merci(prs):
    return _closing(prs)


# ═══════════════════════════════════════════════════════════════════════════
# SUPPORT 1 — GDI (point quotidien / mois en cours) — 10 diapositives
# ═══════════════════════════════════════════════════════════════════════════

def generate_gdi_daily(debut, fin, generated_on):
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    d = _dr2_dataset(debut, fin)
    qs = _dr2_qs(debut, fin)
    period_label = f"{debut.strftime('%d/%m/%Y')} au {fin.strftime('%d/%m/%Y')}"

    # 1. Cover
    _cover(prs)

    # 2. DR2 Trend
    _slide_dr2_trend(prs, d, 'TDR2', period_label)

    # 3. Cas de violation DR2 (période) + répartition escalade + top causes
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS',
            f'CAS DE VIOLATION DR2 {period_label}  —  {d["total_dr2"]} DR2')
    rows = _detail_rows(qs)[:12]
    _table(sl, _DETAIL_HEADERS, rows, col_widths=_DETAIL_COL_W,
           top=CONTENT_TOP, height=Inches(2.9), font_size=7, hdr_size=7)
    esc = _escalade_breakdown(qs, d['total_dr2'])[:6]
    _table(sl, [f'TOTAL DR2 = {d["total_dr2"]}', 'NB', '%'],
           [[e, k, f'{p}%'] for e, k, p in esc],
           left=MARGIN, top=CONTENT_TOP + Inches(3.1), width=Inches(4.2), height=Inches(2.2),
           col_widths=[3, 1, 1], font_size=9)
    causes = _causes_breakdown(qs, 6)
    _table(sl, ['TOP CAUSES', 'NB', '%'],
           [[c[:35], k, f'{p}%'] for c, k, p in causes],
           left=Inches(4.6), top=CONTENT_TOP + Inches(3.1), width=Inches(7.9), height=Inches(2.2),
           col_widths=[5, 1, 1], font_size=9)

    # 4. Top site occurrence DR2
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'TOP SITE OCCURRENCE DR2')
    top_sites = _top_sites(qs, 15)
    if top_sites:
        _bar_chart(sl, [s for s, _ in reversed(top_sites)], [c for _, c in reversed(top_sites)],
                   MARGIN, CONTENT_TOP, SW - 2 * MARGIN, CONTENT_H - Inches(0.1))

    # 5. Répartition DR2 Région / Bases
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'Répartition DR2 Région / Bases')
    regs = [(r['region'], r['dr2']) for r in d['region_rows'] if r['dr2']]
    if regs:
        _pie_chart(sl, [r for r, _ in regs], [c for _, c in regs],
                   MARGIN, CONTENT_TOP, Inches(4.6), CONTENT_H - Inches(0.1))
    bases = _base_breakdown(qs)[:15]
    if bases:
        _bar_chart(sl, [b for b, _ in reversed(bases)], [c for _, c in reversed(bases)],
                   Inches(5.4), CONTENT_TOP, SW - Inches(5.4) - MARGIN, CONTENT_H - Inches(0.1))

    # 6. Efficacité DR2 (par région, vs cible)
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'EFFICACITE DR2')
    rows_eff = [[r['region'], r['dr2'], r['tget'], f"{r['pct_tget']}%"] for r in d['region_rows']]
    fmts = {}
    for i, r in enumerate(d['region_rows']):
        fmts[(i, 3)] = {'red': (C_RED_BG, C_RED_FG), 'yellow': (C_YELL_BG, C_YELL_FG),
                         'green': (C_GREEN_BG, C_GREEN_FG)}[r['color']]
    _table(sl, ['RÉGION', 'DR2', 'CIBLE', '% CIBLE ATTEINTE'], rows_eff,
           col_widths=[3, 2, 2, 2], cell_fmts=fmts, font_size=11,
           top=CONTENT_TOP, height=Inches(3.2))
    labels = [r['region'] for r in d['region_rows']]
    vals = [r['pct_tget'] for r in d['region_rows']]
    if labels:
        _bar_chart(sl, labels, vals, MARGIN, CONTENT_TOP + Inches(3.4), SW - 2 * MARGIN, Inches(1.9), horizontal=False)

    # 7. Points bloquants
    sl = _blank(prs)
    _header(sl, 'RAPPORT GESTION DES INCIDENTS', 'Points Bloquants', extra_right=f'DR2- PB = {sum(k for _, k in _points_bloquants(qs))}')
    pb = _points_bloquants(qs)
    rows_pb = [[p[:55], k] for p, k in pb]
    if rows_pb:
        _table(sl, ['POINT BLOQUANT', 'NB'], rows_pb, col_widths=[8, 2],
               top=CONTENT_TOP, height=Inches(3.4), font_size=10)
    if pb:
        _bar_chart(sl, [p for p, _ in reversed(pb[:8])], [k for _, k in reversed(pb[:8])],
                   MARGIN, CONTENT_TOP + Inches(3.6), SW - 2 * MARGIN, Inches(1.7))

    # 8. Aperçu global et comparatif des tendances DR2
    _slide_apercu_global(prs, d)

    # 9. DR2 comparative mensuel
    sl = _blank(prs)
    _header(sl, 'RAPPORT GESTION DES INCIDENTS', 'DR2 COMPARATIVE MENSUEL')
    months = _monthly_averages()
    if months:
        _bar_chart(sl, [m['label'] for m in months], [m['total'] for m in months],
                   MARGIN, CONTENT_TOP, Inches(7.2), CONTENT_H - Inches(0.1), horizontal=False)
        rows_m = [[m['label'], str(m['moyenne']).replace('.', ',')] for m in months]
        _table(sl, ['MOYENNE-J', ''], rows_m, left=Inches(7.8), top=CONTENT_TOP,
               width=Inches(4.8), height=Inches(2.2), col_widths=[2, 1], font_size=11)

    # 10. Merci
    _slide_merci(prs)

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════
# SUPPORT 2 — RÉUNION hebdomadaire — 1 diapo « Détail DR2 » par jour de la période
# ═══════════════════════════════════════════════════════════════════════════

_DEF_ROWS = [
    ['DR1', "Nombre d'indisponibilité d'une station de base",
     "Nombre de fois qu'une même station de base est restée indisponible pour une durée d'au moins "
     "une heure pendant les 30 derniers jours.", '≤ 2'],
    ['DR2', "Délai d'indisponibilité d'une station de base",
     "Délai d'indisponibilité par jour d'une même station de base quel que soit le lieu de son "
     "implantation sur le territoire national.", '≤ 3H'],
]


def generate_reunion_hebdo(debut, fin, generated_on):
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    d = _dr2_dataset(debut, fin)
    qs_period = _dr2_qs(debut, fin)
    period_label = f"{debut.strftime('%d/%m/%Y')} au {fin.strftime('%d/%m/%Y')}"

    # 1. Cover
    _cover(prs)

    # 2. Définition DR1/DR2 (contenu fixe)
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'Définition DR1/DR2')
    _table(sl, ['Code', 'Indicateur', 'Définition', 'Seuil (2G,3G,4G)'], _DEF_ROWS,
           col_widths=[1, 3, 5, 1.5], top=CONTENT_TOP, height=Inches(3.2), font_size=10)

    # 3. Cas de violation DR1
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'Cas de violation DR1')
    dr1 = _dr1_violations(fin)
    if dr1 is None:
        _txt(sl, "Données ticketing indisponibles pour le calcul DR1 (API injoignable).",
             MARGIN, CONTENT_TOP, SW - 2 * MARGIN, Inches(0.5), size=13, color=C_RED_FG)
    elif not dr1:
        _txt(sl, "Aucun site en violation DR1 sur les 30 derniers jours.",
             MARGIN, CONTENT_TOP, SW - 2 * MARGIN, Inches(0.5), size=13, color=C_RED_FG)
        _kpi_text(sl, 'TDR1 = 0', MARGIN, CONTENT_TOP + Inches(0.6))
    else:
        rows = [[r['site'], r['cnt'], r['region'], str(r['cause'])[:45]] for r in dr1[:12]]
        _table(sl, ['SITE NAME', 'COUNT', 'RÉGION', 'CAUSES'], rows,
               col_widths=[3, 1, 2, 6], top=CONTENT_TOP + Inches(0.6), height=Inches(3.2), font_size=10)
        _kpi_text(sl, f'TDR1 = {len(dr1)} DR1', MARGIN, CONTENT_TOP)

    # 4. DR2 Trend (semaine)
    _slide_dr2_trend(prs, d, 'Nbr DR2', period_label)

    # 5..N. Détail DR2 par jour (1 diapo par jour AVEC violations)
    jours_avec_dr2 = sorted({rec.date for rec in qs_period})
    for idx, day in enumerate(jours_avec_dr2):
        qs_day = qs_period.filter(date=day)
        nb = qs_day.count()
        is_last = (idx == len(jours_avec_dr2) - 1)
        title = 'DETAIL DR2' + (f' {_JOURS_FR[day.weekday()]}' if is_last else '')
        sl = _blank(prs)
        _header(sl, 'REUNION GESTION DES INCIDENTS', title)
        rows = _detail_rows(qs_day)
        tbl_h = Inches(3.1) if is_last else Inches(5.3)
        _table(sl, _DETAIL_HEADERS, rows, col_widths=_DETAIL_COL_W,
               top=CONTENT_TOP, height=tbl_h, font_size=8, hdr_size=8)
        _txt(sl, f'CAS DE VIOLATION DR2 {day.strftime("%d-%m-%Y")}  —  {nb} DR2',
             MARGIN, CONTENT_TOP - Inches(0.32), Inches(9), Inches(0.3), size=11, bold=True, color=C_BLUE3)
        if is_last:
            esc = _escalade_breakdown(qs_period, d['total_dr2'])[:6]
            _table(sl, [f'S = {d["total_dr2"]} DR2', 'DR2', '%'],
                   [[e, k, f'{p}%'] for e, k, p in esc],
                   left=MARGIN, top=CONTENT_TOP + Inches(3.3), width=Inches(4.2), height=Inches(2.0),
                   col_widths=[3, 1, 1], font_size=9)
            causes = _causes_breakdown(qs_period, 6)
            _table(sl, [f'TOP CAUSES RECURRENTES ({d["total_dr2"]})', '', ''],
                   [[c[:35], k, f'{p}%'] for c, k, p in causes],
                   left=Inches(4.6), top=CONTENT_TOP + Inches(3.3), width=Inches(7.9), height=Inches(2.0),
                   col_widths=[5, 1, 1], font_size=9)

    # N+1. Aperçu global et comparatif des tendances DR2
    _slide_apercu_global(prs, d)

    # N+2. Merci
    _slide_merci(prs)

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
