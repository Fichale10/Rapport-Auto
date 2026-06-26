"""
Génération automatique du rapport PowerPoint — Comité Gestion des Incidents
"""

from io import BytesIO
from datetime import date

from pptx import Presentation

_MOIS_FR = [
    '', 'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
    'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE',
]

def _mois_label_fr(d):
    if not d:
        return ''
    return f'{_MOIS_FR[d.month]} {d.year}'
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Palette de couleurs ───────────────────────────────────────────────────────
C_BLUE   = RGBColor(0x00, 0x30, 0x87)
C_BLUE2  = RGBColor(0x00, 0x47, 0xCC)
C_BLUE3  = RGBColor(0x1E, 0x3A, 0x6E)
C_YELL   = RGBColor(0xFF, 0xC7, 0x2C)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_LGRAY  = RGBColor(0xF8, 0xFA, 0xFF)
C_MGRAY  = RGBColor(0xE8, 0xED, 0xF5)
C_DTEXT  = RGBColor(0x33, 0x33, 0x44)
C_BLUE_T = RGBColor(0x2A, 0x4A, 0x80)

C_GREEN_BG  = RGBColor(0xC6, 0xEF, 0xCE)
C_GREEN_FG  = RGBColor(0x27, 0x62, 0x21)
C_YELL_BG   = RGBColor(0xFF, 0xEB, 0x9C)
C_YELL_FG   = RGBColor(0x9C, 0x65, 0x00)
C_RED_BG    = RGBColor(0xFF, 0xC7, 0xCE)
C_RED_FG    = RGBColor(0x9C, 0x00, 0x06)

# ── Dimensions ────────────────────────────────────────────────────────────────
SW = Inches(13.33)
SH = Inches(7.5)
HDR_H = Inches(1.08)
MARGIN = Inches(0.35)
CONTENT_TOP = Inches(1.18)
CONTENT_H   = SH - CONTENT_TOP - Inches(0.25)


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS BAS NIVEAU
# ═══════════════════════════════════════════════════════════════════════════════

def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _rect(slide, l, t, w, h, rgb=None):
    shp = slide.shapes.add_shape(1, l, t, w, h)
    if rgb:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb
    else:
        shp.fill.background()
    shp.line.fill.background()
    return shp


def _txt(slide, text, l, t, w, h, size=11, bold=False, color=C_DTEXT,
         align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = str(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb


def _fmt(seconds):
    if not seconds:
        return '0:00:00'
    s = int(seconds)
    return f'{s//3600}:{(s%3600)//60:02d}:{s%60:02d}'


# ═══════════════════════════════════════════════════════════════════════════════
# COMPOSANTS SLIDES
# ═══════════════════════════════════════════════════════════════════════════════

def _header(slide, title, subtitle='', mois_label=''):
    _rect(slide, 0, 0, SW, HDR_H, C_BLUE)
    _rect(slide, 0, 0, Inches(0.07), HDR_H, C_YELL)    # bande jaune gauche
    tag = 'COMITE GESTION DES INCIDENTS'
    _txt(slide, tag, MARGIN, Inches(0.06), Inches(10), Inches(0.28),
         size=8, bold=True, color=C_YELL)
    _txt(slide, title, MARGIN, Inches(0.30), Inches(10), Inches(0.50),
         size=17, bold=True, color=C_WHITE)
    if subtitle:
        _txt(slide, subtitle, MARGIN, Inches(0.78), Inches(9), Inches(0.25),
             size=9, color=RGBColor(0xB0, 0xC4, 0xE8))
    if mois_label:
        _txt(slide, mois_label, Inches(10.5), Inches(0.32), Inches(2.5), Inches(0.4),
             size=13, bold=True, color=C_YELL, align=PP_ALIGN.RIGHT)


def _table(slide, headers, rows,
           left=MARGIN, top=CONTENT_TOP, width=None, height=None,
           col_widths=None, hdr_bg=C_BLUE3, alt=True, cell_fmts=None,
           font_size=9, hdr_size=9):
    if width  is None: width  = SW - 2 * MARGIN
    if height is None: height = CONTENT_H - Inches(0.1)

    nc = len(headers)
    nr = len(rows) + 1
    tbl = slide.shapes.add_table(nr, nc, left, top, width, height).table

    # Largeurs colonnes
    if col_widths:
        tot = sum(col_widths)
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = int(width * cw / tot)
    else:
        cw = width // nc
        for i in range(nc):
            tbl.columns[i].width = cw

    # Hauteurs lignes
    rh = height // nr
    for i in range(nr):
        tbl.rows[i].height = rh

    def _cell(cell, val, bg=None, fg=C_DTEXT, bold=False,
              align=PP_ALIGN.CENTER, fs=9):
        # Clear existing text
        cell.text = ''
        if bg:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
        else:
            try:
                cell.fill.background()
            except Exception:
                pass
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = str(val) if val is not None else ''
        r.font.size = Pt(fs)
        r.font.bold = bold
        r.font.color.rgb = fg

    # Entête
    for j, h in enumerate(headers):
        _cell(tbl.cell(0, j), h, bg=hdr_bg, fg=C_WHITE, bold=True, fs=hdr_size)

    # Données
    for i, row in enumerate(rows):
        bg_row = C_LGRAY if (alt and i % 2 == 1) else None
        for j, val in enumerate(row):
            fmt = cell_fmts.get((i, j)) if cell_fmts else None
            if fmt:
                cb, cf = fmt
                _cell(tbl.cell(i+1, j), val, bg=cb, fg=cf, bold=True, fs=font_size)
            else:
                is_lbl = (j == 0)
                _cell(tbl.cell(i+1, j), val, bg=bg_row,
                      fg=C_BLUE3 if is_lbl else C_DTEXT,
                      bold=is_lbl,
                      align=PP_ALIGN.LEFT if is_lbl else PP_ALIGN.CENTER,
                      fs=font_size)

    return tbl


def _kpi_bar(slide, kpis):
    """kpis = list of (label, value, unit, color)"""
    n = len(kpis)
    w = (SW - 2 * MARGIN) // n
    for i, (lbl, val, unit, color) in enumerate(kpis):
        x = MARGIN + i * w
        y = CONTENT_TOP
        _rect(slide, x + Inches(0.06), y, w - Inches(0.12), Inches(1.6), color)
        _txt(slide, str(val), x + Inches(0.06), y + Inches(0.25), w - Inches(0.12), Inches(0.8),
             size=36, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        _txt(slide, lbl, x + Inches(0.06), y + Inches(1.1), w - Inches(0.12), Inches(0.35),
             size=10, bold=True, color=RGBColor(0xE8, 0xF0, 0xFF), align=PP_ALIGN.CENTER)
        if unit:
            _txt(slide, unit, x + Inches(0.06), y + Inches(0.08), w - Inches(0.12), Inches(0.25),
                 size=8, color=RGBColor(0xC0, 0xD4, 0xF0), align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES FIXES
# ═══════════════════════════════════════════════════════════════════════════════

def _cover(prs, mois_label, generated_on):
    sl = _blank(prs)
    _rect(sl, 0, 0, SW, SH, C_BLUE)
    _rect(sl, 0, 0, SW, Inches(0.12), C_YELL)
    _rect(sl, 0, SH - Inches(0.12), SW, Inches(0.12), C_YELL)
    # Bandes décoratives
    _rect(sl, SW - Inches(0.12), 0, Inches(0.12), SH, C_YELL)
    _rect(sl, 0, 0, Inches(0.12), SH, C_YELL)

    _txt(sl, 'Yas Togo / DT / DOC / iSOC', MARGIN, Inches(0.5), SW - 2*MARGIN, Inches(0.4),
         size=11, color=RGBColor(0x80, 0xA0, 0xD0), align=PP_ALIGN.CENTER)
    _txt(sl, 'COMITÉ', MARGIN, Inches(1.5), SW - 2*MARGIN, Inches(1.1),
         size=60, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    _txt(sl, 'GESTION DES INCIDENTS', MARGIN, Inches(2.5), SW - 2*MARGIN, Inches(0.9),
         size=40, bold=True, color=C_YELL, align=PP_ALIGN.CENTER)
    _txt(sl, mois_label.upper(), MARGIN, Inches(3.7), SW - 2*MARGIN, Inches(0.8),
         size=32, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    _txt(sl, f'Généré le {generated_on}', MARGIN, SH - Inches(0.8),
         SW - 2*MARGIN, Inches(0.35),
         size=10, color=RGBColor(0x80, 0xA0, 0xD0), align=PP_ALIGN.CENTER)


def _section(prs, num, title, icon=''):
    sl = _blank(prs)
    _rect(sl, 0, 0, SW, SH, C_BLUE)
    _rect(sl, 0, 0, Inches(0.12), SH, C_YELL)
    _rect(sl, SW - Inches(0.12), 0, Inches(0.12), SH, C_YELL)
    _txt(sl, str(num).zfill(2), MARGIN + Inches(0.5), Inches(1.5), Inches(3), Inches(3.5),
         size=140, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.LEFT)
    _txt(sl, icon, Inches(5), Inches(2.0), Inches(2), Inches(2),
         size=72, color=C_WHITE, align=PP_ALIGN.CENTER)
    _txt(sl, title, MARGIN, Inches(5.5), SW - 2*MARGIN, Inches(1.2),
         size=36, bold=True, color=C_YELL, align=PP_ALIGN.CENTER)


def _closing(prs):
    sl = _blank(prs)
    _rect(sl, 0, 0, SW, SH, C_BLUE)
    _rect(sl, 0, 0, SW, Inches(0.12), C_YELL)
    _rect(sl, 0, SH - Inches(0.12), SW, Inches(0.12), C_YELL)
    _txt(sl, 'MERCI', MARGIN, Inches(2.5), SW - 2*MARGIN, Inches(2),
         size=72, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    _txt(sl, 'Yas Togo / DT / DOC / iSOC / GDI',
         MARGIN, SH - Inches(0.8), SW - 2*MARGIN, Inches(0.35),
         size=10, color=RGBColor(0x80, 0xA0, 0xD0), align=PP_ALIGN.CENTER)


def _def_slide(prs, mois_label):
    sl = _blank(prs)
    _header(sl, 'Définition DR1 / DR2', mois_label=mois_label)
    headers = ['Code', 'Indicateur', 'Définition', 'Seuil (2G/3G/4G)']
    rows = [
        ['DR1', "Nombre d'indisponibilités d'une station",
         "Nombre de fois qu'une même station est indisponible sur le mois",
         '≤ 2'],
        ['DR2', "Délai d'indisponibilité d'une station",
         "Délai d'indisponibilité par jour d'une station > 3h constitue un DR2",
         '≤ 3H'],
    ]
    _table(sl, headers, rows,
           col_widths=[1, 3, 6, 2],
           top=CONTENT_TOP + Inches(0.8),
           height=Inches(2.2),
           font_size=11, hdr_size=11)
    _txt(sl, '• Un DR1 = site indisponible plus de 2 fois dans le mois',
         MARGIN, CONTENT_TOP + Inches(3.5), SW - 2*MARGIN, Inches(0.4),
         size=12, color=C_BLUE3)
    _txt(sl, '• Un DR2 = une indisponibilité dont la durée dépasse 3 heures',
         MARGIN, CONTENT_TOP + Inches(4.0), SW - 2*MARGIN, Inches(0.4),
         size=12, color=C_BLUE3)


# ═══════════════════════════════════════════════════════════════════════════════
# FETCH DATA
# ═══════════════════════════════════════════════════════════════════════════════

def _fetch_mobile(mois):
    from .models import Incident
    from django.db.models import Count, Sum

    qs = Incident.objects.filter(domain='mobile')
    if mois:
        qs = qs.filter(mois_rapport=mois)

    total = qs.count()

    # DR1 violations: sites avec > 2 incidents
    dr1_rows = list(
        qs.exclude(site_name='')
        .values('site_name', 'region', 'cause')
        .annotate(cnt=Count('id'))
        .filter(cnt__gt=2)
        .order_by('-cnt')[:15]
    )

    # DR2 data
    dr2_qs = Incident.objects.filter(domain='dr2')
    dr2_mois = dr2_qs.order_by('-mois_rapport').values_list('mois_rapport', flat=True).first()
    if dr2_mois:
        dr2_qs = dr2_qs.filter(mois_rapport=dr2_mois)
    total_dr2 = dr2_qs.count()

    dr2_by_region = {r['region'].upper(): r['cnt'] for r in
                     dr2_qs.exclude(region='').values('region').annotate(cnt=Count('id'))}
    dr2_by_escalade = {e['escalade']: e['cnt'] for e in
                       dr2_qs.exclude(escalade='').values('escalade').annotate(cnt=Count('id'))}

    # Stats par région
    region_stats = []
    for r in qs.exclude(region='').values('region').annotate(
        nb=Count('id'), dur=Sum('duration_sec')
    ).order_by('-nb'):
        reg = r['region'].upper()
        nb = r['nb']
        dur = r['dur'] or 0
        mttr_sec = dur / nb if nb else 0
        dr2_r = dr2_by_region.get(reg, 0)
        eff = round((nb - dr2_r) / nb * 100) if nb else 0
        region_stats.append({
            'region': r['region'], 'nb': nb,
            'mttr': _fmt(mttr_sec), 'dr2': dr2_r, 'eff': eff,
        })

    # Top sites (DR1)
    top_sites = list(
        qs.exclude(site_name='').values('site_name', 'region')
        .annotate(cnt=Count('id')).order_by('-cnt')[:15]
    )

    # Points bloquants
    pb_rows = list(
        qs.exclude(point_bloquant='').values('point_bloquant')
        .annotate(cnt=Count('id')).order_by('-cnt')[:10]
    )

    # Top causes
    cause_rows = list(
        qs.exclude(cause='').values('cause')
        .annotate(cnt=Count('id')).order_by('-cnt')[:10]
    )

    # DR2 par escalade
    esc_rows = list(
        dr2_qs.exclude(escalade='').values('escalade')
        .annotate(nb=Count('id'), dur=Sum('duration_sec')).order_by('-nb')
    )
    for e in esc_rows:
        nb = e['nb']
        dur = e['dur'] or 0
        e['mttr'] = _fmt(dur / nb if nb else 0)
        e['outage'] = _fmt(dur)
        e['pct'] = f"{round(nb / total_dr2 * 100)}%" if total_dr2 else '0%'

    return {
        'total': total, 'total_dr2': total_dr2,
        'tdr1': len(dr1_rows), 'dr1_rows': dr1_rows,
        'region_stats': region_stats, 'top_sites': top_sites,
        'pb_rows': pb_rows, 'cause_rows': cause_rows,
        'esc_rows': esc_rows,
    }


def _fetch_domain(domain, mois):
    from .models import Incident
    from django.db.models import Count, Sum

    qs = Incident.objects.filter(domain=domain)
    if mois:
        qs = qs.filter(mois_rapport=mois)
    if not qs.exists():
        qs = Incident.objects.filter(domain=domain)

    total = qs.count()
    total_dur = qs.aggregate(d=Sum('duration_sec'))['d'] or 0

    by_region = list(
        qs.exclude(region='').values('region')
        .annotate(nb=Count('id'), dur=Sum('duration_sec')).order_by('-nb')
    )
    for r in by_region:
        nb = r['nb']
        r['mttr'] = _fmt((r['dur'] or 0) / nb if nb else 0)
        r['outage'] = _fmt(r['dur'] or 0)

    by_escalade = list(
        qs.exclude(escalade='').values('escalade')
        .annotate(nb=Count('id'), dur=Sum('duration_sec')).order_by('-nb')
    )
    for e in by_escalade:
        nb = e['nb']
        e['mttr']   = _fmt((e['dur'] or 0) / nb if nb else 0)
        e['outage'] = _fmt(e['dur'] or 0)

    by_nature = list(
        qs.exclude(nature='').values('nature')
        .annotate(nb=Count('id')).order_by('-nb')[:10]
    )

    incidents = list(qs.exclude(nature='').values(
        'nature', 'region', 'escalade', 'duration_sec', 'site_name', 'cause', 'status'
    ).order_by('-duration_sec')[:20])
    for inc in incidents:
        inc['dur_fmt'] = _fmt(inc['duration_sec'])

    return {
        'total': total, 'total_dur': _fmt(total_dur),
        'by_region': by_region, 'by_escalade': by_escalade,
        'by_nature': by_nature, 'incidents': incidents,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES RÉSEAU MOBILE
# ═══════════════════════════════════════════════════════════════════════════════

def _slide_dr1(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, f'Cas de violation DR1   —   TDR1 = {d["tdr1"]} sites', mois_label=mois_label)
    rows = d['dr1_rows'][:15]
    tbl_rows = [[r['site_name'], r['cnt'], r['region'],
                 str(r.get('cause', '') or '')[:45]] for r in rows]
    fmts = {}
    for i, r in enumerate(rows):
        if r['cnt'] >= 10:
            fmts[(i, 1)] = (C_RED_BG, C_RED_FG)
        elif r['cnt'] >= 5:
            fmts[(i, 1)] = (C_YELL_BG, C_YELL_FG)
        else:
            fmts[(i, 1)] = (C_GREEN_BG, C_GREEN_FG)
    _table(sl, ['SITE', 'COUNT', 'RÉGION', 'CAUSE'],
           tbl_rows, col_widths=[3, 1, 2, 6],
           cell_fmts=fmts, font_size=9)


def _slide_dr2_overview(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, f'DR2 Trend   —   TDR2 = {d["total_dr2"]} DR2', mois_label=mois_label)
    kpis = [
        ('Total Incidents', d['total'], 'RÉSEAU MOBILE', C_BLUE),
        ('Total DR2', d['total_dr2'], 'VIOLATIONS > 3H', C_RED_FG),
        ('Violations DR1', d['tdr1'], 'SITES > 2 FOIS', RGBColor(0xD9, 0x77, 0x06)),
        ('Efficacité glob.', f"{round((d['total']-d['total_dr2'])/d['total']*100) if d['total'] else 0}%",
         'RÉSOLUTION < 3H', RGBColor(0x05, 0x96, 0x69)),
    ]
    _kpi_bar(sl, kpis)

    # Table par région
    _txt(sl, 'Efficacité DR2 par Région', MARGIN,
         CONTENT_TOP + Inches(1.8), SW - 2*MARGIN, Inches(0.35),
         size=12, bold=True, color=C_BLUE3)
    rows = [[r['region'], r['nb'], r['mttr'], r['dr2'],
             f"{r['eff']}%"] for r in d['region_stats']]
    fmts = {}
    for i, r in enumerate(d['region_stats']):
        eff = r['eff']
        if eff >= 90:
            fmts[(i, 4)] = (C_GREEN_BG, C_GREEN_FG)
        elif eff >= 80:
            fmts[(i, 4)] = (C_YELL_BG, C_YELL_FG)
        else:
            fmts[(i, 4)] = (C_RED_BG, C_RED_FG)
    _table(sl, ['RÉGION', 'Nb INCIDENTS', 'MTTR', 'Nb DR2', '% EFFICACITÉ'],
           rows, col_widths=[3, 2, 2, 2, 2],
           top=CONTENT_TOP + Inches(2.2),
           height=Inches(2.8),
           cell_fmts=fmts, font_size=10)


def _slide_dr2_metier(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, 'Violation DR2 / Métiers', mois_label=mois_label)
    rows = [[e['escalade'], e['nb'], e['mttr'], e['outage'], e['pct']]
            for e in d['esc_rows']]
    _table(sl, ['ESCALADE / MÉTIER', 'Nb DR2', 'MTTR', 'OUTAGE TOTAL', '% du Total'],
           rows, col_widths=[4, 2, 2, 2, 2], font_size=10)


def _slide_top_sites(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, 'Top Sites Récurrents', mois_label=mois_label)
    rows = [[r['site_name'], r['region'], r['cnt']] for r in d['top_sites'][:15]]
    fmts = {}
    for i, r in enumerate(d['top_sites'][:15]):
        if r['cnt'] >= 15:
            fmts[(i, 2)] = (C_RED_BG, C_RED_FG)
        elif r['cnt'] >= 8:
            fmts[(i, 2)] = (C_YELL_BG, C_YELL_FG)
    _table(sl, ['SITE', 'RÉGION', 'NB INCIDENTS'],
           rows, col_widths=[5, 3, 2], cell_fmts=fmts, font_size=10)


def _slide_causes(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, 'Top Causes des Incidents', mois_label=mois_label)
    total = d['total'] or 1
    rows = [[r['cause'][:60], r['cnt'],
             f"{round(r['cnt']/total*100)}%"] for r in d['cause_rows']]
    _table(sl, ['CAUSE', 'NB INCIDENTS', '% du Total'],
           rows, col_widths=[7, 2, 2], font_size=10)


def _slide_points_bloquants(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, 'Points Bloquants', mois_label=mois_label)
    total_pb = sum(r['cnt'] for r in d['pb_rows']) or 1
    rows = [[r['point_bloquant'][:55], r['cnt'],
             f"{round(r['cnt']/total_pb*100)}%"] for r in d['pb_rows']]
    # Ligne total
    rows.append(['TOTAL', sum(r['cnt'] for r in d['pb_rows']), '100%'])
    fmts = {(len(rows)-1, 0): (C_BLUE, C_WHITE),
            (len(rows)-1, 1): (C_BLUE, C_WHITE),
            (len(rows)-1, 2): (C_BLUE, C_WHITE)}
    _table(sl, ['CAUSE', 'COUNT', '% du Total'],
           rows, col_widths=[7, 2, 2], cell_fmts=fmts, font_size=10)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES RÉSEAU FIXE
# ═══════════════════════════════════════════════════════════════════════════════

def _slide_fixe(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, f'Réseau Fixe — {d["total"]} incidents', mois_label=mois_label)
    kpis = [
        ('Total Incidents', d['total'], 'RÉSEAU FIXE', C_BLUE2),
        ('Outage Total', d['total_dur'], '', RGBColor(0xD9, 0x77, 0x06)),
    ]
    _kpi_bar(sl, kpis)

    # Types par nature (résumé)
    if d['by_nature']:
        _txt(sl, 'Types d\'incidents', MARGIN,
             CONTENT_TOP + Inches(1.8), SW / 2 - MARGIN - Inches(0.1), Inches(0.35),
             size=11, bold=True, color=C_BLUE3)
        nat_rows = [[r['nature'][:40], r['nb']] for r in d['by_nature'][:8]]
        _table(sl, ['NATURE', 'Nb'],
               nat_rows, col_widths=[7, 1],
               left=MARGIN, top=CONTENT_TOP + Inches(2.2),
               width=SW // 2 - MARGIN - Inches(0.1),
               height=Inches(3.0), font_size=9)

    # Par région
    if d['by_region']:
        _txt(sl, 'Par Région', SW // 2 + Inches(0.1),
             CONTENT_TOP + Inches(1.8), SW // 2 - MARGIN - Inches(0.1), Inches(0.35),
             size=11, bold=True, color=C_BLUE3)
        reg_rows = [[r['region'], r['nb'], r['mttr']] for r in d['by_region']]
        _table(sl, ['RÉGION', 'Nb', 'MTTR'],
               reg_rows, col_widths=[3, 1, 2],
               left=SW // 2 + Inches(0.1),
               top=CONTENT_TOP + Inches(2.2),
               width=SW // 2 - MARGIN - Inches(0.1),
               height=Inches(3.0), font_size=9)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES TRANSPORT
# ═══════════════════════════════════════════════════════════════════════════════

def _slide_transport(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, f'Transport — {d["total"]} incidents', mois_label=mois_label)
    kpis = [
        ('Total Incidents', d['total'], 'TRANSPORT', C_BLUE2),
        ('Outage Total', d['total_dur'], '', RGBColor(0xD9, 0x77, 0x06)),
    ]
    _kpi_bar(sl, kpis)

    # Par escalade
    if d['by_escalade']:
        _txt(sl, 'Par Métier', MARGIN,
             CONTENT_TOP + Inches(1.8), SW // 2 - MARGIN - Inches(0.1), Inches(0.35),
             size=11, bold=True, color=C_BLUE3)
        esc_rows = [[e['escalade'][:30], e['nb'], e['mttr'], e['outage']]
                    for e in d['by_escalade']]
        _table(sl, ['MÉTIER', 'Nb', 'MTTR', 'OUTAGE'],
               esc_rows, col_widths=[4, 1, 2, 2],
               left=MARGIN, top=CONTENT_TOP + Inches(2.2),
               width=SW // 2 - MARGIN - Inches(0.1),
               height=Inches(3.2), font_size=9)

    # Par région
    if d['by_region']:
        _txt(sl, 'Par Région', SW // 2 + Inches(0.1),
             CONTENT_TOP + Inches(1.8), SW // 2 - MARGIN - Inches(0.1), Inches(0.35),
             size=11, bold=True, color=C_BLUE3)
        reg_rows = [[r['region'], r['nb'], r['mttr'], r['outage']]
                    for r in d['by_region']]
        _table(sl, ['RÉGION', 'Nb', 'MTTR', 'OUTAGE'],
               reg_rows, col_widths=[3, 1, 2, 2],
               left=SW // 2 + Inches(0.1),
               top=CONTENT_TOP + Inches(2.2),
               width=SW // 2 - MARGIN - Inches(0.1),
               height=Inches(3.2), font_size=9)


def _slide_transport_detail(prs, d, mois_label):
    """Slide incidents transport avec impact."""
    sl = _blank(prs)
    _header(sl, 'Détails Incidents Transport', mois_label=mois_label)
    incs = d['incidents'][:15]
    rows = [[i['site_name'][:25] or '—', i['region'],
             i['escalade'][:20], i['dur_fmt'],
             (i['cause'] or '')[:35]] for i in incs]
    _table(sl, ['SITE / LIEN', 'RÉGION', 'MÉTIER', 'DURÉE', 'CAUSE'],
           rows, col_widths=[3, 2, 2, 2, 4], font_size=9)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES IGW
# ═══════════════════════════════════════════════════════════════════════════════

def _slide_igw(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, f'IGW — Disponibilité et Trafic   ({d["total"]} incidents)', mois_label=mois_label)
    kpis = [
        ('Total Incidents', d['total'], 'IGW', C_BLUE2),
        ('Outage Total', d['total_dur'], '', RGBColor(0xD9, 0x77, 0x06)),
    ]
    _kpi_bar(sl, kpis)

    if d['incidents']:
        _txt(sl, 'Incidents critiques IGW', MARGIN,
             CONTENT_TOP + Inches(1.8), SW - 2*MARGIN, Inches(0.35),
             size=11, bold=True, color=C_BLUE3)
        incs = d['incidents'][:12]
        rows = [[i['site_name'][:30] or i['nature'][:30],
                 i['escalade'][:20],
                 (i['cause'] or '')[:35],
                 i['dur_fmt']] for i in incs]
        fmts = {}
        for i, inc in enumerate(incs):
            if (inc['duration_sec'] or 0) > 72*3600:
                fmts[(i, 3)] = (C_RED_BG, C_RED_FG)
            elif (inc['duration_sec'] or 0) > 24*3600:
                fmts[(i, 3)] = (C_YELL_BG, C_YELL_FG)
        _table(sl, ['LIEN / SITE', 'ESCALADE', 'CAUSE', 'DURÉE'],
               rows, col_widths=[3, 2, 5, 2],
               top=CONTENT_TOP + Inches(2.2),
               height=Inches(3.0),
               cell_fmts=fmts, font_size=9)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES CORE
# ═══════════════════════════════════════════════════════════════════════════════

def _slide_core(prs, d, mois_label):
    sl = _blank(prs)
    _header(sl, f'Core — {d["total"]} incidents', mois_label=mois_label)
    if d['incidents']:
        incs = d['incidents'][:10]
        rows = [[i['nature'][:45] or '—',
                 (i['cause'] or '')[:35],
                 i['escalade'][:20],
                 i['dur_fmt']] for i in incs]
        fmts = {}
        for i, inc in enumerate(incs):
            if (inc['duration_sec'] or 0) > 72*3600:
                fmts[(i, 3)] = (C_RED_BG, C_RED_FG)
        _table(sl, ['NATURE INCIDENT', 'CAUSE', 'ESCALADE', 'DURÉE'],
               rows, col_widths=[5, 4, 2, 2], cell_fmts=fmts, font_size=10)
    else:
        _txt(sl, 'Aucun incident enregistré pour cette période.',
             MARGIN, CONTENT_TOP + Inches(1), SW - 2*MARGIN, Inches(0.5),
             size=14, color=C_BLUE3)


# ═══════════════════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

def generate_report(mois_mobile=None, mois_dr2=None, mois_fixe=None,
                    mois_transport=None, mois_igw=None, mois_core=None,
                    generated_on=''):
    """
    Génère le rapport PPTX complet.
    Retourne un BytesIO prêt à être envoyé en réponse HTTP.
    """
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH

    # Label de mois pour l'affichage
    if mois_mobile:
        mois_label = _mois_label_fr(mois_mobile)
    else:
        from datetime import date as _d
        mois_label = _mois_label_fr(_d.today())

    if not generated_on:
        from datetime import date as _d
        generated_on = _d.today().strftime('%d/%m/%Y')

    # ── Couverture ──────────────────────────────────────────────────────────────
    _cover(prs, mois_label, generated_on)

    # ── Section 01 : Réseau Mobile ──────────────────────────────────────────────
    mob = _fetch_mobile(mois_mobile)

    _section(prs, 1, 'RÉSEAU MOBILE', '📡')
    _def_slide(prs, mois_label)
    _slide_dr1(prs, mob, mois_label)
    _slide_dr2_overview(prs, mob, mois_label)
    _slide_dr2_metier(prs, mob, mois_label)
    _slide_top_sites(prs, mob, mois_label)
    _slide_causes(prs, mob, mois_label)
    _slide_points_bloquants(prs, mob, mois_label)

    # ── Section 02 : Réseau Fixe ────────────────────────────────────────────────
    fixe = _fetch_domain('fixe', mois_fixe)
    _section(prs, 2, 'RÉSEAU FIXE', '☎️')
    _slide_fixe(prs, fixe, mois_label)

    # ── Section 03 : Transport ──────────────────────────────────────────────────
    transport = _fetch_domain('transport', mois_transport)
    _section(prs, 3, 'TRANSPORT', '🔗')
    _slide_transport(prs, transport, mois_label)
    if transport['incidents']:
        _slide_transport_detail(prs, transport, mois_label)

    # ── Section 04 : IGW ────────────────────────────────────────────────────────
    igw = _fetch_domain('igw', mois_igw)
    _section(prs, 4, 'IGW', '🔌')
    _slide_igw(prs, igw, mois_label)

    # ── Section 05 : Core ───────────────────────────────────────────────────────
    core = _fetch_domain('core', mois_core)
    _section(prs, 5, 'CORE', '🌐')
    _slide_core(prs, core, mois_label)

    # ── Fermeture ───────────────────────────────────────────────────────────────
    _closing(prs)

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════════
# GÉNÉRATION DEPUIS FICHIER EXCEL CGI (Fixe, Transport, IGW, Core)
# ═══════════════════════════════════════════════════════════════════════════════

def _adapt_fixe_excel(parsed):
    from collections import defaultdict
    rows  = parsed['rows']
    stats = parsed['stats']
    total_dur = sum(r['duration_sec'] for r in rows if r['duration_sec'])

    nat = defaultdict(int)
    for r in rows:
        nat[r['nature'] or 'Autre'] += 1
    by_nature = [{'nature': k, 'nb': v}
                 for k, v in sorted(nat.items(), key=lambda x: -x[1])[:10]]

    by_region = []
    for name, agg in stats['by_plateforme']:
        nb  = agg['nb']
        dur = agg['dur']
        by_region.append({
            'region': name, 'nb': nb,
            'mttr':   _fmt(dur / nb if nb else 0),
            'outage': _fmt(dur),
        })

    incidents = []
    for r in sorted(rows, key=lambda x: -(x['duration_sec'] or 0))[:20]:
        incidents.append({
            'nature':       r['nature'],
            'region':       r['plateforme'],
            'escalade':     r['escalade'],
            'duration_sec': r['duration_sec'],
            'dur_fmt':      r['duration_fmt'],
            'cause':        r['root_cause'],
            'site_name':    r['site_name'],
            'status':       r['status'],
        })

    return {
        'total': stats['total'],
        'total_dur': _fmt(total_dur),
        'by_nature': by_nature,
        'by_region': by_region,
        'by_escalade': [],
        'incidents': incidents,
    }


def _adapt_transport_excel(parsed):
    rows  = parsed['rows']
    stats = parsed['stats']
    total_dur = sum(r['duration_sec'] for r in rows if r['duration_sec'])

    by_escalade = []
    for name, agg in stats['by_escalade']:
        nb  = agg['nb']
        dur = agg['dur']
        by_escalade.append({
            'escalade': name, 'nb': nb,
            'mttr':   _fmt(dur / nb if nb else 0),
            'outage': _fmt(dur),
        })

    by_region = []
    for name, agg in stats['by_region']:
        nb  = agg['nb']
        dur = agg['dur']
        by_region.append({
            'region': name, 'nb': nb,
            'mttr':   _fmt(dur / nb if nb else 0),
            'outage': _fmt(dur),
        })

    incidents = []
    for r in sorted(rows, key=lambda x: -(x['duration_sec'] or 0))[:20]:
        incidents.append({
            'nature':       r['nature'],
            'region':       r['region'],
            'escalade':     r['escalade'],
            'duration_sec': r['duration_sec'],
            'dur_fmt':      r['duration_fmt'],
            'cause':        r['cause'],
            'site_name':    r['site_name'],
            'status':       r['status'],
        })

    return {
        'total': stats['total'],
        'total_dur': _fmt(total_dur),
        'by_escalade': by_escalade,
        'by_region': by_region,
        'incidents': incidents,
    }


def _adapt_igw_excel(parsed):
    rows  = parsed['rows']
    stats = parsed['stats']
    total_dur = sum(r['duration_sec'] for r in rows if r['duration_sec'])

    incidents = []
    for r in sorted(rows, key=lambda x: -(x['duration_sec'] or 0))[:15]:
        incidents.append({
            'site_name':    r['lien'] or r['lien_internet'] or '—',
            'nature':       r['nature'],
            'escalade':     r['escalade'],
            'duration_sec': r['duration_sec'],
            'dur_fmt':      r['duration_fmt'],
            'cause':        r['cause'],
        })

    return {
        'total': stats['total'],
        'total_dur': _fmt(total_dur),
        'incidents': incidents,
    }


def _adapt_core_excel(parsed):
    rows  = parsed['rows']
    stats = parsed['stats']

    incidents = []
    for r in sorted(rows, key=lambda x: -(x['duration_sec'] or 0))[:15]:
        incidents.append({
            'nature':       r['nature'],
            'cause':        r['root_cause'],
            'escalade':     r['escalade'],
            'duration_sec': r['duration_sec'],
            'dur_fmt':      r['duration_fmt'],
            'site_name':    r['espc'],
            'region':       '',
        })

    return {
        'total': stats['total'],
        'incidents': incidents,
    }


def generate_cgi_from_excel(data, mois_label=''):
    """
    Génère le rapport PPTX multi-plateforme depuis cgi_parser.parse_all().
    data: dict {'fixe': {'rows':..,'stats':..}, 'transport': .., 'igw': .., 'core': ..}
    Retourne un BytesIO prêt à envoyer en réponse HTTP.
    """
    from datetime import date as _d
    if not mois_label:
        mois_label = _mois_label_fr(_d.today())
    generated_on = _d.today().strftime('%d/%m/%Y')

    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH

    _cover(prs, mois_label, generated_on)

    section_num = 1

    if 'fixe' in data:
        _section(prs, section_num, 'RÉSEAU FIXE', '☎️')
        section_num += 1
        _slide_fixe(prs, _adapt_fixe_excel(data['fixe']), mois_label)

    if 'transport' in data:
        _section(prs, section_num, 'TRANSPORT', '🔗')
        section_num += 1
        t = _adapt_transport_excel(data['transport'])
        _slide_transport(prs, t, mois_label)
        if t['incidents']:
            _slide_transport_detail(prs, t, mois_label)

    if 'igw' in data:
        _section(prs, section_num, 'IGW', '🔌')
        section_num += 1
        _slide_igw(prs, _adapt_igw_excel(data['igw']), mois_label)

    if 'core' in data:
        _section(prs, section_num, 'CORE', '🌐')
        section_num += 1
        _slide_core(prs, _adapt_core_excel(data['core']), mois_label)

    _closing(prs)

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
