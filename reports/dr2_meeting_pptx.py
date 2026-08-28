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
from calendar import monthrange
from datetime import date, timedelta
from collections import Counter
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (
    XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION,
)

from .pptx_report import (
    _blank as _plain_blank, _rect, _txt, _table as _base_table,
    SW, SH, MARGIN, CONTENT_TOP, CONTENT_H,
    C_BLUE, C_WHITE, C_DTEXT, C_LGRAY, C_BLUE3, C_YELL,
    C_RED_BG, C_RED_FG, C_YELL_BG, C_YELL_FG, C_GREEN_BG, C_GREEN_FG,
)
from .dr2_availability import (
    DR2_REGION_TARGETS, DR2_ESCALADE_ORDER, normalize_dr2_datetimes,
)

C_RED_T = RGBColor(0xC0, 0x00, 0x00)
C_DETAIL_HDR = RGBColor(0x44, 0x54, 0x6A)
C_DETAIL_GREEN = RGBColor(0x70, 0xAD, 0x47)
C_DETAIL_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
C_SUMMARY_GRAY = RGBColor(0xE1, 0xE3, 0xE8)
C_CAUSE_HDR = RGBColor(0xD9, 0xE8, 0xF5)
C_REPORT_NAVY = RGBColor(0x00, 0x20, 0x60)
C_REPORT_BLUE = RGBColor(0x00, 0x30, 0x87)
C_REPORT_CYAN = RGBColor(0x00, 0xB0, 0xF0)
C_REPORT_GREEN = RGBColor(0x63, 0xBE, 0x7B)
C_REPORT_PALE_GREEN = RGBColor(0xE2, 0xF0, 0xD9)
C_REPORT_YELLOW = RGBColor(0xFF, 0xE6, 0x80)
C_REPORT_ORANGE = RGBColor(0xF4, 0xB1, 0x83)
C_REPORT_RED = RGBColor(0xF8, 0x69, 0x6B)

_JOURS_FR = ['LUNDI', 'MARDI', 'MERCREDI', 'JEUDI', 'VENDREDI', 'SAMEDI', 'DIMANCHE']
_MOIS_FR = ['', 'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
            'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE']

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_GDI_REFERENCE = 'presentation a automatiser GDI.pptx'
_REUNION_REFERENCE = 'PRESENTATION REUNION[1] vendredi NEW (9).pptx'


def _table(*args, **kwargs):
    """Tableau GDI homogène, compact et lisible sur écran de réunion."""
    table = _base_table(*args, **kwargs)
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.04)
            cell.margin_right = Inches(0.04)
            cell.margin_top = Inches(0.025)
            cell.margin_bottom = Inches(0.025)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.space_before = Pt(0)
                paragraph.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    if row_idx == 0:
                        run.font.bold = True
    return table


# ═══════════════════════════════════════════════════════════════════════════
# BAS NIVEAU — bandeau titre / logo / diapos fixes
# ═══════════════════════════════════════════════════════════════════════════

def _layout(prs, name):
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == name:
                return layout
    return None


def _remove_slide(prs, slide):
    for slide_id in list(prs.slides._sldIdLst):
        if int(slide_id.id) == slide.slide_id:
            prs.part.drop_rel(slide_id.rId)
            prs.slides._sldIdLst.remove(slide_id)
            return


def _reference_deck(filename):
    """Charge le vrai support YAS et conserve ses gardes natives.

    Les diapositives de contenu sont recréées avec le layout ``Pic1`` du
    modèle. En cas de modèle absent, la génération autonome reste disponible.
    """
    template_path = _PROJECT_ROOT / filename
    if not template_path.is_file():
        prs = Presentation()
        prs.slide_width = SW
        prs.slide_height = SH
        return prs

    prs = Presentation(str(template_path))
    for slide in list(prs.slides):
        if slide.slide_layout.name != 'Pause':
            _remove_slide(prs, slide)
    return prs


def _blank(prs):
    layout = _layout(prs, 'Pic1')
    if layout is None:
        return _plain_blank(prs)
    slide = prs.slides.add_slide(layout)
    for placeholder in list(slide.placeholders):
        slide.shapes._spTree.remove(placeholder._element)
    return slide

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
    uses_reference_layout = slide.slide_layout.name == 'Pic1'
    if not uses_reference_layout:
        _rect(slide, 0, 0, SW, SH, C_WHITE)
    _txt(slide, tag, MARGIN, Inches(0.16), Inches(9.5), Inches(0.45),
         size=20, bold=True, color=C_BLUE)
    if subtitle:
        _txt(slide, subtitle, MARGIN, Inches(0.64), Inches(9.5), Inches(0.4),
             size=20, bold=True, color=C_RED_T)
    if extra_right:
        _txt(slide, extra_right, Inches(8.6), Inches(0.66), Inches(3.0), Inches(0.35),
             size=12, bold=True, color=C_RED_T, align=PP_ALIGN.RIGHT)
    if not uses_reference_layout:
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
    for slide in prs.slides:
        if slide.slide_layout.name == 'Pause':
            return slide
    layout = _layout(prs, 'Pause')
    sl = prs.slides.add_slide(layout) if layout is not None else _plain_blank(prs)
    _rect(sl, 0, 0, SW, SH, C_YELL)
    _logo(sl, l=Inches(4.9), t=Inches(2.35), h=Inches(2.8))
    return sl


def _closing(prs):
    """Diapositive de fin — reproduit la diapo « MERCI » (layout « End1 ») :
    blob bleu en haut à gauche, texte MERCI, mention YAS Togo, logo en bas à droite."""
    layout = _layout(prs, 'End1')
    if layout is not None:
        sl = prs.slides.add_slide(layout)
        placeholders = sorted(sl.placeholders, key=lambda shape: shape.top)
        if placeholders:
            merci = placeholders[0]
            merci.left = Inches(1.0)
            merci.top = Inches(1.65)
            merci.width = Inches(5.0)
            merci.height = Inches(1.2)
            merci.text_frame.word_wrap = False
            merci.text = 'MERCI'
            run = merci.text_frame.paragraphs[0].runs[0]
            run.font.size = Pt(54)
            run.font.bold = True
            run.font.italic = True
            run.font.color.rgb = RGBColor(0x7E, 0xA6, 0xD9)
        footer = next((shape for shape in placeholders if shape.top > Inches(5)), None)
        if footer is not None:
            footer.text = 'YAS Togo / DT / DOC / iSOC / GDI'
        return sl

    sl = _plain_blank(prs)
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


def _set_placeholder(slide, idx, text):
    shape = next(
        (item for item in slide.placeholders if item.placeholder_format.idx == idx),
        None,
    )
    if shape is not None:
        shape.text = text
    return shape


def _meeting_title(prs, meeting_day):
    layout = _layout(prs, 'Title1')
    if layout is None:
        sl = _blank(prs)
        _header(sl, 'REUNION GESTION DES INCIDENTS', meeting_day.strftime('%d/%m/%Y'))
        return sl
    sl = prs.slides.add_slide(layout)
    _set_placeholder(sl, 0, 'REUNION GESTION DES INCIDENTS')
    _set_placeholder(sl, 1, '')
    _set_placeholder(
        sl, 11,
        f'{_JOURS_FR[meeting_day.weekday()].title()} {meeting_day.day} '
        f'{_MOIS_FR[meeting_day.month]} {meeting_day.year}',
    )
    _set_placeholder(sl, 10, 'Yas Togo / DT / DOC / iSOC / GDI')
    return sl


def _section(prs, title, month_day, number):
    layout = _layout(prs, 'Section1')
    if layout is None:
        sl = _blank(prs)
        _header(sl, title, f'{_MOIS_FR[month_day.month]} {month_day.year}')
        return sl
    sl = prs.slides.add_slide(layout)
    _set_placeholder(sl, 0, title)
    _set_placeholder(sl, 1, f'{_MOIS_FR[month_day.month]} {month_day.year}')
    _set_placeholder(sl, 10, f'{number:02d}')
    return sl


def _slide_definitions(prs):
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'Définition DR1/DR2')
    _table(
        sl, ['Code', 'Indicateur', 'Définition', 'Seuil (2G,3G,4G)'],
        _DEF_ROWS, col_widths=[1, 3, 5, 1.5], top=CONTENT_TOP + Inches(1.0),
        height=Inches(3.2), font_size=10,
    )
    return sl


def _summary_card(slide, label, value, detail, left, color):
    top = Inches(1.55)
    width = Inches(2.85)
    height = Inches(1.35)
    _rect(slide, left, top, width, height, RGBColor(0xF5, 0xF7, 0xFA))
    _rect(slide, left, top, Inches(0.08), height, color)
    text_boxes = [
        _txt(slide, label, left + Inches(0.22), top + Inches(0.13),
             width - Inches(0.35), Inches(0.30), size=10, bold=True,
             color=C_REPORT_BLUE, wrap=False),
        _txt(slide, value, left + Inches(0.22), top + Inches(0.43),
             width - Inches(0.35), Inches(0.46), size=25, bold=True,
             color=color, wrap=False),
        _txt(slide, detail, left + Inches(0.22), top + Inches(0.96),
             width - Inches(0.35), Inches(0.28), size=8, color=C_DTEXT,
             wrap=False),
    ]
    for text_box in text_boxes:
        text_box.text_frame.margin_top = 0
        text_box.text_frame.margin_bottom = 0
        for paragraph in text_box.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Arial'


def _slide_executive_summary(prs, month_data, detail_data, qs_month):
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'SYNTHÈSE EXÉCUTIVE DR2')
    _txt(
        sl,
        f"Période analysée : {month_data['debut'].strftime('%d/%m/%Y')} au "
        f"{month_data['fin'].strftime('%d/%m/%Y')}",
        MARGIN, Inches(1.08), SW - 2 * MARGIN, Inches(0.28),
        size=10, color=C_DTEXT,
    )

    recurring_sites = sum(1 for _, count in _top_sites(qs_month, 1000) if count > 1)
    processed_day_label = (
        'jour traité' if month_data['period_days'] == 1 else 'jours traités'
    )
    cards = [
        ('TOTAL DR2', str(month_data['total_dr2']), 'Mois en cours', C_REPORT_BLUE),
        ('MOYENNE / JOUR', str(month_data['moyenne']).replace('.', ','),
         f"{month_data['period_days']} {processed_day_label}", C_REPORT_CYAN),
        ('DERNIER WEEK-END', str(detail_data['total_dr2']),
         str(detail_data['moyenne']).replace('.', ',') + ' DR2 / jour', C_RED_T),
        ('SITES RÉCURRENTS', str(recurring_sites),
         'Au moins 2 occurrences', C_DETAIL_GREEN),
    ]
    for index, (label, value, detail, color) in enumerate(cards):
        _summary_card(sl, label, value, detail,
                      Inches(0.55 + index * 3.15), color)

    _txt(sl, 'POINTS D’ATTENTION', Inches(0.55), Inches(3.30),
         Inches(5.8), Inches(0.32), size=14, bold=True, color=C_REPORT_BLUE)
    _rect(sl, Inches(0.55), Inches(3.68), Inches(12.05), Inches(2.35),
          RGBColor(0xF5, 0xF7, 0xFA))

    region_rows = [row for row in month_data['region_rows'] if row['dr2']]
    worst_region = max(region_rows, key=lambda row: row['dr2'], default=None)
    top_causes = _causes_breakdown(qs_month, 1, total=month_data['total_dr2'])
    blocking_count = sum(count for _, count in _points_bloquants(qs_month))
    attention = [
        (
            'RÉGION PRIORITAIRE',
            f"{worst_region['region']} — {worst_region['dr2']} DR2 "
            f"({worst_region['pct_reg']} % du total)"
            if worst_region else 'Aucune donnée régionale',
            C_RED_T,
        ),
        (
            'CAUSE DOMINANTE',
            f'{top_causes[0][0]} — {top_causes[0][1]} cas ({top_causes[0][2]} %)'
            if top_causes else 'Aucune cause renseignée',
            C_REPORT_CYAN,
        ),
        (
            'POINTS BLOQUANTS',
            f'{blocking_count} DR2 concernés' if blocking_count
            else 'Aucun point bloquant renseigné',
            C_DETAIL_GREEN if not blocking_count else C_REPORT_ORANGE,
        ),
    ]
    for index, (label, value, color) in enumerate(attention):
        top = Inches(3.92 + index * 0.62)
        _rect(sl, Inches(0.82), top, Inches(0.10), Inches(0.38), color)
        _txt(sl, label, Inches(1.08), top - Inches(0.01), Inches(2.15),
             Inches(0.40), size=10, bold=True, color=C_REPORT_BLUE)
        _txt(sl, value, Inches(3.28), top - Inches(0.01), Inches(8.85),
             Inches(0.40), size=11, bold=True, color=C_DTEXT)
    return sl


def _slide_dr1_violations(prs, fin):
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'Cas de violation DR1')
    dr1 = _dr1_violations(fin)
    period_start = fin - timedelta(days=29)
    _txt(
        sl, f'Du {period_start.strftime("%d/%m/%Y")} au {fin.strftime("%d/%m/%Y")}',
        MARGIN, CONTENT_TOP, Inches(4.0), Inches(0.35), size=11, bold=True,
        color=C_RED_T,
    )
    if dr1 is None:
        _txt(sl, 'Données ticketing indisponibles pour le calcul DR1.',
             MARGIN, CONTENT_TOP + Inches(0.55), SW - 2 * MARGIN,
             Inches(0.5), size=13, color=C_RED_FG)
        return sl
    _kpi_text(sl, f'TDR1 = {len(dr1)} DR1', MARGIN, CONTENT_TOP + Inches(0.48))
    if dr1:
        rows = [[r['site'], r['cnt'], r['region'], str(r['cause'])[:45]] for r in dr1[:12]]
        _table(
            sl, ['SITE NAME', 'COUNT', 'RÉGION', 'CAUSES'], rows,
            col_widths=[3, 1, 2, 6], top=CONTENT_TOP + Inches(1.0),
            height=Inches(4.7), font_size=10,
        )
    else:
        _txt(sl, 'Aucun site en violation DR1 sur les 30 derniers jours.',
             MARGIN, CONTENT_TOP + Inches(1.25), SW - 2 * MARGIN,
             Inches(0.5), size=13, color=C_DTEXT)
    return sl


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
    blank_names = {
        row['site_name'] for row in rows
        if not row.get('region') or row.get('region') == '—'
    }
    if blank_names:
        from .models import Site
        region_fallback = {
            site_name: (region or '').strip().upper()
            for site_name, region in Site.objects.filter(
                site_name__in=blank_names,
            ).values_list('site_name', 'region')
        }
        for row in rows:
            if not row.get('region') or row.get('region') == '—':
                row['region'] = region_fallback.get(row['site_name']) or 'INCONNU'
    dataset = _build_dr2_from_rows(rows, debut, fin)
    from .models import Dr2ProcessedDate
    processed_dates = set(Dr2ProcessedDate.objects.filter(
        date__gte=debut, date__lte=fin,
    ).values_list('date', flat=True))
    if not processed_dates and rows:
        processed_dates = {
            date.fromisoformat(row['date']) for row in rows if row.get('date')
        }
    dataset['processed_dates'] = processed_dates
    dataset['period_days'] = len(processed_dates)
    dataset['moyenne'] = (
        round(dataset['total_dr2'] / len(processed_dates), 2)
        if processed_dates else 0
    )
    dataset['nbre_j1'] = sum(
        1 for row in dataset['detail_rows'] if row['date'] == fin
    )
    unknown_rows = [row for row in dataset['detail_rows'] if row['region'] == 'INCONNU']
    if unknown_rows:
        dataset['region_rows'].append({
            'region': 'INCONNU',
            'tget': 0,
            'dr2': len(unknown_rows),
            'pct_reg': round(len(unknown_rows) / dataset['total_dr2'] * 100)
                       if dataset['total_dr2'] else 0,
            'pct_tget': 0,
            'escalades': {},
            'color': 'green',
        })
    return dataset


def _duree_str(alarm_time, cancel_time):
    if not alarm_time:
        return ''
    if not cancel_time:
        return 'EN COURS'
    sec = int((cancel_time - alarm_time).total_seconds())
    if sec < 0:
        return 'EN COURS'
    return f'{sec // 3600}:{(sec % 3600) // 60:02d}:{sec % 60:02d}'


_DETAIL_HEADERS = ['N°', 'Ticket', 'Site parent', 'Site Name', 'Site ID', 'Alarm Time',
                    'Durée', 'Catégorie', 'Cause', 'Root Cause', 'Pt bloquant',
                    'Cancel Time', 'Observation', 'DR2']
_DETAIL_COL_W = [0.30, 1.25, 0.65, 1.0, 0.50, 0.90, 0.65, 0.70,
                 1.20, 1.25, 1.15, 1.15, 1.55, 0.40]


def _detail_rows(qs):
    rows = []
    for i, rec in enumerate(qs, 1):
        alarm_time, cancel_time = normalize_dr2_datetimes(
            rec.date, rec.alarm_time, rec.cancel_time)
        rows.append([
            i,
            rec.numero_ticket or '',
            rec.site_parent or '\xa0',
            rec.site_name,
            rec.site_id or '',
            alarm_time.strftime('%d-%m-%Y %H:%M') if alarm_time else '',
            _duree_str(alarm_time, cancel_time),
            rec.categorie or '—',
            (rec.cause or '')[:38],
            (rec.root_cause or '')[:42],
            rec.point_bloquant or 'N/A',
            cancel_time.strftime('%d-%m-%Y %H:%M') if cancel_time else '\xa0',
            ' '.join((rec.observation or '').split())[:45],
            'OUI',
        ])
    return rows


def _set_header_text_color(table, color):
    for cell in table.rows[0].cells:
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = color


def _detail_table(slide, rows, day, count, height=Inches(2.45)):
    top = CONTENT_TOP + Inches(0.14)
    band_h = Inches(0.22)
    width = SW - 2 * MARGIN
    _rect(slide, MARGIN, top, width, band_h, C_DETAIL_GRAY)
    _txt(slide, f'CAS DE VIOLATION DR2 {day.strftime("%d-%m-%Y")}',
         MARGIN, top, width, band_h, size=9, bold=True, color=C_DTEXT,
         align=PP_ALIGN.CENTER)
    _rect(slide, MARGIN, top + band_h, width, band_h, C_YELL)
    _txt(slide, f'{count} DR2', MARGIN, top + band_h, width, band_h,
         size=9, bold=True, color=C_DTEXT, align=PP_ALIGN.CENTER)

    cell_fmts = {}
    for row_idx, row in enumerate(rows):
        row_bg = C_WHITE if row_idx % 2 == 0 else C_LGRAY
        for col_idx in range(len(_DETAIL_HEADERS)):
            cell_fmts[(row_idx, col_idx)] = (row_bg, C_DTEXT)
        if str(row[2]).strip():
            cell_fmts[(row_idx, 2)] = (C_YELL, C_DTEXT)
        cell_fmts[(row_idx, 7)] = (C_CAUSE_HDR, C_DTEXT)
        cell_fmts[(row_idx, len(_DETAIL_HEADERS) - 1)] = (
            C_DETAIL_GREEN, C_DTEXT,
        )

    table = _table(
        slide, _DETAIL_HEADERS, rows, col_widths=_DETAIL_COL_W,
        top=top + 2 * band_h, height=height, font_size=6, hdr_size=6,
        hdr_bg=C_DETAIL_HDR, alt=False, cell_fmts=cell_fmts,
    )
    observation_index = _DETAIL_HEADERS.index('Observation')
    for row_index in range(1, len(table.rows)):
        row = table.rows[row_index]
        cell = row.cells[observation_index]
        cell.margin_left = Inches(0.03)
        cell.margin_right = Inches(0.03)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.LEFT
            for run in paragraph.runs:
                run.font.size = Pt(5)
                run.font.bold = False
    return table


def _dr2_summary_tables(slide, qs, total, top=Inches(4.75), height=Inches(1.8)):
    esc = _escalade_breakdown(qs, total)[:6]
    summary_rows = [[e, k, f'{p}%'] for e, k, p in esc]
    summary_fmts = {
        (i, j): (C_SUMMARY_GRAY, C_DTEXT)
        for i in range(len(summary_rows))
        for j in range(3)
    }
    summary_table = _table(
        slide, [f'TOTAL DR2 = {total}', 'NBRE INC', '% GENERAL'],
        summary_rows, left=MARGIN, top=top, width=Inches(4.4), height=height,
        col_widths=[3, 1, 1], font_size=8, hdr_size=8, hdr_bg=C_YELL,
        alt=False, cell_fmts=summary_fmts,
    )
    _set_header_text_color(summary_table, C_DTEXT)

    causes = _causes_breakdown(qs, 6, total=total)
    causes_table = _table(
        slide, ['TOP RECURRENT CAUSES', 'Nombre de CAUSE', '% CAUSE'],
        [[c[:35], k, f'{p}%'] for c, k, p in causes],
        left=Inches(5.0), top=top, width=Inches(7.6), height=height,
        col_widths=[5, 1, 1], font_size=8, hdr_size=8,
        hdr_bg=C_CAUSE_HDR, alt=False,
    )
    _set_header_text_color(causes_table, C_DTEXT)


def _escalade_breakdown(qs, total=None):
    cnt = Counter(rec.categorie or '—' for rec in qs)
    total = total or sum(cnt.values()) or 1
    ordered = [e for e in DR2_ESCALADE_ORDER if e in cnt] + [e for e in cnt if e not in DR2_ESCALADE_ORDER]
    return [(e, cnt[e], round(cnt[e] / total * 100)) for e in ordered]


def _causes_breakdown(qs, n=10, total=None):
    cnt = Counter((rec.cause or '—').strip() for rec in qs if (rec.cause or '').strip())
    denominator = total or sum(cnt.values()) or 1
    return [(c, k, round(k / denominator * 100)) for c, k in cnt.most_common(n)]


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


def _shift_month(day, offset):
    month_index = day.year * 12 + day.month - 1 + offset
    year, month_zero = divmod(month_index, 12)
    return date(year, month_zero + 1, 1)


def _gdi_reporting_periods(fin):
    """Retourne les périmètres du support GDI de la réunion du lundi.

    La date de fin choisie peut être postérieure au week-end présenté. Le
    support s'arrête alors au dernier dimanche achevé : tendances depuis le
    premier du mois, détails du vendredi au dimanche, réunion le lundi.
    """
    last_sunday = fin - timedelta(days=(fin.weekday() - 6) % 7)
    return {
        'month_start': last_sunday.replace(day=1),
        'data_end': last_sunday,
        'detail_start': last_sunday - timedelta(days=2),
        'detail_end': last_sunday,
        'meeting_day': last_sunday + timedelta(days=1),
    }


def _monthly_comparison(fin, count=3):
    """Séries journalières des derniers mois, issues des mêmes DR2 que le
    tableau quotidien. La moyenne utilise uniquement les jours traités."""
    from .models import Dr2ProcessedDate

    months = []
    current_start = fin.replace(day=1)
    for offset in range(1 - count, 1):
        start = _shift_month(current_start, offset)
        calendar_end = date(start.year, start.month, monthrange(start.year, start.month)[1])
        end = min(calendar_end, fin) if start == current_start else calendar_end
        counts = Counter(rec.date for rec in _dr2_qs(start, end))
        processed_dates = set(Dr2ProcessedDate.objects.filter(
            date__gte=start, date__lte=end,
        ).values_list('date', flat=True))
        total = sum(counts.values())
        months.append({
            'label': _MOIS_FR[start.month],
            'year': start.year,
            'total': total,
            'moyenne': round(total / len(processed_dates), 2) if processed_dates else None,
            'processed_days': len(processed_dates),
            'values': [
                counts.get(current_day, 0) if current_day in processed_dates else None
                for day in range(1, end.day + 1)
                for current_day in [date(start.year, start.month, day)]
            ],
        })
    return months


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

def _style_count_axes(chart, values):
    """Style commun aux graphes de volumes : aucune graduation fractionnaire."""
    numeric_values = [float(value) for value in values if value is not None]
    maximum = max(numeric_values, default=0)
    if maximum <= 5:
        major_unit = 1
    elif maximum <= 10:
        major_unit = 2
    elif maximum <= 30:
        major_unit = 5
    elif maximum <= 60:
        major_unit = 10
    elif maximum <= 100:
        major_unit = 20
    else:
        major_unit = max(10, int((maximum + 49) // 50) * 10)
    value_axis = chart.value_axis
    value_axis.minimum_scale = 0
    value_axis.major_unit = major_unit
    value_axis.tick_labels.number_format = '0'
    value_axis.tick_labels.font.name = 'Arial'
    value_axis.tick_labels.font.size = Pt(9)
    value_axis.tick_labels.font.color.rgb = C_REPORT_BLUE
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xE2, 0xF3)
    value_axis.format.line.color.rgb = RGBColor(0x9E, 0xAD, 0xC7)
    category_axis = chart.category_axis
    category_axis.tick_labels.font.name = 'Arial'
    category_axis.tick_labels.font.size = Pt(9)
    category_axis.tick_labels.font.color.rgb = C_REPORT_BLUE
    category_axis.format.line.color.rgb = RGBColor(0x9E, 0xAD, 0xC7)


def _show_count_labels(plot, position=XL_DATA_LABEL_POSITION.OUTSIDE_END):
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.show_value = True
    labels.show_legend_key = False
    labels.show_category_name = False
    labels.number_format = '0'
    labels.position = position
    labels.font.name = 'Arial'
    labels.font.size = Pt(9)
    labels.font.bold = True
    labels.font.color.rgb = C_REPORT_BLUE

def _line_chart(slide, categories, values, l, t, w, h, title=''):
    data = CategoryChartData()
    data.categories = [c.strftime('%d/%m') if hasattr(c, 'strftime') else str(c) for c in categories]
    data.add_series('DR2', values)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, l, t, w, h, data)
    chart = gframe.chart
    chart.has_legend = False
    _style_count_axes(chart, values)
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
    plot = chart.plots[0]
    plot.series[0].format.line.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    plot.series[0].format.line.width = Pt(2.5)
    _show_count_labels(plot, XL_DATA_LABEL_POSITION.ABOVE)
    return chart


def _bar_chart(slide, categories, values, l, t, w, h, horizontal=True,
               color=C_BLUE, title='Nombre de DR2', ranked=False):
    data = CategoryChartData()
    data.categories = [str(c) for c in categories]
    data.add_series('Nb', values)
    ctype = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gframe = slide.shapes.add_chart(ctype, l, t, w, h, data)
    chart = gframe.chart
    chart.has_legend = False
    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
        title_font = chart.chart_title.text_frame.paragraphs[0].font
        title_font.name = 'Arial'
        title_font.size = Pt(14)
        title_font.bold = True
        title_font.color.rgb = C_REPORT_BLUE
    _style_count_axes(chart, values)
    plot = chart.plots[0]
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = color
    plot.series[0].format.line.fill.background()
    if ranked:
        plot.gap_width = 90
    elif len(categories) >= 8:
        plot.gap_width = 55
    elif len(categories) >= 4:
        plot.gap_width = 100
    else:
        plot.gap_width = 180
    _show_count_labels(
        plot,
        XL_DATA_LABEL_POSITION.INSIDE_END if ranked else XL_DATA_LABEL_POSITION.OUTSIDE_END,
    )
    if ranked:
        labels = plot.data_labels
        labels.font.color.rgb = C_WHITE
        labels.font.size = Pt(10)
        chart.value_axis.maximum_scale = max(values) + chart.value_axis.major_unit
        chart.value_axis.has_major_gridlines = False
        chart.value_axis.tick_labels.font.size = Pt(1)
        chart.value_axis.tick_labels.font.color.rgb = C_WHITE
        chart.value_axis.format.line.fill.background()
        chart.category_axis.format.line.fill.background()
        chart.category_axis.tick_labels.font.size = Pt(10)
        points = plot.series[0].points
        podium_colors = [
            RGBColor(0x00, 0x8F, 0xC5),
            RGBColor(0x18, 0x72, 0xB8),
            RGBColor(0x26, 0x25, 0x79),
        ]
        for point, point_color in zip(reversed(points), podium_colors):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = point_color
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
    chart.legend.font.name = 'Arial'
    chart.legend.font.size = Pt(9)
    plot = chart.plots[0]
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.show_category_name = True
    labels.show_percentage = True
    labels.show_legend_key = False
    labels.position = XL_DATA_LABEL_POSITION.BEST_FIT
    labels.font.name = 'Arial'
    labels.font.size = Pt(8)
    labels.font.bold = True
    return chart


# ═══════════════════════════════════════════════════════════════════════════
# DIAPOSITIVES PARTAGÉES (utilisées par les 2 supports)
# ═══════════════════════════════════════════════════════════════════════════

def _slide_dr2_trend(prs, d, kpi_label, period_label):
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'DR2 TREND')
    days = _daily_counts(_dr2_qs(d['debut'], d['fin']), d['debut'], d['fin'])
    processed_dates = d.get('processed_dates', set())
    values = [
        count if day in processed_dates or count else None
        for day, count in days
    ]
    _line_chart(sl, [x[0] for x in days], values,
             MARGIN, CONTENT_TOP + Inches(0.45), SW - 2 * MARGIN,
             CONTENT_H - Inches(0.55))
    trend_label = f"TREND DR2 {_MOIS_FR[d['fin'].month]} {d['fin'].year}"
    _rect(sl, Inches(4.7), CONTENT_TOP, Inches(3.2), Inches(0.42), C_YELL)
    _txt(sl, trend_label, Inches(4.7), CONTENT_TOP + Inches(0.03),
        Inches(3.2), Inches(0.34), size=17, bold=True, color=C_DTEXT,
        align=PP_ALIGN.CENTER)
    _rect(sl, Inches(10.2), CONTENT_TOP + Inches(0.65), Inches(2.2), Inches(1.0), C_BLUE)
    if processed_dates:
        kpi_text = f"{kpi_label} = {d['total_dr2']}\nMOY = {str(d['moyenne']).replace('.', ',')}"
    else:
        kpi_text = f'{kpi_label} = N/D\nMOY = N/D'
    _txt(sl, kpi_text, Inches(10.2), CONTENT_TOP + Inches(0.78),
        Inches(2.2), Inches(0.75), size=16, color=C_WHITE,
        align=PP_ALIGN.CENTER)
    return sl


def _report_cell(cell, value='', bg=C_WHITE, fg=C_DTEXT, bold=False,
                 size=8, align=PP_ALIGN.CENTER):
    cell.text = ''
    cell.fill.solid()
    cell.fill.fore_color.rgb = bg
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.025)
    cell.margin_right = Inches(0.025)
    cell.margin_top = Inches(0.015)
    cell.margin_bottom = Inches(0.015)
    paragraph = cell.text_frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = str(value) if value is not None else ''
    run.font.name = 'Arial Narrow'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = fg


def _region_color(row):
    return {
        'green': RGBColor(0xA9, 0xD1, 0x8E),
        'yellow': C_REPORT_YELLOW,
        'red': C_REPORT_RED,
    }[row['color']]


def _use_office_font(text_box, name='Arial'):
    for paragraph in text_box.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = name
    return text_box


def _heat_color(value, maximum):
    if not value:
        return RGBColor(0xF2, 0xF4, 0xF7)
    intensity = value / max(maximum, 1)
    if intensity >= 0.66:
        return RGBColor(0x00, 0x65, 0x8F)
    if intensity >= 0.33:
        return RGBColor(0x00, 0xA6, 0xC8)
    return RGBColor(0xC9, 0xEB, 0xF2)


def _daily_report_stats(d):
    fixed_names = list(DR2_ESCALADE_ORDER) + ['ENVIRONNEMENT']
    extra_names = [
        stat['escalade'] for stat in d['esc_stats']
        if stat['escalade'] not in fixed_names
    ]
    rows = d['detail_rows']
    if any(row['categorie'] == '—' for row in rows):
        extra_names.append('NON RENSEIGNÉ')
    stats = []
    total_metier = sum(1 for row in rows if row['is_resolved'])
    for name in fixed_names + extra_names:
        category_rows = [
            row for row in rows
            if row['categorie'] == name
            or (name == 'NON RENSEIGNÉ' and row['categorie'] == '—')
        ]
        metier = sum(1 for row in category_rows if row['is_resolved'])
        total = len(category_rows)
        stats.append({
            'escalade': name,
            'total': total,
            'metier': metier,
            'pct_tgt': round(metier / total * 100) if total else None,
            'pct_tt': round(metier / total_metier * 100) if total_metier else 0,
        })
    return stats


def _category_matches(row, category):
    return row['categorie'] == category or (
        category == 'NON RENSEIGNÉ' and row['categorie'] == '—'
    )


def _slide_apercu_global(prs, d, subtitle='Aperçu global et comparatif des tendances DR2'):
    sl = _blank(prs)
    _header(sl, 'RAPPORT GESTION DES INCIDENTS', subtitle)
    esc_stats = _daily_report_stats(d)
    esc_names = [stat['escalade'] for stat in esc_stats]

    left = Inches(0.15)
    width = SW - Inches(0.30)
    _rect(sl, Inches(0.50), Inches(1.42), SW - Inches(0.65), Inches(0.40), C_REPORT_NAVY)
    _txt(sl, 'DR2 / DAILY REPORT', Inches(0.50), Inches(1.45),
         SW - Inches(0.65), Inches(0.32), size=18, bold=True, color=C_WHITE,
         align=PP_ALIGN.CENTER)
    _txt(sl, f"Du {d['debut'].strftime('%d/%m/%Y')} au {d['fin'].strftime('%d/%m/%Y')}",
         Inches(0.50), Inches(1.84), SW - Inches(0.65), Inches(0.28),
         size=12, bold=True, color=RGBColor(0xFF, 0x00, 0x00),
         align=PP_ALIGN.CENTER)

    n_cols = 6 + len(esc_names)
    n_rows = 12
    table_top = Inches(2.18)
    table_height = Inches(4.70)
    table = sl.shapes.add_table(n_rows, n_cols, left, table_top, width, table_height).table
    fixed_widths = [1.18, 0.58, 0.68, 0.66, 0.82, 0.28]
    remaining = 13.03 - sum(fixed_widths)
    widths = fixed_widths + [remaining / max(len(esc_names), 1)] * len(esc_names)
    for idx, cell_width in enumerate(widths):
        table.columns[idx].width = Inches(cell_width)
    row_heights = [0.48, 0.34] + [0.38] * 6 + [0.42, 0.42, 0.38, 0.38]
    for idx, row_height in enumerate(row_heights):
        table.rows[idx].height = Inches(row_height)

    left_headers = ['NBRE DR2 / J-1', 'TGET', 'DR2 REG', '% REG', '% TGET REG']
    for col_idx, label in enumerate(left_headers):
        _report_cell(table.cell(0, col_idx), label, C_WHITE, C_DTEXT, True, 8)
    for col_idx in range(5):
        _report_cell(table.cell(1, col_idx), '', C_REPORT_NAVY, C_WHITE, True, 9)
    _report_cell(table.cell(1, 0), d['nbre_j1'], C_WHITE, C_DTEXT, True, 12)

    for row_idx in range(n_rows):
        _report_cell(table.cell(row_idx, 5), '', C_YELL, C_DTEXT, True, 7)

    for esc_idx, stat in enumerate(esc_stats):
        col_idx = 6 + esc_idx
        _report_cell(table.cell(0, col_idx), stat['escalade'], C_DETAIL_GRAY,
                     C_DTEXT, True, 7)
        _report_cell(table.cell(1, col_idx), stat['total'], C_REPORT_NAVY,
                     C_WHITE, True, 10)

    matrix_max = max(
        [value for row in d['region_rows'] for value in row['escalades'].values()] or [0]
    )
    for region_idx, region in enumerate(d['region_rows']):
        row_idx = 2 + region_idx
        values = [region['region'], region['tget'], region['dr2'],
                  f"{region['pct_reg']}%", f"{region['pct_tget']}%"]
        _report_cell(table.cell(row_idx, 0), values[0], C_WHITE, C_DTEXT,
                     True, 9, PP_ALIGN.LEFT)
        _report_cell(table.cell(row_idx, 1), values[1], C_REPORT_NAVY, C_WHITE, True, 9)
        region_bg = _region_color(region)
        for col_idx in (2, 3, 4):
            _report_cell(table.cell(row_idx, col_idx), values[col_idx], region_bg,
                         C_DTEXT, True, 9)
        for esc_idx, esc_name in enumerate(esc_names):
            value = sum(
                1 for row in d['detail_rows']
                if row['region'] == region['region']
                and _category_matches(row, esc_name)
            )
            _report_cell(table.cell(row_idx, 6 + esc_idx), value,
                         _heat_color(value, matrix_max),
                         C_WHITE if value >= max(5, matrix_max * 0.5) else C_DTEXT,
                         True, 9)

    total_pct = round(d['total_dr2'] / d['total_tget'] * 100) if d['total_tget'] else 0
    total_row = ['TOTAL DR2', d['total_tget'], d['total_dr2'], '100%', f'{total_pct}%']
    for col_idx, value in enumerate(total_row):
        _report_cell(table.cell(8, col_idx), value,
                     C_REPORT_NAVY if col_idx != 2 else RGBColor(0x70, 0xAD, 0x47),
                     C_WHITE if col_idx != 2 else C_DTEXT, True, 9)
    for esc_idx, stat in enumerate(esc_stats):
        _report_cell(table.cell(8, 6 + esc_idx), stat['total'],
                     C_REPORT_CYAN, C_WHITE, True, 10)

    average_values = ['MOYENNE DR2', '', str(d['moyenne']).replace('.', ','), '', '% TGET/MÉTIER']
    for col_idx, value in enumerate(average_values):
        bg = C_REPORT_CYAN if col_idx == 2 else C_WHITE
        _report_cell(table.cell(9, col_idx), value, bg,
                     C_WHITE if col_idx == 2 else C_DTEXT, True, 9)
    for esc_idx, stat in enumerate(esc_stats):
        value = f"{stat['pct_tgt']}%" if stat['pct_tgt'] is not None else '—'
        pct = stat['pct_tgt'] or 0
        bg = C_REPORT_RED if pct >= 50 else (C_REPORT_YELLOW if pct >= 10 else C_REPORT_GREEN)
        _report_cell(table.cell(9, 6 + esc_idx), value, bg, C_DTEXT, True, 9)

    _report_cell(table.cell(10, 0), 'DR2 / MÉTIER', C_WHITE, C_DTEXT, True, 9)
    for col_idx in range(1, 5):
        _report_cell(table.cell(10, col_idx), '', C_WHITE)
    for esc_idx, stat in enumerate(esc_stats):
        _report_cell(table.cell(10, 6 + esc_idx), stat['metier'],
                     C_REPORT_PALE_GREEN, C_DTEXT, True, 9)

    _report_cell(table.cell(11, 0), '% MÉTIER / TT DR2', C_WHITE, C_DTEXT, True, 8)
    for col_idx in range(1, 5):
        _report_cell(table.cell(11, col_idx), '', C_WHITE)
    for esc_idx, stat in enumerate(esc_stats):
        _report_cell(table.cell(11, 6 + esc_idx), f"{stat['pct_tt']}%",
                     C_REPORT_PALE_GREEN, C_DTEXT, True, 9)
    return sl


def _slide_region_performance(prs, d):
    sl = _blank(prs)
    _header(sl, 'RAPPORT GESTION DES INCIDENTS', 'PERFORMANCE DR2 PAR RÉGION')
    _txt(
        sl, f"Période : {d['debut'].strftime('%d/%m/%Y')} au {d['fin'].strftime('%d/%m/%Y')}",
        MARGIN, Inches(1.08), Inches(6.8), Inches(0.28), size=10, color=C_DTEXT,
    )
    region_rows = sorted(d['region_rows'], key=lambda row: row['dr2'], reverse=True)
    rows = [
        [row['region'], row['tget'], row['dr2'], f"{row['pct_reg']} %", f"{row['pct_tget']} %"]
        for row in region_rows
    ]
    cell_fmts = {}
    for row_index, row in enumerate(region_rows):
        cell_fmts[(row_index, 4)] = (
            {'red': C_REPORT_RED, 'yellow': C_REPORT_YELLOW,
             'green': C_REPORT_PALE_GREEN}[row['color']],
            C_DTEXT,
        )
    _table(
        sl, ['RÉGION', 'PARC', 'DR2', 'PART DU TOTAL', 'TAUX DU PARC'], rows,
        left=Inches(0.50), top=Inches(1.48), width=Inches(7.20),
        height=Inches(4.95), col_widths=[2.0, 1.15, 0.9, 1.55, 1.6],
        font_size=10, hdr_size=9, cell_fmts=cell_fmts,
    )
    chart_rows = list(reversed([row for row in region_rows if row['dr2']]))
    if chart_rows:
        chart_height = min(Inches(4.95), Inches(max(1.50, 0.65 * len(chart_rows))))
        chart_top = Inches(1.48) + int((Inches(4.95) - chart_height) / 2)
        _bar_chart(
            sl, [row['region'] for row in chart_rows],
            [row['dr2'] for row in chart_rows],
            Inches(8.0), chart_top, Inches(4.75), chart_height,
            title='Volume DR2', ranked=True,
        )
    return sl


def _slide_business_matrix(prs, d):
    sl = _blank(prs)
    _header(sl, 'RAPPORT GESTION DES INCIDENTS', 'MATRICE RÉGION × MÉTIER')
    _use_office_font(_txt(
        sl, f"Période : {d['debut'].strftime('%d/%m/%Y')} au {d['fin'].strftime('%d/%m/%Y')}",
        MARGIN, Inches(1.08), Inches(6.8), Inches(0.28), size=10, color=C_DTEXT,
    ))
    active_stats = [stat for stat in _daily_report_stats(d) if stat['total']]
    active_stats = sorted(active_stats, key=lambda stat: stat['total'], reverse=True)[:8]
    business_names = [stat['escalade'] for stat in active_stats]
    region_rows = [row for row in d['region_rows'] if row['dr2']]
    matrix_rows = []
    for region in region_rows:
        counts = [
            sum(
                1 for row in d['detail_rows']
                if row['region'] == region['region'] and _category_matches(row, name)
            )
            for name in business_names
        ]
        matrix_rows.append([region['region'], region['dr2'], *counts])
    matrix_rows.append([
        'TOTAL', d['total_dr2'], *[stat['total'] for stat in active_stats],
    ])

    matrix_max = max(
        [value for row in matrix_rows[:-1] for value in row[2:]] or [0]
    )
    cell_fmts = {}
    for row_index, row in enumerate(matrix_rows):
        is_total = row_index == len(matrix_rows) - 1
        for col_index in range(len(row)):
            if is_total:
                cell_fmts[(row_index, col_index)] = (C_REPORT_NAVY, C_WHITE)
            elif col_index == 0:
                cell_fmts[(row_index, col_index)] = (C_WHITE, C_DTEXT)
            elif col_index == 1:
                cell_fmts[(row_index, col_index)] = (C_REPORT_CYAN, C_WHITE)
            else:
                value = row[col_index]
                background = _heat_color(value, matrix_max)
                foreground = C_WHITE if value / max(matrix_max, 1) >= 0.33 else C_DTEXT
                cell_fmts[(row_index, col_index)] = (background, foreground)

    table_width = min(12.45, 2.45 + 1.65 * max(len(business_names), 1))
    table_height = min(4.95, max(2.40, 0.65 * (len(matrix_rows) + 1)))
    business_width = (table_width - 2.45) / max(len(business_names), 1)
    _table(
        sl, ['RÉGION', 'TOTAL', *business_names], matrix_rows,
        left=int((SW - Inches(table_width)) / 2),
        top=Inches(1.55) + int((Inches(4.95) - Inches(table_height)) / 2),
        width=Inches(table_width), height=Inches(table_height),
        col_widths=[1.55, 0.9] + [business_width] * len(business_names),
        font_size=8, hdr_size=7, alt=False, cell_fmts=cell_fmts,
    )
    return sl


def _share_bar_chart(slide, categories, values, left, top, width, height,
                     title, color):
    if not values or not sum(values):
        return None
    ranked = sorted(zip(categories, values), key=lambda item: item[1])
    data = CategoryChartData()
    data.categories = [category for category, _ in ranked]
    data.add_series('DR2', [value for _, value in ranked])
    available_height = height - Inches(0.28)
    chart_height = min(
        available_height,
        Inches(max(1.15, 0.38 * len(ranked) + 0.30)),
    )
    chart_top = top + Inches(0.28) + int((available_height - chart_height) / 2)
    frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, left, chart_top, width, chart_height, data,
    )
    chart = frame.chart
    chart.has_title = False
    chart.has_legend = False
    _style_count_axes(chart, [value for _, value in ranked])
    plot = chart.plots[0]
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = color
    plot.series[0].format.line.fill.background()
    plot.gap_width = 160 if len(ranked) <= 3 else 70
    _show_count_labels(plot)
    chart.value_axis.has_major_gridlines = False
    chart.value_axis.format.line.fill.background()
    chart.value_axis.tick_labels.font.size = Pt(1)
    chart.value_axis.tick_labels.font.color.rgb = C_WHITE
    chart.value_axis.maximum_scale = (
        max(value for _, value in ranked) + chart.value_axis.major_unit
    )
    chart.category_axis.format.line.fill.background()
    chart.category_axis.tick_labels.font.size = Pt(8)
    _use_office_font(_txt(
        slide, title, left, top, width, Inches(0.28), size=9, bold=True,
        color=C_REPORT_NAVY, align=PP_ALIGN.CENTER,
    ))
    return chart


def _percent_bar_chart(slide, categories, values, left, top, width, height, title):
    data = CategoryChartData()
    data.categories = categories
    data.add_series('Taux', values)
    frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, data,
    )
    chart = frame.chart
    chart.has_legend = False
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    title_font = chart.chart_title.text_frame.paragraphs[0].font
    title_font.name = 'Arial'
    title_font.size = Pt(14)
    title_font.bold = True
    title_font.color.rgb = C_REPORT_BLUE
    _style_count_axes(chart, values)
    chart.value_axis.maximum_scale = max(max(values, default=0) + 1, 2)
    chart.value_axis.has_major_gridlines = False
    chart.value_axis.format.line.fill.background()
    chart.value_axis.tick_labels.font.size = Pt(1)
    chart.value_axis.tick_labels.font.color.rgb = C_WHITE
    chart.category_axis.format.line.fill.background()
    chart.category_axis.tick_labels.font.size = Pt(10)
    plot = chart.plots[0]
    plot.gap_width = 90
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = C_REPORT_BLUE
    plot.series[0].format.line.fill.background()
    _show_count_labels(plot)
    plot.data_labels.number_format = '0" %"'
    return chart


def _efficiency_dashboard(slide, region_rows):
    def dashboard_text(*args, **kwargs):
        text_box = _txt(*args, **kwargs)
        text_box.text_frame.margin_top = 0
        text_box.text_frame.margin_bottom = 0
        for paragraph in text_box.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Arial'
        return text_box

    rows = list(region_rows)
    left = Inches(0.55)
    top = Inches(1.42)
    width = Inches(12.20)
    row_height = Inches(0.72)
    gauge_left = Inches(4.05)
    gauge_width = Inches(7.25)
    gauge_height = Inches(0.16)

    dashboard_text(slide, 'RÉGION', left + Inches(0.18), top,
                   Inches(1.45), Inches(0.28), size=8, bold=True,
                   color=C_REPORT_BLUE)
    dashboard_text(slide, 'VOLUME / CIBLE', left + Inches(1.72), top,
                   Inches(1.70), Inches(0.28), size=8, bold=True,
                   color=C_REPORT_BLUE)
    dashboard_text(slide, 'CONSOMMATION DE LA CIBLE', gauge_left, top,
                   gauge_width, Inches(0.28), size=8, bold=True,
                   color=C_REPORT_BLUE)
    dashboard_text(slide, 'TAUX', Inches(11.58), top,
                   Inches(0.92), Inches(0.28), size=8, bold=True,
                   color=C_REPORT_BLUE, align=PP_ALIGN.RIGHT)

    rows_top = top + Inches(0.34)
    for index, row in enumerate(rows):
        row_top = rows_top + index * row_height
        background = C_WHITE if index % 2 == 0 else RGBColor(0xF5, 0xF7, 0xFA)
        _rect(slide, left, row_top, width, row_height - Inches(0.04), background)
        dashboard_text(
            slide, row['region'], left + Inches(0.18), row_top + Inches(0.19),
            Inches(1.40), Inches(0.26), size=11, bold=True,
            color=C_REPORT_BLUE, wrap=False,
        )
        dashboard_text(
            slide, f"{row['dr2']} DR2  /  {row['tget']}",
            left + Inches(1.72), row_top + Inches(0.19),
            Inches(1.70), Inches(0.26), size=10, color=C_DTEXT, wrap=False,
        )

        gauge_top = row_top + Inches(0.25)
        _rect(slide, gauge_left, gauge_top, gauge_width, gauge_height,
              RGBColor(0xE4, 0xE9, 0xF1))
        status_color = {
            'red': C_REPORT_RED,
            'yellow': C_REPORT_YELLOW,
            'green': C_DETAIL_GREEN,
        }[row['color']]
        filled_width = max(
            Inches(0.06),
            int(gauge_width * min(max(row['pct_tget'], 0), 100) / 100),
        )
        _rect(slide, gauge_left, gauge_top, filled_width, gauge_height, status_color)
        dashboard_text(
            slide, f"{row['pct_tget']}%", Inches(11.58), row_top + Inches(0.17),
            Inches(0.92), Inches(0.30), size=12, bold=True, color=status_color,
            align=PP_ALIGN.RIGHT, wrap=False,
        )

    legend_top = rows_top + len(rows) * row_height + Inches(0.10)
    legend = [
        (C_DETAIL_GREEN, '< 70%  Maîtrisé'),
        (C_REPORT_YELLOW, '70–99%  Vigilance'),
        (C_REPORT_RED, '≥ 100%  Cible dépassée'),
    ]
    for index, (color, label) in enumerate(legend):
        item_left = Inches(3.15 + index * 2.75)
        _rect(slide, item_left, legend_top + Inches(0.04), Inches(0.12), Inches(0.12), color)
        dashboard_text(
            slide, label, item_left + Inches(0.20), legend_top,
            Inches(2.35), Inches(0.24), size=8, color=C_DTEXT, wrap=False,
        )


def _slide_monthly_comparison(prs, d):
    sl = _blank(prs)
    _header(sl, 'RAPPORT GESTION DES INCIDENTS', 'DR2 COMPARATIVE MENSUEL')
    months = _monthly_comparison(d['fin'])
    if not months:
        return sl

    chart_months = [month for month in months if any(
        value is not None for value in month['values']
    )]
    if chart_months:
        single_month = len(chart_months) == 1
        day_count = len(chart_months[0]['values']) if single_month else 31
        data = CategoryChartData()
        data.categories = list(range(1, day_count + 1))
        for month in chart_months:
            data.add_series(
                month['label'],
                month['values'] + [None] * (day_count - len(month['values'])),
            )
        chart_type = (
            XL_CHART_TYPE.COLUMN_CLUSTERED
            if single_month else XL_CHART_TYPE.LINE_MARKERS
        )
        frame = sl.shapes.add_chart(
            chart_type, Inches(0.65), Inches(1.28),
            Inches(10.85), Inches(3.32), data,
        )
        chart = frame.chart
        chart.has_legend = not single_month
        chart.has_title = False
        _style_count_axes(
            chart,
            [value for month in chart_months for value in month['values']],
        )
        chart.category_axis.tick_labels.font.size = Pt(8)
        if single_month:
            series = chart.plots[0].series[0]
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = C_REPORT_BLUE
            series.format.line.fill.background()
            chart.plots[0].gap_width = 65
            _use_office_font(_txt(
                sl, f"ÉVOLUTION JOURNALIÈRE — {chart_months[0]['label']}",
                Inches(0.65), Inches(1.02), Inches(10.85), Inches(0.28),
                size=10, bold=True, color=C_REPORT_BLUE,
                align=PP_ALIGN.CENTER,
            ))
        else:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.name = 'Arial'
            chart.legend.font.size = Pt(9)
            line_colors = [
                RGBColor(0x5B, 0x9B, 0xD5),
                RGBColor(0xFF, 0xC0, 0x00),
                RGBColor(0xFF, 0x00, 0x00),
            ]
            color_by_label = dict(zip(
                [month['label'] for month in months], line_colors,
            ))
            for series, month in zip(chart.plots[0].series, chart_months):
                color = color_by_label[month['label']]
                series.format.line.color.rgb = color
                series.format.line.width = Pt(2.25)
    else:
        _txt(
            sl, 'Comparatif indisponible : aucun mois traité.',
            Inches(0.65), Inches(2.25), Inches(10.85), Inches(0.50),
            size=16, bold=True, color=C_DTEXT, align=PP_ALIGN.CENTER,
        )

    average_table = sl.shapes.add_table(4, 2, Inches(11.65), Inches(1.55),
                                        Inches(1.45), Inches(1.45)).table
    header = average_table.cell(0, 0)
    header.merge(average_table.cell(0, 1))
    _report_cell(header, 'MOYENNE', RGBColor(0xBD, 0xD7, 0xEE), C_DTEXT, True, 10)
    avg_colors = [C_REPORT_YELLOW, C_REPORT_RED, C_REPORT_GREEN]
    for row_idx, (month, color) in enumerate(zip(months, avg_colors), 1):
        _report_cell(average_table.cell(row_idx, 0), month['label'], C_WHITE,
                     C_DTEXT, True, 9, PP_ALIGN.LEFT)
        average = (
            str(month['moyenne']).replace('.', ',')
            if month['moyenne'] is not None else 'N/D'
        )
        _report_cell(average_table.cell(row_idx, 1),
                     average, color if month['moyenne'] is not None else C_DETAIL_GRAY,
                     C_DTEXT, True, 9)

    region_rows = [row for row in d['region_rows'] if row['dr2']]
    _share_bar_chart(
        sl, [row['region'] for row in region_rows], [row['dr2'] for row in region_rows],
        Inches(0.45), Inches(4.58), Inches(5.80), Inches(2.55),
        'RÉPARTITION PAR RÉGION', C_REPORT_CYAN,
    )
    active_stats = [stat for stat in _daily_report_stats(d) if stat['total']]
    _share_bar_chart(
        sl, [stat['escalade'] for stat in active_stats],
        [stat['total'] for stat in active_stats],
        Inches(6.15), Inches(4.58), Inches(5.80), Inches(2.55),
        'RÉPARTITION PAR MÉTIER', C_REPORT_BLUE,
    )
    return sl


def _slide_merci(prs):
    return _closing(prs)


# ═══════════════════════════════════════════════════════════════════════════
# SUPPORT 1 — GDI (point quotidien / mois en cours) — 10 diapositives
# ═══════════════════════════════════════════════════════════════════════════

def generate_gdi_daily(debut, fin, generated_on):
    prs = _reference_deck(_GDI_REFERENCE)

    periods = _gdi_reporting_periods(fin)
    month_start = periods['month_start']
    data_end = periods['data_end']
    detail_start = periods['detail_start']
    detail_end = periods['detail_end']
    meeting_day = periods['meeting_day']

    month_data = _dr2_dataset(month_start, data_end)
    detail_data = _dr2_dataset(detail_start, detail_end)
    qs_month = _dr2_qs(month_start, data_end)
    qs_detail = _dr2_qs(detail_start, detail_end)
    detail_label = f"{detail_start.strftime('%d/%m/%Y')} au {detail_end.strftime('%d/%m/%Y')}"

    # 1. Cover
    _cover(prs)

    # 2. Page de titre datée
    _meeting_title(prs, meeting_day)

    # 3. Synthèse de décision — mois en cours et dernier week-end traité
    _slide_executive_summary(prs, month_data, detail_data, qs_month)

    # 4. Définitions réglementaires
    _slide_definitions(prs)

    # 5-6. DR1
    _section(prs, 'TENDANCE DR1', data_end, 1)
    _slide_dr1_violations(prs, data_end)

    # 7-8. DR2 du dernier vendredi au dimanche achevé
    _section(prs, 'TENDANCE DR2', data_end, 2)
    _slide_dr2_trend(prs, detail_data, 'TDR2', detail_label)

    # 9-11. Une diapositive lisible par journée, y compris à zéro DR2
    day = detail_start
    while day <= detail_end:
        qs_day = qs_detail.filter(date=day)
        sl = _blank(prs)
        _header(sl, 'REUNION GESTION DES INCIDENTS', f'DETAIL DR2 {_JOURS_FR[day.weekday()]}')
        if day in detail_data['processed_dates'] or qs_day.exists():
            _detail_table(sl, _detail_rows(qs_day), day, qs_day.count(), height=Inches(4.9))
        else:
            _txt(
                sl, f'Journée du {day.strftime("%d/%m/%Y")} non traitée',
                MARGIN, CONTENT_TOP + Inches(1.7), SW - 2 * MARGIN,
                Inches(0.6), size=20, bold=True, color=C_RED_T,
                align=PP_ALIGN.CENTER,
            )
            _txt(
                sl, 'Importez les fichiers de disponibilité 2G et 3G avant de générer le support GDI.',
                MARGIN, CONTENT_TOP + Inches(2.35), SW - 2 * MARGIN,
                Inches(0.5), size=12, color=C_DTEXT, align=PP_ALIGN.CENTER,
            )
        day += timedelta(days=1)

    # 12. Synthèse du week-end
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'SYNTHESE DR2')
    _txt(sl, detail_label, MARGIN, CONTENT_TOP, SW - 2 * MARGIN,
         Inches(0.35), size=11, bold=True, color=C_RED_T,
         align=PP_ALIGN.CENTER)
    if detail_data['processed_dates']:
        _dr2_summary_tables(
            sl, qs_detail, detail_data['total_dr2'],
            top=CONTENT_TOP + Inches(0.55), height=Inches(3.0),
        )
    else:
        _txt(
            sl, 'Synthèse indisponible : aucune journée traitée sur ce week-end.',
            MARGIN, CONTENT_TOP + Inches(2.0), SW - 2 * MARGIN,
            Inches(0.6), size=18, bold=True, color=C_RED_T,
            align=PP_ALIGN.CENTER,
        )

    # 13. Top sites — mois jusqu'au dimanche de référence
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'CLASSEMENT DES SITES RÉCURRENTS DR2')
    top_sites = _top_sites(qs_month, 10)
    if top_sites:
        ranked_sites = [
            (f'{rank:02d}   {site}', count)
            for rank, (site, count) in enumerate(top_sites, 1)
        ]
        available_height = CONTENT_H - Inches(0.35)
        chart_height = min(
            available_height,
            Inches(max(2.2, 0.52 * len(ranked_sites) + 0.45)),
        )
        chart_top = CONTENT_TOP + Inches(0.15) + int(
            (available_height - chart_height) / 2
        )
        _bar_chart(
            sl,
            [site for site, _ in reversed(ranked_sites)],
            [count for _, count in reversed(ranked_sites)],
            Inches(0.8), chart_top,
            SW - Inches(1.35), chart_height,
            title='', ranked=True,
        )

    # 14. Répartition région / bases — mois
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'Répartition DR2 Région / Bases')
    regs = [(r['region'], r['dr2']) for r in month_data['region_rows'] if r['dr2']]
    if regs:
        _share_bar_chart(
            sl, [r for r, _ in regs], [c for _, c in regs],
            MARGIN, CONTENT_TOP, Inches(4.6), CONTENT_H - Inches(0.1),
            'DR2 PAR RÉGION', C_REPORT_CYAN,
        )
    bases = _base_breakdown(qs_month)[:15]
    if bases:
        _share_bar_chart(
            sl, [base for base, _ in bases], [count for _, count in bases],
            Inches(5.4), CONTENT_TOP, SW - Inches(5.4) - MARGIN,
            CONTENT_H - Inches(0.1), 'DR2 PAR BASE', C_REPORT_BLUE,
        )

    # 15. Efficacité — formule actuelle conservée jusqu'à validation métier
    sl = _blank(prs)
    _header(sl, 'REUNION GESTION DES INCIDENTS', 'EFFICACITE DR2 — ATTEINTE DES CIBLES')
    _efficiency_dashboard(sl, month_data['region_rows'])

    # 16. Points bloquants — mois
    sl = _blank(prs)
    pb = _points_bloquants(qs_month)
    with_blocking_point = sum(k for _, k in pb)
    without_blocking_point = max(month_data['total_dr2'] - with_blocking_point, 0)
    _header(
        sl, 'RAPPORT GESTION DES INCIDENTS', 'POINTS BLOQUANTS',
    )
    _summary_card(
        sl, 'AVEC POINT BLOQUANT', str(with_blocking_point),
        'DR2 concernés', Inches(3.65), C_RED_T,
    )
    _summary_card(
        sl, 'SANS POINT BLOQUANT', str(without_blocking_point),
        'DR2-PB', Inches(6.80), C_DETAIL_GREEN,
    )
    if pb:
        ranked_points = [[f'{rank:02d}', label, count]
                         for rank, (label, count) in enumerate(pb, 1)]
        available_height = Inches(3.45)
        table_height = min(
            available_height,
            Inches(max(1.70, 0.32 * (len(ranked_points) + 1))),
        )
        blocking_table = _table(
            sl, ['RANG', 'POINT BLOQUANT', 'NB'], ranked_points,
            left=Inches(0.75),
            top=Inches(3.25),
            width=Inches(11.85), height=table_height,
            col_widths=[0.8, 9.9, 1.15],
            font_size=max(6, min(10, 55 // (len(ranked_points) + 1))),
            hdr_size=9,
        )
        for row_index in range(1, len(blocking_table.rows)):
            point_cell = blocking_table.rows[row_index].cells[1]
            point_cell.margin_left = Inches(0.08)
            for paragraph in point_cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT
    else:
        _rect(sl, Inches(0.75), Inches(3.45), Inches(11.85), Inches(1.75),
              RGBColor(0xF5, 0xF7, 0xFA))
        _txt(sl, 'AUCUN POINT BLOQUANT RENSEIGNÉ',
             Inches(1.0), Inches(3.85), Inches(11.35), Inches(0.45),
             size=20, bold=True, color=C_DETAIL_GREEN, align=PP_ALIGN.CENTER)
        _txt(sl, f"{without_blocking_point} DR2 sans point bloquant déclaré",
             Inches(1.0), Inches(4.37), Inches(11.35), Inches(0.35),
             size=12, color=C_DTEXT, align=PP_ALIGN.CENTER)

    # 17-18. Lecture régionale puis matrice Région / Métiers — mois
    _slide_region_performance(prs, month_data)
    _slide_business_matrix(prs, month_data)

    # 19. Comparatif mensuel
    _slide_monthly_comparison(prs, month_data)

    # 20. Merci
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
    prs = _reference_deck(_REUNION_REFERENCE)

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
        separate_summary = is_last and len(rows) > 3
        table_height = Inches(4.8) if separate_summary or not is_last else Inches(2.75)
        _detail_table(sl, rows, day, nb, height=table_height)
        if is_last:
            if separate_summary:
                sl = _blank(prs)
                _header(sl, 'REUNION GESTION DES INCIDENTS', 'SYNTHESE DR2')
                _dr2_summary_tables(
                    sl, qs_period, d['total_dr2'],
                    top=CONTENT_TOP + Inches(0.55), height=Inches(3.0),
                )
            else:
                _dr2_summary_tables(sl, qs_period, d['total_dr2'])

    # N+1. Aperçu global et comparatif des tendances DR2
    _slide_apercu_global(prs, d)

    # N+2. Merci
    _slide_merci(prs)

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
