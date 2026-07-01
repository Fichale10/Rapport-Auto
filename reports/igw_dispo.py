"""
Rapport « Disponibilité Lien IGW » (Image 2)

Lit le fichier « RAPPORT DE TAUX D'INDISPONIBILITE DES LIENS INTERNATIONAUX »
(colonne LIENS INTERNATIONAUX + durées d'incident par lien) et calcule, pour
chaque lien :

    Taux de disponibilité = (intervalle_min − indispo_min) / intervalle_min × 100

L'intervalle (période du rapport) est détecté depuis le titre « DU … AU … ».
Le rapport final combine ce tableau avec le TOP 3 des incidents critiques issus
de la page Core (rapport GDI « Incidents core »).
"""

from io import BytesIO
from datetime import datetime, timedelta, time as _time, date as _date
import re

import openpyxl

_MOIS_FR = {
    'JANVIER': 1, 'FEVRIER': 2, 'FÉVRIER': 2, 'MARS': 3, 'AVRIL': 4,
    'MAI': 5, 'JUIN': 6, 'JUILLET': 7, 'AOUT': 8, 'AOÛT': 8,
    'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DECEMBRE': 12,
    'DÉCEMBRE': 12,
}
_MOIS_FR_NOM = ['', 'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
                'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE']


def _to_seconds(val):
    """Convertit une durée (timedelta / time / 'HH:MM[:SS]' / nombre) en secondes."""
    if val is None:
        return 0
    if isinstance(val, timedelta):
        return int(val.total_seconds())
    if isinstance(val, _time):
        return val.hour * 3600 + val.minute * 60 + val.second
    if isinstance(val, datetime):
        return val.hour * 3600 + val.minute * 60 + val.second
    if isinstance(val, (int, float)):
        # Excel peut stocker une durée en fraction de jour
        if 0 < val < 10:
            return int(round(val * 86400))
        return int(val)
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'nat', 'none'):
        return 0
    parts = s.split(':')
    try:
        if len(parts) == 3:
            h, m, sec = (int(float(p)) for p in parts)
            return h * 3600 + m * 60 + sec
        if len(parts) == 2:
            h, m = (int(float(p)) for p in parts)
            return h * 3600 + m * 60
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def _clean(val):
    if val is None:
        return ''
    s = str(val).strip()
    if s.lower() in ('nan', 'nat', 'none'):
        return ''
    return ' '.join(p.strip() for p in s.splitlines() if p.strip())


def _short_name(name):
    """Raccourcit le nom de lien pour l'affichage (TRANSIT-TERACO-LOME-10GE-1 → TERACO-1)."""
    n = _clean(name)
    n = re.sub(r'^TRANSIT[\s\-]+', '', n, flags=re.I)
    n = n.replace('-CACA-10GE', '').replace('-LOME-10GE', '')
    n = n.replace('-10GE', '').replace(' 10GE', '')
    return n.strip()


def _parse_interval(title, fallback_dates=None):
    """Détecte la période depuis « DU 25 AVRIL AU 01 MAI 2026 ».
    Retourne (interval_min, period_label, month_label)."""
    label = ''
    days = None
    end_month = None
    end_year = None
    if title:
        m = re.search(
            r'DU\s+(\d{1,2})\s*([A-ZÉÛ]+)?\s+AU\s+(\d{1,2})\s+([A-ZÉÛ]+)\s+(\d{4})',
            title.upper())
        if m:
            d1 = int(m.group(1))
            mo1 = m.group(2)
            d2 = int(m.group(3))
            mo2 = m.group(4)
            year = int(m.group(5))
            month2 = _MOIS_FR.get(mo2)
            month1 = _MOIS_FR.get(mo1) if mo1 else month2
            if month1 and month2:
                try:
                    start = _date(year, month1, d1)
                    end = _date(year, month2, d2)
                    if end < start:  # plage à cheval sur l'année précédente
                        start = _date(year - 1, month1, d1)
                    days = (end - start).days + 1
                    end_month, end_year = month2, year
                    label = m.group(0).replace('DU', 'Du').strip()
                except ValueError:
                    pass

    if days is None and fallback_dates:
        valid = [d for d in fallback_dates if d]
        if valid:
            dmin, dmax = min(valid), max(valid)
            days = (dmax.date() - dmin.date()).days + 1
            end_month, end_year = dmax.month, dmax.year
            label = f'Du {dmin.strftime("%d/%m/%Y")} au {dmax.strftime("%d/%m/%Y")}'

    if days is None or days <= 0:
        days = 30  # défaut : un mois
    interval_min = days * 24 * 60
    month_label = ''
    if end_month:
        month_label = _MOIS_FR_NOM[end_month]
        if end_year:
            month_label = f'{month_label} {end_year}'
    if not label:
        label = month_label or f'{days} jour(s)'
    return interval_min, label, month_label


def parse_igw_dispo(file_path, filename=''):
    """
    Lit le fichier de taux d'indisponibilité et retourne :
    { 'links':[{name, short, n_inc, downtime_sec, downtime_min, availability}],
      'period_label', 'month_label', 'interval_min',
      'global_availability', 'total_inc' }
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)

    # Choisit la feuille contenant un en-tête « LIENS INTERNATIONAUX » avec le
    # plus de durées d'incident renseignées (la plus complète).
    def _sheet_rows(ws):
        return list(ws.iter_rows(values_only=True))

    def _is_liens_header(c):
        cu = _clean(c).upper()
        # En-tête exact (et non le titre qui contient aussi cette expression)
        return cu == 'LIENS INTERNATIONAUX' or (
            'LIENS INTERNATIONAUX' in cu and len(cu) <= 25)

    def _find_header(rows):
        for i, r in enumerate(rows):
            if any(_is_liens_header(c) for c in r):
                return i
        return None

    best = None  # (score, rows, header_idx)
    for ws in wb.worksheets:
        rows = _sheet_rows(ws)
        hidx = _find_header(rows)
        if hidx is None:
            continue
        score = sum(1 for r in rows[hidx + 1:] for c in r if _to_seconds(c) > 0)
        if best is None or score > best[0]:
            best = (score, rows, hidx)

    if best is None:
        raise ValueError("Colonne « LIENS INTERNATIONAUX » introuvable dans le fichier.")
    rows = best[1]
    header_idx = best[2]

    # Mappe les colonnes depuis la ligne d'en-tête
    hdr = [_clean(c).upper() for c in rows[header_idx]]

    def _find_col(pred, default=None):
        for j, c in enumerate(hdr):
            if pred(c):
                return j
        return default

    liens_col = _find_col(lambda c: c == 'LIENS INTERNATIONAUX' or
                          ('LIENS INTERNATIONAUX' in c and len(c) <= 25))
    nature_col = _find_col(lambda c: c.startswith('NATURE'))
    cause_col = _find_col(lambda c: c.startswith('CAUSE'))
    debut_col = _find_col(lambda c: c.startswith('DEBUT'))
    fin_col = _find_col(lambda c: c.startswith('FIN'))
    duree_col = _find_col(lambda c: c.startswith('DUR'))  # 1re « Durée incident »

    # Titre (période) : lignes situées au-dessus de l'en-tête
    title = ''
    for i in range(header_idx):
        joined = ' '.join(_clean(c) for c in rows[i])
        mt = re.search(r'DU\s+\d.*\d{4}', joined.upper())
        if mt:
            title = mt.group(0)
            break
        if 'RAPPORT' in joined.upper() and not title:
            title = joined

    if liens_col is None:
        raise ValueError("Colonne « LIENS INTERNATIONAUX » introuvable dans le fichier.")
    if duree_col is None:
        duree_col = liens_col + 5  # position attendue par défaut

    # 2) Parcours des lignes de données
    links = []
    current = None
    fallback_dates = []
    for r in rows[header_idx + 1:]:
        liens = _clean(r[liens_col]) if liens_col < len(r) else ''
        dur_val = r[duree_col] if duree_col < len(r) else None
        dur_sec = _to_seconds(dur_val)
        if debut_col is not None and debut_col < len(r) and isinstance(r[debut_col], datetime):
            fallback_dates.append(r[debut_col])
        if fin_col is not None and fin_col < len(r) and isinstance(r[fin_col], datetime):
            fallback_dates.append(r[fin_col])

        if liens:
            up = liens.upper()
            if 'GLOBAL' in up:
                current = None
                continue
            current = {
                'name': liens,
                'short': _short_name(liens),
                'n_inc': 0,
                'downtime_sec': 0,
            }
            links.append(current)
            if dur_sec > 0:
                current['n_inc'] += 1
                current['downtime_sec'] += dur_sec
        else:
            # ligne de continuation (incident supplémentaire du lien courant)
            if current is not None and dur_sec > 0:
                current['n_inc'] += 1
                current['downtime_sec'] += dur_sec

    interval_min, period_label, month_label = _parse_interval(title, fallback_dates)

    total_inc = 0
    for lk in links:
        lk['downtime_min'] = lk['downtime_sec'] / 60.0
        avail = (interval_min - lk['downtime_min']) / interval_min * 100 if interval_min else 100.0
        lk['availability'] = max(0.0, min(100.0, avail))
        total_inc += lk['n_inc']

    if links:
        global_availability = sum(lk['availability'] for lk in links) / len(links)
    else:
        global_availability = 100.0

    return {
        'links': links,
        'period_label': period_label,
        'month_label': month_label or 'PÉRIODE',
        'interval_min': interval_min,
        'global_availability': global_availability,
        'total_inc': total_inc,
    }


def fmt_pct(v):
    return f'{v:.2f}'.replace('.', ',') + ' %'


# ═══════════════════════════════════════════════════════════════════════════
# CALCUL DEPUIS LE FICHIER BRUT DE TICKETING (CORE_ET_IGW_*.xlsx)
# ═══════════════════════════════════════════════════════════════════════════

# (clé, mots-clés détectés dans « Nature de l'incident », indices valides,
#  gabarit nom court, gabarit nom complet). Ordre = ordre d'affichage des liens.
_FAMILIES = [
    ('TERACO',  ['TERACO'],                              [1, 2, 3], 'TERACO-{i}',               'TRANSIT-TERACO-LOME-10GE-{i}'),
    ('PARIS',   ['TELMA PARIS', 'TELMA-PARIS', 'PARIS(WACS'], [1, 2], 'PARIS(WACS-LONDON)-{i}', 'TRANSIT-PARIS(WACS-LONDON)-10GE-{i}'),
    ('BICS',    ['BICS'],                                [1, 2],    'BICS-{i}',                 'TRANSIT-BICS-CACA-10GE-{i}'),
    ('MTN',     ['MTN'],                                 [1, 2, 3], 'MTN-{i}',                  'TRANSIT MTN-CACA-10GE-{i}'),
    ('EQUIANO', ['EQUIANO'],                             [1, 2, 3], 'EQUIANO-{i}',              'TRANSIT EQUIANO-CACA-10GE-{i}'),
    ('COGENT',  ['COGENT'],                              [1, 2],    'COGENT-LOME-{i}',          'COGENT-LOME-{i}'),
    ('GOOGLE',  ['GOOGLE-PNI', 'GOOGLE PNI', 'GOOGLE_PNI'], [1, 2], 'GOOGLE-PNI-0{i}',         'GOOGLE-PNI-0{i}'),
    ('TATA',    ['TATA'],                                [1, 2],    'TATA-{i}',                 'TATA-{i}'),
]


def _match_links(nature):
    """Retourne l'ensemble des (famille, index) de liens IGW concernés par un ticket,
    déduits du texte « Nature de l'incident »."""
    up = (nature or '').upper()
    hits = []
    for fam, kws, idxs, _s, _f in _FAMILIES:
        for kw in kws:
            for m in re.finditer(re.escape(kw), up):
                hits.append((m.start(), fam, tuple(idxs)))
    hits.sort(key=lambda h: h[0])
    res = set()
    for j, (pos, fam, idxs) in enumerate(hits):
        end = hits[j + 1][0] if j + 1 < len(hits) else len(up)
        clause = up[pos:end]
        nums = [int(x) for x in re.findall(r'\d+', clause)]
        valid = [n for n in nums if n in idxs]
        if not valid:
            valid = list(idxs)
        for n in valid:
            res.add((fam, n))
    return res


def _core_columns(df):
    lower = {str(c).strip().lower(): c for c in df.columns}

    def col(*names):
        for n in names:
            if n in lower:
                return lower[n]
        return None

    return {
        'nature': col("nature de l'incident", 'nature'),
        'alarm':  col('alarm time', 'date alarme', 'alarm_time'),
        'cancel': col('cancel time', 'date cloture', 'cancel', 'cancel_time'),
        'duration': col('duration', 'durée', 'duree'),
    }


def _core_period(filename, df, cols):
    """Détermine [début, fin) de la période depuis le nom de fichier
    (_YYYYMMDD_YYYYMMDD) ou, à défaut, depuis les dates d'alarme/clôture."""
    import pandas as pd
    m = re.search(r'(\d{8})_(\d{8})', filename or '')
    if m:
        try:
            s = datetime.strptime(m.group(1), '%Y%m%d')
            e = datetime.strptime(m.group(2), '%Y%m%d') + timedelta(days=1)
            return s, e
        except ValueError:
            pass
    dates = []
    for key in ('alarm', 'cancel'):
        if cols.get(key):
            ser = pd.to_datetime(df[cols[key]], errors='coerce', dayfirst=True).dropna()
            dates += [d.to_pydatetime() for d in ser]
    if dates:
        s = min(dates).replace(hour=0, minute=0, second=0, microsecond=0)
        e = max(dates).replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
        return s, e
    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    return today.replace(day=1), today + timedelta(days=1)


def _clipped_seconds(rec, cols, start, end):
    """Durée d'indisponibilité du ticket bornée à la période [start, end]."""
    import pandas as pd
    from .gdi_core import _dur_to_sec
    a = pd.to_datetime(rec.get(cols['alarm']), errors='coerce', dayfirst=True) if cols.get('alarm') else None
    c = pd.to_datetime(rec.get(cols['cancel']), errors='coerce', dayfirst=True) if cols.get('cancel') else None
    if a is not None and not pd.isna(a):
        a = a.to_pydatetime()
        cc = c.to_pydatetime() if (c is not None and not pd.isna(c)) else end
        lo = max(a, start)
        hi = min(cc, end)
        return max(0, (hi - lo).total_seconds())
    # repli : colonne Duration (non bornée)
    if cols.get('duration'):
        return _dur_to_sec(_clean(rec.get(cols['duration'])))
    return 0


def _period_labels(start, end):
    last = end - timedelta(days=1)
    period_label = f'Du {start.strftime("%d/%m/%Y")} au {last.strftime("%d/%m/%Y")}'
    if start.month == last.month and start.year == last.year:
        month_label = f'{_MOIS_FR_NOM[start.month]} {start.year}'
    else:
        month_label = f'{_MOIS_FR_NOM[start.month]}-{_MOIS_FR_NOM[last.month]} {last.year}'
    return period_label, month_label


def parse_core_to_dispo(file_path, filename=''):
    """
    Calcule le tableau de disponibilité des liens IGW directement depuis le
    fichier brut de ticketing (CORE_ET_IGW_*.xlsx) :
      • mappe chaque ticket aux liens via « Nature de l'incident »,
      • somme les durées d'indispo (bornées à la période) et compte les incidents,
      • taux = (intervalle − indispo) / intervalle × 100.
    Retourne la même structure que parse_igw_dispo + 'top_incidents'.
    """
    import pandas as pd
    df = pd.read_excel(file_path)
    cols = _core_columns(df)
    if not cols.get('nature'):
        raise ValueError("Colonne « Nature de l'incident » introuvable dans le fichier.")

    start, end = _core_period(filename, df, cols)
    interval_min = (end - start).total_seconds() / 60.0
    period_label, month_label = _period_labels(start, end)

    # Initialise les 19 liens (ordre fixe), même ceux sans incident
    links_map = {}
    order = []
    for fam, kws, idxs, sfmt, ffmt in _FAMILIES:
        for i in idxs:
            key = (fam, i)
            links_map[key] = {
                'name': ffmt.format(i=i),
                'short': sfmt.format(i=i),
                'n_inc': 0,
                'downtime_sec': 0,
            }
            order.append(key)

    for _, rec in df.iterrows():
        nature = _clean(rec.get(cols['nature']))
        if not nature:
            continue
        assigns = _match_links(nature)
        if not assigns:
            continue
        down = _clipped_seconds(rec, cols, start, end)
        for key in assigns:
            if key in links_map:
                links_map[key]['n_inc'] += 1
                links_map[key]['downtime_sec'] += down

    links = []
    total_inc = 0
    for key in order:
        lk = links_map[key]
        lk['downtime_min'] = lk['downtime_sec'] / 60.0
        avail = (interval_min - lk['downtime_min']) / interval_min * 100 if interval_min else 100.0
        lk['availability'] = max(0.0, min(100.0, avail))
        total_inc += lk['n_inc']
        links.append(lk)

    global_availability = (sum(lk['availability'] for lk in links) / len(links)) if links else 100.0

    # TOP 3 incidents critiques (depuis le même fichier)
    top_incidents = _core_top_incidents(file_path, filename)

    return {
        'links': links,
        'period_label': period_label,
        'month_label': month_label,
        'interval_min': interval_min,
        'global_availability': global_availability,
        'total_inc': total_inc,
        'top_incidents': top_incidents,
    }


def _core_top_incidents(file_path, filename='', n=3):
    """TOP n des incidents (durée décroissante) depuis le fichier core."""
    try:
        from .gdi_core import parse_gdi_core
        rep = parse_gdi_core(file_path, filename=filename)
        return [
            {'nature': r.get('nature', ''), 'impact': r.get('impact', ''),
             'cause': r.get('cause', ''), 'escalade': r.get('escalade', ''),
             'duration': r.get('duration', '')}
            for r in rep.get('rows', [])[:n]
        ]
    except Exception:
        return []


# ═══════════════════════════════════════════════════════════════════════════
# EXPORT PNG (Pillow) — rapport combiné Image 2
# ═══════════════════════════════════════════════════════════════════════════

def build_png(report, top_incidents=None, generated_on=''):
    """Rend le rapport « Disponibilité et trafic IGW » (tableau dispo + TOP 3)."""
    from PIL import Image, ImageDraw
    from .gdi_core import yas_logo_bytes, _load_font, _wrap_text

    links = report.get('links', [])
    month_label = report.get('month_label', '')
    period_label = report.get('period_label', '')
    top_incidents = top_incidents or []

    NAVY = (13, 36, 97)
    BLUE = (0, 48, 135)
    RED = (227, 0, 19)
    YELLOW = (255, 199, 44)
    GRAY = (232, 236, 244)
    GREEN = (34, 197, 94)
    ORANGE = (245, 192, 0)
    WHITE = (255, 255, 255)

    margin = 40
    gap = 30
    # Colonne gauche (dispo) / droite (top 3)
    left_cols = [300, 110, 130]          # Lien | Nombre Inc | %
    left_w = sum(left_cols)
    right_cols = [250, 175, 200, 150, 110]   # Nature | Impact | Cause | Escalade | Duration
    right_w = sum(right_cols)
    img_w = margin * 2 + left_w + gap + right_w

    f_title = _load_font(30, bold=True)
    f_sub = _load_font(20, bold=True)
    f_tag = _load_font(17, bold=True)
    f_hdr = _load_font(14, bold=True)
    f_cell = _load_font(13, bold=False)
    f_cellb = _load_font(13, bold=True)
    f_foot = _load_font(12, bold=False)

    pad = 9
    line_h = 18
    top_block = 165

    def wrap(text, font, width):
        return _wrap_text(text, font, width - 2 * pad)

    # Hauteurs lignes gauche
    left_hdr = ['Lien', 'Nombre Inc', month_label or 'Taux']
    left_hdr_lines = [wrap(h, f_hdr, left_cols[i]) for i, h in enumerate(left_hdr)]
    left_hdr_h = max(len(l) for l in left_hdr_lines) * line_h + 2 * pad
    left_row_h = line_h + 2 * pad

    # Hauteurs lignes droite (TOP 3)
    right_hdr = ["Nature de l'incident", 'Impact - Service', 'Cause', 'Escalade', 'Duration']
    right_hdr_lines = [wrap(h, f_hdr, right_cols[i]) for i, h in enumerate(right_hdr)]
    right_hdr_h = max(len(l) for l in right_hdr_lines) * line_h + 2 * pad
    right_cells = []
    right_heights = []
    for inc in top_incidents[:3]:
        vals = [inc.get('nature', ''), inc.get('impact', ''), inc.get('cause', ''),
                inc.get('escalade', ''), inc.get('duration', '')]
        cl = [wrap(v, f_cellb if i == 0 else f_cell, right_cols[i])
              for i, v in enumerate(vals)]
        h = max(len(c) for c in cl) * line_h + 2 * pad
        right_cells.append(cl)
        right_heights.append(max(h, left_row_h * 2))

    left_table_h = left_hdr_h + left_row_h * max(len(links), 1)
    right_block_h = right_hdr_h + (sum(right_heights) if right_heights else 80)
    body_h = max(left_table_h, right_block_h + 50)
    img_h = top_block + body_h + 60

    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)

    # Titres
    d.text((margin, 26), 'COMITE GESTION DES INCIDENTS', font=f_title, fill=BLUE)
    d.text((margin, 66), 'Disponibilité et trafic IGW', font=f_sub, fill=RED)

    # Logo Yas
    logo = yas_logo_bytes()
    if logo:
        try:
            lg = Image.open(BytesIO(logo)).convert('RGBA')
            lw = 130
            lh = int(lg.height * lw / lg.width)
            lg = lg.resize((lw, lh))
            img.paste(lg, (img_w - margin - lw, 30), lg)
        except Exception:
            pass

    # ── Tableau gauche : Disponibilité Lien IGW ──
    lx = margin
    ly = top_block
    tag_label = f'Disponibilité Lien IGW {month_label}'.strip()
    d.rectangle([lx, ly - 42, lx + left_w, ly - 8], fill=YELLOW)
    d.text((lx + 12, ly - 36), tag_label, font=f_tag, fill=BLUE)

    x = lx
    for i, lines in enumerate(left_hdr_lines):
        d.rectangle([x, ly, x + left_cols[i], ly + left_hdr_h], fill=NAVY,
                    outline=WHITE, width=2)
        ty = ly + pad
        for ln in lines:
            tw = d.textlength(ln, font=f_hdr)
            d.text((x + (left_cols[i] - tw) / 2, ty), ln, font=f_hdr, fill=WHITE)
            ty += line_h
        x += left_cols[i]
    y = ly + left_hdr_h

    if not links:
        d.text((lx + pad, y + 16), 'Aucun lien dans le fichier importé.',
               font=f_cell, fill=NAVY)
    for lk in links:
        x = lx
        avail = lk['availability']
        cells = [
            (lk['short'], GRAY, BLUE, f_cellb, 'left'),
            (str(lk['n_inc']), GRAY, BLUE, f_cellb, 'center'),
            (fmt_pct(avail), GREEN if avail >= 99.999 else ORANGE,
             WHITE if avail >= 99.999 else NAVY, f_cellb, 'center'),
        ]
        for i, (val, bg, fg, font, al) in enumerate(cells):
            d.rectangle([x, y, x + left_cols[i], y + left_row_h], fill=bg,
                        outline=WHITE, width=2)
            tw = d.textlength(val, font=font)
            if al == 'center':
                tx = x + (left_cols[i] - tw) / 2
            else:
                tx = x + pad
            d.text((tx, y + pad), val, font=font, fill=fg)
            x += left_cols[i]
        y += left_row_h

    # ── Bloc droit : TOP 3 Incidents critiques du mois ──
    rx = margin + left_w + gap
    ry = top_block
    d.rectangle([rx, ry - 42, rx + right_w, ry - 8], fill=NAVY)
    ttl = 'TOP 3 Incidents critiques du mois'
    tw = d.textlength(ttl, font=f_tag)
    d.text((rx + (right_w - tw) / 2, ry - 36), ttl, font=f_tag, fill=WHITE)

    x = rx
    for i, lines in enumerate(right_hdr_lines):
        d.rectangle([x, ry, x + right_cols[i], ry + right_hdr_h], fill=YELLOW,
                    outline=WHITE, width=2)
        ty = ry + pad
        for ln in lines:
            d.text((x + pad, ty), ln, font=f_hdr, fill=BLUE)
            ty += line_h
        x += right_cols[i]
    y = ry + right_hdr_h

    if not right_cells:
        d.text((rx + pad, y + 16),
               'Importez le fichier de tickets sur la page Core (TOP 3).',
               font=f_cell, fill=NAVY)
    for ridx, cl in enumerate(right_cells):
        rh = right_heights[ridx]
        x = rx
        for i, lines in enumerate(cl):
            bg, fg, font = (NAVY, WHITE, f_cellb) if i == 0 else (GRAY, BLUE, f_cell)
            d.rectangle([x, y, x + right_cols[i], y + rh], fill=bg,
                        outline=WHITE, width=2)
            ty = y + pad
            for ln in lines:
                d.text((x + pad, ty), ln, font=font, fill=fg)
                ty += line_h
            x += right_cols[i]
        y += rh

    if generated_on:
        d.text((margin, img_h - 28),
               f'Généré le {generated_on} — Yas Togo / DT / DOC / iSOC  —  {period_label}',
               font=f_foot, fill=NAVY)

    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf
