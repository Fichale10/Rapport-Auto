"""
Rapport GDI Core — « Disponibilité et trafic IGW »

Parsing du fichier brut de ticketing (CORE_ET_IGW_*.xlsx) et rendu PNG du
tableau « Incidents core » (Nature de l'incident / Impact - Service / Cause /
Escalade / Duration).
"""

from io import BytesIO
from datetime import datetime

import pandas as pd

_MOIS_FR = [
    '', 'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
    'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE',
]# Correspondance souple des en-têtes attendues
_COL_ALIASES = {
    'nature':   ["nature de l'incident", 'nature'],
    'impact':   ['impact - service', 'impact service', 'impact-service'],
    'cause':    ['cause'],
    'escalade': ['escalade'],
    'duration': ['duration', 'durée', 'duree'],
    'alarm':    ['alarm time', 'date alarme', 'alarm_time'],
}


def _clean(val):
    if val is None:
        return ''
    if isinstance(val, float) and pd.isna(val):
        return ''
    s = str(val).strip()
    if s.lower() in ('nan', 'nat', 'none'):
        return ''
    # Normalise les retours à la ligne multiples
    return ' '.join(part.strip() for part in s.splitlines() if part.strip())


def _dur_to_sec(val):
    """Convertit 'HH:MM:SS' (HH pouvant dépasser 24) en secondes."""
    s = _clean(val)
    if not s:
        return 0
    parts = s.split(':')
    try:
        if len(parts) == 3:
            h, m, sec = (int(float(p)) for p in parts)
            return h * 3600 + m * 60 + sec
        if len(parts) == 2:
            m, sec = (int(float(p)) for p in parts)
            return m * 60 + sec
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def _resolve_columns(df):
    """Mappe les clés logiques vers les noms réels de colonnes (insensible casse)."""
    lower = {str(c).strip().lower(): c for c in df.columns}
    resolved = {}
    for key, aliases in _COL_ALIASES.items():
        for alias in aliases:
            if alias in lower:
                resolved[key] = lower[alias]
                break
    return resolved


def parse_gdi_core(file_path, filename=''):
    """
    Lit le fichier Excel et retourne :
    { 'rows': [ {nature, impact, cause, escalade, duration, duration_sec} ],
      'period_label': str, 'total': int, 'fermes': int, 'ouverts': int }
    """
    df = pd.read_excel(file_path)
    cols = _resolve_columns(df)
    if 'nature' not in cols:
        raise ValueError(
            "Colonne « Nature de l'incident » introuvable dans le fichier."
        )

    rows = []
    for _, rec in df.iterrows():
        nature = _clean(rec.get(cols.get('nature')))
        if not nature:
            continue
        dur = _clean(rec.get(cols.get('duration'))) if 'duration' in cols else ''
        rows.append({
            'nature':       nature,
            'impact':       _clean(rec.get(cols.get('impact'))) if 'impact' in cols else '',
            'cause':        _clean(rec.get(cols.get('cause'))) if 'cause' in cols else '',
            'escalade':     _clean(rec.get(cols.get('escalade'))) if 'escalade' in cols else '',
            'duration':     dur,
            'duration_sec': _dur_to_sec(dur),
        })

    # Tri par durée décroissante (incidents les plus longs en tête)
    rows.sort(key=lambda r: r['duration_sec'], reverse=True)

    period_label = _period_label(df, cols, filename)

    return {
        'rows':         rows,
        'period_label': period_label,
        'total':        len(rows),
    }


def _period_label(df, cols, filename):
    # 1) depuis le nom de fichier : ..._YYYYMMDD_YYYYMMDD.xlsx
    import re
    m = re.search(r'(\d{4})(\d{2})\d{2}_\d{8}', filename or '')
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return f'{_MOIS_FR[mo]} {y}'
    # 2) depuis la colonne Alarm Time
    if 'alarm' in cols:
        try:
            ser = pd.to_datetime(df[cols['alarm']], errors='coerce').dropna()
            if not ser.empty:
                d = ser.min()
                return f'{_MOIS_FR[d.month]} {d.year}'
        except Exception:
            pass
    return datetime.today().strftime('%B %Y').upper()


# ═══════════════════════════════════════════════════════════════════════════════
# LOGO YAS
# ═══════════════════════════════════════════════════════════════════════════════

_LOGO_CACHE = {}


def yas_logo_bytes():
    """Décode le logo Yas (JPEG base64) depuis reports/static/reports/yas_logo.js.
    Retourne des bytes ou None si introuvable."""
    if 'data' in _LOGO_CACHE:
        return _LOGO_CACHE['data']
    import os
    import base64
    import re
    path = os.path.join(os.path.dirname(__file__),
                        'static', 'reports', 'yas_logo.js')
    data = None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        m = re.search(r'base64,([A-Za-z0-9+/=]+)', content)
        if m:
            data = base64.b64decode(m.group(1))
    except Exception:
        data = None
    _LOGO_CACHE['data'] = data
    return data


# ═══════════════════════════════════════════════════════════════════════════════
# EXPORT PNG (Pillow)
# ═══════════════════════════════════════════════════════════════════════════════

def build_png(report, generated_on=''):
    """Rend le tableau « Incidents core » en image PNG. Retourne un BytesIO."""
    from PIL import Image, ImageDraw, ImageFont

    rows = report.get('rows', [])
    period_label = report.get('period_label', '')

    NAVY  = (13, 36, 97)
    BLUE  = (0, 48, 135)
    RED   = (227, 0, 19)
    YELLOW = (255, 199, 44)
    GRAY  = (232, 236, 244)
    WHITE = (255, 255, 255)
    DARK  = (26, 35, 64)

    headers = ["Nature de l'incident", 'Impact - Service', 'Cause', 'Escalade', 'Duration']
    col_w = [430, 320, 360, 230, 150]
    table_w = sum(col_w)
    margin = 40
    img_w = table_w + 2 * margin

    f_title = _load_font(28, bold=True)
    f_sub   = _load_font(20, bold=True)
    f_tag   = _load_font(17, bold=True)
    f_hdr   = _load_font(15, bold=True)
    f_cell  = _load_font(13, bold=False)
    f_cellb = _load_font(13, bold=True)
    f_foot  = _load_font(12, bold=False)

    pad = 10
    line_h = 18

    # Pré-calcul des hauteurs de lignes (wrap par colonne)
    def wrapped(text, font, width):
        return _wrap_text(text, font, width - 2 * pad)

    header_lines = [wrapped(h, f_hdr, col_w[i]) for i, h in enumerate(headers)]
    header_h = max(len(l) for l in header_lines) * line_h + 2 * pad

    row_cells = []
    row_heights = []
    for r in rows:
        vals = [r['nature'], r['impact'], r['cause'], r['escalade'], r['duration']]
        cell_lines = [wrapped(v, f_cellb if i == 0 else f_cell, col_w[i])
                      for i, v in enumerate(vals)]
        h = max(len(cl) for cl in cell_lines) * line_h + 2 * pad
        row_cells.append(cell_lines)
        row_heights.append(h)

    top_block = 150  # titres + tag
    table_h = header_h + sum(row_heights) if rows else header_h + 60
    img_h = top_block + table_h + 60

    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)

    # Titres
    d.text((margin, 24), 'COMITE GESTION DES INCIDENTS', font=f_title, fill=BLUE)
    d.text((margin, 62), 'Disponibilité et trafic IGW', font=f_sub, fill=RED)
    if period_label:
        tw = d.textlength(period_label, font=f_sub)
        d.text((img_w - margin - tw, 28), period_label, font=f_sub, fill=BLUE)

    # Logo Yas (haut-droite)
    logo = yas_logo_bytes()
    if logo:
        try:
            lg = Image.open(BytesIO(logo)).convert('RGBA')
            lw = 120
            lh = int(lg.height * lw / lg.width)
            lg = lg.resize((lw, lh))
            img.paste(lg, (img_w - margin - lw, 64), lg)
        except Exception:
            pass

    # Étiquette « Incidents core »
    tag_y = 100
    d.rectangle([margin, tag_y, margin + 170, tag_y + 32], fill=YELLOW)
    d.text((margin + 14, tag_y + 7), 'Incidents core', font=f_tag, fill=BLUE)

    # Tableau
    x0 = margin
    y = top_block

    # En-tête
    x = x0
    for i, lines in enumerate(header_lines):
        d.rectangle([x, y, x + col_w[i], y + header_h], fill=YELLOW,
                    outline=WHITE, width=2)
        ty = y + pad
        for ln in lines:
            d.text((x + pad, ty), ln, font=f_hdr, fill=BLUE)
            ty += line_h
        x += col_w[i]
    y += header_h

    if not rows:
        d.text((x0 + pad, y + 20), 'Aucun incident dans le fichier importé.',
               font=f_cell, fill=DARK)
    else:
        for ridx, cell_lines in enumerate(row_cells):
            rh = row_heights[ridx]
            x = x0
            for i, lines in enumerate(cell_lines):
                if i == 0:
                    bg, fg, font = NAVY, WHITE, f_cellb
                else:
                    bg, fg, font = GRAY, BLUE, f_cell
                d.rectangle([x, y, x + col_w[i], y + rh], fill=bg,
                            outline=WHITE, width=2)
                ty = y + pad
                for ln in lines:
                    d.text((x + pad, ty), ln, font=font, fill=fg)
                    ty += line_h
                x += col_w[i]
            y += rh

    if generated_on:
        d.text((margin, img_h - 30),
               f'Généré le {generated_on} — Yas Togo / DT / DOC / iSOC',
               font=f_foot, fill=NAVY)

    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


def _load_font(size, bold=False):
    from PIL import ImageFont
    candidates = (
        ['arialbd.ttf', 'segoeuib.ttf', 'DejaVuSans-Bold.ttf'] if bold
        else ['arial.ttf', 'segoeui.ttf', 'DejaVuSans.ttf']
    )
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _wrap_text(text, font, max_width):
    """Découpe `text` en lignes tenant dans max_width (pixels)."""
    text = text or ''
    if not text:
        return ['']
    from PIL import ImageDraw, Image
    d = ImageDraw.Draw(Image.new('RGB', (10, 10)))
    words = text.split()
    lines, cur = [], ''
    for w in words:
        trial = f'{cur} {w}'.strip()
        if d.textlength(trial, font=font) <= max_width or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    # Limite à 6 lignes pour éviter des cellules géantes
    if len(lines) > 6:
        lines = lines[:6]
        lines[-1] += ' …'
    return lines
