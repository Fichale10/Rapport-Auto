"""
Rapport FTTH — Réseau Fixe — page /reporting/fixe/rapport-ftth/

Parse le fichier brut d'incidents RESEAU_FIXE_YYYYMMDD_YYYYMMDD.xlsx et produit
les rapports « Gestion des incidents réseau d'accès fixe » :
  • Image 1 : « Types incidents par régions » (carte Togo + PON / OLT / CARTE).
  • Image 2 : « Inc / Métier / MTTR » (barres count + courbe MTTR par métier).
  • Image 3 : « Incidents / Régions Vs MTTR » (barres count + courbe MTTR par région).
  • Image 4 : « Statistiques / causes » (barres horizontales par cause).

Conventions :
  • Type d'incident déduit de la colonne « Alarm text » (PON LOSS / CARTE MISSING /
    RESSOURCE ISOLATION) puis repli sur la nature.
  • Métier = colonne « Escalade ». MTTR = colonne « Duration » (HH:MM:SS).
  • Un incident = une ligne (un ticket).
"""

from io import BytesIO
from datetime import datetime
import math
import os
import json
import re

import pandas as pd

from .gdi_core import _clean, _dur_to_sec, yas_logo_bytes, _load_font, _wrap_text

_MOIS_FR = [
    '', 'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
    'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE',
]

# Régions (LOME séparée de MARITIME pour les marqueurs de la carte)
REGIONS_FIX = ['SAVANES', 'KARA', 'CENTRALE', 'PLATEAUX', 'MARITIME', 'LOME']

_COLS = {
    'ticket':     ['numero du ticket', 'numéro du ticket', 'numero ticket', 'ticket'],
    'nature':     ["nature de l'incident", 'nature'],
    'alarm_time': ['alarm time', 'alarm_time', 'date alarme'],
    'alarm_text': ['alarm text', 'alarm_text', 'texte alarme'],
    'region':     ['région', 'region', 'rúgion'],
    'equip':      ['impact - equipement', 'impact equipement', 'impact-equipement'],
    'impact':     ['impact - service', 'impact service', 'impact-service'],
    'plateforme': ['plateforme', 'plateform'],
    'techno':     ['technologies', 'technologie'],
    'cause':      ['cause'],
    'escalade':   ['escalade'],
    'duration':   ['duration', 'durée', 'duree', 'durúe'],
    'status':     ['status', 'statut'],
    'site':       ['site name', 'site id', 'site parent'],
}


# ── Couleurs (rendu PNG) ────────────────────────────────────────────────────
BLUE = (0, 48, 135)
RED = (227, 0, 19)
NAVY = (13, 36, 97)
YELLOW = (255, 199, 44)
ORANGE = (237, 125, 49)
GRAY = (232, 236, 244)
LIGHTBLUE = (91, 155, 213)
WHITE = (255, 255, 255)
INK = (26, 35, 64)
GRID = (210, 217, 232)
PON_GREEN = (33, 140, 70)
CARTE_TEAL = (0, 150, 150)


def _resolve(df):
    lower = {str(c).strip().lower(): c for c in df.columns}
    out = {}
    for key, aliases in _COLS.items():
        for a in aliases:
            if a in lower:
                out[key] = lower[a]
                break
    return out


def _norm(s):
    return _clean(s).upper()


def fmt_hms(sec):
    sec = int(round(sec or 0))
    h, r = divmod(sec, 3600)
    m, s = divmod(r, 60)
    return f'{h:02d}:{m:02d}:{s:02d}'


def _avg(durs):
    vals = [d for d in durs if d > 0]
    return sum(vals) / len(vals) if vals else 0


def _region_key(raw):
    s = _norm(raw)
    if 'SAVANE' in s:
        return 'SAVANES'
    if 'KARA' in s:
        return 'KARA'
    if 'CENTRAL' in s:
        return 'CENTRALE'
    if 'PLATEAU' in s:
        return 'PLATEAUX'
    if 'LOME' in s or 'LOMÉ' in s:
        return 'LOME'
    if 'MARITIME' in s:
        return 'MARITIME'
    return None


def _inc_type(alarm_text, nature):
    a = _norm(alarm_text)
    if 'PON LOSS' in a:
        return 'PON'
    if 'CARTE' in a:
        return 'CARTE'
    if a:
        return 'OLT'
    n = _norm(nature)
    if 'CARTE' in n and 'PON' not in n:
        return 'CARTE'
    if 'PON' in n:
        return 'PON'
    return 'OLT'


def _metier(escalade):
    m = _norm(escalade)
    m = re.sub(r'\s+', ' ', m).strip()
    return m or '(vide)'


def _period(filename, df, cols):
    m = re.search(r'(\d{8})_(\d{8})', filename or '')
    if m:
        try:
            s = datetime.strptime(m.group(1), '%Y%m%d')
            e = datetime.strptime(m.group(2), '%Y%m%d')
            return s, e
        except ValueError:
            pass
    dates = []
    if cols.get('alarm_time'):
        ser = pd.to_datetime(df[cols['alarm_time']], errors='coerce', dayfirst=True).dropna()
        dates += list(ser)
    if dates:
        return min(dates).to_pydatetime(), max(dates).to_pydatetime()
    today = datetime.today()
    return today, today


def parse_reseau_fixe(file_path, filename=''):
    """Lit le fichier RESEAU_FIXE_*.xlsx et agrège les données des 4 images."""
    df = pd.read_excel(file_path)
    cols = _resolve(df)
    if 'nature' not in cols and 'region' not in cols:
        raise ValueError("Fichier RESEAU_FIXE invalide : colonnes introuvables.")

    rows = []
    for _, rec in df.iterrows():
        nature = _clean(rec.get(cols.get('nature')))
        region_raw = _clean(rec.get(cols.get('region')))
        if not nature and not region_raw:
            continue
        alarm_text = _clean(rec.get(cols.get('alarm_text')))
        dur = _clean(rec.get(cols.get('duration')))
        rows.append({
            'nature':       nature,
            'region':       _region_key(region_raw),
            'type':         _inc_type(alarm_text, nature),
            'metier':       _metier(rec.get(cols.get('escalade'))),
            'cause':        _norm(rec.get(cols.get('cause'))) or 'N/A',
            'duration_sec': _dur_to_sec(dur),
        })
    rows = [r for r in rows if r['nature']]
    if not rows:
        raise ValueError("Aucun incident détecté dans le fichier.")

    start, end = _period(filename, df, cols)
    period_label = f'Du {start.strftime("%d/%m/%Y")} au {end.strftime("%d/%m/%Y")}'
    month_label = f'{_MOIS_FR[start.month]} {start.year}'

    total = len(rows)

    # ── Image 1 : par type + par région ────────────────────────────────────
    type_counts = {'PON': 0, 'OLT': 0, 'CARTE': 0}
    region_counts = {k: {'PON': 0, 'OLT': 0, 'CARTE': 0, 'total': 0} for k in REGIONS_FIX}
    for r in rows:
        type_counts[r['type']] += 1
        k = r['region']
        if k in region_counts:
            region_counts[k][r['type']] += 1
            region_counts[k]['total'] += 1
    image1 = {
        'total': total,
        'pon': type_counts['PON'],
        'olt': type_counts['OLT'],
        'carte': type_counts['CARTE'],
        'regions': region_counts,
    }

    # ── Image 2 : par métier ───────────────────────────────────────────────
    metiers = {}
    for r in rows:
        m = r['metier']
        metiers.setdefault(m, []).append(r['duration_sec'])
    metier_rows = [
        {'label': m, 'count': len(durs), 'mttr_sec': _avg(durs)}
        for m, durs in metiers.items()
    ]
    metier_rows.sort(key=lambda x: x['count'], reverse=True)
    image2 = {
        'cats': metier_rows,
        'notes': _top_natures(rows, 10),
    }

    # ── Image 3 : par région ───────────────────────────────────────────────
    reg_groups = {}
    for r in rows:
        k = r['region'] or '(vide)'
        reg_groups.setdefault(k, []).append(r)
    order = [k for k in REGIONS_FIX if k in reg_groups] + \
            [k for k in reg_groups if k not in REGIONS_FIX]
    region_rows = [
        {'label': k,
         'count': len(reg_groups[k]),
         'mttr_sec': _avg([x['duration_sec'] for x in reg_groups[k]])}
        for k in order
    ]
    # Boîtes d'annotation : 2 régions au plus fort MTTR
    ann = sorted(region_rows, key=lambda x: x['mttr_sec'], reverse=True)[:2]
    notes3 = []
    for a in ann:
        natures = _top_natures(reg_groups[a['label']], 3)
        notes3.append({'title': a['label'], 'lines': natures})
    image3 = {'cats': region_rows, 'notes': notes3}

    # ── Image 4 : par cause ────────────────────────────────────────────────
    cause_counts = {}
    for r in rows:
        cause_counts[r['cause']] = cause_counts.get(r['cause'], 0) + 1
    cause_rows = sorted(
        ({'label': c, 'count': n} for c, n in cause_counts.items()),
        key=lambda x: x['count'], reverse=True)
    image4 = {'causes': cause_rows}

    return {
        'period_label': period_label,
        'month_label': month_label,
        'total': total,
        'image1': image1,
        'image2': image2,
        'image3': image3,
        'image4': image4,
    }


def _top_natures(rows, n):
    srt = sorted(rows, key=lambda r: r['duration_sec'], reverse=True)
    out = []
    for r in srt[:n]:
        txt = r['nature']
        if len(txt) > 70:
            txt = txt[:68] + '…'
        out.append(txt)
    return out


# ═══════════════════════════════════════════════════════════════════════════
# Helpers de rendu (Pillow)
# ═══════════════════════════════════════════════════════════════════════════

def _logo(img, img_w, margin, top=24, w=120):
    data = yas_logo_bytes()
    if not data:
        return
    try:
        from PIL import Image
        lg = Image.open(BytesIO(data)).convert('RGBA')
        h = int(lg.height * w / lg.width)
        lg = lg.resize((w, h))
        img.paste(lg, (img_w - margin - w, top), lg)
    except Exception:
        pass


def _header(d, img, img_w, margin, title, subtitle):
    f_title = _load_font(30, bold=True)
    f_sub = _load_font(20, bold=True)
    d.text((margin, 26), title, font=f_title, fill=BLUE)
    d.text((margin, 66), subtitle, font=f_sub, fill=RED)
    _logo(img, img_w, margin)


def _ctext(d, center, text, font, fill):
    tw = d.textlength(text, font=font)
    asc, desc = font.getmetrics()
    d.text((center[0] - tw / 2, center[1] - (asc + desc) / 2), text, font=font, fill=fill)


def _glyph_x(d, cx, cy, color, s=9, w=4):
    d.line([cx - s, cy - s, cx + s, cy + s], fill=color, width=w)
    d.line([cx - s, cy + s, cx + s, cy - s], fill=color, width=w)


def _glyph_plus(d, cx, cy, color, s=9, w=4):
    d.line([cx - s, cy, cx + s, cy], fill=color, width=w)
    d.line([cx, cy - s, cx, cy + s], fill=color, width=w)


def _nice(v, ticks=4):
    if v <= 0:
        return ticks, 1
    raw = v / ticks
    mag = 10 ** math.floor(math.log10(raw))
    step = mag * 10
    for mult in (1, 2, 2.5, 5, 10):
        if mag * mult >= raw:
            step = mag * mult
            break
    return step * ticks, step


# ── Carte du Togo : polygones + centroïdes par région ───────────────────────
_FIX_GEO_CACHE = None


def _fix_region_polys():
    global _FIX_GEO_CACHE
    if _FIX_GEO_CACHE is not None:
        return _FIX_GEO_CACHE
    regions = {}
    try:
        path = os.path.join(os.path.dirname(__file__), 'static', 'reports',
                            'togo_geo.js')
        with open(path, 'r', encoding='utf-8') as fh:
            txt = fh.read()
        txt = txt[txt.index('{'):].rstrip().rstrip(';')
        data = json.loads(txt)
        for feat in data.get('features', []):
            name = (feat.get('properties', {}) or {}).get('shapeName', '')
            key = _region_key(name)
            if not key:
                continue
            geom = feat.get('geometry', {}) or {}
            gtype = geom.get('type')
            coords = geom.get('coordinates', [])
            rings = regions.setdefault(key, [])
            if gtype == 'Polygon':
                rings.append(coords[0])
            elif gtype == 'MultiPolygon':
                for poly in coords:
                    rings.append(poly[0])
    except Exception:
        pass
    _FIX_GEO_CACHE = regions
    return regions


def _draw_region_map(d, x0, y0, w, h, region_counts):
    """Dessine la carte du Togo colorée par région. Retourne les centroïdes px."""
    regions = _fix_region_polys()
    if not regions:
        return {}
    allpts = [p for rings in regions.values() for ring in rings for p in ring]
    xs = [p[0] for p in allpts]
    ys = [p[1] for p in allpts]
    minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
    gw = (maxx - minx) or 1
    gh = (maxy - miny) or 1
    scale = min(w / gw, h / gh)
    ox = x0 + (w - gw * scale) / 2
    oy = y0 + (h - gh * scale) / 2

    def proj(pt):
        return (ox + (pt[0] - minx) * scale,
                oy + (maxy - pt[1]) * scale)

    # Échelle de couleur (vert clair → vert foncé) par total région
    geo_totals = {}
    for k in regions:
        t = region_counts.get(k, {}).get('total', 0)
        if k == 'MARITIME':                       # LOME compté dans le polygone Maritime
            t += region_counts.get('LOME', {}).get('total', 0)
        geo_totals[k] = t
    maxc = max(list(geo_totals.values()) + [1])

    def reg_color(c):
        if not c:
            return (224, 235, 224)
        t = max(0.25, c / maxc)
        a = (200, 224, 196)
        b = (27, 94, 32)
        return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))

    f_reg = _load_font(15, bold=True)
    centroids = {}
    for k, rings in regions.items():
        col = reg_color(geo_totals.get(k, 0))
        cxs, cys = [], []
        for ring in rings:
            pts = [proj(p) for p in ring]
            d.polygon(pts, fill=col, outline=WHITE)
            cxs += [p[0] for p in pts]
            cys += [p[1] for p in pts]
        cx = sum(cxs) / len(cxs)
        cy = sum(cys) / len(cys)
        centroids[k] = (cx, cy)
        label = 'Lomé / Maritime' if k == 'MARITIME' else k.capitalize()
        lw = d.textlength(label, font=f_reg)
        d.text((cx - lw / 2, cy - h * 0.05), label, font=f_reg, fill=WHITE,
               stroke_width=2, stroke_fill=NAVY)
    # LOME : sur (ou sous) le centroïde Maritime selon les données
    if 'MARITIME' in centroids:
        mc = centroids['MARITIME']
        mar_total = region_counts.get('MARITIME', {}).get('total', 0)
        centroids['LOME'] = (mc[0], mc[1] + h * 0.10) if mar_total else mc
    return centroids


def _region_marker(d, cx, cy, olt, pon, carte, f_num, with_carte=False):
    s = 9
    _glyph_x(d, cx, cy, RED, s)                       # marqueur OLT (✖)
    px = cx + 32
    _glyph_plus(d, px, cy, PON_GREEN, s)             # marqueur PON (✛)
    r = 18
    # OLT (compteur, cercle rouge à gauche)
    lx = cx - 76
    d.line([lx + r, cy, cx - s, cy], fill=RED, width=2)
    d.ellipse([lx - r, cy - r, lx + r, cy + r], fill=WHITE, outline=RED, width=3)
    _ctext(d, (lx, cy), str(olt), f_num, RED)
    # PON (compteur, cercle rouge à droite)
    rx = cx + 92
    d.line([px + s, cy, rx - r, cy], fill=RED, width=2)
    d.ellipse([rx - r, cy - r, rx + r, cy + r], fill=WHITE, outline=RED, width=3)
    _ctext(d, (rx, cy), str(pon), f_num, RED)
    if with_carte:
        cyy = cy + 36
        _glyph_plus(d, cx, cyy, CARTE_TEAL, s)        # marqueur CARTE (➕)
        d.ellipse([cx + 18, cyy - r, cx + 18 + 2 * r, cyy + r], fill=WHITE,
                  outline=YELLOW, width=3)
        _ctext(d, (cx + 18 + r, cyy), str(carte), f_num, (180, 140, 0))


# ═══════════════════════════════════════════════════════════════════════════
# IMAGE 1 — Types incidents par régions
# ═══════════════════════════════════════════════════════════════════════════

def build_png_image1(report, generated_on=''):
    from PIL import Image, ImageDraw
    im1 = report['image1']
    img_w, img_h = 1600, 900
    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)
    margin = 40
    _header(d, img, img_w, margin,
            'RAPPORT GESTION DES INCIDENTS', 'Types incidents par régions')

    f_box_lbl = _load_font(20, bold=True)
    f_box_val = _load_font(26, bold=True)
    f_total = _load_font(54, bold=True)
    f_tab = _load_font(18, bold=True)
    f_num = _load_font(17, bold=True)
    f_lbl = _load_font(14, bold=True)
    f_foot = _load_font(12, bold=False)

    # ── 3 boîtes (gauche) ──────────────────────────────────────────────────
    boxes = [
        ('PON Loss', im1['pon'], NAVY),
        ('OLT Down', im1['olt'], NAVY),
        ('CARTE Down', im1['carte'], NAVY),
    ]
    bx, bw, bh = margin, 230, 92
    gap = 56
    ys = [180, 180 + bh + gap, 180 + 2 * (bh + gap)]
    mids = [by + bh / 2 for by in ys]
    trunk_y = (mids[0] + mids[-1]) / 2
    circle = (560, int(trunk_y), 92)
    junction_x = bx + bw + 70
    for (lbl, val, col), by in zip(boxes, ys):
        d.rounded_rectangle([bx, by, bx + bw, by + bh], radius=10,
                            outline=col, width=3, fill=WHITE)
        _ctext(d, (bx + bw / 2, by + 28), lbl, f_box_lbl, col)
        _ctext(d, (bx + bw / 2, by + 62), str(val), f_box_val, col)
        # stub horizontal depuis chaque boîte
        d.line([bx + bw, by + bh / 2, junction_x, by + bh / 2], fill=NAVY, width=2)
    # accolade : épine verticale + tronc vers le cercle
    d.line([junction_x, mids[0], junction_x, mids[-1]], fill=NAVY, width=2)
    d.line([junction_x, trunk_y, circle[0] - circle[2], trunk_y], fill=NAVY, width=2)

    # ── Cercle central (total) ─────────────────────────────────────────────
    ccx, ccy, ccr = circle
    d.ellipse([ccx - ccr, ccy - ccr, ccx + ccr, ccy + ccr], fill=NAVY)
    _ctext(d, (ccx, ccy), str(im1['total']), f_total, YELLOW)

    # ── Onglets en-tête au-dessus de la carte ──────────────────────────────
    map_x0 = 740
    for label, tx in (('OLT Down', map_x0 + 40), ('PON Loss', map_x0 + 470)):
        tw = d.textlength(label, font=f_tab) + 28
        d.rounded_rectangle([tx, 120, tx + tw, 152], radius=6, fill=YELLOW)
        _ctext(d, (tx + tw / 2, 136), label, f_tab, NAVY)

    # ── Carte ──────────────────────────────────────────────────────────────
    centroids = _draw_region_map(d, map_x0, 165, 820, 615, im1['regions'])
    rc = im1['regions']
    for k in ['SAVANES', 'KARA', 'CENTRALE', 'PLATEAUX', 'MARITIME', 'LOME']:
        if k not in centroids:
            continue
        cnt = rc.get(k, {'OLT': 0, 'PON': 0, 'CARTE': 0, 'total': 0})
        if cnt.get('total', 0) == 0:          # pas de marqueur si aucun incident
            continue
        cx, cy = centroids[k]
        _region_marker(d, cx, cy, cnt['OLT'], cnt['PON'], cnt['CARTE'],
                       f_num, with_carte=(cnt['CARTE'] > 0))

    # ── Légende (centrée sous la carte) ────────────────────────────────────
    ly = img_h - 56
    lx = map_x0 + 170
    items = [('x', RED, 'OLT'), ('+', CARTE_TEAL, 'CARTE'), ('+', PON_GREEN, 'PON')]
    for gtype, col, name in items:
        if gtype == 'x':
            _glyph_x(d, lx + 11, ly + 11, col, 9)
        else:
            _glyph_plus(d, lx + 11, ly + 11, col, 9)
        d.rectangle([lx + 30, ly, lx + 98, ly + 22], fill=NAVY)
        _ctext(d, (lx + 64, ly + 11), name, f_lbl, WHITE)
        lx += 160

    if generated_on:
        d.text((margin, img_h - 26),
               f"Généré le {generated_on} — Yas Togo / DT / DCO  —  {report['period_label']}",
               font=f_foot, fill=NAVY)

    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════
# Graphe combiné barres + courbe (Images 2 & 3)
# ═══════════════════════════════════════════════════════════════════════════

def _draw_combo(d, x0, y0, w, h, cats, notes_boxes=None):
    """cats : liste {label, count, mttr_sec}. Dessine barres count + courbe MTTR."""
    f_axis = _load_font(13, bold=False)
    f_val = _load_font(13, bold=True)
    f_mttr = _load_font(12, bold=True)
    f_lbl = _load_font(13, bold=True)

    n = len(cats)
    if n == 0:
        return
    counts = [c['count'] for c in cats]
    mttrs_h = [c['mttr_sec'] / 3600 for c in cats]
    cmax, cstep = _nice(max(counts + [1]))
    mmax, mstep = _nice(max(mttrs_h + [1]))

    # Grille + axes
    ticks = 4
    for i in range(ticks + 1):
        gy = y0 + h - (h * i / ticks)
        d.line([x0, gy, x0 + w, gy], fill=GRID, width=1)
        d.text((x0 - 34, gy - 8), str(int(cmax * i / ticks)), font=f_axis, fill=INK)
        mh = mmax * i / ticks
        d.text((x0 + w + 8, gy - 8), fmt_hms(mh * 3600)[:5], font=f_axis, fill=ORANGE)
    d.line([x0, y0, x0, y0 + h], fill=INK, width=2)
    d.line([x0 + w, y0, x0 + w, y0 + h], fill=ORANGE, width=2)
    d.line([x0, y0 + h, x0 + w, y0 + h], fill=INK, width=2)

    slot = w / n
    bw = min(slot * 0.5, 70)
    pts = []
    for i, c in enumerate(cats):
        cx = x0 + slot * i + slot / 2
        bh = (c['count'] / cmax) * h if cmax else 0
        d.rectangle([cx - bw / 2, y0 + h - bh, cx + bw / 2, y0 + h], fill=NAVY)
        _ctext(d, (cx, y0 + h - bh - 12), str(c['count']), f_val, NAVY)
        # label X (tronqué)
        lbl = c['label']
        if d.textlength(lbl, font=f_axis) > slot - 4:
            while lbl and d.textlength(lbl + '…', font=f_axis) > slot - 4:
                lbl = lbl[:-1]
            lbl += '…'
        _ctext(d, (cx, y0 + h + 16), lbl, f_axis, INK)
        py = y0 + h - (mttrs_h[i] / mmax) * h if mmax else y0 + h
        pts.append((cx, py))
    # Courbe MTTR
    if len(pts) >= 2:
        d.line(pts, fill=ORANGE, width=3)
    for i, p in enumerate(pts):
        d.ellipse([p[0] - 5, p[1] - 5, p[0] + 5, p[1] + 5], fill=ORANGE)
        d.text((p[0] + 8, p[1] - 16), fmt_hms(cats[i]['mttr_sec']),
               font=f_mttr, fill=ORANGE)

    # Boîtes d'annotation
    if notes_boxes:
        f_note_t = _load_font(13, bold=True)
        f_note = _load_font(11, bold=True)
        boxes = notes_boxes if isinstance(notes_boxes, list) else [notes_boxes]
        bx = x0 + 40
        for box in boxes:
            lines = box.get('lines', [])
            title = box.get('title', '')
            ww = 360
            wrapped = []
            for ln in lines:
                wrapped += _wrap_text('• ' + ln, f_note, ww - 24)
            bh2 = 16 + (20 if title else 0) + len(wrapped) * 15 + 12
            by = y0 + 6
            d.rectangle([bx, by, bx + ww, by + bh2], fill=YELLOW, outline=ORANGE, width=2)
            ty = by + 8
            if title:
                d.text((bx + 12, ty), title, font=f_note_t, fill=RED)
                ty += 20
            for ln in wrapped:
                d.text((bx + 12, ty), ln, font=f_note, fill=NAVY)
                ty += 15
            bx += ww + 30


def build_png_image2(report, generated_on=''):
    from PIL import Image, ImageDraw
    im2 = report['image2']
    img_w, img_h = 1600, 900
    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)
    margin = 40
    _header(d, img, img_w, margin,
            'RAPPORT GESTION DES INCIDENTS', 'Inc / Métier / MTTR')
    f_ct = _load_font(18, bold=True)
    _ctext(d, (img_w / 2, 130), 'Count Inc By METIER  Vs  MTTR by METIER', f_ct, INK)

    notes = {'title': 'Incidents (MTTR le plus élevé)', 'lines': im2['notes']}
    _draw_combo(d, 110, 200, img_w - 260, 560, im2['cats'], notes_boxes=[notes])

    f_foot = _load_font(12, bold=False)
    if generated_on:
        d.text((margin, img_h - 26),
               f"Généré le {generated_on} — Yas Togo / DT / DCO  —  {report['period_label']}",
               font=f_foot, fill=NAVY)
    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


def build_png_image3(report, generated_on=''):
    from PIL import Image, ImageDraw
    im3 = report['image3']
    img_w, img_h = 1600, 900
    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)
    margin = 40
    _header(d, img, img_w, margin,
            'RAPPORT GESTION DES INCIDENTS', 'Incidents / Régions Vs MTTR')
    f_ct = _load_font(18, bold=True)
    _ctext(d, (img_w / 2, 130), 'Count Inc By Région  Vs  MTTR by REGION', f_ct, INK)

    _draw_combo(d, 110, 200, img_w - 260, 560, im3['cats'], notes_boxes=im3['notes'])

    f_foot = _load_font(12, bold=False)
    if generated_on:
        d.text((margin, img_h - 26),
               f"Généré le {generated_on} — Yas Togo / DT / DCO  —  {report['period_label']}",
               font=f_foot, fill=NAVY)
    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════
# IMAGE 4 — Statistiques / causes (barres horizontales)
# ═══════════════════════════════════════════════════════════════════════════

def build_png_image4(report, generated_on=''):
    from PIL import Image, ImageDraw
    causes = report['image4']['causes']
    img_w, img_h = 1600, 900
    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)
    margin = 40
    _header(d, img, img_w, margin,
            'RAPPORT GESTION DES INCIDENTS', 'Statistiques / causes')
    f_ct = _load_font(18, bold=True)
    _ctext(d, (img_w / 2, 132), 'INCIDENTS OUVERTS PAR CAUSES RESEAUX FIXE', f_ct, NAVY)

    f_lbl = _load_font(13, bold=True)
    f_val = _load_font(13, bold=True)
    f_axis = _load_font(12, bold=False)

    data = causes[:18]
    n = len(data)
    if n == 0:
        buf = BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)
        return buf

    x0, y0 = 470, 190
    plot_w = img_w - x0 - 80
    plot_h = 660
    row_h = plot_h / n
    bar_h = min(row_h * 0.62, 30)
    maxc = max(c['count'] for c in data) or 1
    vmax, vstep = _nice(maxc)

    # Axe X (grille verticale)
    ticks = int(vmax / vstep) if vstep else 5
    for i in range(ticks + 1):
        gx = x0 + (plot_w * i / ticks)
        d.line([gx, y0, gx, y0 + plot_h], fill=GRID, width=1)
        _ctext(d, (gx, y0 + plot_h + 14), str(int(vmax * i / ticks)), f_axis, INK)

    def bar_color(i):
        if i == 0:
            return NAVY
        if i == 1:
            return RED
        if i <= 5:
            return YELLOW
        return LIGHTBLUE

    for i, c in enumerate(data):
        cy = y0 + row_h * i + row_h / 2
        bl = (c['count'] / vmax) * plot_w
        col = bar_color(i)
        d.rectangle([x0, cy - bar_h / 2, x0 + bl, cy + bar_h / 2], fill=col)
        # libellé cause (gauche)
        lbl = c['label']
        if d.textlength(lbl, font=f_lbl) > x0 - margin - 12:
            while lbl and d.textlength(lbl + '…', font=f_lbl) > x0 - margin - 12:
                lbl = lbl[:-1]
            lbl += '…'
        asc, desc = f_lbl.getmetrics()
        d.text((x0 - 10 - d.textlength(lbl, font=f_lbl), cy - (asc + desc) / 2),
               lbl, font=f_lbl, fill=INK)
        # valeur (à droite de la barre)
        d.text((x0 + bl + 8, cy - (asc + desc) / 2), str(c['count']),
               font=f_val, fill=col if col != YELLOW else (180, 140, 0))

    d.line([x0, y0, x0, y0 + plot_h], fill=INK, width=2)

    f_foot = _load_font(12, bold=False)
    if generated_on:
        d.text((margin, img_h - 26),
               f"Généré le {generated_on} — Yas Togo / DT / DCO  —  {report['period_label']}",
               font=f_foot, fill=NAVY)
    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf
