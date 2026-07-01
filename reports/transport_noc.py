"""
Rapport NOC TRANSMISSION — page /reporting/transmission/rapport-noc/

Parse le fichier brut d'incidents TRANSMISSION_YYYYMMDD_YYYYMMDD.xlsx et produit
les rapports du Comité Gestion des Incidents :
  • Image 1 : « Détails Incident transport » (Backhaul / Backbone, avec/sans impact).
  • Image 2 : « Count Inc & MTTR par Métier et par Régions ».
  • Image 3 : « Disponibilité clients IPT & IPLC » (à venir — liste client fixe).

Règles (validées) :
  • MTTR = moyenne des durées par incident.
  • Totaux (global / Backhaul / Backbone) comptés par ticket unique.
  • Répartition par région : un incident est compté dans chaque région qu'il touche.
"""

from io import BytesIO
from datetime import datetime, timedelta
import re

import pandas as pd

from .gdi_core import _clean, _dur_to_sec, yas_logo_bytes, _load_font, _wrap_text

_MOIS_FR_NOM = [
    '', 'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
    'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE',
]

REGIONS = ['LOME', 'MARITIME', 'PLATEAUX', 'CENTRALE', 'KARA', 'SAVANES']
METIERS = ['ENERGIE', 'PROJET', 'TRANS FO', 'TRANS IP']

# Liste fixe des clients IPT & IPLC (Image 3) : (libellé affiché, [clés de match]).
# Une clé est recherchée (contient) dans la colonne « Nature de l'incident »
# des incidents partenaires (LIAISON PARTENAIRE).
CLIENTS_IPT = [
    ('ECOBANK ETI',              ['ECOBANK ETI']),
    ('ECOBANK ASSIGAME',         ['ECOBANK ASSIGAME']),
    ('EGOOV-1',                  ['EGOOV-1', 'EGOOV 1']),
    ('EGOOV-2',                  ['EGOOV-2', 'EGOOV 2']),
    ('SIP YAS - MOOV',           ['SIP YAS', 'SIP MOOV']),
    ('NSIA BENIN',               ['NSIA BENIN', 'NSIA BANK']),
    ('SBIN IPT CSquared (40G)',  ['SBIN IPT CSQUARED', 'SBIN IPT C SQUARED']),
    ('SBIN IPT WACS (20G)',      ['SBIN IPT WACS']),
    ('WACREN-IPT (1G)',          ['WACREN-IPT', 'WACREN IPT']),
    ('WACREN-IPLC (1G)',         ['WACREN-IPLC', 'WACREN IPLC']),
    ('SITA NVLE AEROGARE (10G)', ['SITA']),
]

# ── Résolution des colonnes (insensible à la casse / accents manquants) ──
_COLS = {
    'ticket':     ['numero du ticket', 'numéro du ticket', 'numero ticket', 'ticket'],
    'nature':     ["nature de l'incident", 'nature'],
    'alarm':      ['alarm time', 'alarm_time', 'date alarme'],
    'cancel':     ['cancel time', 'cancel_time'],
    'site':       ['site name', 'site id', 'site parent'],
    'region':     ['région', 'region', 'rúgion'],
    'impact':     ['impact - service', 'impact service', 'impact-service'],
    'categorie':  ['technologies', 'technologie'],          # BACKHAUL / BACKBONE / LIAISON PARTENAIRE
    'equip':      ['plateforme', 'plateform'],              # LIEN DWDM / ROUTEUR IP/MPLS ...
    'cause':      ['cause'],
    'escalade':   ['escalade'],
    'duration':   ['duration', 'durée', 'duree', 'durúe'],
    'status':     ['status', 'statut'],
}


def _resolve(df):
    lower = {str(c).strip().lower(): c for c in df.columns}
    out = {}
    for key, aliases in _COLS.items():
        for a in aliases:
            if a in lower:
                out[key] = lower[a]
                break
    return out


def fmt_hms(sec):
    """Secondes → 'HH:MM:SS' (HH peut dépasser 24)."""
    sec = int(round(sec or 0))
    h, r = divmod(sec, 3600)
    m, s = divmod(r, 60)
    return f'{h:02d}:{m:02d}:{s:02d}'


def _norm(s):
    return _clean(s).upper()


def _has_impact(impact):
    """Vrai si le champ « Impact - Service » dénote un service réellement impacté."""
    v = _norm(impact)
    if not v:
        return False
    if 'AUCUN' in v:      # AUCUN IMPACT
        return False
    return True


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
    if cols.get('alarm'):
        ser = pd.to_datetime(df[cols['alarm']], errors='coerce', dayfirst=True).dropna()
        dates += list(ser)
    if dates:
        return min(dates).to_pydatetime(), max(dates).to_pydatetime()
    today = datetime.today()
    return today, today


def _period_labels(start, end):
    period_label = f'Du {start.strftime("%d/%m/%Y")} au {end.strftime("%d/%m/%Y")}'
    if start.month == end.month and start.year == end.year:
        month_label = f'{_MOIS_FR_NOM[start.month]} {start.year}'
    else:
        month_label = f'{_MOIS_FR_NOM[start.month]}-{_MOIS_FR_NOM[end.month]} {end.year}'
    return period_label, month_label


def _avg(durations):
    vals = [d for d in durations if d and d > 0]
    return sum(vals) / len(vals) if vals else 0


def parse_transport_noc(file_path, filename=''):
    df = pd.read_excel(file_path)
    cols = _resolve(df)
    if not cols.get('ticket') and not cols.get('nature'):
        raise ValueError("Colonnes « Numero du ticket » / « Nature » introuvables.")

    start, end = _period(file_path and filename or filename, df, cols)
    period_label, month_label = _period_labels(start, end)

    def g(rec, key):
        c = cols.get(key)
        return _clean(rec.get(c)) if c else ''

    # ── Regroupement par ticket ──
    tickets = {}          # tid -> dict
    for idx, rec in df.iterrows():
        tid = g(rec, 'ticket') or f'__row{idx}'
        plate = _norm(g(rec, 'categorie'))       # BACKHAUL / BACKBONE / ...
        region = _norm(g(rec, 'region'))
        metier = _norm(g(rec, 'escalade'))
        dur = _dur_to_sec(g(rec, 'duration'))
        impact = g(rec, 'impact')
        techno = _norm(g(rec, 'equip'))           # LIEN DWDM / ROUTEUR IP/MPLS ...
        nature = g(rec, 'nature')
        site = g(rec, 'site')
        cause = g(rec, 'cause')

        t = tickets.get(tid)
        if t is None:
            t = {
                'tid': tid, 'plateformes': set(), 'regions': set(),
                'metiers': set(), 'duration': 0, 'avec_impact': False,
                'nature': nature, 'site': site, 'cause': cause,
                'impact': '', 'techno': set(), 'region_metier': set(),
                'match_text': '', 'nature_text': '', 'is_partner': False,
            }
            tickets[tid] = t
        t['match_text'] += ' ' + _norm(nature) + ' ' + _norm(site)
        t['nature_text'] += ' ' + _norm(nature)
        if 'PARTENAIRE' in plate or 'PARTENAIRE' in techno:
            t['is_partner'] = True
        if plate:
            t['plateformes'].add(plate)
        if region:
            t['regions'].add(region)
        if metier:
            t['metiers'].add(metier)
        if techno:
            t['techno'].add(techno)
        if dur > t['duration']:
            t['duration'] = dur
        if _has_impact(impact):
            t['avec_impact'] = True
            if not t['impact']:
                t['impact'] = _clean(impact)
        if region and metier:
            t['region_metier'].add((region, metier))
        if not t['nature']:
            t['nature'] = nature
        if not t['cause']:
            t['cause'] = cause

    def categorie(t):
        if 'BACKHAUL' in t['plateformes']:
            return 'backhaul'
        if 'BACKBONE' in t['plateformes']:
            return 'backbone'
        return None        # LIAISON PARTENAIRE / PLATEFORME DE SUP → hors Image 1

    # ── Image 1 : Backhaul / Backbone ──
    def make_cat():
        return {'tickets': [], 'sans': [], 'avec': []}
    cats = {'backhaul': make_cat(), 'backbone': make_cat()}
    for t in tickets.values():
        c = categorie(t)
        if not c:
            continue
        cats[c]['tickets'].append(t)
        (cats[c]['avec'] if t['avec_impact'] else cats[c]['sans']).append(t)

    def block(lst):
        return {
            'inc': len(lst),
            'mttr_sec': _avg([t['duration'] for t in lst]),
        }

    image1 = {}
    for c in ('backhaul', 'backbone'):
        cc = cats[c]
        image1[c] = {
            **block(cc['tickets']),
            'sans_impact': block(cc['sans']),
            'avec_impact': block(cc['avec']),
        }
    image1['total_inc'] = image1['backhaul']['inc'] + image1['backbone']['inc']
    # Détails Backbone (Liens/Site, Impact-Service, Cause)
    image1['backbone_details'] = [
        {
            'lien': t['site'] or t['nature'],
            'impact': t['impact'] or '—',
            'cause': t['cause'] or '—',
        }
        for t in sorted(cats['backbone']['tickets'],
                        key=lambda x: x['duration'], reverse=True)
    ]

    # ── Image 2 : par région × métier ──
    # Un ticket compte dans chaque (région, métier) qu'il touche.
    rm_tickets = {}       # (region, metier) -> list[ticket]
    regions_present = []
    for t in tickets.values():
        for (region, metier) in t['region_metier']:
            rm_tickets.setdefault((region, metier), []).append(t)
            if region not in regions_present:
                regions_present.append(region)

    ordered_regions = [r for r in REGIONS] + \
        [r for r in regions_present if r not in REGIONS]

    regions_out = []
    for region in ordered_regions:
        rows = []
        reg_inc = 0
        reg_durs = []
        for metier in METIERS:
            lst = rm_tickets.get((region, metier), [])
            durs = [t['duration'] for t in lst]
            rows.append({'metier': metier, 'inc': len(lst),
                         'mttr_sec': _avg(durs)})
            reg_inc += len(lst)
            reg_durs += durs
        regions_out.append({
            'region': region, 'metiers': rows,
            'total_inc': reg_inc, 'total_mttr_sec': _avg(reg_durs),
            'has_data': reg_inc > 0, 'canonical': region in REGIONS,
        })

    # Box BACKBONE DWDM
    dwdm = [t for t in tickets.values()
            if 'BACKBONE' in t['plateformes']
            and any('DWDM' in x for x in t['techno'] | {_norm(t['nature'])})]
    dwdm_services = [t['impact'] for t in dwdm if t['impact']]
    backbone_dwdm = {
        'count': len(dwdm),
        'mttr_sec': _avg([t['duration'] for t in dwdm]),
        'services': '; '.join(dwdm_services) if dwdm_services
        else 'Aucun service impacté',
    }

    image2 = {'regions': regions_out, 'backbone_dwdm': backbone_dwdm}

    # ── Image 3 : Disponibilité clients IPT & IPLC ──
    # On ne garde que les incidents PARTENAIRE (LIAISON PARTENAIRE) ; un client
    # correspond si son nom apparaît dans la colonne « Nature de l'incident ».
    inclusive_days = (end.date() - start.date()).days + 1
    period_sec = max(inclusive_days, 1) * 86400
    partner_tickets = [t for t in tickets.values() if t['is_partner']]
    clients_out = []
    for disp, keys in CLIENTS_IPT:
        nkeys = [_norm(k) for k in keys if k]
        matched = [t for t in partner_tickets
                   if any(nk in t['nature_text'] for nk in nkeys)]
        durs = [t['duration'] for t in matched]
        downtime = sum(durs)
        taux = max(0.0, (period_sec - downtime) / period_sec * 100.0)
        clients_out.append({
            'name': disp, 'inc': len(matched),
            'durations': durs, 'taux': taux,
        })
    image3 = {'clients': clients_out, 'period_sec': period_sec}

    return {
        'period_label': period_label,
        'month_label': month_label,
        'total_inc': image1['total_inc'],
        'image1': image1,
        'image2': image2,
        'image3': image3,
    }


# ═══════════════════════════════════════════════════════════════════════════
# RENDU PNG (Pillow)
# ═══════════════════════════════════════════════════════════════════════════

NAVY = (13, 36, 97)
BLUE = (0, 48, 135)
RED = (227, 0, 19)
YELLOW = (255, 199, 44)
GRAY = (232, 236, 244)
WHITE = (255, 255, 255)
DARKBLUE = (22, 45, 114)


def _logo(img, img_w, margin, top=22, w=120):
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


_TOGO_CACHE = None


def _togo_polys():
    """Charge et met en cache les anneaux extérieurs des régions du Togo."""
    global _TOGO_CACHE
    if _TOGO_CACHE is not None:
        return _TOGO_CACHE
    import os
    import json
    polys = []
    try:
        path = os.path.join(os.path.dirname(__file__), 'static', 'reports',
                            'togo_geo.js')
        with open(path, 'r', encoding='utf-8') as fh:
            txt = fh.read()
        txt = txt[txt.index('{'):].rstrip().rstrip(';')
        data = json.loads(txt)
        for feat in data.get('features', []):
            geom = feat.get('geometry', {}) or {}
            gtype = geom.get('type')
            coords = geom.get('coordinates', [])
            if gtype == 'Polygon':
                polys.append(coords[0])
            elif gtype == 'MultiPolygon':
                for poly in coords:
                    polys.append(poly[0])
    except Exception:
        pass
    _TOGO_CACHE = polys
    return polys


def _draw_togo(d, cx, cy, cr):
    """Dessine la carte du Togo (régions) centrée dans le cercle."""
    polys = _togo_polys()
    if not polys:
        return
    xs = [p[0] for poly in polys for p in poly]
    ys = [p[1] for poly in polys for p in poly]
    minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
    w = (maxx - minx) or 1
    h = (maxy - miny) or 1
    box = cr * 1.7
    scale = min(box / w, box / h)
    ox = cx - (w * scale) / 2
    oy = cy - (h * scale) / 2

    def proj(pt):
        return (ox + (pt[0] - minx) * scale,
                oy + (maxy - pt[1]) * scale)        # inversion latitude

    for poly in polys:
        pts = [proj(p) for p in poly]
        d.polygon(pts, fill=(236, 239, 246), outline=WHITE)


def _backbone_circle(img, cx, cy, cr):
    """Colle l'image officielle Backbone (carte) découpée en cercle.

    Retourne True si l'image a été collée, False sinon (repli sur la
    carte dessinée + texte)."""
    import os
    from PIL import Image, ImageDraw
    path = os.path.join(os.path.dirname(__file__), 'static', 'reports',
                        'backbone_togo.png')
    if not os.path.exists(path):
        return False
    try:
        src = Image.open(path).convert('RGBA')
        size = cr * 2
        # recadrage centré (cover) puis redimensionnement carré
        sw, sh = src.size
        scale = max(size / sw, size / sh)
        src = src.resize((int(sw * scale), int(sh * scale)))
        sw, sh = src.size
        left = (sw - size) // 2
        top = (sh - size) // 2
        src = src.crop((left, top, left + size, top + size))
        # masque circulaire
        mask = Image.new('L', (size, size), 0)
        ImageDraw.Draw(mask).ellipse([0, 0, size - 1, size - 1], fill=255)
        img.paste(src, (cx - cr, cy - cr), mask)
        return True
    except Exception:
        return False


def _block_arrow(d, x0, x1, cy, fill, sh=11, hh=22, hl=34):
    """Flèche bloc orientée vers la droite."""
    pts = [(x0, cy - sh), (x1 - hl, cy - sh), (x1 - hl, cy - hh),
           (x1, cy), (x1 - hl, cy + hh), (x1 - hl, cy + sh), (x0, cy + sh)]
    d.polygon(pts, fill=fill)


def build_png_image1(report, generated_on=''):
    """Image 1 — Détails Incident transport."""
    from PIL import Image, ImageDraw

    im1 = report['image1']
    img_w, img_h = 1560, 1080
    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)

    f_title = _load_font(30, bold=True)
    f_sub = _load_font(20, bold=True)
    f_big = _load_font(25, bold=True)
    f_node = _load_font(23, bold=True)
    f_lbl = _load_font(16, bold=True)
    f_cell = _load_font(13, bold=False)
    f_cellb = _load_font(13, bold=True)
    f_foot = _load_font(12, bold=False)

    margin = 40
    d.text((margin, 24), 'COMITE GESTION DES INCIDENTS', font=f_title, fill=BLUE)
    d.text((margin, 64), 'Détails Incident transport', font=f_sub, fill=RED)

    # Total (texte simple, sans encadré)
    d.text((margin, 128),
           f"{report['total_inc']} incidents enregistrés",
           font=f_big, fill=NAVY)

    # Cercle central « Backbone de transmission » : image officielle si dispo,
    # sinon cercle jaune + carte dessinée + texte.
    cx, cy, cr = 220, 470, 150
    if not _backbone_circle(img, cx, cy, cr):
        d.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=YELLOW)
        _draw_togo(d, cx, cy, cr)
        for i, ln in enumerate(['Backbone de', 'transmission', 'De Yas']):
            tw = d.textlength(ln, font=f_node)
            d.text((cx - tw / 2, cy - 42 + i * 30), ln, font=f_node, fill=NAVY)

    box_w, box_h = 300, 84

    def draw_branch(nx, ny, label, blk):
        r = 92
        # lien depuis le cercle central
        d.line([cx + cr - 8, cy, nx - r + 8, ny], fill=NAVY, width=4)
        # noeud
        d.ellipse([nx - r, ny - r, nx + r, ny + r], fill=NAVY)
        tw = d.textlength(label, font=f_node)
        d.text((nx - tw / 2, ny - 14), label, font=f_node, fill=WHITE)
        # Inc / MTTR à droite du noeud
        tx = nx + r + 22
        d.text((tx, ny - 32), f"Inc : {blk['inc']}", font=f_big, fill=NAVY)
        d.text((tx, ny + 4), f"MTTR : {fmt_hms(blk['mttr_sec'])}",
               font=f_big, fill=RED)
        # flèche bloc vers les boîtes
        bx = nx + r + 360
        _block_arrow(d, nx + r + 250, bx - 6, ny, NAVY)
        sans_y = ny - box_h - 14
        avec_y = ny + 14

        def ibox(by, title, b):
            d.rectangle([bx, by, bx + box_w, by + box_h], fill=NAVY)
            d.text((bx + 14, by + 10), title, font=f_lbl, fill=YELLOW)
            d.text((bx + 14, by + 36), f"Nbre : {b['inc']}", font=f_cellb, fill=WHITE)
            d.text((bx + 14, by + 58), f"MTTR : {fmt_hms(b['mttr_sec'])}",
                   font=f_cellb, fill=WHITE)
        ibox(sans_y, 'INC SANS IMPACT', blk['sans_impact'])
        ibox(avec_y, 'INC AVEC IMPACT', blk['avec_impact'])

    draw_branch(620, 300, 'Backhaul', im1['backhaul'])
    draw_branch(620, 640, 'BackBone', im1['backbone'])

    # Table détails Backbone (bas, pleine largeur gauche)
    details = im1['backbone_details']
    if details:
        tx0, ty0 = margin, 800
        d.text((tx0, ty0 - 28), 'Détails incidents BACKBONE', font=f_lbl, fill=NAVY)
        cols_w = [330, 260, 230]
        headers = ["Nature de l'incident", 'Impact - Service', 'Cause']
        x = tx0
        for i, h in enumerate(headers):
            d.rectangle([x, ty0, x + cols_w[i], ty0 + 30], fill=DARKBLUE,
                        outline=WHITE, width=1)
            d.text((x + 6, ty0 + 7), h, font=f_cellb, fill=YELLOW)
            x += cols_w[i]
        y = ty0 + 30
        pad = 6
        line_h = 16
        for row in details[:6]:
            vals = [row['lien'], row['impact'], row['cause']]
            wrapped = [_wrap_text(v, f_cell, cols_w[i] - 2 * pad)
                       for i, v in enumerate(vals)]
            rh = max(len(w) for w in wrapped) * line_h + 2 * pad
            x = tx0
            for i, w in enumerate(wrapped):
                bg = NAVY if i == 0 else GRAY
                fg = WHITE if i == 0 else NAVY
                d.rectangle([x, y, x + cols_w[i], y + rh], fill=bg,
                            outline=WHITE, width=1)
                ty = y + pad
                for ln in w:
                    d.text((x + pad, ty), ln, font=f_cell, fill=fg)
                    ty += line_h
                x += cols_w[i]
            y += rh

    if generated_on:
        d.text((margin, img_h - 26),
               f"Généré le {generated_on} — Yas Togo / DT / DOC / iSOC  —  {report['period_label']}",
               font=f_foot, fill=NAVY)

    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


def build_png_image3(report, generated_on=''):
    """Image 3 — Disponibilité clients IPT & IPLC."""
    from PIL import Image, ImageDraw

    im3 = report['image3']
    clients = im3['clients']

    f_title = _load_font(30, bold=True)
    f_sub = _load_font(20, bold=True)
    f_band = _load_font(18, bold=True)
    f_hdr = _load_font(15, bold=True)
    f_cell = _load_font(15, bold=True)
    f_dur = _load_font(14, bold=False)
    f_foot = _load_font(12, bold=False)

    margin = 40
    cols_w = [300, 110, 150, 230]      # LIENS | Nbre Inc | Durée | TAUX
    table_w = sum(cols_w)
    hdr_h = 56
    line_h = 24
    pad = 8

    # hauteur de chaque ligne = max(1, nb durées) lignes de texte
    def row_lines(c):
        return max(1, len(c['durations']))

    body_h = sum(row_lines(c) * line_h + 2 * pad for c in clients)
    img_w = margin * 2 + table_w
    img_h = 150 + hdr_h + body_h + 60

    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)
    d.text((margin, 24), 'COMITE GESTION DES INCIDENTS', font=f_title, fill=BLUE)
    d.text((margin, 64), 'Disponibilité clients IPT et IPLC', font=f_sub, fill=RED)
    # bandeau « Clients IPT & IPLC »
    d.rectangle([margin, 104, margin + 260, 138], fill=NAVY)
    d.text((margin + 16, 110), 'Clients IPT & IPLC', font=f_band, fill=WHITE)
    _logo(img, img_w, margin)

    # En-tête du tableau
    ty0 = 150
    headers = ['LIENS', 'Nbre Inc', 'Durée', 'TAUX DE\nDISPONIBILITE']
    x = margin
    for i, h in enumerate(headers):
        d.rectangle([x, ty0, x + cols_w[i], ty0 + hdr_h], fill=GRAY,
                    outline=NAVY, width=1)
        lines = h.split('\n')
        for j, ln in enumerate(lines):
            tw = d.textlength(ln, font=f_hdr)
            d.text((x + (cols_w[i] - tw) / 2,
                    ty0 + hdr_h / 2 - len(lines) * 9 + j * 18),
                   ln, font=f_hdr, fill=NAVY)
        x += cols_w[i]

    # Lignes
    y = ty0 + hdr_h
    for c in clients:
        nl = row_lines(c)
        rh = nl * line_h + 2 * pad
        # LIENS
        d.rectangle([margin, y, margin + cols_w[0], y + rh], fill=WHITE,
                    outline=NAVY, width=1)
        tw = d.textlength(c['name'], font=f_cell)
        d.text((margin + (cols_w[0] - tw) / 2, y + rh / 2 - 9),
               c['name'], font=f_cell, fill=NAVY)
        # Nbre Inc (jaune si > 0)
        x1 = margin + cols_w[0]
        nb_bg = YELLOW if c['inc'] > 0 else WHITE
        d.rectangle([x1, y, x1 + cols_w[1], y + rh], fill=nb_bg,
                    outline=NAVY, width=1)
        nb = str(c['inc'])
        tw = d.textlength(nb, font=f_cell)
        d.text((x1 + (cols_w[1] - tw) / 2, y + rh / 2 - 9), nb,
               font=f_cell, fill=RED if c['inc'] > 0 else NAVY)
        # Durée (une ligne par incident)
        x2 = x1 + cols_w[1]
        d.rectangle([x2, y, x2 + cols_w[2], y + rh], fill=WHITE,
                    outline=NAVY, width=1)
        durs = c['durations'] if c['durations'] else [0]
        yy = y + pad
        for dsec in durs:
            txt = fmt_hms(dsec)
            tw = d.textlength(txt, font=f_dur)
            d.text((x2 + (cols_w[2] - tw) / 2, yy), txt, font=f_dur, fill=NAVY)
            yy += line_h
        # TAUX
        x3 = x2 + cols_w[2]
        d.rectangle([x3, y, x3 + cols_w[3], y + rh], fill=WHITE,
                    outline=NAVY, width=1)
        taux = f"{c['taux']:.2f}".replace('.', ',') + ' %'
        tw = d.textlength(taux, font=f_cell)
        d.text((x3 + (cols_w[3] - tw) / 2, y + rh / 2 - 9), taux,
               font=f_cell, fill=NAVY)
        y += rh

    if generated_on:
        d.text((margin, img_h - 26),
               f"Généré le {generated_on} — Yas Togo / DT / DOC / iSOC  —  {report['period_label']}",
               font=f_foot, fill=NAVY)

    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


def build_png_image2(report, generated_on=''):
    """Image 2 — Count Inc & MTTR par Métier et par Régions."""
    from PIL import Image, ImageDraw

    im2 = report['image2']
    regions = [r for r in im2['regions'] if r['canonical'] or r['has_data']]

    margin = 40
    f_title = _load_font(28, bold=True)
    f_sub = _load_font(19, bold=True)
    f_reg = _load_font(15, bold=True)
    f_hdr = _load_font(13, bold=True)
    f_cell = _load_font(13, bold=False)
    f_cellb = _load_font(13, bold=True)
    f_box = _load_font(15, bold=True)
    f_foot = _load_font(12, bold=False)

    # Disposition : 2 colonnes de tables régionales
    band_w = 108
    cell_w = [150, 90, 150]      # Métier | Inc | MTTR
    table_w = band_w + sum(cell_w)
    cols_x = [margin, margin + table_w + 45]
    row_h = 28
    hdr_h = 28
    reg_gap = 26

    # Pré-calcul hauteur de chaque table
    def table_h(reg):
        return hdr_h + row_h * len(reg['metiers'])

    # Répartit les régions en 2 colonnes
    n = len(regions)
    left = regions[:(n + 1) // 2]
    right = regions[(n + 1) // 2:]
    col_height = lambda lst: sum(table_h(r) + reg_gap for r in lst)
    body_h = max(col_height(left), col_height(right), 300)
    img_w = cols_x[1] + table_w + 40 + 320      # +box droite
    img_h = 150 + body_h + 60

    img = Image.new('RGB', (img_w, img_h), WHITE)
    d = ImageDraw.Draw(img)
    d.text((margin, 24), 'COMITE GESTION DES INCIDENTS', font=f_title, fill=BLUE)
    d.text((margin, 60), 'Count Inc & MTTR par Métier et par Régions', font=f_sub, fill=RED)
    _logo(img, img_w, margin)

    def draw_table(x0, y0, reg):
        th = hdr_h + row_h * len(reg['metiers'])
        # bandeau région (nom centré verticalement)
        d.rectangle([x0, y0, x0 + band_w, y0 + th], fill=YELLOW)
        rtw = d.textlength(reg['region'], font=f_reg)
        d.text((x0 + (band_w - rtw) / 2, y0 + th / 2 - 9), reg['region'],
               font=f_reg, fill=NAVY)
        tx = x0 + band_w
        # en-tête
        headers = ['Métier', 'Inc', 'MTTR']
        x = tx
        for i, h in enumerate(headers):
            d.rectangle([x, y0, x + cell_w[i], y0 + hdr_h], fill=GRAY,
                        outline=WHITE, width=1)
            d.text((x + 8, y0 + 6), h, font=f_hdr, fill=NAVY)
            x += cell_w[i]
        y = y0 + hdr_h
        for r in reg['metiers']:
            x = tx
            mttr = fmt_hms(r['mttr_sec']) if r['inc'] else '0:00:00'
            hot = r['mttr_sec'] >= 5 * 3600
            cells = [(r['metier'], WHITE, NAVY, f_cell, 'left'),
                     (str(r['inc']), WHITE, NAVY, f_cellb, 'center'),
                     (mttr, RED if hot else WHITE, WHITE if hot else NAVY,
                      f_cellb, 'center')]
            for i, (val, bg, fg, font, al) in enumerate(cells):
                d.rectangle([x, y, x + cell_w[i], y + row_h], fill=bg,
                            outline=GRAY, width=1)
                tw = d.textlength(val, font=font)
                tx2 = x + (cell_w[i] - tw) / 2 if al == 'center' else x + 8
                d.text((tx2, y + 6), val, font=font, fill=fg)
                x += cell_w[i]
            y += row_h
        return y

    for lst, x0 in ((left, cols_x[0]), (right, cols_x[1])):
        y = 150
        for reg in lst:
            end = draw_table(x0, y, reg)
            y = end + reg_gap

    # Box BACKBONE DWDM (droite)
    bx = cols_x[1] + table_w + 40
    by = 150
    bw, bh = 300, 122
    dwdm = im2['backbone_dwdm']
    d.rounded_rectangle([bx, by, bx + bw, by + bh], radius=10, fill=NAVY)
    txt = (f"{dwdm['count']} Indisponibilité(s) du BACKBONE DWDM\n"
           f"MTTR : {fmt_hms(dwdm['mttr_sec'])}\n{dwdm['services']}")
    yy = by + 16
    for ln in txt.split('\n'):
        for w in _wrap_text(ln, f_box, bw - 28):
            d.text((bx + 14, yy), w, font=f_box, fill=WHITE)
            yy += 22

    if generated_on:
        d.text((margin, img_h - 26),
               f"Généré le {generated_on} — Yas Togo / DT / DOC / iSOC  —  {report['period_label']}",
               font=f_foot, fill=NAVY)

    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf
