# -*- coding: utf-8 -*-
"""Convertit « Cartes Liens Backbones .drawio.xml » en SVG autonome.

Usage :  python _drawio2svg.py
Sortie :  reports/static/reports/carte_liens_backbones.svg

À relancer (puis redémarrer le serveur pour collectstatic) chaque fois que
la carte drawio est mise à jour. Aucune dépendance externe (stdlib only).
"""
import xml.etree.ElementTree as ET
import re
import html as htmlmod

SRC = 'Cartes Liens Backbones .drawio.xml'
DST = 'reports/static/reports/carte_liens_backbones.svg'


def style_dict(s):
    d = {}
    for part in (s or '').split(';'):
        if '=' in part:
            k, v = part.split('=', 1)
            d[k] = v
        elif part:
            d[part] = '1'
    return d


def clean_color(c, default='none'):
    if not c or c == 'none':
        return default if c is None else 'none'
    m = re.match(r'light-dark\(([^,]+),', c)
    if m:
        return m.group(1).strip()
    return c


def parse_label(value):
    """HTML drawio -> (lignes de texte, font-size, couleur, bold)."""
    if not value:
        return [], 12.0, '#000000', False
    sizes = [float(s) for s in re.findall(r'font-size:\s*([\d.]+)px', value)]
    sizes = [s for s in sizes if s > 2]  # ignorer les wrappers font-size:1px
    size = max(sizes) if sizes else 12.0
    color = '#000000'
    m = re.search(r'color:\s*(#[0-9a-fA-F]{6})', value)
    if m:
        color = m.group(1)
    else:
        m = re.search(r'color:\s*rgb\((\d+),\s*(\d+),\s*(\d+)\)', value)
        if m:
            color = '#%02x%02x%02x' % tuple(int(x) for x in m.groups())
    bold = '<b>' in value or 'font-weight' in value and 'bold' in value
    txt = re.sub(r'<br\s*/?>', '\n', value)
    txt = re.sub(r'<[^>]+>', '', txt)
    txt = htmlmod.unescape(txt)
    lines = [l.strip() for l in txt.split('\n') if l.strip()]
    return lines, size, color, bold


def esc(t):
    return (t.replace('&', '&amp;').replace('<', '&lt;')
             .replace('>', '&gt;').replace('"', '&quot;'))


def main():
    tree = ET.parse(SRC)
    root_cells = tree.getroot().find('.//mxGraphModel/root')

    vertices = {}      # id -> (x, y, w, h)
    elements = []      # liste ordonnée de fragments SVG (ordre du document)
    labels = []        # rendus en dernier (au-dessus des lignes)
    edges = []

    for el in root_cells:
        mx = el if el.tag == 'mxCell' else el.find('mxCell')
        if mx is None:
            continue
        cid = mx.get('id') or el.get('id')
        st = style_dict(mx.get('style', ''))
        g = mx.find('mxGeometry')
        value = mx.get('value') or el.get('label') or ''

        if mx.get('edge') == '1':
            edges.append((mx, st, g))
            continue
        if g is None:
            continue
        x = float(g.get('x', 0)); y = float(g.get('y', 0))
        w = float(g.get('width', 0)); h = float(g.get('height', 0))
        if cid:
            vertices[cid] = (x, y, w, h)

        # Image de fond (carte)
        img = st.get('image', '')
        if 'image' in st and img.startswith('data:'):
            # data:image/jpg,<base64> -> data URI valide avec ;base64
            m = re.match(r'data:image/([a-zA-Z]+),(.*)', img, re.S)
            if m:
                href = f'data:image/{"jpeg" if m.group(1)=="jpg" else m.group(1)};base64,{m.group(2)}'
            else:
                href = img
            elements.append(
                f'<image x="{x:.2f}" y="{y:.2f}" width="{w:.2f}" height="{h:.2f}" '
                f'xlink:href="{href}" preserveAspectRatio="xMidYMid meet"/>')
        elif w > 0 and h > 0:
            fill = clean_color(st.get('fillColor'), 'none')
            stroke = clean_color(st.get('strokeColor'), 'none')
            sw = st.get('strokeWidth', '1')
            # cadre « Arrière-plan » : on saute (aucun intérêt visuel)
            if (el.get('tags') or '') == 'Arrière-plan':
                pass
            elif 'ellipse' in st:
                elements.append(
                    f'<ellipse cx="{x + w/2:.2f}" cy="{y + h/2:.2f}" rx="{w/2:.2f}" ry="{h/2:.2f}" '
                    f'fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>')
            elif fill != 'none' or stroke != 'none':
                elements.append(
                    f'<rect x="{x:.2f}" y="{y:.2f}" width="{w:.2f}" height="{h:.2f}" '
                    f'fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>')

        # Label éventuel
        lines, size, color, bold = parse_label(value)
        if lines:
            cx = x + w / 2
            cy = y + h / 2
            n = len(lines)
            lh = size * 1.2
            y0 = cy - (n - 1) * lh / 2
            weight = '700' if bold else '400'
            tspans = ''.join(
                f'<tspan x="{cx:.2f}" y="{y0 + i * lh:.2f}">{esc(l)}</tspan>'
                for i, l in enumerate(lines))
            labels.append(
                f'<text text-anchor="middle" dominant-baseline="middle" '
                f'font-family="Calibri, Arial, sans-serif" font-size="{size:.2f}" '
                f'fill="{color}" font-weight="{weight}">{tspans}</text>')

    # ── Arêtes ────────────────────────────────────────────────────────────
    edge_frags = []
    for mx, st, g in edges:
        def endpoint(ref_id, rx_key, ry_key, point_name):
            ref = mx.get(ref_id)
            rx, ry = st.get(rx_key), st.get(ry_key)
            if ref and ref in vertices and rx is not None and ry is not None:
                vx, vy, vw, vh = vertices[ref]
                return vx + float(rx) * vw, vy + float(ry) * vh
            if ref and ref in vertices and (rx is None or ry is None):
                vx, vy, vw, vh = vertices[ref]
                return vx + vw / 2, vy + vh / 2
            if g is not None:
                p = g.find(f"mxPoint[@as='{point_name}']")
                if p is not None:
                    return float(p.get('x', 0)), float(p.get('y', 0))
            return None

        p1 = endpoint('source', 'exitX', 'exitY', 'sourcePoint')
        p2 = endpoint('target', 'entryX', 'entryY', 'targetPoint')
        if not p1 or not p2:
            continue
        pts = [p1]
        if g is not None:
            arr = g.find("Array[@as='points']")
            if arr is not None:
                for p in arr.findall('mxPoint'):
                    pts.append((float(p.get('x', 0)), float(p.get('y', 0))))
        pts.append(p2)
        stroke = clean_color(st.get('strokeColor'), '#000000')
        sw = st.get('strokeWidth', '1')
        dash = ' stroke-dasharray="8 6"' if st.get('dashed') == '1' else ''
        d = 'M ' + ' L '.join(f'{px:.2f} {py:.2f}' for px, py in pts)
        edge_frags.append(
            f'<path d="{d}" fill="none" stroke="{stroke}" stroke-width="{sw}"'
            f'{dash} stroke-linecap="round"/>')

    # ── Bounding box ──────────────────────────────────────────────────────
    xs, ys = [], []
    for vx, vy, vw, vh in vertices.values():
        xs += [vx, vx + vw]; ys += [vy, vy + vh]
    margin = 12
    minx, miny = min(xs) - margin, min(ys) - margin
    maxx, maxy = max(xs) + margin, max(ys) + margin
    W, H = maxx - minx, maxy - miny

    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'viewBox="{minx:.2f} {miny:.2f} {W:.2f} {H:.2f}" '
        f'width="{W:.0f}" height="{H:.0f}">\n'
        f'<rect x="{minx:.2f}" y="{miny:.2f}" width="{W:.2f}" height="{H:.2f}" fill="#ffffff"/>\n'
        + '\n'.join(elements) + '\n'
        + '\n'.join(edge_frags) + '\n'
        + '\n'.join(labels) + '\n</svg>\n'
    )
    with open(DST, 'w', encoding='utf-8') as f:
        f.write(svg)
    print(f'OK -> {DST}  ({len(svg)} octets, {len(edge_frags)} liens, {len(labels)} labels)')


if __name__ == '__main__':
    main()
