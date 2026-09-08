# -*- coding: utf-8 -*-
"""Aperçu autonome du cadre bleu : architecture (haut) + carte Togo (bas)."""
import django, os, json
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'rapport_automatic.settings')
django.setup()
from reports.models import Site

sites = [
    {'name': s.site_name, 'lat': s.latitude, 'lon': s.longitude, 'region': s.region}
    for s in Site.objects.exclude(latitude__isnull=True).exclude(longitude__isnull=True)
]

html = """<!DOCTYPE html><html><head><meta charset="utf-8">
<link rel="stylesheet" href="http://localhost:8000/static/reports/leaflet/leaflet.css">
<style>
  body { margin:0; background:#f0f2f8; font-family:sans-serif; padding:24px; }
  .col { width:520px; margin:0 auto; display:flex; flex-direction:column; gap:10px; }
  .title { font-size:12px; font-weight:700; color:#003087; }
  .archi-geo-frame {
    display:flex; flex-direction:column; border-radius:12px; overflow:hidden;
    background:radial-gradient(circle at 30% 20%,#1c3f96 0%,#102a6e 55%,#0b1f55 100%);
    box-shadow:inset 0 0 40px rgba(0,0,0,.25);
  }
  .site-archi-canvas { min-height:280px; display:flex; align-items:center; justify-content:center;
    color:rgba(255,255,255,.5); font-style:italic; font-size:13px; }
  .archi-geo-divider {
    display:flex; align-items:center; justify-content:space-between; gap:8px; flex-wrap:wrap;
    padding:8px 14px; background:rgba(0,0,0,.28);
    border-top:1px solid rgba(255,255,255,.14); border-bottom:1px solid rgba(255,255,255,.14);
  }
  .archi-geo-divider-title { font-size:11px; font-weight:800; letter-spacing:1px; text-transform:uppercase;
    color:#FFC72C; display:flex; align-items:center; gap:6px; }
  .archi-geo-divider-title small { color:rgba(255,255,255,.55); font-weight:500; letter-spacing:0; text-transform:none; font-size:10.5px; }
  .geo-mini-btn { display:inline-flex; align-items:center; gap:4px; font-size:11px; font-weight:700;
    padding:4px 10px; border-radius:8px; border:1px solid rgba(255,255,255,.3); cursor:pointer;
    background:rgba(255,255,255,.1); color:#fff; }
  .geo-mini-btn.active { background:#FFC72C; color:#003087; border-color:#FFC72C; }
  #geoMap { height:380px; background:#0b1f55; position:relative; z-index:0; }
  #geoMap .leaflet-container { background:#0b1f55; }
  #geoMap .leaflet-tile-pane { filter: invert(1) hue-rotate(200deg) brightness(.72) contrast(.95) saturate(.55); }
  .geo-pin-img { filter: drop-shadow(0 2px 3px rgba(0,0,0,.55)); }
  .site-geo-legend { display:flex; flex-wrap:wrap; gap:12px; font-size:11px; color:#6b7a99; }
  .site-geo-legend span { display:inline-flex; align-items:center; gap:5px; }
  .site-geo-legend i { width:10px; height:10px; border-radius:50%; display:inline-block; }
</style></head><body>
<div class="col">
  <div class="title">🏗️ Architecture du site <span style="font-weight:500;color:#8896b3">— qui porte qui</span></div>
  <div class="archi-geo-frame">
    <div class="site-archi-canvas">[ digraphe architecture ici ]</div>
    <div class="archi-geo-divider">
      <div class="archi-geo-divider-title">🗺️ Géolocalisation <small>— carte du Togo</small></div>
      <div>
        <button class="geo-mini-btn active">🇹🇬 Tous les sites (1214)</button>
        <button class="geo-mini-btn">🎯 ANFAME2</button>
      </div>
    </div>
    <div id="geoMap"></div>
  </div>
  <div class="site-geo-legend">
    <span><i style="background:#ff3b30;"></i>Site du réseau</span>
    <span><i style="background:#FFC72C;border:2px solid #003087;width:12px;height:12px;"></i>ANFAME2 (consulté)</span>
    <span><i style="background:#1d2a63;border:1.5px solid #7d86ff;border-radius:3px;width:13px;height:9px;"></i>Territoire du Togo</span>
  </div>
</div>
<script src="http://localhost:8000/static/reports/leaflet/leaflet.js"></script>
<script src="http://localhost:8000/static/reports/togo_geo.js"></script>
<script>
var SITES_GEO = __SITES__;
var CURRENT = 'ANFAME2';
var TOGO_BOUNDS = [[6.05, -0.15], [11.15, 1.85]];
var map = L.map('geoMap', { zoomControl:true, scrollWheelZoom:false, attributionControl:false });
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom:18, minZoom:6 }).addTo(map);
map.fitBounds(TOGO_BOUNDS, { padding:[10,10] });
if (window.TOGO_GEOJSON) {
  L.geoJSON(window.TOGO_GEOJSON, {
    style: { color:'#7d86ff', weight:2.2, opacity:.95, fillColor:'#1d2a63', fillOpacity:.78 },
    interactive:false
  }).addTo(map);
}
function pinIcon(fill, stroke, w, h) {
  var svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 32'>" +
    "<path d='M12 1C6 1 1.4 5.6 1.4 11.4c0 8.2 10.6 19.2 10.6 19.2s10.6-11 10.6-19.2C22.6 5.6 18 1 12 1z'" +
    " fill='" + fill + "' stroke='" + stroke + "' stroke-width='1.6'/>" +
    "<circle cx='12' cy='11.4' r='4.2' fill='#fff'/></svg>";
  return L.icon({ iconUrl:'data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg),
    iconSize:[w,h], iconAnchor:[w/2,h], popupAnchor:[0,-h+4], className:'geo-pin-img' });
}
var ICON_SITE  = pinIcon('#ff3b30', '#8f0d00', 17, 23);
var ICON_FOCUS = pinIcon('#FFC72C', '#003087', 30, 40);
var currentSite = null;
SITES_GEO.forEach(function(s){
  if (typeof s.lat !== 'number' || typeof s.lon !== 'number') return;
  if (s.name === CURRENT) { currentSite = s; return; }
  L.marker([s.lat, s.lon], { icon: ICON_SITE, keyboard:false }).addTo(map);
});
if (currentSite) L.marker([currentSite.lat, currentSite.lon], { icon: ICON_FOCUS, zIndexOffset:1000 }).addTo(map);
setTimeout(function(){ map.invalidateSize(); }, 200);
</script></body></html>"""

with open('_geo_preview.html', 'w', encoding='utf-8') as f:
    f.write(html.replace('__SITES__', json.dumps(sites)))
print('OK', len(sites), 'sites')
