# -*- coding: utf-8 -*-
import django, os, re
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'rapport_automatic.settings')
django.setup()
from django.test import Client
from django.contrib.auth.models import User

c = Client()
c.force_login(User.objects.get(username='admin'))
r = c.get('/site-info/?q=ABLOGAME')
html = r.content.decode()
print('status:', r.status_code)
for token in ['id="tabCarte"', 'id="tabArchi"', 'id="geoView"', 'id="archiView"',
              'id="geoMap"', 'id="geoBtnNear"', 'id="geoBtnAll"', 'id="archiCanvas"',
              'NEAR_COUNT', 'showNear()']:
    print(token, '->', token in html)
print('site-geo-section (ancienne) :', 'site-geo-section' in html)
m = re.search(r'var SITES_GEO = (\[.*?\]);', html, re.S)
print('nb sites:', m.group(1).count('"name"') if m else 0)
