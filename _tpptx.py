import os, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'rapport_automatic.settings')
django.setup()
from reports.transport_noc import parse_transport_noc, build_png_image1, build_png_image2
from reports.pptx_report import generate_image_slide

f = r"c:\Users\user\Videos\Rapport-Auto\TRANSMISSION_20260525_20260531.xlsx"
r = parse_transport_noc(f, filename=os.path.basename(f))
p1 = build_png_image1(r, '08/06/2026')
p2 = build_png_image2(r, '08/06/2026')
b1 = generate_image_slide(p1, '08/06/2026', r['period_label'])
b2 = generate_image_slide(p2, '08/06/2026', r['period_label'])
print('pptx1 bytes', len(b1.getvalue()))
print('pptx2 bytes', len(b2.getvalue()))
print('OK')
