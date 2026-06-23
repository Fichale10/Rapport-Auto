import pandas as pd
from django.core.management.base import BaseCommand
from reports.models import Site


def _clean(val):
    if pd.isna(val):
        return ''
    s = str(val).strip()
    return '' if s.lower() in ('nan', 'none', 'n/a') else s


class Command(BaseCommand):
    help = 'Marque les sites solaires NetEco depuis le fichier "250 SITES SOLAIRES NET ECO.xlsx"'

    def add_arguments(self, parser):
        parser.add_argument('filepath', type=str, help='Chemin vers le fichier Excel des sites solaires')

    def handle(self, *args, **options):
        path = options['filepath']
        self.stdout.write(f'Lecture de {path}...')
        df = pd.read_excel(path)

        updated = not_found = 0
        for _, row in df.iterrows():
            name = _clean(row.get('SITE NAME', ''))
            if not name:
                continue

            defaults = {'site_solaire_neteco': 'OUI'}

            zone = _clean(row.get('ZONE', ''))
            if zone:
                defaults['zone'] = zone

            lon = row.get('LONGITUDE')
            lat = row.get('LATITUDE')
            if lon is not None and pd.notna(lon):
                try:
                    defaults['longitude'] = float(lon)
                except (ValueError, TypeError):
                    pass
            if lat is not None and pd.notna(lat):
                try:
                    defaults['latitude'] = float(lat)
                except (ValueError, TypeError):
                    pass

            typo_avant = _clean(row.get('TYPOLOGIE AVANT Projet', ''))
            if typo_avant:
                defaults['typologie_avant'] = typo_avant
            typo_apres = _clean(row.get('TYPOLOGIE Apres Projet', ''))
            if typo_apres:
                defaults['typologie_apres'] = typo_apres

            count = Site.objects.filter(site_name=name).update(**defaults)
            if count:
                updated += 1
            else:
                not_found += 1
                self.stdout.write(self.style.WARNING(f'  Site non trouvé en base : {name}'))

        self.stdout.write(self.style.SUCCESS(
            f'Terminé — {updated} sites marqués OUI, {not_found} non trouvés en base.'
        ))
