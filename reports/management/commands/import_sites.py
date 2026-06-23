import pandas as pd
from django.core.management.base import BaseCommand
from reports.models import Site


def _clean(val):
    if pd.isna(val):
        return ''
    s = str(val).strip()
    return '' if s.lower() in ('nan', 'none', 'n/a') else s


class Command(BaseCommand):
    help = 'Import sites from Excel file into the Site table'

    def add_arguments(self, parser):
        parser.add_argument('filepath', type=str, help='Path to the Excel file')

    def handle(self, *args, **options):
        path = options['filepath']
        self.stdout.write(f'Reading {path}...')
        df = pd.read_excel(path)

        created = updated = 0
        for _, row in df.iterrows():
            name = _clean(row.get('SITE NAME', ''))
            if not name:
                continue

            date_mes = None
            raw_date = row.get('DATE MES')
            if pd.notna(raw_date):
                try:
                    date_mes = pd.to_datetime(raw_date).date()
                except Exception:
                    pass

            # Longitude/Latitude : accepter majuscules ou minuscules
            lon = row.get('LONGITUDE') if pd.notna(row.get('LONGITUDE', float('nan'))) else row.get('Longitude')
            lat = row.get('LATITUDE')  if pd.notna(row.get('LATITUDE',  float('nan'))) else row.get('Latitude')

            defaults = dict(
                date_mes=date_mes,
                site_id=_clean(row.get('SITE ID', '')),
                region=_clean(row.get('REGION', '')),
                base=_clean(row.get('BASE', '')),
                olt=_clean(row.get('OLT', '')),
                longitude=float(lon) if lon is not None and pd.notna(lon) else None,
                latitude=float(lat) if lat is not None and pd.notna(lat) else None,
                config=_clean(row.get('CONFIG', '')),
                techno=_clean(row.get('Techno', '')),
                typ_trans=_clean(row.get('TYP TRANS', '')),
                typ_energie=_clean(row.get('TYP ENERGIE', '')),
                ge_auto=_clean(row.get('GE AUTO', '')),
                site_lithium=_clean(row.get('SITE LITHIUM', '')),
                site_esm=_clean(row.get('SITE ESM', '')),
                site_solaire_neteco=_clean(row.get('SITE SOLAIRE NETECO', '')),
                config_2g=_clean(row.get('CONFIG 2G', '')),
                config_3g=_clean(row.get('CONFIG 3G', '')),
                config_4g=_clean(row.get('CONFIG 4G', '')),
                classif_tech=_clean(row.get('CLASSIF TECH', '')),
                type_site=_clean(row.get('TYPE SITE', '')),
                numero_agent=_clean(row.get('NUMERO AGENT', '')),
                societe_gardiens=_clean(row.get('Société GARDIENS ', '')),
                contacts_surveillants=_clean(row.get('Contacts des surveillants', '')),
            )
            # Champs optionnels présents uniquement dans certains fichiers
            zone = _clean(row.get('ZONE', ''))
            if zone:
                defaults['zone'] = zone
            typo_avant = _clean(row.get('TYPOLOGIE AVANT Projet', ''))
            if typo_avant:
                defaults['typologie_avant'] = typo_avant
            typo_apres = _clean(row.get('TYPOLOGIE Apres Projet', ''))
            if typo_apres:
                defaults['typologie_apres'] = typo_apres

            obj, is_new = Site.objects.update_or_create(
                site_name=name,
                defaults=defaults,
            )
            if is_new:
                created += 1
            else:
                updated += 1

        self.stdout.write(self.style.SUCCESS(
            f'Done — {created} créés, {updated} mis à jour.'
        ))
