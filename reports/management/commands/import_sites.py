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

            lon = row.get('Longitude')
            lat = row.get('Latitude')

            obj, is_new = Site.objects.update_or_create(
                site_name=name,
                defaults=dict(
                    date_mes=date_mes,
                    site_id=_clean(row.get('SITE ID', '')),
                    region=_clean(row.get('REGION', '')),
                    base=_clean(row.get('BASE', '')),
                    olt=_clean(row.get('OLT', '')),
                    longitude=float(lon) if pd.notna(lon) else None,
                    latitude=float(lat) if pd.notna(lat) else None,
                    config=_clean(row.get('CONFIG', '')),
                    techno=_clean(row.get('Techno', '')),
                    typ_trans=_clean(row.get('TYP TRANS', '')),
                    typ_energie=_clean(row.get('TYP ENERGIE', '')),
                    ge_auto=_clean(row.get('GE AUTO', '')),
                    site_lithium=_clean(row.get('SITE LITHIUM', '')),
                    site_esm=_clean(row.get('SITE ESM', '')),
                    config_2g=_clean(row.get('CONFIG 2G', '')),
                    config_3g=_clean(row.get('CONFIG 3G', '')),
                    config_4g=_clean(row.get('CONFIG 4G', '')),
                    classif_tech=_clean(row.get('CLASSIF TECH', '')),
                    type_site=_clean(row.get('TYPE SITE', '')),
                    numero_agent=_clean(row.get('NUMERO AGENT', '')),
                    societe_gardiens=_clean(row.get('Société GARDIENS ', '')),
                    contacts_surveillants=_clean(row.get('Contacts des surveillants', '')),
                ),
            )
            if is_new:
                created += 1
            else:
                updated += 1

        self.stdout.write(self.style.SUCCESS(
            f'Done — {created} créés, {updated} mis à jour.'
        ))
