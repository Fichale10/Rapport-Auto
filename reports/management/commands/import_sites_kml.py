import xml.etree.ElementTree as ET

from django.core.management.base import BaseCommand, CommandError

from reports.models import Site

KML_NS = '{http://www.opengis.net/kml/2.2}'


def _clean(val):
    if val is None:
        return ''
    s = str(val).strip()
    return '' if s.lower() in ('nan', 'none', 'n/a') else s


class Command(BaseCommand):
    help = 'Import sites (nom, latitude, longitude, config) depuis un fichier KML dans la table Site'

    def add_arguments(self, parser):
        parser.add_argument('filepath', type=str, help='Chemin vers le fichier .kml')

    def handle(self, *args, **options):
        path = options['filepath']
        self.stdout.write(f'Lecture de {path}...')

        try:
            tree = ET.parse(path)
        except ET.ParseError as exc:
            raise CommandError(f'Fichier KML invalide : {exc}')

        root = tree.getroot()
        placemarks = root.iter(f'{KML_NS}Placemark')

        created = updated = skipped = 0

        for pm in placemarks:
            name_el = pm.find(f'{KML_NS}name')
            name = _clean(name_el.text if name_el is not None else '')
            if not name:
                skipped += 1
                continue

            data = {}
            ext_data = pm.find(f'{KML_NS}ExtendedData')
            if ext_data is not None:
                for sd in ext_data.iter(f'{KML_NS}SimpleData'):
                    data[sd.get('name')] = sd.text

            lat_raw = data.get('Latitude')
            lon_raw = data.get('Longitude')

            lat = lon = None
            if lat_raw is not None and lon_raw is not None:
                try:
                    lat = float(lat_raw)
                    lon = float(lon_raw)
                except ValueError:
                    pass

            # Fallback : lire <Point><coordinates>lon,lat,alt</coordinates>
            if lat is None or lon is None:
                point = pm.find(f'{KML_NS}Point')
                if point is not None:
                    coords_el = point.find(f'{KML_NS}coordinates')
                    if coords_el is not None and coords_el.text:
                        parts = coords_el.text.strip().split(',')
                        if len(parts) >= 2:
                            try:
                                lon = float(parts[0])
                                lat = float(parts[1])
                            except ValueError:
                                pass

            if lat is None or lon is None:
                skipped += 1
                continue

            site_id = _clean(data.get('ID', ''))
            config = _clean(data.get('CONFIG', ''))

            defaults = {'latitude': lat, 'longitude': lon}
            if site_id:
                defaults['site_id'] = site_id
            if config:
                defaults['config'] = config

            obj, is_new = Site.objects.update_or_create(
                site_name=name,
                defaults=defaults,
            )
            if is_new:
                created += 1
            else:
                updated += 1

        self.stdout.write(self.style.SUCCESS(
            f'Terminé — {created} créés, {updated} mis à jour, {skipped} ignorés (sans coordonnées).'
        ))
