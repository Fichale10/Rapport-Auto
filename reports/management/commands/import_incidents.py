"""Import du fichier multi-feuilles BASES DES INCIDENTS vers le modèle Incident."""
import pandas as pd
from datetime import date
from django.core.management.base import BaseCommand
from django.utils import timezone
from reports.models import Incident


def _clean(val, default=''):
    if val is None:
        return default
    if pd.isna(val) if not isinstance(val, str) else False:
        return default
    s = str(val).strip()
    return default if s.lower() in ('nan', 'none', 'n/a', 'na') else s


def _parse_dt(val):
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return None
    try:
        ts = pd.to_datetime(str(val), dayfirst=True, errors='coerce')
        if ts is pd.NaT or pd.isna(ts):
            return None
        naive = ts.to_pydatetime()
        return timezone.make_aware(naive, timezone.get_current_timezone())
    except Exception:
        return None


def _duration_sec(alarm, cancel):
    if alarm is None or cancel is None:
        return None
    try:
        delta = (cancel - alarm).total_seconds()
        return delta if delta >= 0 else None
    except Exception:
        return None


def _read_sheet(path, sheet, header_row):
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df = df.dropna(how='all')
    # Nettoyer les noms de colonnes
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _mois_from_df(df, col_date):
    """Déduit le mois du rapport depuis la première date valide."""
    for val in df[col_date]:
        dt = _parse_dt(val)
        if dt is not None and not pd.isna(dt):
            return date(dt.year, dt.month, 1)
    return None


# ── Parsers par domaine ──────────────────────────────────────────────────────

def parse_mobile(path, source_file):
    df = _read_sheet(path, 'Reseau mobile ', header_row=3)
    # Supprimer la ligne 0 qui contient parfois les noms de colonnes
    if 'Ingénieur NOC' in str(df.iloc[0].tolist()):
        df = df.iloc[1:].reset_index(drop=True)

    col_map = {
        'Numero du ticket':           'numero_ticket',
        "Nature de l'incident":       'nature',
        'Alarm Time':                 'alarm_time',
        'Cancel Time':                'cancel_time',
        'Site Parent':                'site_parent',
        'Site Name':                  'site_name',
        'Site ID':                    'site_id',
        'Région':                     'region',
        'Base':                       'base',
        'Impact - Equipement':        'impact_equipement',
        'Impact - Service':           'impact_service',
        'Plateforme':                 'plateforme',
        'Technologies':               'technologies',
        'Escalade':                   'escalade',
        'Cause':                      'cause',
        'Root Cause':                 'root_cause',
        'Action':                     'action',
        'Technicien Informé':         'technicien_informe',
        'Technicien de maintenance':  'technicien_maint',
        'Point bloquant':             'point_bloquant',
        'Observation':                'observation',
        'Status':                     'status',
    }
    df = df.rename(columns={c: v for c, v in col_map.items() if c in df.columns})
    mois = _mois_from_df(df, 'alarm_time') if 'alarm_time' in df.columns else None
    return _rows_to_incidents(df, 'mobile', mois, source_file)


def parse_dr2(path, source_file):
    df = _read_sheet(path, 'DR2', header_row=1)
    col_map = {
        'Numero ticket':    'numero_ticket',
        'DATE DR2 ':        'alarm_time',
        'DATE DR2':         'alarm_time',
        'SITE PARENT':      'site_parent',
        'Site Name':        'site_name',
        'Site ID':          'site_id',
        'REGION':           'region',
        'BASE':             'base',
        'Alarm Time':       'cancel_time',   # DR2: "Alarm Time" = début panne
        'DUREE':            '_duree_str',
        'CATEGORIE':        'escalade',
        'CAUSE':            'cause',
        'POINT BLOQUANTS':  'point_bloquant',
        'Cancel Time':      'cancel_time',
        'OBSERVATION':      'observation',
    }
    df = df.rename(columns={c: v for c, v in col_map.items() if c in df.columns})
    # Dans DR2 "Alarm Time" dans le fichier = debut de la panne
    # on renomme manuellement
    raw = pd.read_excel(path, sheet_name='DR2', header=1)
    raw.columns = [str(c).strip() for c in raw.columns]
    raw = raw.dropna(how='all')
    col_map2 = {
        'Numero ticket':   'numero_ticket',
        'DATE DR2 ':       'date_dr2',
        'DATE DR2':        'date_dr2',
        'SITE PARENT':     'site_parent',
        'Site Name':       'site_name',
        'Site ID':         'site_id',
        'REGION':          'region',
        'BASE':            'base',
        'Alarm Time':      'alarm_time',
        'DUREE':           '_duree_str',
        'CATEGORIE':       'escalade',
        'CAUSE':           'cause',
        'POINT BLOQUANTS': 'point_bloquant',
        'Cancel Time':     'cancel_time',
        'OBSERVATION':     'observation',
    }
    raw = raw.rename(columns={c: v for c, v in col_map2.items() if c in raw.columns})
    mois = _mois_from_df(raw, 'alarm_time') if 'alarm_time' in raw.columns else None
    return _rows_to_incidents(raw, 'dr2', mois, source_file)


def parse_fixe(path, source_file):
    df = _read_sheet(path, 'Reseau Fixe', header_row=4)
    col_map = {
        "Nature de l'incident":       'nature',
        'Alarm Time':                 'alarm_time',
        'Cancel Time':                'cancel_time',
        'Site Name':                  'site_name',
        'Plateforme':                 'plateforme',
        'Impact - Equipement':        'impact_equipement',
        'Impact - Service':           'impact_service',
        'Nbre de client Impactés':    'nbre_clients',
        'Escalade':                   'escalade',
        'Root Cause':                 'root_cause',
        'Action':                     'action',
        'Technicien de maintenance':  'technicien_maint',
        'Status':                     'status',
        'Commentaire':                'observation',
    }
    df = df.rename(columns={c: v for c, v in col_map.items() if c in df.columns})
    mois = _mois_from_df(df, 'alarm_time') if 'alarm_time' in df.columns else None
    return _rows_to_incidents(df, 'fixe', mois, source_file)


def parse_transport(path, source_file):
    df = _read_sheet(path, 'Transport', header_row=1)
    col_map = {
        'Numero du ticket':           'numero_ticket',
        "Nature de l'incident":       'nature',
        'Alarm Time':                 'alarm_time',
        'Cancel Time':                'cancel_time',
        'Site Parent':                'site_parent',
        'Site Name':                  'site_name',
        'Site ID':                    'site_id',
        'Région':                     'region',
        'Impact - Equipement':        'impact_equipement',
        'Impact - Service':           'impact_service',
        'Plateforme':                 'plateforme',
        'Technologies':               'technologies',
        'Cause':                      'cause',
        'Escalade':                   'escalade',
        'Technicien Informé':         'technicien_informe',
        'Action':                     'action',
        'Technicien de maintenance':  'technicien_maint',
        'Root Cause':                 'root_cause',
        'Observation':                'observation',
        'Point bloquant':             'point_bloquant',
        'Status':                     'status',
    }
    df = df.rename(columns={c: v for c, v in col_map.items() if c in df.columns})
    mois = _mois_from_df(df, 'alarm_time') if 'alarm_time' in df.columns else None
    return _rows_to_incidents(df, 'transport', mois, source_file)


def parse_igw(path, source_file):
    df = _read_sheet(path, 'IGW', header_row=4)
    col_map = {
        'ALARM TIME':                'alarm_time',
        "NATURE DE L'INCIDENT ":    'nature',
        "NATURE DE L'INCIDENT":     'nature',
        'LIEN':                     'site_name',
        'LIEN INTERNET':            'plateforme',
        'IMPACTS ':                 'impact_service',
        'IMPACTS':                  'impact_service',
        'ESCALADE':                 'escalade',
        "CAUSES DE L'INCIDENT ":    'cause',
        "CAUSES DE L'INCIDENT":     'cause',
        'CANCEL TIME':              'cancel_time',
        'ACTIONS DONE':             'action',
        'STATUS':                   'status',
        'OBSERVATIONS':             'observation',
    }
    df = df.rename(columns={c: v for c, v in col_map.items() if c in df.columns})
    mois = _mois_from_df(df, 'alarm_time') if 'alarm_time' in df.columns else None
    return _rows_to_incidents(df, 'igw', mois, source_file)


def parse_core(path, source_file):
    df = _read_sheet(path, 'Core', header_row=4)
    col_map = {
        "Nature de l'incident":      'nature',
        'Alarm Time':                'alarm_time',
        'Cancel Time':               'cancel_time',
        'ESPC':                      'site_name',
        'Impact - Service':          'impact_service',
        'Escalade':                  'escalade',
        'Action':                    'action',
        'Technicien Informé':        'technicien_informe',
        'Root Cause':                'root_cause',
        'Status':                    'status',
        'Commentaire ':              'observation',
        'Commentaire':               'observation',
    }
    df = df.rename(columns={c: v for c, v in col_map.items() if c in df.columns})
    mois = _mois_from_df(df, 'alarm_time') if 'alarm_time' in df.columns else None
    return _rows_to_incidents(df, 'core', mois, source_file)


# ── Convertisseur générique ──────────────────────────────────────────────────

INCIDENT_FIELDS = [
    'numero_ticket', 'nature', 'alarm_time', 'cancel_time', 'site_parent',
    'site_name', 'site_id', 'region', 'base', 'plateforme', 'technologies',
    'impact_equipement', 'impact_service', 'escalade', 'cause', 'root_cause',
    'action', 'technicien_informe', 'technicien_maint', 'point_bloquant',
    'observation', 'status', 'nbre_clients',
]


def _rows_to_incidents(df, domain, mois, source_file):
    incidents = []
    for _, row in df.iterrows():
        alarm  = _parse_dt(row.get('alarm_time'))
        cancel = _parse_dt(row.get('cancel_time'))
        if alarm is not None and pd.isna(alarm):
            alarm = None
        if cancel is not None and pd.isna(cancel):
            cancel = None

        kwargs = {
            'domain':      domain,
            'mois_rapport': mois,
            'source_file': source_file,
            'alarm_time':  alarm,
            'cancel_time': cancel,
            'duration_sec': _duration_sec(alarm, cancel),
        }
        for f in INCIDENT_FIELDS:
            if f in ('alarm_time', 'cancel_time'):
                continue
            val = row.get(f)
            kwargs[f] = _clean(val)

        # Ignorer les lignes sans aucune info utile
        if not kwargs.get('nature') and not kwargs.get('site_name') and not kwargs.get('alarm_time'):
            continue

        incidents.append(Incident(**kwargs))
    return incidents


# ── Commande Django ──────────────────────────────────────────────────────────

PARSERS = {
    'mobile':    parse_mobile,
    'dr2':       parse_dr2,
    'fixe':      parse_fixe,
    'transport': parse_transport,
    'igw':       parse_igw,
    'core':      parse_core,
}


class Command(BaseCommand):
    help = 'Importe le fichier multi-feuilles BASES DES INCIDENTS vers Incident'

    def add_arguments(self, parser):
        parser.add_argument('filepath', type=str, help='Chemin vers le fichier Excel')
        parser.add_argument(
            '--domain', nargs='+',
            choices=list(PARSERS.keys()),
            default=list(PARSERS.keys()),
            help='Domaines à importer (défaut: tous)',
        )
        parser.add_argument(
            '--clear-mois', action='store_true',
            help='Supprimer les incidents existants du même mois avant import',
        )

    def handle(self, *args, **options):
        path = options['filepath']
        domains = options['domain']
        source = path.split('\\')[-1].split('/')[-1]

        self.stdout.write(f'Fichier : {path}')

        total_created = 0
        for domain in domains:
            self.stdout.write(f'  -> {domain} ...', ending=' ')
            try:
                incidents = PARSERS[domain](path, source)
            except Exception as e:
                self.stdout.write(self.style.ERROR(f'ERREUR: {e}'))
                continue

            if not incidents:
                self.stdout.write('0 lignes')
                continue

            mois = incidents[0].mois_rapport

            if options['clear_mois'] and mois:
                deleted, _ = Incident.objects.filter(domain=domain, mois_rapport=mois).delete()
                self.stdout.write(f'(supprimé {deleted}) ', ending='')

            Incident.objects.bulk_create(incidents, batch_size=500, ignore_conflicts=False)
            total_created += len(incidents)
            self.stdout.write(self.style.SUCCESS(f'{len(incidents)} incidents ({mois})'))

        self.stdout.write(self.style.SUCCESS(f'\nTotal : {total_created} incidents importés.'))
