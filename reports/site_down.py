"""Traitement automatique « SITE DOWN » (micro-coupures réseau mobile).

Version intégrée du script autonome ``site_down_auto.py`` :

- Les fichiers d'alarmes NetAct proviennent soit du partage réseau ISOC
  (collecte automatique planifiée), soit d'un upload manuel via la page web.
- Les régions sont lues depuis le modèle :class:`reports.models.Site`
  (plus besoin de ``BASE TECH.csv``).
- Les Cause / Escalade sont lues depuis le modèle
  :class:`reports.models.Incident` (domaine mobile) — plus besoin de relire
  les rapports journaliers Excel.
- Les alarmes traitées sont stockées dans :class:`reports.models.SiteDownAlarm`
  (contrainte unique site + alarm_time → pas de doublons).
- Un fichier Excel consolidé ``SITE_DOWN_AAAA-MM.xlsx`` est produit par mois
  dans ``MEDIA_ROOT/site_down/traites/`` (onglet Données + onglet Cumul).
"""

import logging
import os
import re
import shutil
from datetime import datetime, time

import pandas as pd
from django.conf import settings
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Répertoires de travail (sous MEDIA_ROOT)
# ──────────────────────────────────────────────────────────────────────────────
def _base_dir():
    return os.path.join(settings.MEDIA_ROOT, 'site_down')

def folder_a_traiter():
    return os.path.join(_base_dir(), 'a_traiter')

def folder_traites():
    return os.path.join(_base_dir(), 'traites')

def folder_erreurs():
    return os.path.join(_base_dir(), 'erreurs')

def _journal_alarmes():
    return os.path.join(_base_dir(), '.journal_alarmes.txt')

def ensure_dirs():
    for d in (folder_a_traiter(), folder_traites(), folder_erreurs()):
        os.makedirs(d, exist_ok=True)


# ──────────────────────────────────────────────────────────────────────────────
# Constantes
# ──────────────────────────────────────────────────────────────────────────────
ALARM_FILTER = 'WCDMA BASE STATION OUT OF USE'

NOMS_MOIS_COURTS = {
    '01': 'janv', '02': 'févr', '03': 'mars', '04': 'avr',
    '05': 'mai',  '06': 'juin', '07': 'juil', '08': 'août',
    '09': 'sept', '10': 'oct',  '11': 'nov',  '12': 'déc',
}
NOMS_MOIS_COMPLETS = {
    '01': 'Janvier', '02': 'Février', '03': 'Mars',      '04': 'Avril',
    '05': 'Mai',     '06': 'Juin',    '07': 'Juillet',   '08': 'Août',
    '09': 'Septembre', '10': 'Octobre', '11': 'Novembre', '12': 'Décembre',
}
NOMS_MOIS_DOSSIER = {
    1: 'JANVIER', 2: 'FEVRIER', 3: 'MARS', 4: 'AVRIL',
    5: 'MAI', 6: 'JUIN', 7: 'JUILLET', 8: 'AOUT',
    9: 'SEPTEMBRE', 10: 'OCTOBRE', 11: 'NOVEMBRE', 12: 'DECEMBRE',
}

JAUNE_HEADER     = 'FFFF00'
NOIR_TEXTE       = '000000'
GRIS_LIGNE_PAIRE = 'F2F2F2'
VERT_TOTAL       = '92D050'
ORANGE_TOTAL     = 'FFA500'
ROUGE_TOTAL      = 'FF4040'

BORDER_THIN = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'),  bottom=Side(style='thin'),
)


# ──────────────────────────────────────────────────────────────────────────────
# Collecte réseau (facultative — dépend de SITE_DOWN_NETWORK_BASES)
# ──────────────────────────────────────────────────────────────────────────────
def _network_bases():
    """Liste des racines réseau candidates (ex: \\\\10.228.15.80\\isoc)."""
    return [b for b in getattr(settings, 'SITE_DOWN_NETWORK_BASES', []) if b]


def _source_alarmes_candidates():
    now = datetime.now()
    mois, annee = NOMS_MOIS_DOSSIER[now.month], now.year
    return [
        os.path.join(base, 'RAPPORT RESEAU MOBILE', 'ALARME NETACT SITES DOWN',
                     f'SITE DOWN {mois} {annee}')
        for base in _network_bases()
    ]


def _lire_journal(path):
    if not os.path.exists(path):
        return set()
    with open(path, 'r', encoding='utf-8') as f:
        return {l.strip() for l in f if l.strip()}


def _ecrire_journal(path, noms):
    with open(path, 'a', encoding='utf-8') as f:
        for nom in noms:
            f.write(nom + '\n')


def extraire_mois_annee(filename):
    """`... 05-07-2026 ...` → ``2026-07`` (None si pas de date dans le nom)."""
    m = re.search(r'(\d{2})-(\d{2})-(\d{4})', filename)
    return f"{m.group(3)}-{m.group(2)}" if m else None


def _dates_deja_traitees_cache():
    """Dates (YYYY-MM-DD) déjà présentes dans les fichiers consolidés, par mois."""
    cache = {}

    def get(mois_annee):
        if mois_annee in cache:
            return cache[mois_annee]
        path = os.path.join(folder_traites(), f'SITE_DOWN_{mois_annee}.xlsx')
        dates = set()
        if os.path.exists(path):
            try:
                df = pd.read_excel(path, sheet_name='Données', usecols=['Alarm Time'])
                df['Alarm Time'] = pd.to_datetime(df['Alarm Time'], errors='coerce')
                dates = set(df['Alarm Time'].dt.strftime('%Y-%m-%d').dropna().unique())
            except Exception:
                logger.warning("site_down : lecture dates existantes %s échouée", path,
                               exc_info=True)
        cache[mois_annee] = dates
        return dates

    return get


def collecter_alarmes():
    """Copie les nouveaux fichiers d'alarmes du partage réseau vers ``a_traiter``.

    Returns:
        int: nombre de fichiers copiés (0 si le réseau est inaccessible).
    """
    ensure_dirs()
    source = next((s for s in _source_alarmes_candidates() if os.path.exists(s)), None)
    if source is None:
        logger.warning("site_down : source alarmes inaccessible (%s)",
                       _source_alarmes_candidates())
        return 0

    deja_copies = _lire_journal(_journal_alarmes())
    dates_traitees = _dates_deja_traitees_cache()

    fichiers = [
        f for f in os.listdir(source)
        if os.path.isfile(os.path.join(source, f))
        and f.lower().endswith(('.xlsx', '.xls', '.csv'))
        and not f.startswith(('~$', '.'))
    ]

    candidats = []
    for f in sorted(fichiers):
        ma = extraire_mois_annee(f)
        m = re.search(r'(\d{2})-(\d{2})-(\d{4})', f)
        date_fichier = f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None
        if ma and date_fichier:
            if date_fichier not in dates_traitees(ma):
                candidats.append(f)
        elif f not in deja_copies:
            candidats.append(f)

    copies = []
    for nom in candidats:
        try:
            shutil.copy2(os.path.join(source, nom), os.path.join(folder_a_traiter(), nom))
            copies.append(nom)
        except Exception:
            logger.exception("site_down : erreur copie %s", nom)

    if copies:
        _ecrire_journal(_journal_alarmes(), copies)
        logger.info("site_down : %d fichier(s) alarme copié(s)", len(copies))
    return len(copies)


# ──────────────────────────────────────────────────────────────────────────────
# Mappings depuis la base Django
# ──────────────────────────────────────────────────────────────────────────────
def charger_regions_map():
    """SITE NAME (majuscules) → région, depuis le modèle Site."""
    from .models import Site
    return {
        (name or '').strip().upper(): (region or '').strip()
        for name, region in Site.objects.exclude(region='').values_list('site_name', 'region')
    }


def charger_causes_escalades_map(mois_annee=None):
    """(site, alarm_time arrondi à la minute) → {Cause, Escalade}.

    Construit depuis le modèle Incident (domaine mobile), au lieu de relire
    les rapports journaliers Excel.
    """
    from .models import Incident
    qs = Incident.objects.filter(domain=Incident.DOMAIN_MOBILE).exclude(alarm_time=None)
    if mois_annee:
        annee, mois = mois_annee.split('-')
        qs = qs.filter(alarm_time__year=int(annee), alarm_time__month=int(mois))

    mapping = {}
    for site, alarm_time, cause, escalade in qs.values_list(
            'site_name', 'alarm_time', 'cause', 'escalade'):
        if not site:
            continue
        at = alarm_time
        if at.tzinfo is not None:
            at = at.astimezone().replace(tzinfo=None)
        key = (site.strip(), at.replace(second=0, microsecond=0))
        mapping[key] = {'Cause': cause or '', 'Escalade': escalade or ''}
    return mapping


def ajouter_cause_escalade(df, mapping):
    causes, escalades = [], []
    for _, row in df.iterrows():
        cause = escalade = ''
        at = row['Alarm Time']
        if pd.notna(at):
            key = (str(row['Name']).strip(),
                   at.to_pydatetime().replace(second=0, microsecond=0))
            entry = mapping.get(key)
            if entry:
                cause, escalade = entry['Cause'], entry['Escalade']
        causes.append(cause)
        escalades.append(escalade)
    df['Cause'] = causes
    df['Escalade'] = escalades
    return df


# ──────────────────────────────────────────────────────────────────────────────
# Lecture d'un fichier d'alarmes brut
# ──────────────────────────────────────────────────────────────────────────────
def _float_to_duration(val):
    if pd.isna(val) or val == 0:
        return '00:00:00'
    try:
        total = float(val) * 24 * 3600
        return f"{int(total // 3600):02d}:{int((total % 3600) // 60):02d}:{int(total % 60):02d}"
    except (TypeError, ValueError):
        return str(val)


def traiter_fichier(file_path):
    """Lit un fichier d'alarmes (xlsx/xls/csv) et normalise ses colonnes.

    Returns:
        DataFrame avec colonnes Name, Alarm Time, Cancel Time, Duration, Alarm Text.
    """
    ext = file_path.lower().rsplit('.', 1)[-1]
    if ext in ('xlsx', 'xls'):
        df_raw = pd.read_excel(file_path, header=None)
        header_row = 0
        for i, row in df_raw.iterrows():
            vals = ' '.join(str(v).strip().lower() for v in row.tolist())
            if any(k in vals for k in ('alarm time', 'cancel time', 'name', 'site')):
                header_row = i
                break
        df = pd.read_excel(file_path, header=header_row)
    elif ext == 'csv':
        df = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8-sig')
    else:
        raise ValueError(f"Format non supporté : {ext}")

    df.columns = df.columns.str.strip().str.replace('\ufeff', '')

    mapping = {}
    for col in df.columns:
        cl = col.lower().strip()
        if cl == 'site' or cl == 'site name':
            mapping['site_name'] = col
        elif cl == 'name':
            mapping.setdefault('site_name', col)
        elif 'alarm' in cl and 'text' in cl:
            mapping['alarm_text'] = col
        elif 'alarm' in cl and 'time' in cl:
            mapping['alarm_time'] = col
        elif 'cancel' in cl and 'time' in cl:
            mapping['cancel_time'] = col
        elif 'duration' in cl or 'durée' in cl or cl == 'd':
            mapping.setdefault('duration', col)

    missing = [c for c in ('site_name', 'alarm_time', 'cancel_time', 'duration')
               if c not in mapping]
    if missing:
        raise ValueError(f"Colonnes manquantes : {missing}")

    cols      = [mapping['site_name'], mapping['alarm_time'],
                 mapping['cancel_time'], mapping['duration']]
    col_names = ['Name', 'Alarm Time', 'Cancel Time', 'Duration']
    if 'alarm_text' in mapping:
        cols.append(mapping['alarm_text'])
        col_names.append('Alarm Text')

    df = df[cols].copy()
    df.columns = col_names
    if 'Alarm Text' not in df.columns:
        df['Alarm Text'] = ''
    df = df.dropna(subset=['Name'])
    df['Alarm Time']  = pd.to_datetime(df['Alarm Time'],  errors='coerce', dayfirst=True)
    df['Cancel Time'] = pd.to_datetime(df['Cancel Time'], errors='coerce', dayfirst=True)

    if df['Duration'].dtype == 'float64':
        df['Duration'] = df['Duration'].apply(_float_to_duration)
    elif pd.api.types.is_timedelta64_dtype(df['Duration']):
        df['Duration'] = df['Duration'].apply(
            lambda x: str(x).split(' days ')[-1] if pd.notna(x) else '')
    elif df['Duration'].apply(lambda x: isinstance(x, time)).any():
        df['Duration'] = df['Duration'].apply(
            lambda x: x.strftime('%H:%M:%S') if isinstance(x, time) else str(x))
    return df


def dur_to_min(val):
    if pd.isna(val) or val in ('', '00:00:00'):
        return 0
    try:
        parts = str(val).split(':')
        if len(parts) == 3:
            return int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60
    except (TypeError, ValueError):
        pass
    return 0


def _clean_str(val):
    s = str(val).strip() if pd.notna(val) else ''
    return '' if s.lower() in ('nan', 'none', '') else s


# ──────────────────────────────────────────────────────────────────────────────
# Feuille Cumul (pivot sites × jours)
# ──────────────────────────────────────────────────────────────────────────────
def _duree_reelle_par_site_jour(df_site):
    """Durée réelle (min) par jour pour un site : découpe par minuit puis
    fusionne les plages chevauchantes pour éviter le double comptage."""
    plages_par_jour = {}
    for _, row in df_site.iterrows():
        alarm_dt, cancel_dt = row['Alarm Time'], row['Cancel Time']
        if pd.isna(alarm_dt) or pd.isna(cancel_dt) or cancel_dt <= alarm_dt:
            continue
        current = alarm_dt
        while current < cancel_dt:
            minuit_suivant = pd.Timestamp(current.date()) + pd.Timedelta(days=1)
            fin_tranche = min(cancel_dt, minuit_suivant)
            if fin_tranche > current:
                plages_par_jour.setdefault(current.day, []).append((current, fin_tranche))
            current = minuit_suivant

    result = {}
    for jour, plages in plages_par_jour.items():
        merged = []
        for start, end in sorted(plages, key=lambda x: x[0]):
            if merged and start <= merged[-1][1]:
                merged[-1][1] = max(merged[-1][1], end)
            else:
                merged.append([start, end])
        result[jour] = int(round(sum((e - s).total_seconds() / 60 for s, e in merged)))
    return result


def creer_feuille_cumul(df_consolide, mois_annee, regions_map=None):
    base_cols = ['Name', 'Alarm Time', 'Cancel Time', 'Duration', 'Cause', 'Escalade']
    if 'Alarm Text' in df_consolide.columns:
        base_cols.append('Alarm Text')
    df = df_consolide[base_cols].copy()
    if 'Alarm Text' not in df.columns:
        df['Alarm Text'] = ''

    # Si Alarm Text absent, toutes les lignes comptent dans Nb
    if not (df['Alarm Text'].notna().any() and (df['Alarm Text'].astype(str).str.strip() != '').any()):
        df['Alarm Text'] = ALARM_FILTER

    df['Alarm Time']  = pd.to_datetime(df['Alarm Time'],  errors='coerce')
    df['Cancel Time'] = pd.to_datetime(df['Cancel Time'], errors='coerce')
    df = df.sort_values('Alarm Time')

    rows_nb = [
        {'Name': row['Name'], 'jour': row['Alarm Time'].day,
         'is_wcdma': ALARM_FILTER in str(row['Alarm Text']).strip().upper()}
        for _, row in df.iterrows() if pd.notna(row['Alarm Time'])
    ]
    df_nb = pd.DataFrame(rows_nb, columns=['Name', 'jour', 'is_wcdma'])
    df_wcdma = df_nb[df_nb['is_wcdma']] if len(df_nb) else df_nb

    all_sites = df['Name'].unique()
    durees_par_site, all_jours_set = {}, set()
    for site in all_sites:
        d = _duree_reelle_par_site_jour(df[df['Name'] == site])
        durees_par_site[site] = d
        all_jours_set.update(d.keys())
    all_jours = sorted(all_jours_set)

    if len(df_wcdma) > 0:
        pivot_count = df_wcdma.pivot_table(index='Name', columns='jour',
                                           values='is_wcdma', aggfunc='count', fill_value=0)
    else:
        pivot_count = pd.DataFrame()
    pivot_count = pivot_count.reindex(index=all_sites, columns=all_jours, fill_value=0)

    pivot_duree = pd.DataFrame(
        [{**{'Name': s}, **{j: durees_par_site[s].get(j, 0) for j in all_jours}}
         for s in all_sites]
    ).set_index('Name') if len(all_sites) else pd.DataFrame()
    if len(pivot_duree):
        pivot_duree = pivot_duree.reindex(index=all_sites, columns=all_jours, fill_value=0)

    _, mois_num = mois_annee.split('-')
    nom_mois = NOMS_MOIS_COURTS.get(mois_num, mois_num)

    result = pd.DataFrame()
    result['Site'] = list(all_sites)
    result['Région'] = result['Site'].apply(
        lambda s: regions_map.get(str(s).strip().upper(), '') if regions_map else '')

    for jour in all_jours:
        prefix = f"{int(jour)}-{nom_mois}"
        nb_vals, dur_vals = [], []
        for site in all_sites:
            count = int(pivot_count.loc[site, jour]) if site in pivot_count.index else 0
            duree = int(round(pivot_duree.loc[site, jour])) if site in pivot_duree.index else 0
            nb_vals.append(count or None)
            dur_vals.append(duree or None)
        result[f"{prefix} Nb"] = nb_vals
        result[f"{prefix} Durée"] = dur_vals

    total_nb, total_dur, col_cause, col_escalade = [], [], [], []
    for site in all_sites:
        t_count = int(pivot_count.loc[site].sum()) if site in pivot_count.index else 0
        t_duree = int(round(pivot_duree.loc[site].sum())) if site in pivot_duree.index else 0
        site_data = df[df['Name'] == site].sort_values('Alarm Time')
        derniere_cause    = _clean_str(site_data.iloc[-1]['Cause'])    if len(site_data) else ''
        derniere_escalade = _clean_str(site_data.iloc[-1]['Escalade']) if len(site_data) else ''
        total_nb.append(t_count or None)
        total_dur.append(t_duree or None)
        col_cause.append(derniere_cause or None)
        col_escalade.append(derniere_escalade or None)

    result['Total Nb']    = total_nb
    result['Total Durée'] = total_dur
    result['Cause']       = col_cause
    result['Escalade']    = col_escalade

    result['_sort'] = result['Total Nb'].fillna(0)
    return (result.sort_values('_sort', ascending=False)
                  .drop(columns=['_sort']).reset_index(drop=True))


# ──────────────────────────────────────────────────────────────────────────────
# Formatage Excel
# ──────────────────────────────────────────────────────────────────────────────
def _style_header(cell):
    cell.font      = Font(name='Arial', bold=True, color=NOIR_TEXTE, size=10)
    cell.fill      = PatternFill('solid', start_color=JAUNE_HEADER)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border    = BORDER_THIN


def _style_cell(cell, row_index, align='left'):
    bg             = GRIS_LIGNE_PAIRE if row_index % 2 == 0 else 'FFFFFF'
    cell.font      = Font(name='Arial', size=9)
    cell.fill      = PatternFill('solid', start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical='center')
    cell.border    = BORDER_THIN


def formater_feuille(ws, df, sheet_name):
    col_widths = ({'Name': 28, 'Alarm Time': 22, 'Cancel Time': 22,
                   'Duration': 14, 'Cause': 25, 'Escalade': 18}
                  if sheet_name == 'Données' else {})

    for col_idx in range(1, len(df.columns) + 1):
        _style_header(ws.cell(row=1, column=col_idx))
    ws.row_dimensions[1].height = 30

    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            col_name = df.columns[col_idx - 1] if col_idx <= len(df.columns) else ''
            align = 'left' if (col_idx == 1 or 'Cause' in col_name or 'Escalade' in col_name) else 'center'
            _style_cell(ws.cell(row=row_idx, column=col_idx), row_idx, align=align)
        ws.row_dimensions[row_idx].height = 16

    for col_idx, col_name in enumerate(df.columns, start=1):
        letter = get_column_letter(col_idx)
        if col_name in col_widths:
            ws.column_dimensions[letter].width = col_widths[col_name]
        elif sheet_name == 'Données':
            ws.column_dimensions[letter].width = 15
        elif col_idx == 1:
            ws.column_dimensions[letter].width = 30
        elif col_name == 'Région':
            ws.column_dimensions[letter].width = 16
        elif 'Cause' in col_name or 'Escalade' in col_name:
            ws.column_dimensions[letter].width = 22
        elif 'Nb' in col_name:
            ws.column_dimensions[letter].width = 8
        elif 'Durée' in col_name:
            ws.column_dimensions[letter].width = 10
        else:
            ws.column_dimensions[letter].width = 14

    safe_name = re.sub(r'[^A-Za-z0-9_]', '_', sheet_name)
    table = Table(displayName=f"Tbl_{safe_name}",
                  ref=f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}")
    table.tableStyleInfo = TableStyleInfo(
        name='TableStyleMedium2', showFirstColumn=False,
        showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)
    ws.freeze_panes = 'B2'

    if sheet_name == 'Données':
        for col_idx, col_name in enumerate(df.columns, start=1):
            if col_name == 'Alarm Text':
                ws.column_dimensions[get_column_letter(col_idx)].hidden = True
                break


def colorier_totaux_cumul(ws, df):
    if 'Total Nb' not in df.columns:
        return

    jours = []
    for col in df.columns:
        m = re.match(r'^(\d+)-\w+ Nb$', col)
        if m:
            jours.append((int(m.group(1)), col))
    jours.sort(key=lambda x: x[0])
    if not jours:
        return

    dernier_jour = max(j for j, _ in jours)
    cols_rouge  = [c for j, c in jours if dernier_jour - 3 <= j <= dernier_jour]
    cols_orange = [c for j, c in jours if dernier_jour - 2 <= j <= dernier_jour]

    col_total_nb  = next((i for i, c in enumerate(df.columns, 1) if c == 'Total Nb'), None)
    col_total_dur = next((i for i, c in enumerate(df.columns, 1) if c == 'Total Durée'), None)
    if col_total_nb is None:
        return

    def tous_en_panne(row, cols):
        return bool(cols) and all(pd.notna(row.get(c)) and row.get(c, 0) > 0 for c in cols)

    def aucune_panne(row, cols):
        return all(not (pd.notna(row.get(c)) and row.get(c, 0) > 0) for c in cols)

    for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        if len(cols_rouge) == 4 and tous_en_panne(row, cols_rouge):
            couleur, font_c = ROUGE_TOTAL, Font(name='Arial', size=9, bold=True, color='FFFFFF')
        elif len(cols_orange) == 3 and tous_en_panne(row, cols_orange):
            couleur, font_c = ORANGE_TOTAL, Font(name='Arial', size=9, bold=True, color='000000')
        elif aucune_panne(row, cols_orange):
            couleur, font_c = VERT_TOTAL, Font(name='Arial', size=9, bold=True, color='000000')
        else:
            couleur, font_c = None, Font(name='Arial', size=9)

        for col_idx in (col_total_nb, col_total_dur):
            if col_idx:
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.fill = PatternFill('solid', start_color=couleur) if couleur else PatternFill()
                cell.font = font_c
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = BORDER_THIN


# ──────────────────────────────────────────────────────────────────────────────
# Persistance ORM
# ──────────────────────────────────────────────────────────────────────────────
def _sauvegarder_orm(df, source_file, regions_map):
    """Upsert des alarmes dans SiteDownAlarm (clé : site + alarm_time)."""
    from django.utils import timezone as tz
    from .models import SiteDownAlarm

    current_tz = tz.get_current_timezone()

    def aware(ts):
        if pd.isna(ts):
            return None
        dt = ts.to_pydatetime()
        return tz.make_aware(dt, current_tz) if dt.tzinfo is None else dt

    created = updated = 0
    for _, row in df.iterrows():
        alarm_time = aware(row['Alarm Time'])
        if alarm_time is None:
            continue
        site = str(row['Name']).strip()
        defaults = {
            'cancel_time':  aware(row['Cancel Time']),
            'duration_min': dur_to_min(row.get('Duration')) or None,
            'alarm_text':   _clean_str(row.get('Alarm Text', ''))[:255],
            'cause':        _clean_str(row.get('Cause', '')),
            'escalade':     _clean_str(row.get('Escalade', ''))[:80],
            'region':       (regions_map or {}).get(site.upper(), '')[:50],
            'source_file':  source_file[:255],
        }
        _, was_created = SiteDownAlarm.objects.update_or_create(
            site_name=site, alarm_time=alarm_time, defaults=defaults)
        created += was_created
        updated += (not was_created)
    return created, updated


# ──────────────────────────────────────────────────────────────────────────────
# Pipeline principal
# ──────────────────────────────────────────────────────────────────────────────
def fichiers_mensuels():
    """Liste [(nom, chemin, taille, mtime)] des fichiers consolidés produits."""
    ensure_dirs()
    out = []
    for f in sorted(os.listdir(folder_traites()), reverse=True):
        if re.fullmatch(r'SITE_DOWN_\d{4}-\d{2}\.xlsx', f):
            p = os.path.join(folder_traites(), f)
            st = os.stat(p)
            out.append({'name': f, 'path': p, 'size': st.st_size,
                        'mtime': datetime.fromtimestamp(st.st_mtime)})
    return out


def process_pending_files():
    """Traite tous les fichiers présents dans ``a_traiter``.

    Returns:
        dict résumé : {processed, errors, months, created, updated, messages}
    """
    ensure_dirs()
    regions_map = charger_regions_map()

    files = sorted(f for f in os.listdir(folder_a_traiter())
                   if os.path.isfile(os.path.join(folder_a_traiter(), f)))

    summary = {'processed': 0, 'errors': 0, 'months': [],
               'created': 0, 'updated': 0, 'messages': []}

    if not files:
        summary['messages'].append("Aucun nouveau fichier à traiter.")
        return summary

    fichiers_par_mois = {}
    for f in files:
        ma = extraire_mois_annee(f)
        if ma:
            fichiers_par_mois.setdefault(ma, []).append(f)
        else:
            summary['messages'].append(f"Mois indéterminé, ignoré : {f}")

    for mois_annee in sorted(fichiers_par_mois):
        causes_map = charger_causes_escalades_map(mois_annee)
        annee, mois_num = mois_annee.split('-')
        nom_feuille_cumul = f"Cumul_{NOMS_MOIS_COMPLETS.get(mois_num, mois_num)}_{annee}"
        output_file = os.path.join(folder_traites(), f'SITE_DOWN_{mois_annee}.xlsx')

        df_existant = None
        if os.path.exists(output_file):
            try:
                df_existant = pd.read_excel(output_file, sheet_name='Données')
                df_existant['Alarm Time']  = pd.to_datetime(df_existant['Alarm Time'],  errors='coerce')
                df_existant['Cancel Time'] = pd.to_datetime(df_existant['Cancel Time'], errors='coerce')
            except Exception:
                logger.warning("site_down : lecture existant %s échouée", output_file, exc_info=True)

        dfs_mois, fichiers_ok = [], []
        for file in sorted(fichiers_par_mois[mois_annee]):
            file_path = os.path.join(folder_a_traiter(), file)
            try:
                df = traiter_fichier(file_path)
                df = ajouter_cause_escalade(df, causes_map)
                dfs_mois.append(df)
                fichiers_ok.append((file, file_path))
                summary['processed'] += 1
            except Exception as exc:
                logger.exception("site_down : erreur traitement %s", file)
                summary['errors'] += 1
                summary['messages'].append(f"Erreur {file} : {exc}")
                try:
                    shutil.move(file_path, os.path.join(folder_erreurs(), file))
                except OSError:
                    pass

        if not dfs_mois:
            continue

        df_nouvelles = pd.concat(dfs_mois, ignore_index=True)

        if df_existant is not None and len(df_existant) > 0:
            df_existant = ajouter_cause_escalade(df_existant, causes_map)
            df_consolide = pd.concat([df_existant, df_nouvelles], ignore_index=True)
            df_consolide = df_consolide.drop_duplicates(subset=['Name', 'Alarm Time'], keep='last')
        else:
            df_consolide = df_nouvelles

        df_consolide = df_consolide.sort_values('Alarm Time').reset_index(drop=True)
        df_cumul = creer_feuille_cumul(df_consolide, mois_annee, regions_map)

        cols_donnees = [c for c in df_consolide.columns if c not in ('Cause', 'Escalade')]
        df_donnees = df_consolide[cols_donnees]

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_donnees.to_excel(writer, sheet_name='Données', index=False)
            df_cumul.to_excel(writer, sheet_name=nom_feuille_cumul, index=False)

        wb = load_workbook(output_file)
        formater_feuille(wb['Données'], df_donnees, 'Données')
        formater_feuille(wb[nom_feuille_cumul], df_cumul, nom_feuille_cumul)
        colorier_totaux_cumul(wb[nom_feuille_cumul], df_cumul)
        wb.save(output_file)

        # Persistance ORM (uniquement les nouvelles lignes traitées)
        created, updated = _sauvegarder_orm(
            df_nouvelles, ', '.join(f for f, _ in fichiers_ok)[:255], regions_map)
        summary['created'] += created
        summary['updated'] += updated
        summary['months'].append(mois_annee)
        summary['messages'].append(
            f"{mois_annee} : {len(df_donnees)} lignes, {len(df_cumul)} sites "
            f"({created} créées, {updated} mises à jour en base)")

        for _, fp in fichiers_ok:
            try:
                os.remove(fp)
            except OSError:
                pass

    return summary


def actualiser_fichiers_existants():
    """Ré-applique Cause/Escalade + cumul sur les fichiers mensuels existants."""
    ensure_dirs()
    regions_map = charger_regions_map()
    summary = {'refreshed': 0, 'errors': 0, 'messages': []}

    for info in fichiers_mensuels():
        m = re.search(r'SITE_DOWN_(\d{4})-(\d{2})\.xlsx', info['name'])
        if not m:
            continue
        mois_annee = f"{m.group(1)}-{m.group(2)}"
        nom_feuille_cumul = f"Cumul_{NOMS_MOIS_COMPLETS.get(m.group(2), m.group(2))}_{m.group(1)}"
        try:
            df = pd.read_excel(info['path'], sheet_name='Données')
            df['Alarm Time']  = pd.to_datetime(df['Alarm Time'],  errors='coerce')
            df['Cancel Time'] = pd.to_datetime(df['Cancel Time'], errors='coerce')
            df = ajouter_cause_escalade(df, charger_causes_escalades_map(mois_annee))
            df_cumul = creer_feuille_cumul(df, mois_annee, regions_map)

            cols_donnees = [c for c in df.columns if c not in ('Cause', 'Escalade')]
            df_donnees = df[cols_donnees]

            with pd.ExcelWriter(info['path'], engine='openpyxl') as writer:
                df_donnees.to_excel(writer, sheet_name='Données', index=False)
                df_cumul.to_excel(writer, sheet_name=nom_feuille_cumul, index=False)

            wb = load_workbook(info['path'])
            formater_feuille(wb['Données'], df_donnees, 'Données')
            formater_feuille(wb[nom_feuille_cumul], df_cumul, nom_feuille_cumul)
            colorier_totaux_cumul(wb[nom_feuille_cumul], df_cumul)
            wb.save(info['path'])
            summary['refreshed'] += 1
        except Exception as exc:
            logger.exception("site_down : actualisation %s échouée", info['name'])
            summary['errors'] += 1
            summary['messages'].append(f"Erreur {info['name']} : {exc}")

    return summary


def run_auto():
    """Point d'entrée complet : collecte réseau puis traitement (scheduler)."""
    nb = collecter_alarmes()
    summary = process_pending_files()
    summary['collected'] = nb
    if nb == 0 and summary['processed'] == 0:
        # Rien de neuf : rafraîchit quand même Cause/Escalade des fichiers du mois
        refresh = actualiser_fichiers_existants()
        summary['messages'].append(
            f"Actualisation : {refresh['refreshed']} fichier(s), {refresh['errors']} erreur(s)")
    return summary
