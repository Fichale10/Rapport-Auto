# -*- coding: utf-8 -*-
"""Calcul automatique DR2 à partir des fichiers bruts de disponibilité
horaire 2G / 3G envoyés quotidiennement par mail — remplace la saisie
manuelle du fichier DR2 Excel (colonne DR2 = OUI/NON déjà renseignée).

Fichiers source (ex. « 2G_AVAILABILITY_PER_HOUR_@_DAY-...xlsx »,
« 3G_AVAILABILITY_PER_HOUR_@_DAY-...xlsx ») : une ligne par site et par
heure (colonne « Period start time »).

Règle DR2 (définie par l'utilisateur, août 2026) : un site est DR2 ssi,
pour une même heure, sa disponibilité est nulle (0 %) simultanément :
  - en 3G, colonne « Cell Availability, excluding blocked by user state (BLU) »
  - en 2G, colonne « TCH availability ratio »
et que le nombre total d'heures instables ainsi comptées (cumulées, même
non consécutives — ex. 1h + 1h + 1h) est >= 3 sur la journée.

Le rapprochement entre nom de site 2G (BCF name) et 3G (WBTS name) se fait
par normalisation/préfixe du libellé (les deux techno ne partagent pas
exactement le même nom de site).
"""
from __future__ import annotations

import logging
import re
from datetime import date, datetime

import pandas as pd

logger = logging.getLogger(__name__)

DR2_THRESHOLD_HOURS = 3

# Tokens techno / séparateurs retirés pour rapprocher un nom de site 2G (BCF)
# et son équivalent 3G (WBTS) — rapprochement par préfixe du libellé.
_TECH_TOKENS = re.compile(
    r'(2G|3G|4G|_BCF|_WBTS|_MRBTS|_BTS|_NODEB|_NODE-B|_CELL|[^A-Z0-9])',
    re.IGNORECASE,
)


def normalize_site_key(name) -> str:
    """Normalise un nom de site (BCF/WBTS/Site Name du ticket) pour
    rapprochement par préfixe : majuscule, tokens techno retirés."""
    if not name:
        return ''
    return _TECH_TOKENS.sub('', str(name).upper()).strip()


def _find_col(columns, *candidates):
    norm_cols = {re.sub(r'\s+', ' ', str(c)).strip().upper(): c for c in columns}
    for cand in candidates:
        cand_u = cand.upper()
        for norm, orig in norm_cols.items():
            if cand_u in norm:
                return orig
    return None


def _read_availability(fileobj, site_candidates, kpi_candidates):
    """Lit un fichier de disponibilité horaire → DataFrame normalisé avec
    colonnes : period, site_raw, site_key, value (float, %)."""
    df = pd.read_excel(fileobj, engine='openpyxl')
    period_col = _find_col(df.columns, 'PERIOD START TIME', 'PERIOD START', 'DATE TIME', 'DATETIME')
    site_col   = _find_col(df.columns, *site_candidates)
    kpi_col    = _find_col(df.columns, *kpi_candidates)
    missing = [n for n, c in (('période', period_col), ('site', site_col), ('KPI', kpi_col)) if not c]
    if missing:
        raise ValueError(f"Colonne(s) introuvable(s) dans le fichier : {', '.join(missing)}")

    out = df[[period_col, site_col, kpi_col]].copy()
    out.columns = ['period', 'site_raw', 'value']
    out['period'] = pd.to_datetime(out['period'], errors='coerce')
    out['site_raw'] = out['site_raw'].astype(str).str.strip()
    out['value'] = pd.to_numeric(
        out['value'].apply(lambda v: str(v).replace(',', '.') if pd.notna(v) else v),
        errors='coerce',
    )
    out = out.dropna(subset=['period', 'site_raw'])
    out = out[out['site_raw'] != '']
    out['site_key'] = out['site_raw'].apply(normalize_site_key)
    return out


def parse_2g_file(fileobj) -> pd.DataFrame:
    """Fichier 2G_AVAILABILITY_PER_HOUR : site = BCF name, KPI = TCH availability ratio."""
    return _read_availability(fileobj, ['BCF NAME', 'BCF'], ['TCH AVAILABILITY RATIO', 'TCH AVAILABILITY'])


def parse_3g_file(fileobj) -> pd.DataFrame:
    """Fichier 3G_AVAILABILITY_PER_HOUR : site = WBTS name, KPI = Cell Availability excl. BLU."""
    return _read_availability(
        fileobj, ['WBTS NAME', 'WBTS'],
        ['CELL AVAILABILITY, EXCLUDING BLOCKED BY USER STATE (BLU)',
         'CELL AVAILABILITY EXCLUDING BLOCKED', 'BLU'],
    )


def compute_dr2_sites(df_2g: pd.DataFrame, df_3g: pd.DataFrame) -> list[dict]:
    """Sites en violation DR2 : disponibilité nulle simultanément en 2G ET
    3G sur >= DR2_THRESHOLD_HOURS heures cumulées (même non consécutives).

    Retourne une liste de dicts triés par nb d'heures décroissant :
    {site_key, site_name_2g, site_name_3g, hours_down, down_hours: [datetime, …]}.
    """
    if df_2g is None or df_2g.empty or df_3g is None or df_3g.empty:
        return []

    down_2g = df_2g[df_2g['value'] == 0]
    down_3g = df_3g[df_3g['value'] == 0]
    if down_2g.empty or down_3g.empty:
        return []

    merged = down_2g.merge(down_3g, on=['site_key', 'period'], suffixes=('_2g', '_3g'), how='inner')
    if merged.empty:
        return []

    result = []
    for site_key, grp in merged.groupby('site_key'):
        if not site_key:
            continue
        hours = sorted(grp['period'].unique())
        if len(hours) < DR2_THRESHOLD_HOURS:
            continue
        result.append({
            'site_key':     site_key,
            'site_name_2g': grp['site_raw_2g'].iloc[0],
            'site_name_3g': grp['site_raw_3g'].iloc[0],
            'hours_down':   len(hours),
            'down_hours':   [pd.Timestamp(h).to_pydatetime() for h in hours],
        })
    return sorted(result, key=lambda r: -r['hours_down'])


def _match_ticket_row(site_key: str, mobile_df: pd.DataFrame | None):
    """Cherche, dans le dataframe des incidents mobile du jour (colonnes FR
    d'origine), la ligne dont le Site Name se rapproche le plus du site_key
    DR2 (normalisation identique). Retourne la 1re ligne matching ou None."""
    if mobile_df is None or mobile_df.empty or 'Site Name' not in mobile_df.columns:
        return None
    keys = mobile_df['Site Name'].astype(str).apply(normalize_site_key)
    matches = mobile_df[keys == site_key]
    if matches.empty:
        return None
    return matches.iloc[0]


def build_dr2_rows(sites: list[dict], mobile_df: pd.DataFrame | None, day: date) -> list[dict]:
    """Construit, pour chaque site DR2 détecté, un dict prêt à persister en
    base (`Dr2ViolationRecord`) — enrichi avec les infos du ticket
    correspondant quand celui-ci est retrouvé dans les incidents du jour."""
    rows = []
    for s in sites:
        ticket = _match_ticket_row(s['site_key'], mobile_df)

        def _val(col):
            if ticket is None or col not in ticket or pd.isna(ticket[col]):
                return ''
            return str(ticket[col]).strip()

        alarm_time = cancel_time = None
        is_resolved = False
        if ticket is not None:
            at = ticket.get('Alarm Time')
            ct = ticket.get('Cancel Time')
            alarm_time = pd.to_datetime(at, errors='coerce') if pd.notna(at) else None
            cancel_time = pd.to_datetime(ct, errors='coerce') if pd.notna(ct) else None
            is_resolved = cancel_time is not None and pd.notna(cancel_time)

        rows.append({
            'date':           day,
            'site_name':      _val('Site Name') or s['site_name_3g'] or s['site_name_2g'],
            'site_name_2g':   s['site_name_2g'],
            'site_name_3g':   s['site_name_3g'],
            'site_id':        _val('Site ID'),
            'site_parent':    _val('Site Parent'),
            'region':         _val('Région') or _val('Region'),
            'numero_ticket':  _val('Numero du ticket'),
            'categorie':      _val('Escalade'),
            'cause':          _val('Cause'),
            'point_bloquant': _val('Point bloquant'),
            'observation':    _val('Observation'),
            'alarm_time':     alarm_time.to_pydatetime() if alarm_time is not None and pd.notna(alarm_time) else None,
            'cancel_time':    cancel_time.to_pydatetime() if cancel_time is not None and pd.notna(cancel_time) else None,
            'hours_down':     s['hours_down'],
            'is_resolved':    bool(is_resolved),
        })
    return rows


def save_dr2_day(day: date, rows: list[dict], filename_2g: str = '', filename_3g: str = '', user=None):
    """Persiste (upsert) les violations DR2 calculées pour `day` : supprime
    les anciens enregistrements du jour puis recrée, et marque le jour
    comme traité (`Dr2ProcessedDate`) — même si `rows` est vide (0 DR2 ce
    jour-là, à distinguer d'un jour non traité)."""
    from .models import Dr2ProcessedDate, Dr2ViolationRecord

    Dr2ViolationRecord.objects.filter(date=day).delete()
    Dr2ViolationRecord.objects.bulk_create([Dr2ViolationRecord(**r) for r in rows])
    Dr2ProcessedDate.objects.update_or_create(
        date=day,
        defaults={
            'sites_count': len(rows),
            'filename_2g': filename_2g,
            'filename_3g': filename_3g,
            'uploaded_by': user if user and user.is_authenticated else None,
        },
    )
