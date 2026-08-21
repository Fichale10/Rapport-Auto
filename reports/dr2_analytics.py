"""Calculs et export du tableau de bord Analytics DR2."""
from __future__ import annotations

import io
from collections import Counter
from datetime import date, timedelta

import pandas as pd


EMPTY_LABEL = '(Non renseigné)'


def _site_metadata(names: set[str]) -> dict[str, dict]:
    from .models import Site

    return {
        row['site_name'].strip().upper(): row
        for row in Site.objects.filter(site_name__in=names).values(
            'site_name', 'region', 'base', 'classif_tech'
        )
    }


def _records_dataframe(debut: date, fin: date) -> pd.DataFrame:
    from .models import Dr2ViolationRecord

    fields = [
        'date', 'site_name', 'site_name_2g', 'site_name_3g', 'site_id',
        'site_parent', 'region', 'numero_ticket', 'categorie', 'cause',
        'root_cause', 'point_bloquant', 'observation', 'alarm_time',
        'cancel_time', 'hours_down', 'is_resolved',
    ]
    records = list(
        Dr2ViolationRecord.objects.filter(date__gte=debut, date__lte=fin).values(*fields)
    )
    df = pd.DataFrame(records, columns=fields)
    if df.empty:
        return df

    metadata = _site_metadata(set(df['site_name'].dropna().astype(str)))
    for index, row in df.iterrows():
        site = str(row['site_name'] or '').strip().upper()
        meta = metadata.get(site, {})
        if not str(row['region'] or '').strip():
            df.at[index, 'region'] = meta.get('region') or EMPTY_LABEL
        df.at[index, 'base'] = meta.get('base') or EMPTY_LABEL
        df.at[index, 'classification'] = meta.get('classif_tech') or EMPTY_LABEL

    for column in ('region', 'base', 'classification', 'categorie', 'cause', 'root_cause'):
        df[column] = df[column].fillna('').astype(str).str.strip().replace('', EMPTY_LABEL)
    df['site_name'] = df['site_name'].fillna('').astype(str).str.strip()
    df['hours_down'] = pd.to_numeric(df['hours_down'], errors='coerce').fillna(0).astype(int)
    df['is_resolved'] = df['is_resolved'].fillna(False).astype(bool)
    return df


def _processed_dates(debut: date, fin: date) -> set[date]:
    from .models import Dr2ProcessedDate

    return set(
        Dr2ProcessedDate.objects.filter(date__gte=debut, date__lte=fin)
        .values_list('date', flat=True)
    )


def compute(
    debut: date,
    fin: date,
    *,
    regions=(),
    sites=(),
    escalades=(),
    causes=(),
    status: str = '',
) -> dict:
    """Construit les KPIs, séries et tableaux du dashboard DR2."""
    df_all = _records_dataframe(debut, fin)
    processed = _processed_dates(debut, fin)

    filters = {
        'regions': sorted(x for x in df_all.get('region', pd.Series(dtype=str)).unique()
                          if x != EMPTY_LABEL),
        'sites': sorted(x for x in df_all.get('site_name', pd.Series(dtype=str)).unique() if x),
        'escalades': sorted(x for x in df_all.get('categorie', pd.Series(dtype=str)).unique()
                            if x != EMPTY_LABEL),
        'causes': sorted(x for x in df_all.get('cause', pd.Series(dtype=str)).unique()
                         if x != EMPTY_LABEL),
    }

    df = df_all
    if not df.empty:
        if regions:
            df = df[df['region'].isin(regions)]
        if sites:
            df = df[df['site_name'].isin(sites)]
        if escalades:
            df = df[df['categorie'].isin(escalades)]
        if causes:
            df = df[df['cause'].isin(causes)]
        if status == 'resolved':
            df = df[df['is_resolved']]
        elif status == 'open':
            df = df[~df['is_resolved']]

    total = int(len(df))
    resolved = int(df['is_resolved'].sum()) if total else 0
    site_counts = Counter(df['site_name']) if total else Counter()
    processed_count = len(processed)
    kpi = {
        'violations': total,
        'sites': int(df['site_name'].nunique()) if total else 0,
        'processed_days': processed_count,
        'average': round(total / processed_count, 2) if processed_count else 0,
        'resolved': resolved,
        'resolution_rate': round(resolved / total * 100, 1) if total else 0,
        'average_hours': round(float(df['hours_down'].mean()), 1) if total else 0,
        'recurrent_sites': sum(1 for count in site_counts.values() if count > 1),
    }

    by_day = Counter(df['date']) if total else Counter()
    trend = []
    day = debut
    while day <= fin:
        trend.append({
            'date': day.isoformat(),
            'label': day.strftime('%d/%m'),
            'count': int(by_day.get(day, 0)) if day in processed else None,
            'processed': day in processed,
        })
        day += timedelta(days=1)

    def grouped(column: str, limit: int = 15) -> list[dict]:
        if not total:
            return []
        group = (df.groupby(column).agg(
            violations=(column, 'size'),
            sites=('site_name', 'nunique'),
            hours=('hours_down', 'sum'),
        ).sort_values(['violations', 'hours'], ascending=False).head(limit))
        return [
            {
                'label': str(label), 'violations': int(row['violations']),
                'sites': int(row['sites']), 'hours': int(row['hours']),
                'pct': round(int(row['violations']) / total * 100, 1),
            }
            for label, row in group.iterrows()
        ]

    top_sites = []
    if total:
        grouped_sites = (df.groupby('site_name').agg(
            violations=('site_name', 'size'),
            hours=('hours_down', 'sum'),
            region=('region', 'first'),
            escalade=('categorie', lambda values: values.value_counts().index[0]),
            resolved=('is_resolved', 'sum'),
        ).sort_values(['violations', 'hours'], ascending=False).head(20))
        top_sites = [
            {
                'site': str(site), 'violations': int(row['violations']),
                'hours': int(row['hours']), 'region': str(row['region']),
                'escalade': str(row['escalade']), 'resolved': int(row['resolved']),
            }
            for site, row in grouped_sites.iterrows()
        ]

    details = []
    if total:
        detail_df = df.sort_values(['date', 'hours_down'], ascending=[False, False]).head(500)
        for row in detail_df.to_dict('records'):
            details.append({
                'date': row['date'].isoformat(),
                'site': row['site_name'],
                'site_id': row['site_id'] or '',
                'region': row['region'],
                'ticket': row['numero_ticket'] or '',
                'escalade': row['categorie'],
                'cause': row['cause'],
                'hours': row['hours_down'],
                'resolved': bool(row['is_resolved']),
            })

    return {
        'empty': total == 0,
        'filters': filters,
        'kpi': kpi,
        'trend': trend,
        'regions': grouped('region'),
        'escalades': grouped('categorie'),
        'causes': grouped('cause'),
        'top_sites': top_sites,
        'details': details,
        'unprocessed_days': (fin - debut).days + 1 - processed_count,
    }


def build_excel(result: dict, debut: date, fin: date) -> io.BytesIO:
    """Exporte exactement le périmètre filtré présenté par Analytics DR2."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        kpi = result['kpi']
        pd.DataFrame([
            ('Période', f'{debut:%d/%m/%Y} au {fin:%d/%m/%Y}'),
            ('Violations DR2', kpi['violations']),
            ('Sites concernés', kpi['sites']),
            ('Jours traités', kpi['processed_days']),
            ('Moyenne par jour traité', kpi['average']),
            ('Taux de résolution', f"{kpi['resolution_rate']} %"),
            ('Heures indisponibles moyennes', kpi['average_hours']),
            ('Sites récurrents', kpi['recurrent_sites']),
        ], columns=['Indicateur', 'Valeur']).to_excel(writer, sheet_name='KPIs', index=False)
        pd.DataFrame(result['details']).rename(columns={
            'date': 'Date', 'site': 'Site', 'site_id': 'Site ID', 'region': 'Région',
            'ticket': 'Ticket', 'escalade': 'Escalade', 'cause': 'Cause',
            'hours': 'Heures indisponibles', 'resolved': 'Résolu',
        }).to_excel(writer, sheet_name='Détail DR2', index=False)
        pd.DataFrame(result['top_sites']).to_excel(writer, sheet_name='Sites récurrents', index=False)
        pd.DataFrame(result['regions']).to_excel(writer, sheet_name='Régions', index=False)
        pd.DataFrame(result['escalades']).to_excel(writer, sheet_name='Escalades', index=False)
        pd.DataFrame(result['causes']).to_excel(writer, sheet_name='Causes', index=False)
    buffer.seek(0)
    return buffer