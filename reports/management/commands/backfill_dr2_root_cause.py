from __future__ import annotations

from datetime import datetime, time

import pandas as pd
from django.core.management.base import BaseCommand

from reports.dr2_availability import normalize_site_key
from reports.models import Dr2ViolationRecord


def _clean(value) -> str:
    if value is None or pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text.lower() in ('nan', 'nat', 'none', 'null') else text


def _match_root_cause(record, tickets: pd.DataFrame) -> str:
    if tickets is None or tickets.empty or 'Root Cause' not in tickets.columns:
        return ''

    candidates = tickets
    ticket_number = _clean(record.numero_ticket)
    if ticket_number and 'Numero du ticket' in tickets.columns:
        exact = tickets[
            tickets['Numero du ticket'].map(_clean).str.upper() == ticket_number.upper()
        ]
        roots = exact['Root Cause'].map(_clean)
        if (roots != '').any():
            return roots[roots != ''].iloc[0]

    if 'Site Name' not in tickets.columns:
        return ''
    site_key = normalize_site_key(record.site_name)
    if not site_key:
        return ''
    candidates = candidates[
        candidates['Site Name'].map(normalize_site_key) == site_key
    ].copy()
    if candidates.empty:
        return ''

    if 'Alarm Time' in candidates.columns:
        alarm = pd.to_datetime(candidates['Alarm Time'], dayfirst=True,
                               format='mixed', errors='coerce')
        cancel = pd.to_datetime(
            candidates.get('Cancel Time', pd.Series(index=candidates.index, dtype=object)),
            dayfirst=True, format='mixed', errors='coerce')
        day_start = pd.Timestamp(datetime.combine(record.date, time.min))
        day_end = pd.Timestamp(datetime.combine(record.date, time(23, 59, 59)))
        active = alarm.notna() & (alarm <= day_end) & (cancel.isna() | (cancel >= day_start))
        candidates = candidates[active].copy()
        alarm = alarm[active]
        if candidates.empty:
            return ''
        if record.alarm_time is not None:
            target = pd.Timestamp(record.alarm_time.replace(tzinfo=None))
            candidates['_distance'] = (alarm.dt.tz_localize(None) - target).abs()
            candidates = candidates.sort_values('_distance')

    roots = candidates['Root Cause'].map(_clean)
    return roots[roots != ''].iloc[0] if (roots != '').any() else ''


class Command(BaseCommand):
    help = "Complète Root Cause sur les anciens Dr2ViolationRecord depuis l'API mobile"

    def add_arguments(self, parser):
        parser.add_argument('--dry-run', action='store_true',
                            help="N'écrit aucune modification")

    def handle(self, *args, **options):
        candidates = list(
            Dr2ViolationRecord.objects.filter(root_cause='').order_by('date', 'site_name')
        )
        if not candidates:
            self.stdout.write(self.style.SUCCESS('Aucun DR2 à compléter.'))
            return

        date_start = candidates[0].date
        date_end = candidates[-1].date
        self.stdout.write(
            f'Chargement API mobile du {date_start:%d/%m/%Y} au {date_end:%d/%m/%Y} '
            f'pour {len(candidates)} DR2...')

        from reports.api_import import fetch_api_excel

        buffer, _filename = fetch_api_excel(
            date_start.isoformat(), date_end.isoformat(), 'mobile')
        tickets = pd.read_excel(buffer)

        updated = []
        for record in candidates:
            root_cause = _match_root_cause(record, tickets)
            if not root_cause:
                continue
            record.root_cause = root_cause
            updated.append(record)

        if updated and not options['dry_run']:
            Dr2ViolationRecord.objects.bulk_update(updated, ['root_cause'])

        missing = len(candidates) - len(updated)
        suffix = ' [DRY-RUN]' if options['dry_run'] else ''
        self.stdout.write(self.style.SUCCESS(
            f'Terminé : {len(updated)} complété(s), {missing} sans correspondance{suffix}'))