from __future__ import annotations

import pandas as pd
from django.core.management.base import BaseCommand
from django.utils import timezone

from reports.dr2_availability import normalize_site_key, parse_ticket_datetime
from reports.models import Dr2ViolationRecord


def _clean(value) -> str:
    if value is None or pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text.lower() in ('nan', 'nat', 'none', 'null') else text


def _matching_ticket(record, tickets: pd.DataFrame):
    if tickets is None or tickets.empty or 'Numero du ticket' not in tickets.columns:
        return None
    ticket_number = _clean(record.numero_ticket).upper()
    if not ticket_number:
        return None
    matches = tickets[
        tickets['Numero du ticket'].map(_clean).str.upper() == ticket_number
    ]
    if matches.empty:
        return None
    if 'Site Name' in matches.columns:
        site_key = normalize_site_key(record.site_name)
        same_site = matches[
            matches['Site Name'].map(normalize_site_key) == site_key
        ]
        if not same_site.empty:
            matches = same_site
    return matches.iloc[0]


def _aware(value):
    parsed = parse_ticket_datetime(value)
    if parsed is None:
        return None
    result = parsed.to_pydatetime()
    if timezone.is_naive(result):
        result = timezone.make_aware(result, timezone.get_current_timezone())
    return result


class Command(BaseCommand):
    help = "Répare Alarm Time / Cancel Time des DR2 avec le format français JJ-MM-AA"

    def add_arguments(self, parser):
        parser.add_argument('--dry-run', action='store_true',
                            help="Affiche les corrections sans les sauvegarder")

    def handle(self, *args, **options):
        candidates = list(
            Dr2ViolationRecord.objects.exclude(numero_ticket='')
            .order_by('date', 'site_name')
        )
        if not candidates:
            self.stdout.write(self.style.SUCCESS('Aucun DR2 horodaté à vérifier.'))
            return

        from reports.api_import import fetch_api_excel

        date_start = candidates[0].date
        date_end = candidates[-1].date
        self.stdout.write(
            f'Chargement API mobile du {date_start:%d/%m/%Y} au {date_end:%d/%m/%Y}...')
        buffer, _filename = fetch_api_excel(
            date_start.isoformat(), date_end.isoformat(), 'mobile')
        tickets = pd.read_excel(buffer)

        updated = []
        missing = 0
        for record in candidates:
            ticket = _matching_ticket(record, tickets)
            if ticket is None:
                missing += 1
                continue
            alarm_time = _aware(ticket.get('Alarm Time'))
            cancel_time = _aware(ticket.get('Cancel Time'))
            if alarm_time is None:
                missing += 1
                continue
            if record.alarm_time == alarm_time and record.cancel_time == cancel_time:
                continue
            self.stdout.write(
                f'  {record.site_name}: {record.alarm_time} → {alarm_time} | '
                f'{record.cancel_time} → {cancel_time}')
            record.alarm_time = alarm_time
            record.cancel_time = cancel_time
            record.is_resolved = cancel_time is not None
            updated.append(record)

        if updated and not options['dry_run']:
            Dr2ViolationRecord.objects.bulk_update(
                updated, ['alarm_time', 'cancel_time', 'is_resolved'])

        suffix = ' [DRY-RUN]' if options['dry_run'] else ''
        self.stdout.write(self.style.SUCCESS(
            f'Terminé : {len(updated)} correction(s), {missing} sans correspondance{suffix}'))