from django.db import migrations


def _swapped_day_month(value, report_day, allow_future=False):
    if value is None or report_day is None:
        return None
    if value.year != report_day.year or value.month == report_day.month:
        return None
    if value.day != report_day.month or value.month > 12:
        return None
    try:
        candidate = value.replace(month=value.day, day=value.month)
    except ValueError:
        return None
    original_gap = abs((value.date() - report_day).days)
    candidate_gap = abs((candidate.date() - report_day).days)
    if not allow_future and candidate.date() > report_day:
        return None
    if candidate_gap <= 7 and candidate_gap + 14 < original_gap:
        return candidate
    return None


def _normalize(report_day, alarm_time, cancel_time):
    normalized_alarm = _swapped_day_month(alarm_time, report_day) or alarm_time
    normalized_cancel = (
        _swapped_day_month(cancel_time, report_day, allow_future=True)
        or cancel_time
    )
    if (normalized_alarm is not None and normalized_cancel is not None
            and normalized_cancel < normalized_alarm):
        return alarm_time, cancel_time
    return normalized_alarm, normalized_cancel


def repair_dr2_datetimes(apps, schema_editor):
    Dr2ViolationRecord = apps.get_model('reports', 'Dr2ViolationRecord')
    updates = []
    for record in Dr2ViolationRecord.objects.all().iterator(chunk_size=500):
        alarm_time, cancel_time = _normalize(
            record.date, record.alarm_time, record.cancel_time)
        if alarm_time == record.alarm_time and cancel_time == record.cancel_time:
            continue
        record.alarm_time = alarm_time
        record.cancel_time = cancel_time
        record.is_resolved = cancel_time is not None
        updates.append(record)
        if len(updates) == 500:
            Dr2ViolationRecord.objects.bulk_update(
                updates, ['alarm_time', 'cancel_time', 'is_resolved'])
            updates = []
    if updates:
        Dr2ViolationRecord.objects.bulk_update(
            updates, ['alarm_time', 'cancel_time', 'is_resolved'])


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0039_dr2violationrecord_root_cause'),
    ]

    operations = [
        migrations.RunPython(repair_dr2_datetimes, migrations.RunPython.noop),
    ]