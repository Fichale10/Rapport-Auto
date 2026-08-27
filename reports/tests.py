from datetime import date, datetime, timedelta, timezone

import pandas as pd
from django.test import SimpleTestCase, TestCase

from .dr2_analytics import compute
from .dr2_availability import (
	compute_dr2_sites, filter_availability_day, normalize_dr2_datetimes,
)
from .dr2_meeting_pptx import (
	_causes_breakdown, _dr2_dataset, _gdi_reporting_periods,
)
from .models import Dr2ProcessedDate, Dr2ViolationRecord


class Dr2AvailabilityTests(SimpleTestCase):
	def test_selected_day_excludes_adjacent_hours(self):
		periods = pd.to_datetime([
			'2026-08-19 22:00', '2026-08-19 23:00',
			'2026-08-20 00:00', '2026-08-20 01:00',
		])
		source = pd.DataFrame({
			'period': periods,
			'site_raw': ['SITE_A'] * 4,
			'site_key': ['SITEA'] * 4,
			'value': [0] * 4,
		})

		selected = filter_availability_day(source, date(2026, 8, 20), '2G')

		self.assertEqual(len(selected), 2)
		self.assertEqual(compute_dr2_sites(selected, selected), [])

	def test_normalize_dr2_datetimes_repairs_day_month_inversion(self):
		alarm = datetime(2026, 5, 8, 23, 0, tzinfo=timezone.utc)
		cancel = datetime(2026, 6, 8, 2, 0, tzinfo=timezone.utc)

		normalized_alarm, normalized_cancel = normalize_dr2_datetimes(
			date(2026, 8, 6), alarm, cancel)

		self.assertEqual(
			normalized_alarm,
			datetime(2026, 8, 5, 23, 0, tzinfo=timezone.utc),
		)
		self.assertEqual(
			normalized_cancel,
			datetime(2026, 8, 6, 2, 0, tzinfo=timezone.utc),
		)
		self.assertEqual(normalized_cancel - normalized_alarm, timedelta(hours=3))

	def test_normalize_dr2_datetimes_preserves_legitimate_long_outage(self):
		alarm = datetime(2026, 7, 1, 8, 0, tzinfo=timezone.utc)
		cancel = datetime(2026, 7, 11, 8, 0, tzinfo=timezone.utc)

		self.assertEqual(
			normalize_dr2_datetimes(date(2026, 7, 11), alarm, cancel),
			(alarm, cancel),
		)

	def test_normalize_dr2_datetimes_rejects_distant_swap(self):
		alarm = datetime(2026, 7, 8, 8, 0, tzinfo=timezone.utc)

		self.assertEqual(
			normalize_dr2_datetimes(date(2026, 8, 31), alarm, None),
			(alarm, None),
		)

	def test_normalize_dr2_datetimes_does_not_move_alarm_after_report_day(self):
		alarm = datetime(2026, 7, 8, 8, 0, tzinfo=timezone.utc)

		self.assertEqual(
			normalize_dr2_datetimes(date(2026, 8, 6), alarm, None),
			(alarm, None),
		)

	def test_gdi_reporting_periods_use_last_completed_weekend(self):
		periods = _gdi_reporting_periods(date(2026, 8, 26))

		self.assertEqual(periods['month_start'], date(2026, 8, 1))
		self.assertEqual(periods['detail_start'], date(2026, 8, 21))
		self.assertEqual(periods['detail_end'], date(2026, 8, 23))
		self.assertEqual(periods['meeting_day'], date(2026, 8, 24))


class Dr2AnalyticsTests(TestCase):
	def test_gdi_cause_percentages_include_records_without_cause(self):
		day = date(2026, 8, 23)
		for site, cause in (
			('SITE-A', 'RADIO ODU HS'),
			('SITE-B', 'RADIO ODU HS'),
			('SITE-C', ''),
			('SITE-D', ''),
		):
			Dr2ViolationRecord.objects.create(
				date=day, site_name=site, cause=cause, hours_down=3,
			)

		result = _causes_breakdown(
			Dr2ViolationRecord.objects.filter(date=day), total=4,
		)

		self.assertEqual(result, [('RADIO ODU HS', 2, 50)])

	def test_gdi_dataset_uses_processed_days_and_selected_period_end(self):
		for day in (date(2026, 8, 21), date(2026, 8, 22), date(2026, 8, 23)):
			Dr2ProcessedDate.objects.create(date=day, sites_count=0)
		Dr2ViolationRecord.objects.create(
			date=date(2026, 8, 21), site_name='SITE-A', region='LOME',
			categorie='ENERGIE', hours_down=3,
		)
		Dr2ViolationRecord.objects.create(
			date=date(2026, 8, 23), site_name='SITE-B', region='KARA',
			categorie='RAN-FIELD O', hours_down=4,
		)

		result = _dr2_dataset(date(2026, 8, 1), date(2026, 8, 23))

		self.assertEqual(result['period_days'], 3)
		self.assertEqual(result['moyenne'], 0.67)
		self.assertEqual(result['nbre_j1'], 1)

	def test_average_and_trend_use_processed_days(self):
		for day, count in ((date(2026, 8, 10), 2), (date(2026, 8, 11), 1),
						   (date(2026, 8, 12), 0)):
			Dr2ProcessedDate.objects.create(date=day, sites_count=count)
		Dr2ViolationRecord.objects.create(
			date=date(2026, 8, 10), site_name='SITE-A', region='LOME',
			categorie='ENERGIE', cause='PANNE', hours_down=3, is_resolved=True,
		)
		Dr2ViolationRecord.objects.create(
			date=date(2026, 8, 10), site_name='SITE-B', region='KARA',
			categorie='RAN', cause='LIEN', hours_down=4, is_resolved=False,
		)
		Dr2ViolationRecord.objects.create(
			date=date(2026, 8, 11), site_name='SITE-A', region='LOME',
			categorie='ENERGIE', cause='PANNE', hours_down=5, is_resolved=False,
		)

		result = compute(date(2026, 8, 10), date(2026, 8, 13))

		self.assertEqual(result['kpi']['violations'], 3)
		self.assertEqual(result['kpi']['processed_days'], 3)
		self.assertEqual(result['kpi']['average'], 1.0)
		self.assertEqual(result['kpi']['recurrent_sites'], 1)
		self.assertEqual([point['count'] for point in result['trend']], [2, 1, 0, None])
