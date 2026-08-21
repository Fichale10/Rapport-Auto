from datetime import date

import pandas as pd
from django.test import SimpleTestCase, TestCase

from .dr2_analytics import compute
from .dr2_availability import compute_dr2_sites, filter_availability_day
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


class Dr2AnalyticsTests(TestCase):
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
