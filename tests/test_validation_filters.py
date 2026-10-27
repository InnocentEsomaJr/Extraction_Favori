from datetime import date, datetime, timezone
import unittest

from validation_filters import date_range_months, monthly_bounds, normalize_date, reference_month


class ValidationDateFiltersTests(unittest.TestCase):
    def test_exact_range_covers_all_intersecting_months(self):
        self.assertEqual(date_range_months(date(2026, 7, 27), date(2026, 10, 27)),
                         ["202607", "202608", "202609", "202610"])

    def test_same_day_and_leap_day_cover_one_month(self):
        self.assertEqual(date_range_months("2024-02-29", "2024-02-29"), ["202402"])
        self.assertEqual(monthly_bounds(["202402"]), (date(2024, 2, 1), date(2024, 2, 29)))

    def test_cross_year_months_and_reference_month(self):
        self.assertEqual(date_range_months("2025-12-31", "2026-02-01"), ["202512", "202601", "202602"])
        self.assertEqual(reference_month(date(2026, 1, 27)), "202512")
        self.assertEqual(reference_month("2026-07-27"), "202606")

    def test_dates_datetimes_and_iso_strings_keep_their_civil_day(self):
        self.assertEqual(normalize_date(datetime(2026, 10, 27, 23, 30, tzinfo=timezone.utc)), date(2026, 10, 27))
        self.assertEqual(normalize_date("2026-10-27T23:30:00+00:00"), date(2026, 10, 27))
        self.assertEqual(normalize_date(" 2026-10-27 "), date(2026, 10, 27))

    def test_monthly_defaults_accept_unsorted_months_and_a_single_period(self):
        self.assertEqual(monthly_bounds(["202601", "202512", "202602"]), (date(2025, 12, 1), date(2026, 2, 28)))
        self.assertEqual(monthly_bounds("202607"), (date(2026, 7, 1), date(2026, 7, 31)))

    def test_invalid_missing_or_reversed_dates_are_rejected(self):
        for first, last in [(None, "2026-10-27"), ("bad", "2026-10-27"), ("2026-10-28", "2026-10-27")]:
            with self.subTest(first=first, last=last), self.assertRaises(ValueError):
                date_range_months(first, last)
        for periods in [None, [], ["202613"], ["2026Q1"], [None], 202601]:
            with self.subTest(periods=periods), self.assertRaises(ValueError):
                monthly_bounds(periods)
        with self.assertRaises(ValueError):
            reference_month(None)


if __name__ == "__main__":
    unittest.main()
