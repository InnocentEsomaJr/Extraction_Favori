from datetime import date, timedelta
import unittest

from reporting_periods import (
    PERIODICITIES,
    analytics_periods,
    coverage_months,
    period_options,
    select_period_range,
    selection_bounds,
)
from validation_filters import monthly_bounds, reference_month


class ReportingPeriodsTests(unittest.TestCase):
    def test_monthly_options_preserve_month_ids_labels_and_leap_day(self):
        months = period_options("Mensuel", "2024")
        self.assertEqual(len(months), 12)
        self.assertEqual((months[0].id, months[0].label), ("202401", "Janvier"))
        self.assertEqual((months[-1].id, months[-1].label), ("202412", "Décembre"))
        self.assertEqual(selection_bounds([months[1]]), (date(2024, 2, 1), date(2024, 2, 29)))

    def test_quarters_cover_standard_calendar_quarters(self):
        quarters = period_options("Trimestriel", 2026)
        self.assertEqual([period.id for period in quarters], ["2026Q1", "2026Q2", "2026Q3", "2026Q4"])
        self.assertEqual([selection_bounds([period]) for period in quarters], [
            (date(2026, 1, 1), date(2026, 3, 31)),
            (date(2026, 4, 1), date(2026, 6, 30)),
            (date(2026, 7, 1), date(2026, 9, 30)),
            (date(2026, 10, 1), date(2026, 12, 31)),
        ])
        self.assertEqual(quarters[0].label, "T1 — Janvier à Mars")
        self.assertEqual(quarters[-1].label, "T4 — Octobre à Décembre")

    def test_semesters_start_in_january_and_july(self):
        semesters = period_options("Semestriel", 2026)
        self.assertEqual([period.id for period in semesters], ["2026S1", "2026S2"])
        self.assertEqual([selection_bounds([period]) for period in semesters], [
            (date(2026, 1, 1), date(2026, 6, 30)),
            (date(2026, 7, 1), date(2026, 12, 31)),
        ])
        self.assertEqual(semesters[0].label, "S1 — Janvier à Juin")
        self.assertEqual(semesters[-1].label, "S2 — Juillet à Décembre")

    def test_annual_option_is_one_complete_calendar_year(self):
        year = period_options("Annuel", 2024)
        self.assertEqual(len(year), 1)
        self.assertEqual((year[0].id, year[0].label), ("2024", "2024"))
        self.assertEqual(selection_bounds(year), (date(2024, 1, 1), date(2024, 12, 31)))
        self.assertEqual(coverage_months(year), [f"2024{month:02d}" for month in range(1, 13)])

    def test_iso_weeks_include_cross_year_dates_and_valid_week_53(self):
        weeks = period_options("Hebdomadaire", 2020)
        self.assertEqual(len(weeks), 53)
        self.assertEqual(weeks[0].id, "2020W1")
        self.assertEqual(selection_bounds([weeks[0]]), (date(2019, 12, 30), date(2020, 1, 5)))
        self.assertEqual(weeks[-1].id, "2020W53")
        self.assertEqual(selection_bounds([weeks[-1]]), (date(2020, 12, 28), date(2021, 1, 3)))
        self.assertEqual(weeks[0].label, "Semaine 1 — 30/12/2019 au 05/01/2020")
        self.assertEqual(weeks[-1].label, "Semaine 53 — 28/12/2020 au 03/01/2021")
        for index, week in enumerate(weeks, 1):
            self.assertEqual(week.start_date.isocalendar(), (2020, index, 1))
            self.assertEqual(week.end_date - week.start_date, timedelta(days=6))

    def test_year_with_52_iso_weeks_has_no_week_53(self):
        weeks = period_options("Hebdomadaire", 2021)
        self.assertEqual(len(weeks), 52)
        self.assertEqual(weeks[-1].id, "2021W52")
        self.assertEqual(selection_bounds([weeks[-1]]), (date(2021, 12, 27), date(2022, 1, 2)))
        with self.assertRaises(ValueError):
            select_period_range("Hebdomadaire", 2021, "2021W52", "2021W53")

    def test_interval_is_inclusive_and_orders_weeks_numerically(self):
        selected = select_period_range("Hebdomadaire", 2026, "2026W9", "2026W11")
        self.assertEqual([period.id for period in selected], ["2026W9", "2026W10", "2026W11"])
        quarters = select_period_range("Trimestriel", 2026, "2026Q2", "2026Q3")
        self.assertEqual([period.id for period in quarters], ["2026Q2", "2026Q3"])
        self.assertEqual(coverage_months(quarters), [f"2026{month:02d}" for month in range(4, 10)])
        self.assertEqual(selection_bounds(quarters), monthly_bounds(coverage_months(quarters)))

    def test_same_period_is_accepted_in_every_periodicity(self):
        for periodicity in PERIODICITIES:
            with self.subTest(periodicity=periodicity):
                first = period_options(periodicity, 2026)[0]
                self.assertEqual(select_period_range(periodicity, 2026, first.id, first.id), (first,))

    def test_reversed_intervals_are_rejected_for_every_multi_period_type(self):
        for periodicity in PERIODICITIES:
            options = period_options(periodicity, 2026)
            if len(options) > 1:
                with self.subTest(periodicity=periodicity), self.assertRaisesRegex(ValueError, "début"):
                    select_period_range(periodicity, 2026, options[-1].id, options[0].id)

    def test_month_coverage_deduplicates_without_replacing_exact_week_bounds(self):
        weeks = select_period_range("Hebdomadaire", 2020, "2020W1", "2020W2")
        self.assertEqual(coverage_months(weeks), ["201912", "202001"])
        self.assertEqual(selection_bounds(weeks), (date(2019, 12, 30), date(2020, 1, 12)))
        self.assertEqual(monthly_bounds(coverage_months(weeks)), (date(2019, 12, 1), date(2020, 1, 31)))
        self.assertEqual(reference_month(selection_bounds(weeks)[0]), "201911")

    def test_analytics_uses_native_period_ids_for_calendar_groups(self):
        cases = [
            ("Mensuel", "202602", "202603", ["202602", "202603"]),
            ("Trimestriel", "2026Q1", "2026Q2", ["2026Q1", "2026Q2"]),
            ("Semestriel", "2026S2", "2026S2", ["2026S2"]),
            ("Annuel", "2026", "2026", ["2026"]),
        ]
        for periodicity, first, last, expected in cases:
            with self.subTest(periodicity=periodicity):
                self.assertEqual(analytics_periods(select_period_range(periodicity, 2026, first, last)), expected)

    def test_weekly_analytics_uses_native_ids_instead_of_covered_months(self):
        weeks = select_period_range("Hebdomadaire", 2020, "2020W52", "2020W53")
        self.assertEqual(analytics_periods(weeks), ["2020W52", "2020W53"])
        self.assertEqual(coverage_months(weeks), ["202012", "202101"])

    def test_no_current_year_or_future_period_cutoff(self):
        self.assertEqual(period_options("Trimestriel", 2099)[-1].id, "2099Q4")
        self.assertEqual(period_options("Mensuel", 1900)[0].id, "190001")
        for periodicity in PERIODICITIES:
            options = period_options(periodicity, 2026)
            selected = select_period_range(periodicity, 2026, options[0].id, options[-1].id)
            self.assertEqual(selected, options)

    def test_invalid_year_periodicity_and_period_ids_are_rejected(self):
        for year in (None, True, 0, -1, 10000, 2026.0, "2026.5", "invalid"):
            with self.subTest(year=year), self.assertRaises(ValueError):
                period_options("Mensuel", year)
        for periodicity in (None, "", "Inconnu"):
            with self.subTest(periodicity=periodicity), self.assertRaises(ValueError):
                period_options(periodicity, 2026)
        for first, last in (("2025Q1", "2026Q2"), ("2026Q1", "2026Q5"), (None, "2026Q1")):
            with self.subTest(first=first, last=last), self.assertRaises(ValueError):
                select_period_range("Trimestriel", 2026, first, last)

    def test_empty_or_invalid_period_lists_fail_clearly(self):
        for helper in (selection_bounds, coverage_months, analytics_periods):
            for periods in (None, [], "202601", [None]):
                with self.subTest(helper=helper.__name__, periods=periods), self.assertRaises(ValueError):
                    helper(periods)

    def test_analytics_deduplicates_native_ids_without_changing_order(self):
        periods = select_period_range("Trimestriel", 2026, "2026Q2", "2026Q3")
        self.assertEqual(analytics_periods((periods[1], periods[0], periods[1])), ["2026Q3", "2026Q2"])


if __name__ == "__main__":
    unittest.main()
