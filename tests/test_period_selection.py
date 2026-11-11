"""Bornes des périodes fixes et sélections hebdomadaires libres."""

from datetime import date
import unittest

from period_selection import (
    canonical_period_ids, period_choices, period_info, period_interval,
    selection_bounds, selection_months, weekly_choices_for_range,
)


class PeriodSelectionTests(unittest.TestCase):
    def test_fixed_month_quarter_semester_and_year_boundaries(self):
        examples = [
            ('20240229', date(2024, 2, 29), date(2024, 2, 29)),
            ('202602', date(2026, 2, 1), date(2026, 2, 28)),
            ('2024Q1', date(2024, 1, 1), date(2024, 3, 31)),
            ('2026S2', date(2026, 7, 1), date(2026, 12, 31)),
            ('2024', date(2024, 1, 1), date(2024, 12, 31)),
        ]
        for identifier, first, last in examples:
            with self.subTest(identifier=identifier):
                record = period_info(identifier)
                self.assertEqual((record['start'], record['end']), (first, last))

    def test_choice_counts_french_labels_and_leap_month(self):
        for kind, count in [('Journalier', 366), ('Mensuel', 12), ('Trimestriel', 4), ('Semestriel', 2), ('Annuel', 1)]:
            with self.subTest(kind=kind):
                self.assertEqual(len(period_choices(kind, '2024')), count)
        self.assertEqual(period_choices('Monthly', 2024)[1]['end'], date(2024, 2, 29))
        self.assertEqual(period_choices('Monthly', 2026)[8]['label'], 'Septembre 2026')
        self.assertEqual(len(period_choices('Daily', 2026)), 365)
        self.assertEqual(period_choices('Jour', 2024)[59]['id'], '20240229')

    def test_iso_week_choices_include_week_53_and_actual_cross_year_dates(self):
        records = period_choices('Weekly', 2026)
        self.assertEqual(len(records), 53)
        self.assertEqual(records[0]['id'], '2026W1')
        self.assertEqual(records[0]['start'], date(2025, 12, 29))
        self.assertEqual(records[-1]['end'], date(2027, 1, 3))
        self.assertEqual(len(period_choices('Hebdomadaire', 2025)), 52)

    def test_selected_week_normalizes_padding_without_changing_its_dates(self):
        record = period_info(' 2026w01 ')
        self.assertEqual(record['id'], '2026W1')
        self.assertIn('2025-12-29 au 2026-01-04', record['label'])

    def test_selection_deduplicates_sorts_and_keeps_gaps(self):
        self.assertEqual(canonical_period_ids(['202609', '202601', '202609']), ('202601', '202609'))
        self.assertEqual(selection_months(['202609', '202601']), ('202601', '202609'))
        self.assertEqual(selection_bounds(['202609', '202601']), (date(2026, 1, 1), date(2026, 9, 30)))

    def test_full_semester_expands_to_six_months(self):
        self.assertEqual(selection_months('2026S1'), tuple(f'2026{month:02d}' for month in range(1, 7)))
        self.assertEqual(selection_months('2026'), tuple(f'2026{month:02d}' for month in range(1, 13)))

    def test_week_selection_uses_sunday_month_without_adding_its_monday_month(self):
        self.assertEqual(selection_months(['2026W1', '2026W36']), ('202601', '202609'))
        self.assertEqual(selection_months([f'2026W{week}' for week in range(36, 40)]), ('202609',))
        self.assertEqual(selection_months(['2026W53']), ('202701',))
        self.assertEqual(selection_bounds(['2026W53']), (date(2026, 12, 28), date(2027, 1, 3)))

    def test_month_interval_crosses_year_and_includes_every_period(self):
        records = period_interval('Mois', '202611', '202702')
        self.assertEqual([record['id'] for record in records], ['202611', '202612', '202701', '202702'])

    def test_quarter_and_semester_intervals_cross_year(self):
        self.assertEqual([record['id'] for record in period_interval('Trimestre', '2026Q4', '2027Q2')],
                         ['2026Q4', '2027Q1', '2027Q2'])
        self.assertEqual([record['id'] for record in period_interval('Semestre', '2026S2', '2027S1')],
                         ['2026S2', '2027S1'])

    def test_week_interval_preserves_iso_week_53_across_year(self):
        self.assertEqual([record['id'] for record in period_interval('Semaine', '2026W52', '2027W2')],
                         ['2026W52', '2026W53', '2027W1', '2027W2'])

    def test_day_interval_is_inclusive_and_crosses_year_without_gaps(self):
        self.assertEqual([record['id'] for record in period_interval('Jour', '20261231', '20270102')],
                         ['20261231', '20270101', '20270102'])
        self.assertEqual(selection_bounds(['20261231', '20270102']),
                         (date(2026, 12, 31), date(2027, 1, 2)))
        self.assertEqual(selection_months(['20261231', '20270102']), ('202612', '202701'))
        self.assertEqual([record['id'] for record in period_interval('Daily', '20240228', '20240301')],
                         ['20240228', '20240229', '20240301'])
        self.assertEqual([record['id'] for record in period_interval('Journalier', '20260906', '20260906')],
                         ['20260906'])

    def test_adjacent_quarters_have_distinct_closed_dates_and_week_membership(self):
        q1, q2 = period_info('2024Q1'), period_info('2024Q2')
        self.assertEqual(q1['end'], date(2024, 3, 31))
        self.assertEqual(q2['start'], date(2024, 4, 1))
        first = {record['id'] for record in weekly_choices_for_range(q1['start'], q1['end'])}
        second = {record['id'] for record in weekly_choices_for_range(q2['start'], q2['end'])}
        self.assertIn('2024W13', first)
        self.assertNotIn('2024W13', second)
        self.assertEqual(first & second, set())
        self.assertEqual(weekly_choices_for_range('2024-04-01', '2024-04-06'), [])

    def test_year_interval_expands_then_can_produce_months(self):
        records = period_interval('Année', '2025', '2027')
        self.assertEqual([record['id'] for record in records], ['2025', '2026', '2027'])
        self.assertEqual(len(selection_months([record['id'] for record in records])), 36)

    def test_september_matches_screenshot_without_automatic_august_carry(self):
        choices = weekly_choices_for_range('2026-09-01', '2026-09-30')
        self.assertEqual([record['id'] for record in choices], [f'2026W{week}' for week in range(36, 40)])
        self.assertEqual(choices[0]['start'], date(2026, 8, 31))

    def test_november_keeps_five_choices_and_december_does_not_force_a_carry(self):
        november = weekly_choices_for_range(date(2026, 11, 1), date(2026, 11, 30))
        december = weekly_choices_for_range(date(2026, 12, 1), date(2026, 12, 31))
        self.assertEqual([record['id'] for record in november], [f'2026W{week}' for week in range(44, 49)])
        self.assertEqual([record['id'] for record in december], [f'2026W{week}' for week in range(49, 53)])

    def test_semester_week_choices_are_not_limited_to_four_or_five(self):
        choices = weekly_choices_for_range('2026-01-01', '2026-06-30')
        self.assertEqual(len(choices), 26)
        self.assertEqual(len({record['id'] for record in choices}), 26)

    def test_short_date_interval_without_a_sunday_has_no_week_choices(self):
        self.assertEqual(weekly_choices_for_range('2026-09-07', '2026-09-12'), [])

    def test_sunday_is_included_at_both_closed_bounds_and_monday_excludes_previous_week(self):
        self.assertEqual([record['id'] for record in weekly_choices_for_range('2026-09-06', '2026-09-06')],
                         ['2026W36'])
        self.assertEqual([record['id'] for record in weekly_choices_for_range('2026-09-06', '2026-09-13')],
                         ['2026W36', '2026W37'])
        self.assertEqual([record['id'] for record in weekly_choices_for_range('2026-09-07', '2026-09-13')],
                         ['2026W37'])
        self.assertEqual(weekly_choices_for_range('2026-09-07', '2026-09-12'), [])

    def test_contained_weeks_keep_full_monday_sunday_within_closed_bounds(self):
        self.assertEqual([record['id'] for record in weekly_choices_for_range('2026-09-01', '2026-09-30', boundary='contained')],
                         ['2026W37', '2026W38', '2026W39'])
        self.assertEqual([record['id'] for record in weekly_choices_for_range('2026-09-07', '2026-09-13', boundary='contained')],
                         ['2026W37'])
        self.assertEqual(weekly_choices_for_range('2026-09-08', '2026-09-13', boundary='contained'), [])
        self.assertEqual(weekly_choices_for_range('2026-09-13', '2026-09-13', boundary='contained'), [])

    def test_week_boundary_policy_preserves_iso_year_rollover(self):
        self.assertEqual([record['id'] for record in weekly_choices_for_range('2026-12-28', '2027-01-10', boundary='contained')],
                         ['2026W53', '2027W1'])
        self.assertEqual([record['id'] for record in weekly_choices_for_range('2027-01-01', '2027-01-03')],
                         ['2026W53'])

    def test_invalid_periods_and_reversed_or_mixed_intervals_are_rejected(self):
        for identifier in ['202600', '202613', '2026Q0', '2026Q5', '2026S3', '2025W53', '20260229', '20260230', '20260001', '20261301', '0000', '', None, '2026-09']:
            with self.subTest(identifier=identifier), self.assertRaises(ValueError):
                period_info(identifier)
        for values in [[], None, ['202601', 'malformed']]:
            with self.subTest(values=values), self.assertRaises(ValueError):
                canonical_period_ids(values)
        for kind, first, last in [('Monthly', '202602', '202601'), ('Monthly', '2026Q1', '2026Q2'),
                                  ('Daily', '20260907', '20260906'), ('Daily', '202609', '202610')]:
            with self.assertRaises(ValueError):
                period_interval(kind, first, last)
        with self.assertRaises(ValueError):
            weekly_choices_for_range('2026-09-30', '2026-09-01')
        for policy in ['overlap', '', None, False]:
            with self.subTest(policy=policy), self.assertRaises(ValueError):
                weekly_choices_for_range('2026-09-01', '2026-09-30', boundary=policy)


if __name__ == '__main__':
    unittest.main()
