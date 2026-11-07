"""Le report conserve les semaines et leur appartenance sans doublon."""

from datetime import date, timedelta
import unittest

from weekly_periods import weekly_period_schedule, weekly_periods_for_months


class WeeklyPeriodTests(unittest.TestCase):
    def test_civil_september_matches_the_four_weeks_in_the_screenshot(self):
        schedule = weekly_period_schedule(['202609'], carry_fifth=False)
        self.assertEqual([row['Semaine'] for row in schedule], ['2026W36', '2026W37', '2026W38', '2026W39'])
        self.assertEqual(schedule[0], {
            'Mois': '202609', 'Semaine': '2026W36', 'Début': '2026-08-31',
            'Fin': '2026-09-06', 'Reportée': False,
        })
        self.assertEqual(schedule[-1]['Fin'], '2026-09-27')
        self.assertNotIn('2026W40', weekly_periods_for_months(['202609'], carry_fifth=False))

    def test_august_fifth_week_is_added_to_september_once_by_default(self):
        august = weekly_period_schedule('202608')
        september = weekly_period_schedule('202609')
        self.assertEqual([row['Semaine'] for row in august], ['2026W31', '2026W32', '2026W33', '2026W34'])
        self.assertEqual([row['Semaine'] for row in september], ['2026W35', '2026W36', '2026W37', '2026W38', '2026W39'])
        self.assertEqual(september[0], {
            'Mois': '202609', 'Semaine': '2026W35', 'Début': '2026-08-24',
            'Fin': '2026-08-30', 'Reportée': True,
        })
        together = weekly_periods_for_months(['202608', '202609'])
        self.assertEqual(together, tuple(f'2026W{week}' for week in range(31, 40)))
        self.assertEqual(len(together), len(set(together)))

    def test_november_fifth_week_is_evaluated_in_december(self):
        november = weekly_period_schedule('202611')
        december = weekly_period_schedule('202612')
        self.assertEqual([row['Semaine'] for row in november], ['2026W44', '2026W45', '2026W46', '2026W47'])
        self.assertEqual([row['Semaine'] for row in december], ['2026W48', '2026W49', '2026W50', '2026W51', '2026W52'])
        self.assertEqual(december[0]['Mois'], '202612')
        self.assertEqual(december[0]['Fin'], '2026-11-29')
        self.assertTrue(december[0]['Reportée'])
        self.assertFalse(any(row['Reportée'] for row in december[1:]))

    def test_civil_policy_keeps_all_five_endings_in_the_original_month(self):
        november = weekly_period_schedule('202611', carry_fifth=False)
        december = weekly_period_schedule('202612', carry_fifth=False)
        self.assertEqual([row['Semaine'] for row in november], ['2026W44', '2026W45', '2026W46', '2026W47', '2026W48'])
        self.assertEqual([row['Semaine'] for row in december], ['2026W49', '2026W50', '2026W51', '2026W52'])
        self.assertFalse(any(row['Reportée'] for row in november + december))

    def test_january_uses_iso_week_year_and_unpadded_week_numbers(self):
        schedule = weekly_period_schedule('202101')
        self.assertEqual([row['Semaine'] for row in schedule], ['2020W53', '2021W1', '2021W2', '2021W3'])
        self.assertEqual(schedule[0]['Début'], '2020-12-28')
        self.assertEqual(schedule[0]['Fin'], '2021-01-03')
        self.assertEqual(weekly_periods_for_months('202102')[0], '2021W4')

    def test_december_fifth_week_can_move_across_a_calendar_year(self):
        december = weekly_period_schedule('202312')
        january = weekly_period_schedule('202401')
        self.assertNotIn('2023W52', [row['Semaine'] for row in december])
        self.assertEqual(january[0], {
            'Mois': '202401', 'Semaine': '2023W52', 'Début': '2023-12-25',
            'Fin': '2023-12-31', 'Reportée': True,
        })

    def test_leap_february_and_multi_month_selection_keep_real_week_dates(self):
        schedule = weekly_period_schedule(['202402', '202403'])
        self.assertEqual(weekly_periods_for_months('202402'), ('2024W5', '2024W6', '2024W7', '2024W8'))
        self.assertEqual([row['Fin'] for row in schedule[:4]], ['2024-02-04', '2024-02-11', '2024-02-18', '2024-02-25'])
        for row in schedule:
            start, end = date.fromisoformat(row['Début']), date.fromisoformat(row['Fin'])
            self.assertEqual(start.weekday(), 0)
            self.assertEqual(end.weekday(), 6)
            self.assertEqual(end - start, timedelta(days=6))

    def test_repeated_unordered_months_and_semicolon_input_are_identical(self):
        expected = weekly_period_schedule(['202611', '202612'])
        self.assertEqual(weekly_period_schedule(['202612', '202611', '202612']), expected)
        self.assertEqual(weekly_period_schedule('202612;202611'), expected)
        self.assertEqual(len(expected), len({row['Semaine'] for row in expected}))

    def test_consecutive_months_preserve_a_continuous_week_sequence(self):
        schedule = weekly_period_schedule([f'2026{month:02d}' for month in range(1, 13)])
        endings = [date.fromisoformat(row['Fin']) for row in schedule]
        self.assertEqual(len(endings), 52)
        self.assertTrue(all(second - first == timedelta(days=7) for first, second in zip(endings, endings[1:])))
        self.assertEqual(len(schedule), len({row['Semaine'] for row in schedule}))

    def test_invalid_months_are_rejected_instead_of_broadening_the_period(self):
        for months in [None, [], '', '202600', '202613', '000001', '20261', '2026-01',
                       '202601;', ['202609', 'not-a-month'], 202609, b'202609']:
            with self.subTest(months=months), self.assertRaises(ValueError):
                weekly_period_schedule(months)

    def test_policy_requires_a_boolean(self):
        for policy in ['false', 1, None]:
            with self.subTest(policy=policy), self.assertRaises(ValueError):
                weekly_period_schedule('202609', carry_fifth=policy)

    def test_minimum_calendar_year_does_not_request_year_zero(self):
        schedule = weekly_period_schedule('000101')
        self.assertEqual(len(schedule), 4)
        self.assertEqual(schedule[0]['Début'], '0001-01-01')
        self.assertEqual(schedule[0]['Semaine'], '0001W1')


if __name__ == '__main__':
    unittest.main()
