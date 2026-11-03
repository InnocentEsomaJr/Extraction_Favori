import unittest

import pandas as pd

from categorisation import calculer_categorisation
from quality_score import build_quality_rows, metrics_from_dashboard_summary, metrics_from_reporting_source


class QualityRowsTests(unittest.TestCase):
    def test_quality_uses_live_reporting_and_joins_validation_only_by_id(self):
        zones = pd.DataFrame([
            {'Zone_ID': 'ZoneAAAAAAA', 'Zone': 'Même zone', 'Province': 'Province A'},
            {'Zone_ID': 'ZoneBBBBBBB', 'Zone': 'Même zone', 'Province': 'Province B'},
        ])
        metrics = pd.DataFrame([
            {'Zone_ID': 'ZoneAAAAAAA', 'Complétude_Globale': 98, 'Promptitude_Globale': 80,
             'Actual_Reports': 98, 'Nombre_Datasets_95': 10},
            {'Zone_ID': 'ZoneBBBBBBB', 'Complétude_Globale': 50, 'Promptitude_Globale': 30,
             'Actual_Reports': 50, 'Nombre_Datasets_95': 4},
        ])
        validation = pd.DataFrame([
            {'Zone_ID': 'ZoneBBBBBBB', 'Zone de santé': 'Même zone', 'Violations_Actuelles': 7,
             'Complétude_Globale': 999, 'Actual_Reports': 999},
            {'Zone_ID': 'OtherZoneAA', 'Violations_Actuelles': 999},
        ])
        result = build_quality_rows(metrics, zones, validation).set_index('Zone_ID')
        self.assertEqual(result.loc['ZoneAAAAAAA', 'Score_Qualite'], 89)
        self.assertEqual(result.loc['ZoneBBBBBBB', 'Score_Qualite'], 40)
        self.assertEqual(result.loc['ZoneBBBBBBB', 'Actual_Reports'], 50)
        self.assertTrue(pd.isna(result.loc['ZoneAAAAAAA', 'Violations_Actuelles']))
        self.assertEqual(result.loc['ZoneBBBBBBB', 'Violations_Actuelles'], 7)
        self.assertNotEqual(result.iloc[0]['Zone de santé'], result.iloc[1]['Zone de santé'])

    def test_missing_inputs_stay_numeric_missing_and_zero_remains_zero(self):
        zones = pd.DataFrame([{'Zone_ID': 'ZoneAAAAAAA', 'Zone': 'A'}])
        missing = build_quality_rows(pd.DataFrame(), zones)
        for column in ['Complétude_Globale', 'Promptitude_Globale', 'Actual_Reports', 'Score_Qualite', 'Violations_Actuelles']:
            self.assertTrue(pd.api.types.is_numeric_dtype(missing[column]))
            self.assertTrue(missing[column].isna().all())
        zeros = build_quality_rows(pd.DataFrame([{
            'Zone_ID': 'ZoneAAAAAAA', 'Complétude_Globale': 0, 'Promptitude_Globale': 0,
            'Actual_Reports': 0, 'Nombre_Datasets_95': 0,
        }]), zones)
        self.assertEqual(zeros.iloc[0]['Score_Qualite'], 0)
        self.assertEqual(zeros.iloc[0]['Actual_Reports'], 0)
        self.assertFalse(zeros.iloc[0]['Controle_Corrections_Effectue'])


class DashboardSummaryMetricsTests(unittest.TestCase):
    def targets(self, ids=('ZoneAAAAAAA',)):
        return pd.DataFrame([{'Zone_ID': identifier, 'Zone': 'Ango'} for identifier in ids])

    def raw(self, identifier='ZoneAAAAAAA', name='Ango'):
        row = {'Organisation unit ID': identifier, 'Organisation unit': name}
        for number in range(9):
            actual, on_time = (10, 7) if number < 8 else (20, 12)
            row.update({
                f'Dataset {number} Actual reports': actual,
                f'Dataset {number} Expected reports': actual,
                f'Dataset {number} Actual reports on time': on_time,
                f'Dataset {number} Reporting rate': 100,
            })
        return row

    def summary(self, name='Ango', **overrides):
        row = {
            'Organisation unit': name, 'Complétude_Globale (%)': 100,
            'Promptitude_Globale (%)': 68, 'Reports_Actual': 100,
            'Nombre des data set complétude >/=95%': 9,
        }
        row.update(overrides)
        return row

    def test_ango_summary_keeps_100_68_and_nine_datasets_despite_other_missing_dataset(self):
        raw = self.raw()
        raw.update({
            'Inactive Actual reports': float('nan'), 'Inactive Expected reports': float('nan'),
            'Inactive Actual reports on time': float('nan'), 'Inactive Reporting rate': float('nan'),
        })
        metrics = metrics_from_dashboard_summary(pd.DataFrame([self.summary()]), pd.DataFrame([raw]), self.targets())
        self.assertEqual(metrics.iloc[0].tolist(), ['ZoneAAAAAAA', 100, 68, 100, 9])
        rows = build_quality_rows(metrics, self.targets())
        calculated = calculer_categorisation(rows, profil='ZS')
        self.assertEqual(calculated.iloc[0]['Score_Qualite'], 84)
        self.assertEqual(calculated.iloc[0]['Score de qualité /20'], 15)

    def test_summary_explicit_ids_resolve_homonyms_without_using_other_zone_values(self):
        source = pd.DataFrame([self.raw(), self.raw('ZoneBBBBBBB')])
        ambiguous = metrics_from_dashboard_summary(pd.DataFrame([self.summary()]), source, self.targets())
        self.assertTrue(ambiguous.iloc[0, 1:].isna().all())
        explicit = pd.DataFrame([
            self.summary(**{'Organisation unit ID': 'ZoneBBBBBBB', 'Complétude_Globale (%)': 50}),
            self.summary(**{'Organisation unit ID': 'ZoneAAAAAAA'}),
        ])
        result = metrics_from_dashboard_summary(explicit, source, self.targets(('ZoneAAAAAAA', 'ZoneBBBBBBB')))
        self.assertEqual(result['Zone_ID'].tolist(), ['ZoneAAAAAAA', 'ZoneBBBBBBB'])
        self.assertEqual(result['Complétude_Globale'].tolist(), [100, 50])

    def test_bridge_requires_one_display_name_per_identifier_and_one_summary_row(self):
        source = pd.DataFrame([self.raw(), self.raw(name='Autre libellé')])
        result = metrics_from_dashboard_summary(pd.DataFrame([self.summary()]), source, self.targets())
        self.assertTrue(result.iloc[0, 1:].isna().all())
        duplicated = pd.DataFrame([self.summary(), self.summary()])
        result = metrics_from_dashboard_summary(duplicated, pd.DataFrame([self.raw()]), self.targets())
        self.assertTrue(result.iloc[0, 1:].isna().all())

    def test_empty_raw_values_do_not_turn_legacy_filled_zeros_into_measurements(self):
        raw = self.raw()
        for column in list(raw):
            if column not in ['Organisation unit', 'Organisation unit ID']:
                raw[column] = float('nan')
        summary = self.summary(**{
            'Complétude_Globale (%)': 0, 'Promptitude_Globale (%)': 0,
            'Reports_Actual': 0, 'Nombre des data set complétude >/=95%': 0,
        })
        result = metrics_from_dashboard_summary(pd.DataFrame([summary]), pd.DataFrame([raw]), self.targets())
        self.assertTrue(result.iloc[0, 1:].isna().all())

    def test_real_zero_reports_with_positive_expected_can_be_reused(self):
        raw = self.raw()
        for column in list(raw):
            if column.endswith('Actual reports') or column.endswith('Actual reports on time') or column.endswith('Reporting rate'):
                raw[column] = 0
        summary = self.summary(**{
            'Complétude_Globale (%)': 0, 'Promptitude_Globale (%)': 0,
            'Reports_Actual': 0, 'Nombre des data set complétude >/=95%': 0,
        })
        result = metrics_from_dashboard_summary(pd.DataFrame([summary]), pd.DataFrame([raw]), self.targets())
        self.assertEqual(result.iloc[0].tolist(), ['ZoneAAAAAAA', 0, 0, 0, 0])

    def test_absent_rows_and_ids_remain_missing_for_each_target(self):
        source = pd.DataFrame([self.raw()])
        summary = pd.DataFrame([self.summary()])
        for invalid_source in [source.iloc[:0], source.drop(columns='Organisation unit ID')]:
            with self.subTest(columns=invalid_source.columns.tolist()):
                result = metrics_from_dashboard_summary(summary, invalid_source, self.targets(('ZoneAAAAAAA', 'ZoneBBBBBBB')))
                self.assertEqual(result['Zone_ID'].tolist(), ['ZoneAAAAAAA', 'ZoneBBBBBBB'])
                self.assertTrue(result.iloc[:, 1:].isna().all().all())
        result = metrics_from_dashboard_summary(summary, source, self.targets(('ZoneBBBBBBB',)))
        self.assertTrue(result.iloc[0, 1:].isna().all())

    def test_missing_on_time_family_keeps_the_other_available_metrics(self):
        source = pd.DataFrame([self.raw()])
        source = source.drop(columns=[column for column in source if column.endswith('Actual reports on time')])
        result = metrics_from_dashboard_summary(pd.DataFrame([self.summary()]), source, self.targets()).iloc[0]
        self.assertEqual(result['Actual_Reports'], 100)
        self.assertEqual(result['Complétude_Globale'], 100)
        self.assertEqual(result['Nombre_Datasets_95'], 9)
        self.assertTrue(pd.isna(result['Promptitude_Globale']))

    def test_missing_expected_family_only_keeps_real_report_counts(self):
        source = pd.DataFrame([self.raw()])
        source = source.drop(columns=[column for column in source if column.endswith('Expected reports')])
        result = metrics_from_dashboard_summary(pd.DataFrame([self.summary()]), source, self.targets()).iloc[0]
        self.assertEqual(result['Actual_Reports'], 100)
        self.assertTrue(result[['Complétude_Globale', 'Promptitude_Globale', 'Nombre_Datasets_95']].isna().all())

    def test_no_mutation_and_invalid_or_zero_expected_do_not_generate_rates(self):
        source = pd.DataFrame([self.raw()])
        summary, targets = pd.DataFrame([self.summary()]), self.targets()
        source_before, summary_before, targets_before = source.copy(deep=True), summary.copy(deep=True), targets.copy(deep=True)
        metrics_from_dashboard_summary(summary, source, targets)
        pd.testing.assert_frame_equal(source, source_before)
        pd.testing.assert_frame_equal(summary, summary_before)
        pd.testing.assert_frame_equal(targets, targets_before)
        for value in [0, -1, float('inf'), 'inconnu']:
            with self.subTest(expected=value):
                changed = source.copy()
                for column in [c for c in source if c.endswith('Expected reports')]:
                    changed[column] = value
                result = metrics_from_dashboard_summary(summary, changed, targets)
                self.assertTrue(result[['Complétude_Globale', 'Promptitude_Globale', 'Nombre_Datasets_95']].isna().all().all())

    def test_fractional_or_unproven_dataset_count_is_not_taken_from_summary(self):
        source = pd.DataFrame([self.raw()])
        for count in [9.5, 10, -1]:
            with self.subTest(count=count):
                result = metrics_from_dashboard_summary(
                    pd.DataFrame([self.summary(**{'Nombre des data set complétude >/=95%': count})]), source, self.targets(),
                )
                self.assertTrue(pd.isna(result.iloc[0]['Nombre_Datasets_95']))


class ReportingFallbackMetricsTests(unittest.TestCase):
    def targets(self):
        return pd.DataFrame([
            {'Zone_ID': 'ProvinceAAA', 'Zone': 'DPS A', 'Profil_Bareme': 'DPS'},
            {'Zone_ID': 'ProvinceBBB', 'Zone': 'DPS B', 'Profil_Bareme': 'DPS'},
        ])

    def raw(self):
        rows = []
        for identifier, dataset_count, actual, on_time in [
            ('ProvinceAAA', 9, 10, 8), ('ProvinceBBB', 6, 5, 4),
        ]:
            row = {'Organisation unit ID': identifier, 'Organisation unit': identifier}
            for number in range(dataset_count):
                row.update({
                    f'Dataset {number} Actual reports': actual,
                    f'Dataset {number} Expected reports': 10,
                    f'Dataset {number} Actual reports on time': on_time,
                    f'Dataset {number} Reporting rate': actual / 10 * 100,
                    f'Dataset {number} Reporting rate on time': on_time / 10 * 100,
                })
            rows.append(row)
        frame = pd.DataFrame(rows)
        for suffix in ['Actual reports', 'Expected reports', 'Actual reports on time', 'Reporting rate', 'Reporting rate on time']:
            frame['Dataset inutilisé ' + suffix] = float('nan')
        return frame

    def test_inactive_datasets_are_excluded_separately_for_each_dps(self):
        source, targets = self.raw(), self.targets()
        before = source.copy(deep=True)
        result = metrics_from_reporting_source(source, targets).set_index('Zone_ID')
        self.assertEqual(result.loc['ProvinceAAA'].tolist(), [100, 80, 90, 9])
        self.assertEqual(result.loc['ProvinceBBB'].tolist(), [50, 40, 30, 0])
        pd.testing.assert_frame_equal(source, before)
        rows = build_quality_rows(result.reset_index(), targets)
        self.assertEqual(rows['Profil_Bareme'].tolist(), ['DPS', 'DPS'])
        scores = calculer_categorisation(rows)
        self.assertEqual(scores['Score datasets /20'].tolist(), [10, 0])

    def test_all_absent_is_missing_but_observed_zeros_remain_zero(self):
        frame = self.raw()
        value_columns = [column for column in frame if column not in ['Organisation unit', 'Organisation unit ID']]
        frame[value_columns] = float('nan')
        result = metrics_from_reporting_source(frame, self.targets())
        self.assertTrue(result.iloc[:, 1:].isna().all().all())
        frame = self.raw()
        for column in frame:
            if column.endswith(('Actual reports', 'Actual reports on time', 'Reporting rate', 'Reporting rate on time')):
                frame[column] = frame[column].where(frame[column].isna(), 0)
        result = metrics_from_reporting_source(frame, self.targets())
        self.assertEqual(result.iloc[:, 1:].to_numpy().tolist(), [[0, 0, 0, 0], [0, 0, 0, 0]])

    def test_missing_one_family_does_not_erase_other_metrics(self):
        source = self.raw().drop(columns=[column for column in self.raw() if column.endswith('Actual reports on time')])
        result = metrics_from_reporting_source(source, self.targets()).iloc[0]
        self.assertEqual(result['Actual_Reports'], 90)
        self.assertEqual(result['Complétude_Globale'], 100)
        self.assertEqual(result['Nombre_Datasets_95'], 9)
        self.assertTrue(pd.isna(result['Promptitude_Globale']))

    def test_missing_cell_in_an_active_dataset_is_not_ignored(self):
        source = self.raw()
        source.loc[0, 'Dataset 0 Actual reports'] = float('nan')
        result = metrics_from_reporting_source(source, self.targets())
        self.assertTrue(result.loc[0, ['Actual_Reports', 'Complétude_Globale', 'Nombre_Datasets_95']].isna().all())
        self.assertEqual(result.loc[0, 'Promptitude_Globale'], 80)
        self.assertEqual(result.loc[1, 'Complétude_Globale'], 50)

    def test_malformed_negative_or_infinite_values_only_affect_their_target(self):
        for value in ['illisible', -1, float('inf')]:
            with self.subTest(value=value):
                source = self.raw().astype(object)
                source.loc[0, 'Dataset 0 Expected reports'] = value
                result = metrics_from_reporting_source(source, self.targets())
                self.assertEqual(result.loc[0, 'Actual_Reports'], 90)
                self.assertTrue(result.loc[0, ['Complétude_Globale', 'Promptitude_Globale', 'Nombre_Datasets_95']].isna().all())
                self.assertEqual(result.loc[1, 'Actual_Reports'], 30)
                self.assertEqual(result.loc[1, 'Promptitude_Globale'], 40)

    def test_missing_id_cannot_be_joined_by_label(self):
        source = self.raw().drop(columns='Organisation unit ID')
        source['Organisation unit'] = ['DPS A', 'DPS B']
        result = metrics_from_reporting_source(source, self.targets())
        self.assertTrue(result.iloc[:, 1:].isna().all().all())


if __name__ == '__main__':
    unittest.main()
