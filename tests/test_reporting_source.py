"""La source commune remplace les métriques sans changer le périmètre demandé."""

from copy import deepcopy
import unittest

import pandas as pd

from reporting_source import replace_reporting_source, recover_monthly_reporting, available_reporting_summary


class ReportingSourceTests(unittest.TestCase):
    def test_recovery_preserves_only_selected_monthly_cells_and_marks_total_incomplete(self):
        source = self.favorite()
        source['Tertiaire Actual reports'] = 999
        source['Surveillance Actual reports'] = 68
        source.attrs['reporting_item_labels'] = {
            'DatasetAAAA.ACTUAL_REPORTS': 'Dataset Actual reports',
            'DatasetAAAA.EXPECTED_REPORTS': 'Dataset Expected reports',
            'TertiaryAAA.ACTUAL_REPORTS': 'Tertiaire Actual reports',
            'WeeklyAAAAA.ACTUAL_REPORTS': 'Surveillance Actual reports',
        }
        recovered = recover_monthly_reporting(source, ['ZoneAAAAAAA'], ['DatasetAAAA'], ['Erreur reçue'])
        self.assertEqual(recovered['Dataset Actual reports'].tolist(), [999])
        self.assertNotIn('Tertiaire Actual reports', recovered)
        self.assertNotIn('Surveillance Actual reports', recovered)
        self.assertEqual(recovered.attrs['failed_units'], ['ZoneAAAAAAA'])
        self.assertEqual(recovered.attrs['failed_cells_by_unit'], {})
        self.assertEqual(recovered.attrs['reporting_warnings'], ['Erreur reçue'])
        self.assertEqual(recovered.attrs['reporting_status'], 'partial')

    def test_recovery_never_uses_name_guesses_or_other_units(self):
        recovered = recover_monthly_reporting(self.favorite(), ['ZoneZZZZZZZ'], ['DatasetAAAA'])
        self.assertEqual(recovered['Organisation unit ID'].tolist(), ['ZoneZZZZZZZ'])
        self.assertEqual(recovered.attrs['reporting_columns'], [])
        self.assertEqual(recovered.attrs['reporting_status'], 'unavailable')

    def test_unselected_dataset_failure_does_not_hide_recovered_monthly_cells(self):
        from reporting_availability import mask_failed_reporting_cells
        source = self.favorite()
        source.attrs.update({
            'reporting_item_labels': {'DatasetAAAA.ACTUAL_REPORTS': 'Dataset Actual reports'},
            'failed_units': ['ZoneAAAAAAA'],
            'failed_metrics_by_unit': {'ZoneAAAAAAA': ['TertiaryAAA.ACTUAL_REPORTS']},
        })
        recovered = recover_monthly_reporting(source, ['ZoneAAAAAAA'], ['DatasetAAAA'])
        self.assertEqual(recovered.attrs['failed_cells_by_unit'], {})
        visible = mask_failed_reporting_cells(recovered, recovered)
        self.assertEqual(visible['Dataset Actual reports'].tolist(), [999])

    def test_available_summary_keeps_received_zero_and_missing_weekly_distinct(self):
        source = self.combined()
        source.loc[1, source.attrs['weekly_reporting_columns']] = float('nan')
        source.attrs['failed_units'] = ['ZoneAAAAAAA']
        summary = available_reporting_summary(source)
        self.assertEqual(summary['Rapports mensuels disponibles'].tolist(), [0, 90])
        self.assertEqual(summary.loc[0, 'Rapports hebdomadaires disponibles'], 0)
        self.assertTrue(pd.isna(summary.loc[1, 'Rapports hebdomadaires disponibles']))
        self.assertEqual(summary['État des sources'].tolist(), ['Complet', 'Incomplet'])

    def favorite(self):
        frame = pd.DataFrame({
            'Organisation unit ID': ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'],
            'Organisation unit': ['Ancien nom A', 'Même nom', 'Hors périmètre'],
            'Dataset Actual reports': [999, 999, 999],
            'Dataset Expected reports': [1000, 1000, 1000],
            'Dataset Reporting rate': [99.9, 99.9, 99.9],
            'Indicateur': [10, 20, 30],
            'Faux Actual reports': [555, 666, 777],
        })
        frame.attrs['reporting_columns'] = [
            'Dataset Actual reports', 'Dataset Expected reports', 'Dataset Reporting rate',
        ]
        return frame

    def combined(self):
        frame = pd.DataFrame({
            'Organisation unit ID': ['ZoneBBBBBBB', 'ZoneAAAAAAA'],
            'Organisation unit': ['Même nom', 'Nom récent A'],
            'Dataset Actual reports': [0.0, 90.0],
            'Dataset Expected reports': [100.0, 100.0],
            'Dataset Reporting rate': [0.0, 90.0],
            'eSurveillance Actual reports': [0.0, 68.0],
            'eSurveillance Expected reports': [68.0, 68.0],
            'eSurveillance Reporting rate': [0.0, 100.0],
        })
        frame.attrs.update({
            'reporting_columns': list(frame.columns[2:]),
            'weekly_reporting_columns': list(frame.columns[-3:]),
            'monthly_reporting_columns': list(frame.columns[2:5]),
            'reporting_periods': {'monthly': ('202609',), 'weekly': ('2026W36', '2026W37')},
            'failed_units': [], 'reporting_warnings': [],
        })
        return frame

    def test_combined_metrics_replace_old_values_in_exact_target_order(self):
        original, combined = self.favorite(), self.combined()
        before_original, before_combined = original.copy(deep=True), combined.copy(deep=True)
        result = replace_reporting_source(original, combined)
        self.assertEqual(result['Organisation unit ID'].tolist(), ['ZoneBBBBBBB', 'ZoneAAAAAAA'])
        self.assertEqual(result['Dataset Actual reports'].tolist(), [0, 90])
        self.assertEqual(result['eSurveillance Actual reports'].tolist(), [0, 68])
        self.assertEqual(result['Indicateur'].tolist(), [20, 10])
        self.assertEqual(result['Faux Actual reports'].tolist(), [666, 555])
        self.assertNotIn('Faux Actual reports', result.attrs['reporting_columns'])
        self.assertEqual(result['Organisation unit'].tolist(), ['Même nom', 'Nom récent A'])
        self.assertNotIn('ZoneCCCCCCC', result['Organisation unit ID'].tolist())
        pd.testing.assert_frame_equal(original, before_original)
        pd.testing.assert_frame_equal(combined, before_combined)

    def test_missing_original_unit_has_missing_generic_value_without_name_join(self):
        original, combined = self.favorite().iloc[:1].copy(), self.combined()
        combined['Organisation unit'] = 'Ancien nom A'
        result = replace_reporting_source(original, combined)
        self.assertTrue(pd.isna(result.loc[0, 'Indicateur']))
        self.assertEqual(result.loc[1, 'Indicateur'], 10)
        self.assertEqual(result.loc[0, 'Dataset Actual reports'], 0)

    def test_combined_failure_attributes_and_periods_survive_and_are_copied(self):
        original, combined = self.favorite(), self.combined()
        combined.attrs['failed_units'] = ['ZoneAAAAAAA']
        combined.attrs['failed_metrics_by_unit'] = {'ZoneAAAAAAA': ['WeeklyAAAAA.ACTUAL_REPORTS']}
        combined.attrs['reporting_warnings'] = ['Surveillance indisponible']
        original.attrs['failed_units'] = ['ZoneBBBBBBB', 'ZoneCCCCCCC']
        original.attrs['reporting_warnings'] = ['Indicateur générique indisponible']
        result = replace_reporting_source(original, combined)
        self.assertEqual(result.attrs['failed_units'], ['ZoneAAAAAAA'])
        self.assertEqual(result.attrs['favorite_failed_units'], ['ZoneBBBBBBB'])
        self.assertEqual(result.attrs['reporting_warnings'], ['Surveillance indisponible'])
        self.assertEqual(result.attrs['favorite_warnings'], ['Indicateur générique indisponible'])
        self.assertEqual(result.attrs['reporting_periods'], combined.attrs['reporting_periods'])
        result.attrs['failed_metrics_by_unit']['ZoneAAAAAAA'].append('other')
        self.assertEqual(combined.attrs['failed_metrics_by_unit']['ZoneAAAAAAA'], ['WeeklyAAAAA.ACTUAL_REPORTS'])

    def test_absent_original_keeps_combined_source_and_attrs(self):
        combined = self.combined()
        result = replace_reporting_source(None, combined)
        pd.testing.assert_frame_equal(result, combined)
        self.assertEqual(result.attrs, combined.attrs)
        result.attrs['failed_units'].append('ZoneBBBBBBB')
        self.assertEqual(combined.attrs['failed_units'], [])

    def test_uncertified_legacy_source_uses_metric_suffixes_only(self):
        original, combined = self.favorite(), self.combined()
        original.attrs = {}
        original['Commentaire sur reporting rate présent'] = ['A', 'B', 'C']
        result = replace_reporting_source(original, combined)
        self.assertNotIn('Faux Actual reports', result)
        self.assertEqual(result['Commentaire sur reporting rate présent'].tolist(), ['B', 'A'])
        self.assertEqual(result['Indicateur'].tolist(), [20, 10])

    def test_explicit_empty_certification_preserves_metric_lookalike(self):
        original = self.favorite()[['Organisation unit ID', 'Faux Actual reports']].copy()
        original.attrs['reporting_columns'] = []
        result = replace_reporting_source(original, self.combined())
        self.assertEqual(result['Faux Actual reports'].tolist(), [666, 555])
        self.assertNotIn('Faux Actual reports', result.attrs['reporting_columns'])

    def test_names_fallback_only_by_exact_id_when_combined_has_no_display_name(self):
        combined = self.combined()
        combined['Organisation unit'] = ['ZoneBBBBBBB', None]
        result = replace_reporting_source(self.favorite(), combined)
        self.assertEqual(result['Organisation unit'].tolist(), ['Même nom', 'Ancien nom A'])

    def test_ambiguous_generic_collision_is_rejected(self):
        original = self.favorite()
        original.attrs['reporting_columns'] = ['Dataset Actual reports', 'Dataset Expected reports']
        with self.assertRaisesRegex(ValueError, 'même libellé'):
            replace_reporting_source(original, self.combined())

    def test_duplicate_or_missing_unit_ids_and_duplicate_columns_are_rejected(self):
        for side in ('original', 'combined'):
            for invalid in ('duplicate', 'missing', 'no_column', 'columns'):
                with self.subTest(side=side, invalid=invalid):
                    original, combined = self.favorite(), self.combined()
                    frame = original if side == 'original' else combined
                    if invalid == 'duplicate':
                        frame.loc[1, 'Organisation unit ID'] = frame.loc[0, 'Organisation unit ID']
                    elif invalid == 'missing':
                        frame.loc[0, 'Organisation unit ID'] = None
                    elif invalid == 'no_column':
                        frame = frame.drop(columns='Organisation unit ID')
                    else:
                        frame.columns = ['Organisation unit ID', *[frame.columns[1]] * (len(frame.columns) - 1)]
                    if side == 'original':
                        original = frame
                    else:
                        combined = frame
                    with self.assertRaises(ValueError):
                        replace_reporting_source(original, combined)

    def test_empty_combined_source_never_reintroduces_old_units_or_reporting_values(self):
        combined = self.combined().iloc[:0].copy()
        combined.attrs['quality_warnings'] = ['Source indisponible']
        result = replace_reporting_source(self.favorite(), combined)
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs['quality_warnings'], ['Source indisponible'])
        self.assertIn('Indicateur', result)
        self.assertEqual(result['Organisation unit ID'].tolist(), [])


if __name__ == '__main__':
    unittest.main()
