"""Un échec de chargement n'est ni un zéro ni une moyenne partielle."""

import unittest

import pandas as pd

from reporting_availability import mask_failed_reporting_cells, mask_failed_reporting_units


class ReportingAvailabilityTests(unittest.TestCase):
    def source(self):
        frame = pd.DataFrame({
            'Organisation unit': ['Aba', 'Autre ZS'],
            'Organisation unit ID': ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
            'Dataset Actual reports': [20.0, 0.0],
        })
        frame.attrs['failed_units'] = ['ZoneAAAAAAA']
        return frame

    def test_failed_unit_partial_numbers_are_masked_while_real_zero_survives(self):
        source = self.source()
        frame = pd.DataFrame({
            'Organisation unit ID': ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
            'Organisation unit': ['Aba', 'Autre ZS'],
            'Reports_Actual': [20.0, 0.0],
            'Complétude_Globale (%)': [100.0, 0.0],
            'Nombre des data set complétude >/=95%': [1, 0],
            'Commentaire': ['partiel', 'réel'],
        })
        before = frame.copy(deep=True)
        result = mask_failed_reporting_units(frame, source)
        self.assertTrue(result.loc[0, [
            'Reports_Actual', 'Complétude_Globale (%)', 'Nombre des data set complétude >/=95%',
        ]].isna().all())
        self.assertEqual(result.loc[1, 'Reports_Actual'], 0)
        self.assertEqual(result['Organisation unit ID'].tolist(), frame['Organisation unit ID'].tolist())
        self.assertEqual(result['Commentaire'].tolist(), ['partiel', 'réel'])
        pd.testing.assert_frame_equal(frame, before)

    def test_quality_metric_ids_use_zone_id_and_ignore_row_order_or_names(self):
        frame = pd.DataFrame({
            'Zone_ID': ['ZoneBBBBBBB', 'ZoneAAAAAAA'],
            'Actual_Reports': [0.0, 20.0], 'Nombre_Datasets_95': [0.0, 1.0],
        })
        result = mask_failed_reporting_units(frame, self.source())
        self.assertEqual(result.loc[0, 'Actual_Reports'], 0)
        self.assertTrue(result.loc[1, ['Actual_Reports', 'Nombre_Datasets_95']].isna().all())

    def test_summary_without_ids_uses_a_bijective_current_source_bridge(self):
        frame = pd.DataFrame({
            'Organisation unit': ['Autre ZS', 'Aba'], 'Reports_Actual': [0.0, 20.0],
        }, index=[15, 8])
        result = mask_failed_reporting_units(frame, self.source())
        self.assertEqual(result.loc[15, 'Reports_Actual'], 0)
        self.assertTrue(pd.isna(result.loc[8, 'Reports_Actual']))

    def test_ambiguous_names_or_missing_source_ids_do_not_invent_a_join(self):
        source = self.source()
        source['Organisation unit'] = 'Même nom'
        frame = pd.DataFrame({'Organisation unit': ['Même nom'], 'Reports_Actual': [20.0]})
        pd.testing.assert_frame_equal(mask_failed_reporting_units(frame, source), frame)
        pd.testing.assert_frame_equal(
            mask_failed_reporting_units(frame, source.drop(columns='Organisation unit ID')), frame,
        )
        source = self.source()
        alias = source.iloc[:1].copy()
        alias['Organisation unit'] = 'Alias Aba'
        source = pd.concat([source, alias], ignore_index=True)
        source.attrs['failed_units'] = ['ZoneAAAAAAA']
        frame = pd.DataFrame({'Organisation unit': ['Aba'], 'Reports_Actual': [20.0]})
        pd.testing.assert_frame_equal(mask_failed_reporting_units(frame, source), frame)

    def test_explicit_ids_take_priority_over_the_display_label(self):
        frame = pd.DataFrame({
            'Organisation unit ID': ['ZoneBBBBBBB', 'ZoneAAAAAAA'],
            'Organisation unit': ['Aba', 'Autre ZS'], 'Reports_Actual': [0.0, 20.0],
        })
        result = mask_failed_reporting_units(frame, self.source())
        self.assertEqual(result.loc[0, 'Reports_Actual'], 0)
        self.assertTrue(pd.isna(result.loc[1, 'Reports_Actual']))

    def test_explicit_metric_list_handles_object_numbers_and_preserves_other_metrics(self):
        frame = pd.DataFrame({
            'Organisation unit ID': ['ZoneAAAAAAA'], 'Organisation unit': ['Aba'],
            'Reports_Actual': ['20'], 'Indicateur indépendant': [99],
        })
        result = mask_failed_reporting_units(frame, self.source(), [
            'Reports_Actual', 'Organisation unit ID', 'Organisation unit', 'Colonne absente',
        ])
        self.assertTrue(pd.isna(result.loc[0, 'Reports_Actual']))
        self.assertEqual(result.loc[0, 'Indicateur indépendant'], 99)
        self.assertEqual(result.loc[0, 'Organisation unit ID'], 'ZoneAAAAAAA')
        self.assertEqual(result.loc[0, 'Organisation unit'], 'Aba')

    def test_failed_attributes_survive_dataframe_copies_and_masking(self):
        source = self.source()
        transformed = source.copy().rename(columns={'Dataset Actual reports': 'Rapports'}).copy()
        self.assertEqual(transformed.attrs['failed_units'], ['ZoneAAAAAAA'])
        result = mask_failed_reporting_units(transformed, transformed)
        self.assertEqual(result.attrs['failed_units'], ['ZoneAAAAAAA'])
        self.assertTrue(pd.isna(result.loc[0, 'Rapports']))
        pd.testing.assert_frame_equal(source, self.source())

    def test_absent_failures_empty_frames_and_unknown_failed_ids_are_safe(self):
        frame = self.source()
        source = frame.copy()
        source.attrs = {}
        pd.testing.assert_frame_equal(mask_failed_reporting_units(frame, source), frame)
        pd.testing.assert_frame_equal(mask_failed_reporting_units(frame.iloc[:0], frame), frame.iloc[:0])
        source.attrs['failed_units'] = ['MissingAAAA']
        pd.testing.assert_frame_equal(mask_failed_reporting_units(frame, source), frame)


class ReportingCellAvailabilityTests(unittest.TestCase):
    def source(self):
        frame = pd.DataFrame({
            'Organisation unit ID': ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
            'Organisation unit': ['Aba', 'Autre ZS'],
            'Mensuel Actual reports': [90.0, 30.0],
            'Mensuel Reporting rate': [90.0, 75.0],
            'Hebdomadaire Actual reports': [0.0, 0.0],
            'Hebdomadaire Reporting rate': [0.0, 0.0],
            'Reports_Actual': [90.0, 30.0],
            'Commentaire': ['réponse hebdomadaire absente', 'zéro réel'],
        })
        frame.attrs.update({
            'failed_units': ['ZoneAAAAAAA'],
            'failed_metrics_by_unit': {'ZoneAAAAAAA': [
                'WeeklyAAAAA.ACTUAL_REPORTS', 'WeeklyAAAAA.REPORTING_RATE',
            ]},
            'reporting_item_labels': {
                'MonthlyAAAA.ACTUAL_REPORTS': 'Mensuel Actual reports',
                'MonthlyAAAA.REPORTING_RATE': 'Mensuel Reporting rate',
                'WeeklyAAAAA.ACTUAL_REPORTS': 'Hebdomadaire Actual reports',
                'WeeklyAAAAA.REPORTING_RATE': 'Hebdomadaire Reporting rate',
            },
            'reporting_columns': [
                'Mensuel Actual reports', 'Mensuel Reporting rate',
                'Hebdomadaire Actual reports', 'Hebdomadaire Reporting rate',
            ],
        })
        return frame

    def test_received_monthly_cells_remain_visible_and_failed_weekly_cells_are_missing(self):
        source = self.source()
        before = source.copy(deep=True)
        result = mask_failed_reporting_cells(source, source)
        self.assertEqual(result.loc[0, 'Mensuel Actual reports'], 90)
        self.assertEqual(result.loc[0, 'Mensuel Reporting rate'], 90)
        self.assertTrue(result.loc[0, [
            'Hebdomadaire Actual reports', 'Hebdomadaire Reporting rate',
        ]].isna().all())
        self.assertEqual(result.loc[0, 'Reports_Actual'], 90)
        self.assertEqual(result.loc[0, 'Commentaire'], 'réponse hebdomadaire absente')
        self.assertEqual(result.attrs, source.attrs)
        pd.testing.assert_frame_equal(source, before)

    def test_other_unit_and_its_genuine_weekly_zero_are_unchanged(self):
        source = self.source()
        result = mask_failed_reporting_cells(source, source)
        pd.testing.assert_series_equal(result.iloc[1], source.iloc[1])
        self.assertEqual(result.loc[1, 'Hebdomadaire Actual reports'], 0)

    def test_strict_unit_mask_is_still_required_for_partial_combined_totals(self):
        source = self.source()
        cells = mask_failed_reporting_cells(source, source)
        totals = mask_failed_reporting_units(cells, source, ['Reports_Actual'])
        self.assertTrue(pd.isna(totals.loc[0, 'Reports_Actual']))
        self.assertEqual(totals.loc[0, 'Mensuel Actual reports'], 90)
        self.assertEqual(totals.loc[1, 'Reports_Actual'], 30)

    def test_explicit_ids_take_priority_over_misleading_names_and_row_order(self):
        source = self.source()
        frame = source.iloc[::-1].copy()
        frame['Organisation unit'] = ['Aba', 'Autre ZS']
        result = mask_failed_reporting_cells(frame, source)
        self.assertEqual(result.iloc[0]['Hebdomadaire Actual reports'], 0)
        self.assertTrue(pd.isna(result.iloc[1]['Hebdomadaire Actual reports']))
        self.assertEqual(result.iloc[1]['Mensuel Actual reports'], 90)

    def test_idless_view_uses_only_a_bijective_source_name_bridge(self):
        source = self.source()
        frame = source.drop(columns='Organisation unit ID').iloc[::-1].copy()
        result = mask_failed_reporting_cells(frame, source)
        self.assertEqual(result.iloc[0]['Hebdomadaire Actual reports'], 0)
        self.assertTrue(pd.isna(result.iloc[1]['Hebdomadaire Actual reports']))
        self.assertEqual(result.iloc[1]['Mensuel Actual reports'], 90)

    def test_homonyms_or_aliases_do_not_invent_a_source_join(self):
        source = self.source()
        source['Organisation unit'] = 'Même nom'
        frame = pd.DataFrame({
            'Organisation unit': ['Même nom'], 'Hebdomadaire Actual reports': [17.0],
        })
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(frame, source), frame)
        source = self.source()
        alias = source.iloc[:1].copy()
        alias['Organisation unit'] = 'Alias Aba'
        combined = pd.concat([source, alias], ignore_index=True)
        combined.attrs = source.attrs.copy()
        frame = source.iloc[:1].drop(columns='Organisation unit ID')
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(frame, combined), frame)

    def test_missing_dx_details_fall_back_to_certified_raw_columns_only(self):
        for missing_key in ['failed_metrics_by_unit', 'reporting_item_labels']:
            with self.subTest(missing_key=missing_key):
                source = self.source()
                source.attrs.pop(missing_key)
                result = mask_failed_reporting_cells(source, source)
                self.assertTrue(result.loc[0, source.attrs['reporting_columns']].isna().all())
                self.assertEqual(result.loc[0, 'Reports_Actual'], 90)
                pd.testing.assert_series_equal(result.iloc[1], source.iloc[1])

    def test_unknown_failed_dx_is_conservative_instead_of_presuming_healthy_raw_cells(self):
        source = self.source()
        source.attrs['failed_metrics_by_unit']['ZoneAAAAAAA'].append('UnknownAAAA.ACTUAL_REPORTS')
        result = mask_failed_reporting_cells(source, source)
        self.assertTrue(result.loc[0, source.attrs['reporting_columns']].isna().all())
        self.assertEqual(result.loc[0, 'Reports_Actual'], 90)

    def test_legacy_failure_without_certification_masks_numeric_cells(self):
        source = self.source()
        source.attrs = {'failed_units': ['ZoneAAAAAAA']}
        result = mask_failed_reporting_cells(source, source)
        self.assertTrue(result.loc[0, [
            'Mensuel Actual reports', 'Hebdomadaire Actual reports', 'Reports_Actual',
        ]].isna().all())
        self.assertEqual(result.loc[0, 'Organisation unit ID'], 'ZoneAAAAAAA')
        self.assertEqual(result.loc[0, 'Commentaire'], source.loc[0, 'Commentaire'])
        pd.testing.assert_series_equal(result.iloc[1], source.iloc[1])

    def test_certified_empty_columns_do_not_mask_derived_aggregates(self):
        source = self.source()
        source.attrs = {'failed_units': ['ZoneAAAAAAA'], 'reporting_columns': []}
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(source, source), source)

    def test_failed_dx_map_is_applied_even_if_failed_units_list_is_missing(self):
        source = self.source()
        source.attrs.pop('failed_units')
        result = mask_failed_reporting_cells(source, source)
        self.assertTrue(pd.isna(result.loc[0, 'Hebdomadaire Actual reports']))
        self.assertEqual(result.loc[0, 'Mensuel Actual reports'], 90)

    def test_non_numeric_failed_cells_are_masked_by_their_exact_dx_label(self):
        source = self.source()
        frame = source.copy()
        frame['Hebdomadaire Actual reports'] = ['17', '0']
        result = mask_failed_reporting_cells(frame, source)
        self.assertTrue(pd.isna(result.loc[0, 'Hebdomadaire Actual reports']))
        self.assertEqual(result.loc[1, 'Hebdomadaire Actual reports'], '0')
        self.assertEqual(result.loc[0, 'Mensuel Actual reports'], 90)

    def test_explicit_empty_cell_override_preserves_raw_fallback_despite_failed_totals(self):
        source = self.source()
        source.attrs['failed_cells_by_unit'] = {}
        result = mask_failed_reporting_cells(source, source)
        pd.testing.assert_frame_equal(result, source)
        totals = mask_failed_reporting_units(result, source, ['Reports_Actual'])
        self.assertTrue(pd.isna(totals.loc[0, 'Reports_Actual']))
        self.assertEqual(totals.loc[0, 'Mensuel Actual reports'], 90)

    def test_cell_override_uses_only_its_units_and_failed_dx_columns(self):
        source = self.source()
        source.attrs['failed_units'] = ['ZoneAAAAAAA', 'ZoneBBBBBBB']
        source.attrs['failed_cells_by_unit'] = {
            'ZoneAAAAAAA': ['WeeklyAAAAA.ACTUAL_REPORTS'],
        }
        result = mask_failed_reporting_cells(source, source)
        self.assertTrue(pd.isna(result.loc[0, 'Hebdomadaire Actual reports']))
        self.assertEqual(result.loc[0, 'Hebdomadaire Reporting rate'], 0)
        self.assertEqual(result.loc[0, 'Mensuel Actual reports'], 90)
        pd.testing.assert_series_equal(result.iloc[1], source.iloc[1])

    def test_empty_frames_absent_sources_and_unknown_failed_units_are_safe(self):
        source = self.source()
        frame = source.iloc[:0]
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(frame, source), frame)
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(source, None), source)
        source.attrs = {'failed_units': ['UnknownAAAA']}
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(source, source), source)


if __name__ == '__main__':
    unittest.main()
