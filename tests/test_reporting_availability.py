"""Un échec de chargement n'est ni un zéro ni une moyenne partielle."""

import unittest

import pandas as pd

from reporting_availability import mask_failed_reporting_units


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


if __name__ == '__main__':
    unittest.main()
