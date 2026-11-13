"""Un échec de chargement n'est ni un zéro ni une moyenne partielle."""

import unittest
from copy import deepcopy

import pandas as pd

from quality_score import metrics_from_reporting_source
from reporting_availability import (
    mask_failed_reporting_cells, mask_failed_reporting_units, select_reporting_source,
)
from scoped_analytics import METRIC_LABELS, analytics_to_reporting_frame


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
    def test_generic_indicator_failure_does_not_erase_healthy_reporting_cells(self):
        labels = {
            'MonthlyAAAA.ACTUAL_REPORTS': 'Mensuel Actual reports',
            'IndicatorAA': 'Indicateur indépendant',
        }
        for rows in [[], [
            ['ZoneAAAAAAA', 'MonthlyAAAA.ACTUAL_REPORTS', 90],
            ['ZoneAAAAAAA', 'IndicatorAA', 25],
        ]]:
            with self.subTest(empty=not rows):
                source = analytics_to_reporting_frame({
                    'headers': [{'name': name} for name in ['ou', 'dx', 'value']],
                    'rows': rows,
                }, labels)
                self.assertEqual(source.attrs['reporting_item_labels'], labels)
                self.assertEqual(source.attrs['reporting_columns'], ['Mensuel Actual reports'])
                if rows:
                    source.attrs.update({
                        'failed_units': ['ZoneAAAAAAA'],
                        'failed_metrics_by_unit': {'ZoneAAAAAAA': ['IndicatorAA']},
                    })
                    result = mask_failed_reporting_cells(source, source)
                    self.assertEqual(result['Mensuel Actual reports'].tolist(), [90])
                    self.assertTrue(result['Indicateur indépendant'].isna().all())

    def source(self):
        frame = pd.DataFrame({
            'Organisation unit ID': ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
            'Organisation unit': ['Aba', 'Autre ZS'],
            'Mensuel Actual reports': [90.0, 0.0],
            'Hebdo Actual reports': [999.0, 0.0],
            'Indicateur indépendant': [25, 30],
        })
        frame.attrs.update({
            'failed_units': ['ZoneAAAAAAA'],
            'failed_metrics_by_unit': {'ZoneAAAAAAA': ['WeeklyAAAAA.ACTUAL_REPORTS']},
            'reporting_item_labels': {
                'MonthlyAAAA.ACTUAL_REPORTS': 'Mensuel Actual reports',
                'WeeklyAAAAA.ACTUAL_REPORTS': 'Hebdo Actual reports',
            },
            'reporting_columns': ['Mensuel Actual reports', 'Hebdo Actual reports'],
        })
        return frame

    def test_only_failed_dataset_cells_are_masked_and_other_unit_zero_is_preserved(self):
        source = self.source()
        before = source.copy(deep=True)
        result = mask_failed_reporting_cells(source, source)
        self.assertEqual(result['Mensuel Actual reports'].tolist(), [90, 0])
        self.assertTrue(pd.isna(result.loc[0, 'Hebdo Actual reports']))
        self.assertEqual(result.loc[1, 'Hebdo Actual reports'], 0)
        self.assertEqual(result['Indicateur indépendant'].tolist(), [25, 30])
        self.assertEqual(result.attrs, source.attrs)
        pd.testing.assert_frame_equal(source, before)

    def test_bijective_display_bridge_handles_reordered_rows(self):
        source = self.source()
        frame = source.drop(columns='Organisation unit ID').iloc[::-1].copy()
        result = mask_failed_reporting_cells(frame, source)
        self.assertEqual(result.loc[1, 'Hebdo Actual reports'], 0)
        self.assertTrue(pd.isna(result.loc[0, 'Hebdo Actual reports']))
        self.assertEqual(result.loc[0, 'Mensuel Actual reports'], 90)

    def test_homonymous_names_do_not_invent_a_unit_match(self):
        source = self.source()
        source['Organisation unit'] = 'Même nom'
        frame = source.drop(columns='Organisation unit ID').copy()
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(frame, source), frame)
        alias = self.source().iloc[:1].copy()
        alias['Organisation unit'] = 'Alias Aba'
        source = pd.concat([self.source(), alias], ignore_index=True)
        source.attrs = self.source().attrs
        frame = self.source().drop(columns='Organisation unit ID')
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(frame, source), frame)

    def test_exact_ids_win_over_wrong_display_names(self):
        source = self.source()
        frame = source.copy()
        frame['Organisation unit'] = ['Autre ZS', 'Aba']
        result = mask_failed_reporting_cells(frame, source)
        self.assertTrue(pd.isna(result.loc[0, 'Hebdo Actual reports']))
        self.assertEqual(result.loc[1, 'Hebdo Actual reports'], 0)

    def test_unmapped_failure_uses_only_certified_columns(self):
        source = self.source()
        source.attrs['reporting_item_labels'] = {}
        result = mask_failed_reporting_cells(source, source)
        self.assertTrue(result.loc[0, ['Mensuel Actual reports', 'Hebdo Actual reports']].isna().all())
        self.assertEqual(result.loc[0, 'Indicateur indépendant'], 25)

    def test_legacy_failure_without_metadata_keeps_conservative_numeric_mask(self):
        source = self.source()
        source.attrs = {'failed_units': ['ZoneAAAAAAA']}
        result = mask_failed_reporting_cells(source, source)
        self.assertTrue(result.loc[0, ['Mensuel Actual reports', 'Hebdo Actual reports', 'Indicateur indépendant']].isna().all())
        self.assertEqual(result.loc[0, 'Organisation unit'], 'Aba')
        self.assertEqual(result.loc[0, 'Organisation unit ID'], 'ZoneAAAAAAA')
        self.assertEqual(result.loc[1, 'Mensuel Actual reports'], 0)

    def test_per_metric_failure_alone_is_used_and_missing_failure_is_noop(self):
        source = self.source()
        source.attrs.pop('failed_units')
        result = mask_failed_reporting_cells(source, source)
        self.assertTrue(pd.isna(result.loc[0, 'Hebdo Actual reports']))
        self.assertEqual(result.loc[0, 'Mensuel Actual reports'], 90)
        source.attrs = {}
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(source, source), source)
        pd.testing.assert_frame_equal(mask_failed_reporting_cells(source.iloc[:0], source), source.iloc[:0])

    def test_raw_preservation_does_not_weaken_the_combined_total_mask(self):
        source = self.source()
        total = pd.DataFrame({'Zone_ID': ['ZoneAAAAAAA', 'ZoneBBBBBBB'], 'Reports_Actual': [90.0, 0.0]})
        result = mask_failed_reporting_units(total, source)
        self.assertTrue(pd.isna(result.loc[0, 'Reports_Actual']))
        self.assertEqual(result.loc[1, 'Reports_Actual'], 0)


class ReportingSourceSelectionTests(unittest.TestCase):
    def source(self):
        labels = {
            f'MonthlyAAAA.{metric}': f'Mensuel {suffix}'
            for metric, suffix in METRIC_LABELS.items()
        }
        # Ce libellé ressemble à un compte, mais son DX n'est pas du reporting.
        labels['IndicatorAA'] = 'Indicateur Actual reports'
        values = {
            'REPORTING_RATE': 90, 'REPORTING_RATE_ON_TIME': 80,
            'ACTUAL_REPORTS': 90, 'ACTUAL_REPORTS_ON_TIME': 80, 'EXPECTED_REPORTS': 100,
        }
        rows = []
        for identifier, zero in [('ZoneAAAAAAA', False), ('ZoneBBBBBBB', True)]:
            for metric, value in values.items():
                rows.append([identifier, f'MonthlyAAAA.{metric}',
                             value if not zero or metric == 'EXPECTED_REPORTS' else 0])
            rows.append([identifier, 'IndicatorAA', 999])
        frame = analytics_to_reporting_frame({
            'headers': [{'name': name} for name in ['ou', 'dx', 'value']],
            'rows': rows,
        }, labels)
        frame['Organisation unit'] = ['Même nom', 'Même nom']
        frame.attrs.update({
            'reporting_status': 'partial', 'reporting_warnings': ['Diagnostic conservé'],
            'quality_source': 'Source courante', 'successful_units': ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
        })
        return frame

    def targets(self):
        return pd.DataFrame([
            {'Zone_ID': identifier, 'Zone': 'Même nom'}
            for identifier in ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC']
        ])

    def metrics(self, source):
        metrics = metrics_from_reporting_source(source, self.targets())
        return mask_failed_reporting_units(metrics, source).set_index('Zone_ID')

    def test_known_generic_failure_keeps_healthy_reporting_and_observed_zeros(self):
        source = self.source()
        source.attrs.update({
            'failed_units': ['ZoneAAAAAAA'],
            'failed_metrics_by_unit': {'ZoneAAAAAAA': ['IndicatorAA']},
        })
        selected = select_reporting_source(source)
        self.assertNotIn('Indicateur Actual reports', selected)
        self.assertEqual(selected.attrs['failed_units'], [])
        self.assertEqual(selected.attrs['failed_metrics_by_unit'], {})
        metrics = self.metrics(selected)
        self.assertEqual(metrics.loc['ZoneAAAAAAA'].tolist(), [90, 80, 90, 0])
        self.assertEqual(metrics.loc['ZoneBBBBBBB'].tolist(), [0, 0, 0, 0])
        self.assertTrue(metrics.loc['ZoneCCCCCCC'].isna().all())
        self.assertEqual(selected.attrs['reporting_warnings'], ['Diagnostic conservé'])
        self.assertEqual(selected.attrs['quality_source'], 'Source courante')

    def test_reporting_failure_remains_strict_for_its_uid_despite_homonyms(self):
        source = self.source()
        source.attrs.update({
            'failed_units': ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
            'failed_metrics_by_unit': {
                'ZoneAAAAAAA': ['IndicatorAA'],
                'ZoneBBBBBBB': ['IndicatorAA', 'MonthlyAAAA.ACTUAL_REPORTS'],
            },
        })
        selected = select_reporting_source(source)
        self.assertEqual(selected.attrs['failed_units'], ['ZoneBBBBBBB'])
        self.assertEqual(selected.attrs['failed_metrics_by_unit'], {
            'ZoneBBBBBBB': ['MonthlyAAAA.ACTUAL_REPORTS'],
        })
        metrics = self.metrics(selected)
        self.assertEqual(metrics.loc['ZoneAAAAAAA'].tolist(), [90, 80, 90, 0])
        self.assertTrue(metrics.loc['ZoneBBBBBBB'].isna().all())
        self.assertTrue(metrics.loc['ZoneCCCCCCC'].isna().all())

    def test_metric_only_reporting_failure_still_marks_its_unit(self):
        source = self.source()
        source.attrs['failed_metrics_by_unit'] = {' ZoneAAAAAAA ': ['MonthlyAAAA.EXPECTED_REPORTS']}
        selected = select_reporting_source(source)
        self.assertEqual(selected.attrs['failed_units'], ['ZoneAAAAAAA'])
        self.assertTrue(self.metrics(selected).loc['ZoneAAAAAAA'].isna().all())

    def test_unknown_or_ambiguous_failure_metadata_remains_strict(self):
        cases = [
            {}, {'ZoneAAAAAAA': []}, {'ZoneAAAAAAA': None},
            {'ZoneAAAAAAA': ['UnmappedAAA']},
            {'ZoneAAAAAAA': ['UnmappedAAA.ACTUAL_REPORTS']},
            {'ZoneAAAAAAA': ['IndicatorAA', None]},
        ]
        for failures in cases:
            with self.subTest(failures=failures):
                source = self.source()
                source.attrs.update({
                    'failed_units': ['ZoneAAAAAAA'], 'failed_metrics_by_unit': failures,
                })
                selected = select_reporting_source(source)
                self.assertEqual(selected.attrs['failed_units'], ['ZoneAAAAAAA'])
                self.assertEqual(selected.attrs['failed_metrics_by_unit'], {})
                self.assertTrue(self.metrics(selected).loc['ZoneAAAAAAA'].isna().all())

    def test_missing_or_conflicting_labels_do_not_clear_a_failure(self):
        for labels in [{}, {'IndicatorAA': 'Mensuel Actual reports'}]:
            with self.subTest(labels=labels):
                source = self.source()
                source.attrs.update({
                    'failed_units': 'ZoneAAAAAAA',
                    'failed_metrics_by_unit': {'ZoneAAAAAAA': 'IndicatorAA'},
                    'reporting_item_labels': labels,
                })
                selected = select_reporting_source(source)
                self.assertEqual(selected.attrs['failed_units'], ['ZoneAAAAAAA'])
                self.assertTrue(self.metrics(selected).loc['ZoneAAAAAAA'].isna().all())

    def test_absent_reporting_values_are_preserved_and_never_become_zero(self):
        source = self.source()
        source[source.attrs['reporting_columns']] = float('nan')
        selected = select_reporting_source(source)
        self.assertTrue(selected[selected.attrs['reporting_columns']].isna().all().all())
        self.assertTrue(self.metrics(selected).isna().all().all())
        self.assertEqual(selected.attrs['failed_units'], [])

    def test_successfully_absent_dataset_does_not_erase_active_dataset_values(self):
        source = self.source()
        for metric, suffix in METRIC_LABELS.items():
            column = f'Inactif {suffix}'
            source[column] = float('nan')
            source.attrs['reporting_columns'].append(column)
            source.attrs['reporting_item_labels'][f'InactiveAAA.{metric}'] = column
        selected = select_reporting_source(source)
        self.assertEqual(self.metrics(selected).loc['ZoneAAAAAAA'].tolist(), [90, 80, 90, 0])
        self.assertTrue(selected[[f'Inactif {suffix}' for suffix in METRIC_LABELS.values()]].isna().all().all())

    def test_missing_uid_is_not_replaced_by_a_homonymous_label(self):
        source = self.source().drop(columns='Organisation unit ID')
        selected = select_reporting_source(source)
        self.assertTrue(self.metrics(selected).isna().all().all())

    def test_legacy_source_without_certification_is_copied_conservatively(self):
        source = self.source()
        source.attrs = {'failed_units': ['ZoneAAAAAAA']}
        selected = select_reporting_source(source)
        pd.testing.assert_frame_equal(selected, source)
        self.assertEqual(selected.attrs, source.attrs)
        self.assertIsNot(selected, source)
        self.assertIn('Indicateur Actual reports', selected)
        self.assertTrue(self.metrics(selected).loc['ZoneAAAAAAA'].isna().all())

    def test_empty_certification_has_no_numeric_reporting_evidence(self):
        source = self.source()
        source.attrs['reporting_columns'] = []
        selected = select_reporting_source(source)
        self.assertEqual(selected.columns.tolist(), ['Organisation unit', 'Organisation unit ID'])
        self.assertTrue(self.metrics(selected).isna().all().all())

    def test_projection_does_not_mutate_source_or_loading_metadata(self):
        source = self.source()
        source.attrs.update({
            'failed_units': ['ZoneAAAAAAA'],
            'failed_metrics_by_unit': {'ZoneAAAAAAA': ['IndicatorAA']},
        })
        before, attrs_before = source.copy(deep=True), deepcopy(source.attrs)
        selected = select_reporting_source(source)
        pd.testing.assert_frame_equal(source, before)
        self.assertEqual(source.attrs, attrs_before)
        self.assertEqual(selected['Organisation unit ID'].tolist(), source['Organisation unit ID'].tolist())
        self.assertEqual(selected['Organisation unit'].tolist(), ['Même nom', 'Même nom'])
        self.assertEqual(selected.attrs['successful_units'], attrs_before['successful_units'])
        empty = select_reporting_source(source.iloc[:0])
        self.assertTrue(empty.empty)
        self.assertEqual(empty.columns.tolist(), selected.columns.tolist())
        self.assertTrue(select_reporting_source(None).empty)


if __name__ == '__main__':
    unittest.main()
