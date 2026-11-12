"""Surveillance hebdomadaire agrégée avec les comptes mensuels, sans faux score."""

from copy import deepcopy
import unittest
from unittest import mock

import pandas as pd
import requests

from quality_reporting import fetch_quality_reporting, split_quality_reporting_query
from quality_score import metrics_from_reporting_source
from reporting_availability import mask_failed_reporting_units
from scoped_analytics import METRIC_LABELS


WEEKS = ['2026W36', '2026W37', '2026W38', '2026W39']
UNITS = ['ZoneAAAAAAA', 'ZoneBBBBBBB']


class QualityReportingTests(unittest.TestCase):
    def selection(self, *, weekly_only=False, same_names=False):
        datasets = [
            {'id': 'MonthlyAAAA', 'name': 'Mensuel', 'periodType': 'Monthly', 'is_weekly': False},
            {'id': 'WeeklyAAAAA', 'name': 'eSurveillance', 'periodType': 'Weekly', 'is_weekly': True},
        ]
        if weekly_only:
            datasets = datasets[1:]
        labels = {
            f"{record['id']}.{metric}": (
                f"Commun ({record['id']}) {suffix}" if same_names else f"{record['name']} {suffix}"
            )
            for record in datasets for metric, suffix in METRIC_LABELS.items()
        }
        params = [
            ('dimension', 'dx:' + ';'.join(labels)), ('dimension', 'ou:LEVEL-3;ProvinceAAA'),
            ('filter', 'pe:202609'), ('filter', 'CategoryAAA:OptionAAAAA'),
            ('skipRounding', 'false'), ('aggregationType', 'DEFAULT'), ('completedOnly', 'false'),
        ]
        return params, labels, datasets

    def client(self, *, weekly_failure=None, weekly_empty=None, weekly_zero=False):
        calls = []

        def get(endpoint, params=None):
            self.assertEqual(endpoint, 'analytics.json')
            dimensions = {value.split(':', 1)[0]: value.split(':', 1)[1]
                          for key, value in params if key == 'dimension'}
            unit = dimensions['ou']
            items = dimensions['dx'].split(';')
            weekly = items[0].startswith('WeeklyAAAAA.')
            calls.append({'unit': unit, 'items': items, 'params': list(params), 'weekly': weekly})
            if weekly and unit == weekly_failure:
                raise requests.ReadTimeout('private URL password=fixture')
            if weekly and unit == weekly_empty:
                return {'rows': []}
            if weekly:
                actual, expected, timely = (0, 68, 0) if weekly_zero else (68, 68, 67)
            else:
                actual, expected, timely = 90, 100, 80
            metrics = {'ACTUAL_REPORTS': actual, 'EXPECTED_REPORTS': expected,
                       'ACTUAL_REPORTS_ON_TIME': timely,
                       'REPORTING_RATE': actual / expected * 100,
                       'REPORTING_RATE_ON_TIME': timely / expected * 100}
            return {
                'headers': [{'name': name} for name in ['ou', 'dx', 'value']],
                'rows': [[unit, item, metrics[item.rsplit('.', 1)[1]]] for item in items],
                'metaData': {'items': {unit: {'displayName': 'Nom ' + unit}}},
            }
        client = mock.Mock()
        client.get.side_effect = get
        return client, calls

    def targets(self, units=UNITS):
        return pd.DataFrame([{'Zone_ID': identifier, 'Zone': identifier} for identifier in units])

    def test_split_uses_separate_monthly_and_four_week_filters_without_mutating_extra_filters(self):
        params, labels, datasets = self.selection()
        before = deepcopy((params, labels, datasets))
        groups = split_quality_reporting_query(params, labels, datasets, WEEKS)
        self.assertEqual([group['kind'] for group in groups], ['monthly', 'weekly'])
        self.assertEqual(groups[0]['periods'], ('202609',))
        self.assertEqual(groups[1]['periods'], tuple(WEEKS))
        self.assertIn(('filter', 'pe:2026W36;2026W37;2026W38;2026W39'), groups[1]['params'])
        for group in groups:
            for item in params[3:]:
                self.assertIn(item, group['params'])
            self.assertEqual(len(group['labels']), 5)
            self.assertEqual(set(group['labels'].values()), {labels[key] for key in group['labels']})
        self.assertEqual((params, labels, datasets), before)

    def test_weighted_counts_combine_weekly_once_and_keep_each_unit(self):
        params, labels, datasets = self.selection()
        client, calls = self.client()
        progress = mock.Mock()
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS, WEEKS, on_progress=progress)
        self.assertEqual(len(calls), 4)
        self.assertEqual([call['unit'] for call in calls], UNITS + UNITS)
        self.assertEqual([call['weekly'] for call in calls], [False, False, True, True])
        self.assertEqual([call.args[0] for call in progress.call_args_list], ['monthly', 'monthly', 'weekly', 'weekly'])
        self.assertEqual(result['Organisation unit ID'].tolist(), UNITS)
        self.assertEqual(result['eSurveillance Actual reports'].tolist(), [68, 68])
        metrics = metrics_from_reporting_source(result, self.targets())
        self.assertEqual(metrics['Actual_Reports'].tolist(), [158, 158])
        self.assertEqual(metrics['Complétude_Globale'].tolist(), [94.05, 94.05])
        self.assertEqual(metrics['Promptitude_Globale'].tolist(), [87.5, 87.5])
        self.assertEqual(metrics['Nombre_Datasets_95'].tolist(), [1, 1])
        self.assertEqual(result.attrs['reporting_status'], 'complete')
        self.assertEqual(result.attrs['reporting_periods'], {'monthly': ('202609',), 'weekly': tuple(WEEKS)})
        self.assertEqual(result.attrs['weekly_reporting_columns'], [label for item, label in labels.items() if item.startswith('WeeklyAAAAA.')])
        self.assertEqual(result.attrs['monthly_reporting_columns'], [label for item, label in labels.items() if item.startswith('MonthlyAAAA.')])
        self.assertEqual(result.attrs['failed_units'], [])
        self.assertEqual(result.attrs['successful_units'], UNITS)
        self.assertEqual(len(result.attrs['reporting_columns']), 10)
        self.assertEqual({entry['Source'] for entry in result.attrs['reporting_requests']}, {'monthly', 'weekly'})

    def test_weekly_failure_masks_incomplete_unit_without_erasing_other_unit_or_using_zero(self):
        params, labels, datasets = self.selection()
        client, _ = self.client(weekly_failure=UNITS[0])
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS, WEEKS)
        self.assertEqual(result.attrs['failed_units'], UNITS[:1])
        self.assertEqual(result.attrs['successful_units'], UNITS[1:])
        self.assertEqual(result.attrs['reporting_status'], 'partial')
        self.assertEqual(len(result.attrs['failed_metrics_by_unit'][UNITS[0]]), 5)
        metrics = mask_failed_reporting_units(metrics_from_reporting_source(result, self.targets()), result)
        self.assertTrue(metrics.iloc[0, 1:].isna().all())
        self.assertEqual(metrics.loc[1, 'Actual_Reports'], 158)
        self.assertEqual(metrics.loc[1, 'Nombre_Datasets_95'], 1)
        self.assertNotIn('password', str(result.attrs))
        self.assertNotIn('private', str(result.attrs))

    def test_successful_empty_surveillance_is_missing_instead_of_a_complete_monthly_only_score(self):
        params, labels, datasets = self.selection()
        client, _ = self.client(weekly_empty=UNITS[0])
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS, WEEKS)
        self.assertEqual(result.attrs['failed_units'], UNITS[:1])
        self.assertTrue(any('surveillance hebdomadaire indisponible' in text for text in result.attrs['reporting_warnings']))
        metrics = mask_failed_reporting_units(metrics_from_reporting_source(result, self.targets()), result)
        self.assertTrue(metrics.iloc[0, 1:].isna().all())
        self.assertEqual(metrics.loc[1, 'Actual_Reports'], 158)

    def test_genuine_weekly_zero_with_positive_expectation_is_included(self):
        params, labels, datasets = self.selection()
        client, _ = self.client(weekly_zero=True)
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS[:1], WEEKS)
        self.assertEqual(result.attrs['failed_units'], [])
        metrics = metrics_from_reporting_source(result, self.targets(UNITS[:1]))
        self.assertEqual(metrics.loc[0, 'Actual_Reports'], 90)
        self.assertEqual(metrics.loc[0, 'Complétude_Globale'], 53.57)
        self.assertEqual(metrics.loc[0, 'Nombre_Datasets_95'], 0)

    def test_weekly_only_keeps_one_dataset_and_deduplicates_periods(self):
        params, labels, datasets = self.selection(weekly_only=True)
        client, calls = self.client()
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS[:1], WEEKS + WEEKS)
        self.assertEqual(len(calls), 1)
        self.assertIn(('filter', 'pe:' + ';'.join(WEEKS)), calls[0]['params'])
        self.assertEqual(result.attrs['reporting_periods'], {'weekly': tuple(WEEKS)})
        metrics = metrics_from_reporting_source(result, self.targets(UNITS[:1]))
        self.assertEqual(metrics.loc[0, 'Actual_Reports'], 68)
        self.assertEqual(metrics.loc[0, 'Nombre_Datasets_95'], 1)

    def test_weekly_zero_expected_and_actual_is_valid_nonapplicable_not_an_extra_dataset(self):
        params, labels, datasets = self.selection()
        client, _ = self.client()
        original = client.get.side_effect

        def get(endpoint, params=None):
            response = original(endpoint, params)
            for row in response['rows']:
                if row[1].startswith('WeeklyAAAAA.'):
                    row[2] = None if 'RATE' in row[1] else 0
            return response
        client.get.side_effect = get
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS[:1], WEEKS)
        self.assertEqual(result.attrs['failed_units'], [])
        self.assertEqual(result.attrs['reporting_status'], 'complete')
        metrics = metrics_from_reporting_source(result, self.targets(UNITS[:1]))
        self.assertEqual(metrics.loc[0, 'Actual_Reports'], 90)
        self.assertEqual(metrics.loc[0, 'Complétude_Globale'], 90)
        self.assertEqual(metrics.loc[0, 'Nombre_Datasets_95'], 0)

    def test_monthly_only_does_not_require_or_substitute_week_periods(self):
        params, labels, datasets = self.selection()
        monthly_labels = {item: label for item, label in labels.items() if item.startswith('MonthlyAAAA.')}
        params[0] = ('dimension', 'dx:' + ';'.join(monthly_labels))
        client, calls = self.client()
        result = fetch_quality_reporting(client, params, monthly_labels, datasets, UNITS[:1], None)
        self.assertEqual(len(calls), 1)
        self.assertIn(('filter', 'pe:202609'), calls[0]['params'])
        self.assertEqual(result.attrs['weekly_reporting_columns'], [])
        self.assertEqual(result.attrs['reporting_periods'], {'monthly': ('202609',)})

    def test_homonymous_dataset_names_use_complete_query_labels_unchanged(self):
        params, labels, datasets = self.selection(same_names=True)
        client, _ = self.client()
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS[:1], WEEKS)
        self.assertEqual(set(result.columns[2:]), set(labels.values()))
        self.assertEqual(len(result.columns), 12)
        self.assertEqual(metrics_from_reporting_source(result, self.targets(UNITS[:1])).loc[0, 'Actual_Reports'], 158)

    def test_uid_only_labels_are_joined_by_item_id_after_server_names_are_added(self):
        params, labels, datasets = self.selection()
        labels = {item: f'{item.rsplit(".", 1)[0]} {METRIC_LABELS[item.rsplit(".", 1)[1]]}' for item in labels}
        client, _ = self.client()
        original = client.get.side_effect

        def get(endpoint, params=None):
            response = original(endpoint, params)
            response['metaData']['items'].update({
                'MonthlyAAAA': {'displayName': 'Consultations'},
                'WeeklyAAAAA': {'displayName': 'Surveillance enrichie'},
            })
            return response

        client.get.side_effect = get
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS[:1], WEEKS)
        expected = {
            f'{dataset}.{metric}': f'{name} {suffix}'
            for dataset, name in [('MonthlyAAAA', 'Consultations'), ('WeeklyAAAAA', 'Surveillance enrichie')]
            for metric, suffix in METRIC_LABELS.items()
        }
        self.assertEqual(result.attrs['reporting_item_labels'], expected)
        self.assertEqual(result.attrs['weekly_reporting_columns'], [expected[item] for item in expected if item.startswith('WeeklyAAAAA.')])
        self.assertEqual(result[expected['MonthlyAAAA.ACTUAL_REPORTS']].tolist(), [90])
        self.assertEqual(result[expected['WeeklyAAAAA.ACTUAL_REPORTS']].tolist(), [68])
        self.assertEqual(metrics_from_reporting_source(result, self.targets(UNITS[:1])).loc[0, 'Actual_Reports'], 158)

    def test_metadata_homonyms_from_separate_sources_keep_distinct_dataset_families(self):
        params, labels, datasets = self.selection()
        labels = {item: f'{item.rsplit(".", 1)[0]} {METRIC_LABELS[item.rsplit(".", 1)[1]]}' for item in labels}
        client, _ = self.client()
        original = client.get.side_effect

        def get(endpoint, params=None):
            response = original(endpoint, params)
            response['metaData']['items'].update({dataset: {'displayName': 'Commun'} for dataset in ['MonthlyAAAA', 'WeeklyAAAAA']})
            return response

        client.get.side_effect = get
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS[:1], WEEKS)
        expected = {f'{dataset}.{metric}': f'Commun ({dataset}) {suffix}'
                    for dataset in ['MonthlyAAAA', 'WeeklyAAAAA'] for metric, suffix in METRIC_LABELS.items()}
        self.assertEqual(result.attrs['reporting_item_labels'], expected)
        self.assertEqual(len(set(result.columns)), 12)
        self.assertEqual(result[expected['MonthlyAAAA.ACTUAL_REPORTS']].tolist(), [90])
        self.assertEqual(result[expected['WeeklyAAAAA.ACTUAL_REPORTS']].tolist(), [68])
        self.assertEqual(metrics_from_reporting_source(result, self.targets(UNITS[:1])).loc[0, 'Nombre_Datasets_95'], 1)

    def test_one_enriched_metric_supplies_a_coherent_prefix_for_its_whole_family(self):
        params, labels, datasets = self.selection()
        labels = {item: f'{item.rsplit(".", 1)[0]} {METRIC_LABELS[item.rsplit(".", 1)[1]]}' for item in labels}
        client, _ = self.client()
        original = client.get.side_effect

        def get(endpoint, params=None):
            response = original(endpoint, params)
            response['metaData']['items']['MonthlyAAAA.ACTUAL_REPORTS'] = {'displayName': 'Consultations Actual reports'}
            return response

        client.get.side_effect = get
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS[:1], WEEKS)
        mapping = result.attrs['reporting_item_labels']
        self.assertEqual({mapping[f'MonthlyAAAA.{metric}'] for metric in METRIC_LABELS},
                         {f'Consultations {suffix}' for suffix in METRIC_LABELS.values()})
        self.assertEqual(metrics_from_reporting_source(result, self.targets(UNITS[:1])).loc[0, 'Actual_Reports'], 158)

    def test_failed_metrics_keep_their_item_identity_when_other_unit_metadata_enriches_names(self):
        params, labels, datasets = self.selection()
        labels = {item: f'{item.rsplit(".", 1)[0]} {METRIC_LABELS[item.rsplit(".", 1)[1]]}' for item in labels}
        client, _ = self.client(weekly_failure=UNITS[0])
        original = client.get.side_effect

        def get(endpoint, params=None):
            response = original(endpoint, params)
            response['metaData']['items'].update({
                'MonthlyAAAA': {'displayName': 'Mensuel enrichi'},
                'WeeklyAAAAA': {'displayName': 'Hebdomadaire enrichi'},
            })
            return response

        client.get.side_effect = get
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS, WEEKS)
        self.assertEqual(result.attrs['failed_metrics_by_unit'], {UNITS[0]: [f'WeeklyAAAAA.{metric}' for metric in METRIC_LABELS]})
        self.assertTrue(result.loc[0, result.attrs['weekly_reporting_columns']].isna().all())
        self.assertEqual(result.loc[1, result.attrs['reporting_item_labels']['WeeklyAAAAA.ACTUAL_REPORTS']], 68)
        masked = mask_failed_reporting_units(metrics_from_reporting_source(result, self.targets()), result)
        self.assertTrue(masked.iloc[0, 1:].isna().all())
        self.assertEqual(masked.loc[1, 'Actual_Reports'], 158)

    def test_missing_week_selection_preserves_received_monthly_values_without_a_weekly_request(self):
        params, labels, datasets = self.selection()
        client, calls = self.client()
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS[:1], [])
        self.assertEqual(len(calls), 1)
        self.assertFalse(calls[0]['weekly'])
        self.assertEqual(result['Mensuel Actual reports'].tolist(), [90])
        self.assertTrue(result[result.attrs['weekly_reporting_columns']].isna().all().all())
        self.assertEqual(result.attrs['reporting_periods'], {'monthly': ('202609',), 'weekly': ()})
        self.assertEqual(result.attrs['failed_units'], UNITS[:1])
        self.assertEqual(result.attrs['failed_metrics_by_unit'], {UNITS[0]: [f'WeeklyAAAAA.{metric}' for metric in METRIC_LABELS]})
        self.assertTrue(any('Aucune semaine de surveillance sélectionnée' in warning for warning in result.attrs['reporting_warnings']))
        masked = mask_failed_reporting_units(metrics_from_reporting_source(result, self.targets(UNITS[:1])), result)
        self.assertTrue(masked.iloc[0, 1:].isna().all())

    def test_weekly_only_missing_selection_keeps_known_units_and_schema_without_any_request(self):
        params, labels, datasets = self.selection(weekly_only=True)
        client, _ = self.client()
        result = fetch_quality_reporting(client, params, labels, datasets, UNITS, [])
        client.get.assert_not_called()
        self.assertEqual(result['Organisation unit ID'].tolist(), UNITS)
        self.assertEqual(result.attrs['failed_units'], UNITS)
        self.assertEqual(result.attrs['reporting_item_labels'], labels)
        self.assertTrue(result[result.attrs['weekly_reporting_columns']].isna().all().all())
        self.assertEqual(result.attrs['reporting_status'], 'unavailable')

    def test_unsupported_period_type_or_invalid_week_periods_never_call_api(self):
        params, labels, datasets = self.selection()
        client, _ = self.client()
        invalid = deepcopy(datasets)
        invalid[0]['periodType'] = 'Quarterly'
        for records, weeks in [(invalid, WEEKS), (datasets[:1], WEEKS), (datasets, ['2026W54'])]:
            with self.subTest(records=records, weeks=weeks):
                with self.assertRaises(ValueError):
                    fetch_quality_reporting(client, params, labels, records, UNITS, weeks)
        client.get.assert_not_called()

    def test_no_requested_units_never_expands_scope(self):
        params, labels, datasets = self.selection()
        client, _ = self.client()
        result = fetch_quality_reporting(client, params, labels, datasets, [], WEEKS)
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs['reporting_status'], 'unavailable')
        client.get.assert_not_called()


if __name__ == '__main__':
    unittest.main()
