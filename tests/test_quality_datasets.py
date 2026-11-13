"""Sélection des datasets accessibles, indépendante du contenu du favori."""

from copy import deepcopy
import unittest
from unittest import mock

from quality_datasets import build_dataset_catalog, favorite_dataset_ids, select_quality_datasets
from scoped_analytics import METRIC_LABELS, build_quality_reporting_query


class DatasetCatalogTests(unittest.TestCase):
    def test_all_96_datasets_are_retained_regardless_of_names_or_periods(self):
        datasets = [
            {'id': f'Data{number:07d}', 'displayName': f'Dataset {number}', 'periodType': 'Monthly'}
            for number in range(96)
        ]
        datasets[0].update(displayName='Service tertiaire', periodType='Weekly')
        datasets[1]['displayName'] = 'Surveillance hebdomadaire'
        response = {'dataSets': datasets, 'pager': {'total': 96}}
        before = deepcopy(response)
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            catalog = build_dataset_catalog(response)
        self.assertEqual(len(catalog), 96)
        self.assertEqual({entry['id'] for entry in catalog}, {entry['id'] for entry in datasets})
        self.assertEqual(response, before)
        self.assertEqual(next(entry for entry in catalog if entry['id'] == 'Data0000000')['periodType'], 'Weekly')

    def test_homonymous_datasets_keep_separate_uids_and_missing_names_use_uid(self):
        catalog = build_dataset_catalog({'dataSets': [
            {'id': 'DatasetAAAA', 'displayName': 'Même nom'},
            {'id': 'DatasetBBBB', 'displayName': 'Même nom'},
            {'id': 'DatasetCCCC', 'displayName': '  ', 'name': ' Nom DHIS2 '},
            {'id': 'DatasetDDDD'},
        ]})
        self.assertEqual(dict((entry['id'], entry['displayName']) for entry in catalog), {
            'DatasetAAAA': 'Même nom', 'DatasetBBBB': 'Même nom',
            'DatasetCCCC': 'Nom DHIS2', 'DatasetDDDD': 'DatasetDDDD',
        })

    def test_compatible_duplicate_uids_merge_without_dropping_metadata(self):
        catalog = build_dataset_catalog({'dataSets': [
            {'id': 'DatasetAAAA'},
            {'id': 'DatasetAAAA', 'displayName': ' Consultations '},
            {'id': 'DatasetAAAA', 'name': 'Consultations', 'periodType': 'Monthly'},
        ]})
        self.assertEqual(catalog, [{'id': 'DatasetAAAA', 'displayName': 'Consultations', 'periodType': 'Monthly'}])

    def test_conflicting_names_and_period_types_are_rejected(self):
        for other in [
            {'displayName': 'Autre nom'}, {'displayName': 'Consultations', 'periodType': 'Weekly'},
        ]:
            with self.subTest(other=other), self.assertRaises(ValueError):
                build_dataset_catalog({'dataSets': [
                    {'id': 'DatasetAAAA', 'displayName': 'Consultations', 'periodType': 'Monthly'},
                    {'id': 'DatasetAAAA', **other},
                ]})

    def test_invalid_response_entries_and_uid_formats_are_rejected(self):
        responses = [None, [], {}, {'dataSets': None}, {'dataSets': {}}, {'dataSets': 'abc'},
                     {'dataSets': [None]}, {'dataSets': ['DatasetAAAA']}]
        responses += [{'dataSets': [{'id': uid}]} for uid in [
            None, 123, True, '', 'short', '1atasetAAAA', 'DatasetAAA_', ' DatasetAAAA', 'DatasetAAAA ',
        ]]
        for response in responses:
            with self.subTest(response=response), self.assertRaises(ValueError):
                build_dataset_catalog(response)
        self.assertEqual(build_dataset_catalog({'dataSets': []}), [])


class QualityDatasetSelectionTests(unittest.TestCase):
    def favorite(self):
        return {
            'id': 'FavoriteAAA',
            'dataDimensionItems': [
                {'indicator': {'id': 'IndicatorAA', 'displayName': 'Faux Actual reports'}},
                {'reportingRate': {'dataSet': {'id': 'DatasetOLD1'}, 'metric': 'REPORTING_RATE'}},
            ],
            'columns': [{'id': 'dx', 'items': [{'id': 'IndicatorBB'}], 'custom': {'keep': True}}],
            'rows': [
                {'dimension': 'dx', 'items': [{'id': 'DatasetOLD2.EXPECTED_REPORTS'}]},
                {'id': 'ou', 'items': [{'id': 'OldZoneAAAA'}]},
            ],
            'filters': [
                {'dimension': 'dx', 'items': [{'id': 'DatasetOLD3.ACTUAL_REPORTS'}]},
                {'dimension': 'pe', 'items': [{'id': '202401'}]},
                {'dimension': 'CategoryAAA', 'items': [{'id': 'OptionAAAAA'}]},
            ],
            'categoryDimensions': [{'category': {'id': 'CategoryAAA'}, 'categoryOptions': [{'id': 'OptionAAAAA'}]}],
            'organisationUnits': [{'id': 'OldZoneAAAA'}],
            'aggregationType': 'SUM', 'skipRounding': True, 'completedOnly': False,
            'custom': {'nested': ['keep']},
        }

    def catalog(self):
        return build_dataset_catalog({'dataSets': [
            {'id': 'DatasetAAAA', 'displayName': 'Même nom'},
            {'id': 'DatasetBBBB', 'displayName': 'Même nom'},
            {'id': 'DatasetCCCC', 'displayName': 'Service tertiaire'},
            {'id': 'DatasetDDDD'},
        ]})

    def test_exact_five_metrics_replace_all_old_dx_and_preserve_other_filters_options(self):
        original = self.favorite()
        before = deepcopy(original)
        catalog = self.catalog()
        catalog_before = deepcopy(catalog)
        selected = ['DatasetBBBB', 'DatasetAAAA', 'DatasetBBBB']
        result = select_quality_datasets(original, selected, catalog)
        expected = [f'{uid}.{metric}' for uid in ['DatasetBBBB', 'DatasetAAAA'] for metric in METRIC_LABELS]
        self.assertEqual([item['reportingRate']['id'] for item in result['dataDimensionItems']], expected)
        for location in ['columns', 'rows', 'filters']:
            self.assertEqual(result[location][0]['items'], [{'id': identifier} for identifier in expected])
        for field in original:
            if field not in ['dataDimensionItems', 'columns', 'rows', 'filters']:
                self.assertEqual(result[field], original[field])
        self.assertEqual(result['rows'][1:], original['rows'][1:])
        self.assertEqual(result['filters'][1:], original['filters'][1:])
        self.assertEqual(result['columns'][0]['custom'], original['columns'][0]['custom'])
        self.assertEqual(original, before)
        self.assertEqual(catalog, catalog_before)
        self.assertEqual(selected, ['DatasetBBBB', 'DatasetAAAA', 'DatasetBBBB'])
        result['custom']['nested'].append('changed')
        result['filters'][1]['items'][0]['id'] = 'changed'
        self.assertEqual(original, before)
        self.assertNotIn('Indicator', str(result))
        self.assertNotIn('DatasetOLD', str(result))

    def test_homonymous_names_have_consistent_uid_qualified_analytics_labels(self):
        result = select_quality_datasets(self.favorite(), ['DatasetAAAA', 'DatasetBBBB'], self.catalog())
        params, labels = build_quality_reporting_query(result, '202609', 'ZoneAAAAAAA')
        self.assertEqual(len(labels), 10)
        self.assertEqual(len(set(labels.values())), 10)
        for uid in ['DatasetAAAA', 'DatasetBBBB']:
            for metric, suffix in METRIC_LABELS.items():
                self.assertEqual(labels[f'{uid}.{metric}'], f'Même nom ({uid}) {suffix}')
        self.assertIn(('filter', 'CategoryAAA:OptionAAAAA'), params)
        self.assertIn(('aggregationType', 'SUM'), params)
        self.assertIn(('skipRounding', 'true'), params)
        self.assertIn(('completedOnly', 'false'), params)

    def test_any_catalog_dataset_can_be_selected_and_missing_name_uses_uid(self):
        result = select_quality_datasets({}, ('DatasetCCCC', 'DatasetDDDD'), self.catalog())
        self.assertEqual(favorite_dataset_ids(result), ['DatasetCCCC', 'DatasetDDDD'])
        _, labels = build_quality_reporting_query(result, '202609', 'ZoneAAAAAAA')
        self.assertEqual(labels['DatasetCCCC.ACTUAL_REPORTS'], 'Service tertiaire Actual reports')
        self.assertEqual(labels['DatasetDDDD.ACTUAL_REPORTS'], 'DatasetDDDD Actual reports')

    def test_96_selected_datasets_generate_exactly_480_metrics(self):
        catalog = build_dataset_catalog({'dataSets': [
            {'id': f'Data{number:07d}', 'displayName': f'Dataset {number}'} for number in range(96)
        ]})
        result = select_quality_datasets(self.favorite(), [entry['id'] for entry in catalog], catalog)
        self.assertEqual(len(result['dataDimensionItems']), 480)
        self.assertEqual(len(favorite_dataset_ids(result)), 96)

    def test_empty_non_list_invalid_and_unknown_selections_are_rejected(self):
        selections = [[], (), None, '', 'DatasetAAAA', b'DatasetAAAA', {'DatasetAAAA'},
                      {'id': 'DatasetAAAA'}, iter(['DatasetAAAA']), [None], [True], [123], ['short'],
                      ['1atasetAAAA'], ['DatasetAAA_'], [' DatasetAAAA'], ['DatasetAAAA '],
                      ['UnknownAAAA'], ['DatasetAAAA', 'UnknownAAAA'], ['DatasetOLD1']]
        for selected in selections:
            with self.subTest(selected=selected), self.assertRaises(ValueError):
                select_quality_datasets(self.favorite(), selected, self.catalog())

    def test_invalid_definitions_and_catalogs_are_rejected_without_mutation(self):
        for favorite in [None, [], {'columns': 'dx'}, {'rows': [None]}]:
            with self.subTest(favorite=favorite), self.assertRaises(ValueError):
                select_quality_datasets(favorite, ['DatasetAAAA'], self.catalog())
        for catalog in [None, {}, {'dataSets': self.catalog()}, [None], [{'id': 'short'}]]:
            with self.subTest(catalog=catalog), self.assertRaises(ValueError):
                select_quality_datasets({}, ['DatasetAAAA'], catalog)


class FavoriteDatasetIdsTests(unittest.TestCase):
    def test_ids_are_unique_and_follow_all_favorite_dx_locations(self):
        favorite = QualityDatasetSelectionTests().favorite()
        self.assertEqual(favorite_dataset_ids(favorite), ['DatasetOLD1', 'DatasetOLD2', 'DatasetOLD3'])
        favorite['filters'][0]['items'].append({'id': 'DatasetOLD1.EXPECTED_REPORTS'})
        self.assertEqual(favorite_dataset_ids(favorite), ['DatasetOLD1', 'DatasetOLD2', 'DatasetOLD3'])
        self.assertEqual(favorite_dataset_ids({}), [])
        self.assertEqual(favorite_dataset_ids({'dataDimensionItems': [{'indicator': {'id': 'IndicatorAA'}}]}), [])

    def test_reporting_uid_and_favorite_container_are_checked(self):
        invalid = [None, [], {'dataDimensionItems': 'DatasetAAAA.REPORTING_RATE'},
                   {'dataDimensionItems': [{'reportingRate': {'id': 'short.REPORTING_RATE'}}]},
                   {'columns': [{'id': 'dx', 'items': 'DatasetAAAA.REPORTING_RATE'}]}]
        for favorite in invalid:
            with self.subTest(favorite=favorite), self.assertRaises(ValueError):
                favorite_dataset_ids(favorite)


if __name__ == '__main__':
    unittest.main()
