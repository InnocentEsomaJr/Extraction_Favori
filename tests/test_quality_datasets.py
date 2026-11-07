"""Le périmètre des datasets qualité reste explicite et indépendant des valeurs."""

from copy import deepcopy
import unittest

from quality_datasets import (
    quality_dataset_selection, reporting_dataset_ids, select_quality_datasets,
)
from scoped_analytics import METRIC_LABELS, build_quality_reporting_query


class QualityDatasetSelectionTests(unittest.TestCase):
    def favorite(self):
        return {
            'dataDimensionItems': [
                {'reportingRate': {'dataSet': {'id': identifier, 'displayName': name}, 'metric': metric}}
                for identifier, name in [
                    ('DatasetAAAA', 'Consultations'), ('DatasetBBBB', 'Service tertiaire'),
                ] for metric in ['REPORTING_RATE', 'ACTUAL_REPORTS']
            ],
            'columns': [{'id': 'dx'}],
            'rows': [{'dimension': 'ou'}],
            'filters': [{'dimension': 'pe', 'items': [{'id': '202601'}]},
                        {'dimension': 'CategoryAAA', 'items': [{'id': 'OptionAAAAA'}]}],
            'skipRounding': True,
        }

    def test_surveillance_is_added_and_tertiary_excluded_without_fixed_dataset_count(self):
        favorite = self.favorite()
        before = deepcopy(favorite)
        selection = quality_dataset_selection(favorite, {'dataSets': [
            {'id': 'DatasetCCCC', 'displayName': 'Survéillance épidémiologique'},
        ]})
        self.assertEqual(selection['defaults'], ['DatasetAAAA', 'DatasetCCCC'])
        self.assertEqual(selection['favorite_ids'], ['DatasetAAAA', 'DatasetBBBB'])
        self.assertNotIn('DatasetBBBB', selection['options'])
        self.assertIn('DatasetBBBB', [dataset['id'] for dataset in selection['datasets']])
        self.assertEqual(selection['warnings'], [])
        self.assertEqual(favorite, before)

    def test_multiple_surveillance_candidates_require_an_explicit_choice(self):
        selection = quality_dataset_selection(self.favorite(), [
            {'id': 'DatasetCCCC', 'name': 'Surveillance hebdomadaire'},
            {'id': 'DatasetDDDD', 'name': 'Surveillance mensuelle'},
        ])
        self.assertEqual(selection['defaults'], ['DatasetAAAA'])
        self.assertEqual(selection['ambiguous_surveillance_ids'], ['DatasetCCCC', 'DatasetDDDD'])
        self.assertEqual(len(selection['warnings']), 1)
        self.assertIn('sélectionnez', selection['warnings'][0])

    def test_existing_surveillance_is_not_duplicated_or_replaced_by_another_candidate(self):
        favorite = self.favorite()
        favorite['dataDimensionItems'][0]['reportingRate']['dataSet']['displayName'] = 'Surveillance'
        selection = quality_dataset_selection(favorite, [
            {'id': 'DatasetCCCC', 'name': 'Surveillance hebdomadaire'},
            {'id': 'DatasetDDDD', 'name': 'Surveillance mensuelle'},
        ])
        self.assertEqual(selection['defaults'], ['DatasetAAAA'])
        self.assertEqual(selection['ambiguous_surveillance_ids'], [])
        self.assertEqual(selection['warnings'], [])

    def test_tertiary_name_variants_are_removed_but_no_substring_false_match(self):
        favorite = {'dataDimensionItems': [
            {'reportingRate': {'dataSet': {'id': identifier, 'displayName': name}, 'metric': 'REPORTING_RATE'}}
            for identifier, name in [
                ('DatasetAAAA', 'SERVICES TERTIAIRES'), ('DatasetBBBB', 'Tertiary service'),
                ('DatasetCCCC', 'Surveillance'), ('DatasetDDDD', 'Nontertiaire'),
            ]
        ]}
        selection = quality_dataset_selection(favorite)
        self.assertEqual(selection['defaults'], ['DatasetCCCC', 'DatasetDDDD'])

    def test_homonymous_names_stay_separate_by_uid(self):
        selection = quality_dataset_selection(self.favorite(), [
            {'id': 'DatasetCCCC', 'name': 'Consultations'},
        ])
        self.assertEqual(selection['labels']['DatasetAAAA'], 'Consultations [DatasetAAAA]')
        self.assertEqual(selection['labels']['DatasetCCCC'], 'Consultations [DatasetCCCC]')
        self.assertEqual(len(selection['options']), 2)

    def test_metadata_names_enrich_uid_fallback_and_invalid_catalog_entries_are_ignored(self):
        favorite = {'columns': [{'id': 'dx', 'items': [{'id': 'DatasetAAAA.REPORTING_RATE'}]}]}
        selection = quality_dataset_selection(favorite, [
            {'id': 'DatasetAAAA', 'displayName': 'Surveillance'},
            {'id': 'invalid', 'name': 'Surveillance erronée'},
            None,
        ])
        self.assertEqual(selection['defaults'], ['DatasetAAAA'])
        self.assertEqual(selection['labels'], {'DatasetAAAA': 'Surveillance'})
        self.assertEqual(selection['warnings'], [])

    def test_dataset_ids_cover_both_dimension_layouts_and_typed_items(self):
        favorite = self.favorite()
        favorite['rows'].append({'dimension': 'dx', 'items': [
            'DatasetCCCC.REPORTING_RATE', {'id': 'DatasetAAAA.EXPECTED_REPORTS'},
            {'indicator': {'id': 'IndicatorAA', 'name': 'Indicateur'}},
        ]})
        self.assertEqual(reporting_dataset_ids(favorite), ('DatasetAAAA', 'DatasetBBBB', 'DatasetCCCC'))

    def test_no_surveillance_is_explicit_without_invented_uid(self):
        selection = quality_dataset_selection(self.favorite())
        self.assertEqual(selection['defaults'], ['DatasetAAAA'])
        self.assertEqual(selection['surveillance_ids'], [])
        self.assertIn('Aucun dataset', selection['warnings'][0])

    def test_tertiary_cannot_be_reintroduced_into_the_quality_query(self):
        favorite = self.favorite()
        selection = quality_dataset_selection(favorite)
        self.assertIn('DatasetBBBB', selection['favorite_ids'])
        self.assertNotIn('DatasetBBBB', selection['options'])
        tertiary = next(dataset for dataset in selection['datasets'] if dataset['id'] == 'DatasetBBBB')
        self.assertTrue(tertiary['is_tertiary'])
        with self.assertRaisesRegex(ValueError, 'tertiaire.*exclu'):
            select_quality_datasets(favorite, ['DatasetAAAA', 'DatasetBBBB'], selection['datasets'])

    def test_unclassified_uid_only_favorite_has_an_explicit_metadata_warning(self):
        favorite = {'dataDimensionItems': [
            {'reportingRate': {'id': 'DatasetAAAA.REPORTING_RATE'}},
        ]}
        selection = quality_dataset_selection(favorite)
        self.assertEqual(selection['defaults'], ['DatasetAAAA'])
        self.assertTrue(any('type ne peut pas être vérifié' in message and 'DatasetAAAA' in message
                            for message in selection['warnings']))

    def test_exact_surveillance_name_takes_priority_over_unrelated_favorite_and_catalog_candidates(self):
        favorite = self.favorite()
        favorite['dataDimensionItems'].append({'reportingRate': {
            'dataSet': {'id': 'DatasetDDDD', 'name': 'Surveillance hebdomadaire'}, 'metric': 'REPORTING_RATE',
        }})
        selection = quality_dataset_selection(favorite, [
            {'id': 'DatasetCCCC', 'name': '1Y - SURVEILLANCE EPIDEMIOLOGIQUE'},
            {'id': 'DatasetEEEE', 'name': 'Surveillance mensuelle'},
        ], surveillance_name='1y-Surveillance épidémiologique')
        self.assertEqual(selection['defaults'], ['DatasetAAAA', 'DatasetCCCC'])
        self.assertEqual(selection['warnings'], [])
        self.assertIn('DatasetDDDD', selection['options'])
        self.assertIn('DatasetEEEE', selection['options'])
        self.assertTrue(next(dataset for dataset in selection['datasets']
                             if dataset['id'] == 'DatasetCCCC')['is_surveillance'])

    def test_missing_expected_surveillance_does_not_select_another_surveillance(self):
        selection = quality_dataset_selection(self.favorite(), [
            {'id': 'DatasetCCCC', 'name': 'Surveillance hebdomadaire'},
        ], surveillance_name='1y-Surveillance épidémiologique')
        self.assertEqual(selection['defaults'], ['DatasetAAAA'])
        self.assertEqual(selection['ambiguous_surveillance_ids'], [])
        self.assertIn('1y-Surveillance épidémiologique', selection['warnings'][0])
        self.assertIn('aucun autre', selection['warnings'][0])

    def test_identical_expected_names_with_different_uids_require_explicit_disambiguation(self):
        selection = quality_dataset_selection(self.favorite(), [
            {'id': 'DatasetCCCC', 'name': '1y-Surveillance épidémiologique'},
            {'id': 'DatasetDDDD', 'name': '1Y Surveillance Epidemiologique'},
        ], surveillance_name='1y-Surveillance épidémiologique')
        self.assertEqual(selection['defaults'], ['DatasetAAAA'])
        self.assertEqual(selection['ambiguous_surveillance_ids'], ['DatasetCCCC', 'DatasetDDDD'])
        self.assertIn('UID exact', selection['warnings'][0])

    def test_repeated_catalog_entry_for_the_same_surveillance_uid_is_not_ambiguous(self):
        entry = {'id': 'DatasetCCCC', 'name': '1y-Surveillance épidémiologique'}
        selection = quality_dataset_selection(self.favorite(), [entry, dict(entry)],
                                               surveillance_name='1y-Surveillance épidémiologique')
        self.assertEqual(selection['defaults'], ['DatasetAAAA', 'DatasetCCCC'])
        self.assertEqual(selection['ambiguous_surveillance_ids'], [])
        self.assertEqual(selection['warnings'], [])

    def test_esurveillance_weekly_exact_name_replaces_old_monthly_surveillance(self):
        favorite = self.favorite()
        favorite['dataDimensionItems'].append({'reportingRate': {
            'dataSet': {'id': 'DatasetDDDD', 'name': '1y-Surveillance épidémiologique',
                        'periodType': 'Monthly'}, 'metric': 'REPORTING_RATE',
        }})
        selection = quality_dataset_selection(favorite, [
            {'id': 'DatasetCCCC', 'displayName': 'eSurveillance Hebdomadaire-Édition2025',
             'periodType': 'Weekly'},
            {'id': 'DatasetEEEE', 'displayName': 'eSurveillance Hebdomadaire-Edition 2024',
             'periodType': 'Weekly'},
        ], surveillance_name='eSurveillance Hebdomadaire-Edition 2025')
        self.assertEqual(selection['defaults'], ['DatasetAAAA', 'DatasetCCCC'])
        self.assertEqual(selection['warnings'], [])
        records = {dataset['id']: dataset for dataset in selection['datasets']}
        self.assertTrue(records['DatasetCCCC']['is_surveillance'])
        self.assertTrue(records['DatasetCCCC']['is_weekly'])
        self.assertEqual(records['DatasetCCCC']['periodType'], 'Weekly')
        self.assertFalse(records['DatasetDDDD']['is_weekly'])
        self.assertEqual(records['DatasetDDDD']['periodType'], 'Monthly')

    def test_favorite_weekly_metadata_is_retained_and_catalog_type_is_authoritative(self):
        favorite = self.favorite()
        for item in favorite['dataDimensionItems']:
            item['reportingRate']['dataSet']['periodType'] = 'Weekly'
        selection = quality_dataset_selection(favorite, [
            {'id': 'DatasetAAAA', 'name': 'Consultations', 'periodType': 'Monthly'},
        ])
        records = {dataset['id']: dataset for dataset in selection['datasets']}
        self.assertFalse(records['DatasetAAAA']['is_weekly'])
        self.assertEqual(records['DatasetAAAA']['periodType'], 'Monthly')
        self.assertTrue(records['DatasetBBBB']['is_weekly'])
        self.assertEqual(records['DatasetBBBB']['periodType'], 'Weekly')

    def test_weekly_name_fallback_only_for_esurveillance_hebdomadaire_without_explicit_type(self):
        selection = quality_dataset_selection(self.favorite(), [
            {'id': 'DatasetCCCC', 'name': 'eSurveillance Hebdomadaire-Edition 2025'},
            {'id': 'DatasetDDDD', 'name': '1y-Surveillance épidémiologique', 'periodType': 'Monthly'},
            {'id': 'DatasetEEEE', 'name': 'Surveillance mensuelle'},
            {'id': 'DatasetFFFF', 'name': 'eSurveillance Hebdomadaire ancienne', 'periodType': 'Monthly'},
            {'id': 'DatasetGGGG', 'name': 'Autre collecte', 'periodType': 'weekly'},
        ], surveillance_name='eSurveillance Hebdomadaire-Edition 2025')
        records = {dataset['id']: dataset for dataset in selection['datasets']}
        self.assertTrue(records['DatasetCCCC']['is_weekly'])
        self.assertFalse(records['DatasetDDDD']['is_weekly'])
        self.assertFalse(records['DatasetEEEE']['is_weekly'])
        self.assertFalse(records['DatasetFFFF']['is_weekly'])
        self.assertTrue(records['DatasetGGGG']['is_weekly'])
        self.assertEqual(records['DatasetGGGG']['periodType'], 'Weekly')

    def test_selected_query_preserves_weekly_dataset_period_type_without_mutating_favorite(self):
        favorite = self.favorite()
        before = deepcopy(favorite)
        selection = quality_dataset_selection(favorite, [
            {'id': 'DatasetCCCC', 'name': 'eSurveillance Hebdomadaire-Edition 2025', 'periodType': 'Weekly'},
        ], surveillance_name='eSurveillance Hebdomadaire-Edition 2025')
        chosen = select_quality_datasets(favorite, ['DatasetCCCC'], selection['datasets'])
        self.assertEqual(len(chosen['dataDimensionItems']), 5)
        self.assertTrue(all(item['reportingRate']['dataSet']['periodType'] == 'Weekly'
                            for item in chosen['dataDimensionItems']))
        self.assertEqual(favorite, before)

    def test_favorite_period_type_is_retained_when_dataset_is_outside_reporting_identifier(self):
        favorite = {'dataDimensionItems': [{
            'reportingRate': 'DatasetCCCC.REPORTING_RATE',
            'dataSet': {'id': 'DatasetCCCC', 'displayName': 'eSurveillance Hebdomadaire-Edition 2025',
                        'periodType': 'Weekly'},
        }]}
        selection = quality_dataset_selection(favorite)
        dataset = selection['datasets'][0]
        self.assertEqual(dataset['periodType'], 'Weekly')
        self.assertTrue(dataset['is_weekly'])

    def test_exact_selected_scope_replaces_all_old_dx_but_preserves_extra_filters(self):
        favorite = self.favorite()
        favorite['columns'][0]['items'] = [{'id': 'DatasetBBBB.REPORTING_RATE'}]
        favorite['dataDimensionItems'].append({'indicator': {'id': 'IndicatorAA', 'name': 'Indicateur'}})
        before = deepcopy(favorite)
        selection = quality_dataset_selection(favorite, [{'id': 'DatasetCCCC', 'name': 'Surveillance'}])
        chosen = select_quality_datasets(favorite, selection['defaults'], selection['datasets'])
        params, labels = build_quality_reporting_query(chosen, '202609', 'ZoneAAAAAAA')
        self.assertEqual(set(labels), {
            f'{identifier}.{metric}' for identifier in ['DatasetAAAA', 'DatasetCCCC'] for metric in METRIC_LABELS
        })
        self.assertNotIn('DatasetBBBB', str(params))
        self.assertNotIn('IndicatorAA', str(params))
        self.assertIn(('filter', 'CategoryAAA:OptionAAAAA'), params)
        self.assertIn(('filter', 'pe:202609'), params)
        self.assertIn(('dimension', 'ou:ZoneAAAAAAA'), params)
        self.assertEqual(favorite, before)

    def test_selection_can_explicitly_pick_one_ambiguous_surveillance_uid(self):
        favorite = self.favorite()
        selection = quality_dataset_selection(favorite, [
            {'id': 'DatasetCCCC', 'name': 'Surveillance hebdomadaire'},
            {'id': 'DatasetDDDD', 'name': 'Surveillance mensuelle'},
        ])
        chosen = select_quality_datasets(favorite, ['DatasetAAAA', 'DatasetDDDD'], selection['datasets'])
        self.assertEqual(reporting_dataset_ids(chosen), ('DatasetAAAA', 'DatasetDDDD'))

    def test_empty_unknown_or_malformed_selection_is_refused(self):
        for selected in [[], ['DatasetZZZZ'], ['invalid'], 'DatasetAAAA']:
            with self.subTest(selected=selected), self.assertRaises(ValueError):
                select_quality_datasets(self.favorite(), selected)


if __name__ == '__main__':
    unittest.main()
