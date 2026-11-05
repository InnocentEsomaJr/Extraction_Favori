import unittest
from datetime import date
from unittest import mock

import pandas as pd

from geography import build_geography_catalog
from quality_batch import run_quality_batch
from quality_analysis import analyse_quality_period, build_validation_comparison_rows
from quality_scope import assign_quality_units, quality_target_catalog


class QualityScopeTests(unittest.TestCase):
    def setUp(self):
        country = {'id': 'CountryAAAA', 'displayName': 'Pays', 'level': 1}
        pa = {'id': 'ProvinceAAA', 'displayName': 'DPS A', 'level': 2}
        pb = {'id': 'ProvinceBBB', 'displayName': 'DPS B', 'level': 2}
        self.catalog = build_geography_catalog([
            country, {**pa, 'ancestors': [country]}, {**pb, 'ancestors': [country]},
            {'id': 'ZoneAAAAAAA', 'displayName': 'ZS A', 'level': 3, 'ancestors': [country, pa]},
            {'id': 'ZoneBBBBBBB', 'displayName': 'ZS B', 'level': 3, 'ancestors': [country, pa]},
            {'id': 'ZoneCCCCCCC', 'displayName': 'ZS C', 'level': 3, 'ancestors': [country, pb]},
        ])

    def test_country_targets_real_provinces_with_dps_profiles(self):
        result = quality_target_catalog(self.catalog, 'CountryAAAA')
        self.assertEqual(result['Zone_ID'].tolist(), ['ProvinceAAA', 'ProvinceBBB'])
        self.assertEqual(result['Zone'].tolist(), ['DPS A', 'DPS B'])
        self.assertEqual(result['Profil_Bareme'].tolist(), ['DPS', 'DPS'])
        self.assertEqual(result['level'].tolist(), [2, 2])

    def test_province_and_zone_target_only_their_actual_zones(self):
        province = quality_target_catalog(self.catalog, 'CountryAAAA', 'ProvinceAAA')
        self.assertEqual(province['Zone_ID'].tolist(), ['ZoneAAAAAAA', 'ZoneBBBBBBB'])
        self.assertEqual(province['Profil_Bareme'].tolist(), ['ZS', 'ZS'])
        focused = quality_target_catalog(self.catalog, 'CountryAAAA', 'ProvinceAAA', 'ZoneBBBBBBB')
        self.assertEqual(focused['Zone_ID'].tolist(), ['ZoneBBBBBBB'])
        self.assertTrue(quality_target_catalog(self.catalog, 'CountryAAAA', 'ProvinceBBB', 'ZoneBBBBBBB').empty)

    def test_country_can_explicitly_display_all_zones(self):
        result = quality_target_catalog(self.catalog, 'CountryAAAA', target_level=3)
        self.assertEqual(result['Zone_ID'].tolist(), ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'])
        self.assertTrue(result['level'].eq(3).all())

    def test_invalid_or_unknown_country_does_not_enlarge_scope(self):
        self.assertTrue(quality_target_catalog(self.catalog, '').empty)
        self.assertTrue(quality_target_catalog(self.catalog, 'OtherCountry').empty)
        self.assertTrue(quality_target_catalog(self.catalog, 'CountryAAAA', '').empty)
        with self.assertRaises(ValueError):
            quality_target_catalog(self.catalog, 'CountryAAAA', target_level=4)

    def test_dps_mapping_uses_paths_or_ancestors_and_preserves_duplicate_rows(self):
        targets = quality_target_catalog(self.catalog, 'CountryAAAA')
        rows = pd.DataFrame([
            {'organisationUnitId': 'CentreAAAAA', 'leftSideValue': 84},
            {'organisationUnitId': 'CentreAAAAA', 'leftSideValue': 84},
            {'organisationUnitId': 'CentreBBBBB', 'leftSideValue': 20},
        ])
        rows.attrs['request_method'] = 'POST'
        metadata = {
            'CentreAAAAA': {'path': '/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/CentreAAAAA'},
            'CentreBBBBB': {'ancestors': [
                {'id': 'ProvinceBBB', 'level': 2}, {'id': 'ZoneCCCCCCC', 'level': 3},
            ]},
        }
        result, warnings = assign_quality_units(rows, targets, metadata, 'CountryAAAA')
        self.assertEqual(result['_zone_id'].tolist(), ['ProvinceAAA', 'ProvinceAAA', 'ProvinceBBB'])
        self.assertEqual(result['leftSideValue'].tolist(), [84, 84, 20])
        self.assertEqual(result.attrs['request_method'], 'POST')
        self.assertEqual(warnings, [])
        self.assertNotIn('_zone_id', rows)

    def test_one_dps_parent_can_supply_unknown_children_but_never_conflicting_dps(self):
        targets = quality_target_catalog(self.catalog, 'CountryAAAA').iloc[:1]
        rows = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA'}])
        result, warnings = assign_quality_units(rows, targets, parent_id='ProvinceAAA')
        self.assertEqual(result['_zone_id'].tolist(), ['ProvinceAAA'])
        self.assertEqual(warnings, [])
        metadata = {'CentreAAAAA': {'ancestors': [{'id': 'ProvinceBBB', 'level': 2}]}}
        result, warnings = assign_quality_units(rows, targets, metadata, 'ProvinceAAA')
        self.assertEqual(result['_zone_id'].tolist(), [''])
        self.assertTrue(warnings)

    def test_country_unknown_centre_stays_visible_without_invented_dps(self):
        targets = quality_target_catalog(self.catalog, 'CountryAAAA')
        rows = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA'}])
        result, warnings = assign_quality_units(rows, targets, parent_id='CountryAAAA')
        self.assertEqual(len(result), 1)
        self.assertEqual(result.attrs['unassigned_rows'], 1)
        self.assertTrue(warnings)

    def test_zs_mapping_keeps_existing_behavior_and_zero_analysis_is_complete(self):
        targets = quality_target_catalog(self.catalog, 'CountryAAAA', 'ProvinceAAA')
        result, warnings = assign_quality_units(pd.DataFrame(), targets, parent_id='ProvinceAAA')
        self.assertTrue(result.empty)
        self.assertEqual(warnings, [])
        rows = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA'}])
        result, warnings = assign_quality_units(rows, targets.iloc[:1], parent_id='ZoneAAAAAAA')
        self.assertEqual(result['_zone_id'].tolist(), ['ZoneAAAAAAA'])
        self.assertEqual(warnings, [])


class QualityBatchTests(unittest.TestCase):
    def setUp(self):
        self.targets = pd.DataFrame([
            {'Zone_ID': 'ZoneAAAAAAA', 'Zone': 'A'},
            {'Zone_ID': 'ZoneBBBBBBB', 'Zone': 'B'},
            {'Zone_ID': 'ZoneCCCCCCC', 'Zone': 'C'},
        ])

    def test_zero_success_failed_target_and_real_rows_stay_independent_without_retry(self):
        empty = pd.DataFrame()
        empty.attrs['failed_periods'] = []
        failed = pd.DataFrame()
        failed.attrs = {'failed_periods': ['202609'], 'errors': {'202609': ['HTTP 504']}}
        rows = pd.DataFrame([{'organisationUnitId': 'CentreCCCCA', 'leftSideValue': 5}])
        rows.attrs['request_method'] = 'POST'
        fetch = mock.Mock(side_effect=[empty, failed, rows])
        progress = mock.Mock()
        result = run_quality_batch(
            self.targets, ['202609'], 'GroupAAAAAA', fetch, cache_user=('user', 'token'),
            start_date=date(2026, 9, 1), end_date=date(2026, 9, 30), on_progress=progress,
        )
        self.assertEqual(fetch.call_count, 3)
        fetch.assert_any_call('ZoneAAAAAAA', ['202609'], 'GroupAAAAAA', cache_user=('user', 'token'),
                              start_date=date(2026, 9, 1), end_date=date(2026, 9, 30))
        self.assertEqual(result['succeeded_by_id'], {'ZoneAAAAAAA': True, 'ZoneBBBBBBB': False, 'ZoneCCCCCCC': True})
        self.assertEqual(result['requests'][0]['Lignes reçues'], 0)
        self.assertTrue(pd.isna(result['requests'][1]['Lignes reçues']))
        self.assertEqual(result['requests'][2]['Lignes reçues'], 1)
        self.assertIn('HTTP 504', result['warnings'][0])
        self.assertEqual(result['raw']['_analysis_target_id'].tolist(), ['ZoneCCCCCCC'])
        self.assertEqual(result['raw']['leftSideValue'].tolist(), [5])
        self.assertEqual(progress.call_args_list[0], mock.call(1, 3, 'ZoneAAAAAAA', 'A'))
        self.assertEqual(progress.call_count, 3)
        self.assertFalse(result['complete'])

    def test_reference_analysis_is_separate_and_omits_explicit_dates(self):
        fetch = mock.Mock(return_value=pd.DataFrame())
        result = run_quality_batch(self.targets.iloc[:1], ['202608'], None, fetch, label='M-1')
        fetch.assert_called_once_with('ZoneAAAAAAA', ['202608'], None, cache_user=None)
        self.assertTrue(result['complete'])
        self.assertEqual(result['requests'][0]['Période'], 'M-1')
        self.assertEqual(result['requests'][0]['Statut'], 'Terminé')

    def test_exception_and_truncation_are_unavailable_but_do_not_stop_other_targets(self):
        truncated = pd.DataFrame([{'organisationUnitId': 'CentreBBBBB'}])
        truncated.attrs['truncated_periods'] = ['202609']
        fetch = mock.Mock(side_effect=[RuntimeError('private credential value'), truncated, pd.DataFrame()])
        result = run_quality_batch(self.targets, ['202609'], None, fetch)
        self.assertEqual(fetch.call_count, 3)
        self.assertFalse(result['succeeded_by_id']['ZoneAAAAAAA'])
        self.assertFalse(result['succeeded_by_id']['ZoneBBBBBBB'])
        self.assertTrue(result['succeeded_by_id']['ZoneCCCCCCC'])
        self.assertNotIn('private credential value', str(result))
        self.assertEqual(result['requests'][1]['Statut'], 'Partiel')

    def test_one_target_step_saves_its_reply_and_leaves_pending_counts_unknown(self):
        fetch = mock.Mock(return_value=pd.DataFrame())
        checkpoint = mock.Mock()
        result = run_quality_batch(
            self.targets, ['202609'], None, fetch, max_targets=1, on_result=checkpoint,
        )
        fetch.assert_called_once_with('ZoneAAAAAAA', ['202609'], None, cache_user=None)
        self.assertEqual(result['processed_ids'], ['ZoneAAAAAAA'])
        self.assertEqual(result['pending_ids'], ['ZoneBBBBBBB', 'ZoneCCCCCCC'])
        self.assertFalse(result['finished'])
        self.assertFalse(result['complete'])
        self.assertEqual(checkpoint.call_count, 1)
        self.assertEqual(checkpoint.call_args.args[0]['processed_ids'], ['ZoneAAAAAAA'])
        assigned, errors, _ = analyse_quality_period(result, self.targets)
        compared = build_validation_comparison_rows(self.targets, assigned, errors).set_index('Zone_ID')
        self.assertEqual(compared.loc['ZoneAAAAAAA', 'Violations_Actuelles'], 0)
        self.assertTrue(compared.loc[['ZoneBBBBBBB', 'ZoneCCCCCCC'], 'Violations_Actuelles'].isna().all())

    def test_interruption_on_seventh_zone_keeps_six_replies_and_resumes_at_seventh(self):
        targets = pd.DataFrame([
            {'Zone_ID': f'Zone{number:07d}', 'Zone': f'ZS {number}'} for number in range(1, 9)
        ])
        checkpoints = []
        fetch = mock.Mock(side_effect=[pd.DataFrame() for _ in range(6)] + [KeyboardInterrupt()])
        with self.assertRaises(KeyboardInterrupt):
            run_quality_batch(targets, ['202609'], None, fetch, on_result=checkpoints.append)
        self.assertEqual(fetch.call_count, 7)
        self.assertEqual(len(checkpoints), 6)
        saved = checkpoints[-1]
        self.assertEqual(saved['processed_ids'], targets['Zone_ID'].tolist()[:6])
        self.assertEqual(saved['pending_ids'], targets['Zone_ID'].tolist()[6:])
        self.assertFalse(saved['finished'])
        resume = mock.Mock(return_value=pd.DataFrame())
        result = run_quality_batch(targets, ['202609'], None, resume, initial_state=saved)
        self.assertEqual([call.args[0] for call in resume.call_args_list], targets['Zone_ID'].tolist()[6:])
        self.assertEqual(len(result['requests']), 8)
        self.assertEqual(len(set(result['processed_ids'])), 8)
        self.assertEqual(result['pending_ids'], [])
        self.assertTrue(result['finished'])
        self.assertTrue(result['complete'])

    def test_resuming_does_not_retry_an_already_failed_target(self):
        failed = pd.DataFrame()
        failed.attrs = {'failed_periods': ['202609'], 'errors': {'202609': ['HTTP 504']}}
        initial = run_quality_batch(
            self.targets, ['202609'], None, mock.Mock(return_value=failed), max_targets=1,
        )
        resume = mock.Mock(return_value=pd.DataFrame())
        result = run_quality_batch(self.targets, ['202609'], None, resume, initial_state=initial)
        self.assertEqual([call.args[0] for call in resume.call_args_list], ['ZoneBBBBBBB', 'ZoneCCCCCCC'])
        self.assertFalse(result['succeeded_by_id']['ZoneAAAAAAA'])
        self.assertIn('HTTP 504', result['warnings'][0])
        self.assertTrue(result['finished'])
        self.assertFalse(result['complete'])
        already_finished_fetch = mock.Mock(side_effect=AssertionError('Aucune relance attendue'))
        final = run_quality_batch(self.targets, ['202609'], None, already_finished_fetch, initial_state=result)
        already_finished_fetch.assert_not_called()
        self.assertEqual(final['processed_ids'], result['processed_ids'])

    def test_resume_filters_foreign_targets_and_never_reuses_another_period(self):
        rows = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA'}])
        initial = run_quality_batch(self.targets, ['202609'], None, mock.Mock(return_value=rows))
        focused = self.targets.iloc[[1]]
        fetch = mock.Mock(side_effect=AssertionError('La cible existe déjà'))
        result = run_quality_batch(focused, ['202609'], None, fetch, initial_state=initial)
        fetch.assert_not_called()
        self.assertEqual(list(result['frames_by_id']), ['ZoneBBBBBBB'])
        self.assertEqual(result['raw']['_analysis_target_id'].tolist(), ['ZoneBBBBBBB'])
        self.assertEqual([row['Unité parent'] for row in result['requests']], ['ZoneBBBBBBB'])
        next_period = mock.Mock(return_value=pd.DataFrame())
        replaced = run_quality_batch(focused, ['202610'], None, next_period, initial_state=initial)
        next_period.assert_called_once_with('ZoneBBBBBBB', ['202610'], None, cache_user=None)
        self.assertTrue(replaced['raw'].empty)
        self.assertEqual(replaced['context']['periods'], ['202610'])

    def test_checkpoints_and_restored_frames_are_deep_copies(self):
        rows = pd.DataFrame([{'organisationUnit': {'id': 'CentreAAAAA'}}])
        rows.attrs = {'metadata': {'name': 'Source'}}
        saved = []

        def change_snapshot(snapshot):
            saved.append(snapshot)
            snapshot['frames_by_id']['ZoneAAAAAAA'].iloc[0]['organisationUnit']['id'] = 'MutatedFrame'
            snapshot['frames_by_id']['ZoneAAAAAAA'].attrs['metadata']['name'] = 'MutatedMetadata'
            snapshot['raw'].iloc[0]['organisationUnit']['id'] = 'MutatedRaw'

        result = run_quality_batch(
            self.targets.iloc[:1], ['202609'], None, mock.Mock(return_value=rows), on_result=change_snapshot,
        )
        self.assertEqual(result['frames_by_id']['ZoneAAAAAAA'].iloc[0]['organisationUnit']['id'], 'CentreAAAAA')
        self.assertEqual(result['frames_by_id']['ZoneAAAAAAA'].attrs['metadata']['name'], 'Source')
        self.assertEqual(rows.iloc[0]['organisationUnit']['id'], 'CentreAAAAA')
        initial_before = result['frames_by_id']['ZoneAAAAAAA'].iloc[0]['organisationUnit']['id']
        restored = run_quality_batch(self.targets.iloc[:1], ['202609'], None, mock.Mock(), initial_state=result)
        restored['frames_by_id']['ZoneAAAAAAA'].iloc[0]['organisationUnit']['id'] = 'RestoredMutation'
        self.assertEqual(result['frames_by_id']['ZoneAAAAAAA'].iloc[0]['organisationUnit']['id'], initial_before)

    def test_invalid_step_limit_is_rejected_before_fetching(self):
        for value in [0, -1, True, 1.5]:
            with self.subTest(value=value):
                fetch = mock.Mock()
                with self.assertRaises(ValueError):
                    run_quality_batch(self.targets, ['202609'], None, fetch, max_targets=value)
                fetch.assert_not_called()


if __name__ == '__main__':
    unittest.main()
