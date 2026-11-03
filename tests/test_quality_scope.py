import unittest
from datetime import date
from unittest import mock

import pandas as pd

from geography import build_geography_catalog
from quality_batch import run_quality_batch
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


if __name__ == '__main__':
    unittest.main()
