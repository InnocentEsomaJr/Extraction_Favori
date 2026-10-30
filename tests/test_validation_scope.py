import unittest

import pandas as pd

from validation_scope import assign_validation_zones


class ValidationScopeTests(unittest.TestCase):
    def setUp(self):
        self.zones = pd.DataFrame([
            {'Zone_ID': 'ZoneAAAAAAA', 'Zone': 'Ango', 'Province': 'Bas Uele'},
            {'Zone_ID': 'ZoneBBBBBBB', 'Zone': 'B', 'Province': 'Bas Uele'},
        ])

    def test_named_centres_without_paths_use_metadata_and_keep_every_source_row(self):
        rows = pd.DataFrame([
            {'organisationUnitId': 'CentreAAAAA', 'organisationUnitDisplayName': 'Centre A', 'leftSideValue': 84},
            {'organisationUnitId': 'CentreAAAAA', 'organisationUnitDisplayName': 'Centre A', 'leftSideValue': 84},
            {'organisationUnitId': 'CentreBBBBB', 'leftSideValue': 20},
        ])
        rows.attrs['failed_periods'] = []
        metadata = {
            'CentreAAAAA': {'path': '/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/CentreAAAAA'},
            'CentreBBBBB': {'ancestors': [{'id': 'ZoneBBBBBBB', 'level': 3}]},
        }
        result, warnings = assign_validation_zones(rows, self.zones, metadata, 'CountryAAAA')
        self.assertEqual(result['_zone_id'].tolist(), ['ZoneAAAAAAA', 'ZoneAAAAAAA', 'ZoneBBBBBBB'])
        self.assertEqual(result['leftSideValue'].tolist(), [84, 84, 20])
        self.assertEqual(result.attrs['failed_periods'], [])
        self.assertEqual(warnings, [])
        self.assertNotIn('_zone_id', rows)

    def test_zone_parent_is_sufficient_when_optional_metadata_is_missing(self):
        rows = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA'}])
        result, warnings = assign_validation_zones(rows, self.zones.iloc[:1], parent_id='ZoneAAAAAAA')
        self.assertEqual(result['_zone_id'].tolist(), ['ZoneAAAAAAA'])
        self.assertEqual(warnings, [])

    def test_country_unknown_centre_is_kept_but_counts_are_marked_incomplete(self):
        rows = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA', 'leftSideValue': 0}])
        result, warnings = assign_validation_zones(rows, self.zones, parent_id='CountryAAAA')
        self.assertEqual(len(result), 1)
        self.assertEqual(result.attrs['unassigned_rows'], 1)
        self.assertEqual(result['leftSideValue'].tolist(), [0])
        self.assertTrue(warnings)

    def test_conflicting_known_path_does_not_default_to_selected_zone(self):
        rows = pd.DataFrame([{
            'organisationUnitId': 'CentreAAAAA',
            'organisationUnitPath': '/CountryAAAA/OtherZoneAA/CentreAAAAA',
        }])
        result, warnings = assign_validation_zones(rows, self.zones.iloc[:1], parent_id='ZoneAAAAAAA')
        self.assertEqual(result['_zone_id'].tolist(), [''])
        self.assertTrue(warnings)

    def test_empty_analysis_stays_complete(self):
        result, warnings = assign_validation_zones(pd.DataFrame(), self.zones, parent_id='CountryAAAA')
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs['unassigned_rows'], 0)
        self.assertEqual(warnings, [])

    def test_conflicting_zone_ancestor_blocks_the_default_parent_assignment(self):
        rows = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA'}])
        metadata = {'CentreAAAAA': {'ancestors': [{'id': 'OtherZoneAA', 'level': 3}]}}
        result, warnings = assign_validation_zones(rows, self.zones.iloc[:1], metadata, 'ZoneAAAAAAA')
        self.assertEqual(len(result), 1)
        self.assertEqual(result['_zone_id'].tolist(), [''])
        self.assertTrue(warnings)


if __name__ == '__main__':
    unittest.main()
