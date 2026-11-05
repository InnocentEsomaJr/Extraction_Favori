import unittest

import pandas as pd

from geography import build_geography_catalog
from validation_details import build_validation_details
from validation_detail_view import (
    CORRECTION_COL, enrich_detail_geography, filter_detail_geography, mark_previous_corrections,
)


def signatures(frame):
    return [
        (row['organisationUnitId'], row['validationRuleId'])
        for row in frame.to_dict('records')
        if pd.notna(row.get('organisationUnitId')) and pd.notna(row.get('validationRuleId'))
    ]


class DetailGeographyTests(unittest.TestCase):
    def setUp(self):
        country = {'id': 'CountryAAAA', 'displayName': 'RDC', 'level': 1}
        province_a = {'id': 'ProvinceAAA', 'displayName': 'Haut Uele', 'level': 2}
        province_b = {'id': 'ProvinceBBB', 'displayName': 'Bas Uele', 'level': 2}
        self.catalog = build_geography_catalog([
            country,
            {**province_a, 'ancestors': [country]},
            {**province_b, 'ancestors': [country]},
            {'id': 'ZoneAAAAAAA', 'displayName': 'ZS commune', 'level': 3, 'ancestors': [country, province_a]},
            {'id': 'ZoneBBBBBBB', 'displayName': 'ZS commune', 'level': 3, 'ancestors': [country, province_b]},
        ])

    def raw(self):
        return pd.DataFrame([
            {'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleAAAAAAA', '_zone_id': 'ProvinceAAA',
             'organisationUnitPath': '/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/CentreAAAAA'},
            {'organisationUnitId': 'CentreBBBBB', 'validationRuleId': 'RuleBBBBBBB', '_zone_id': 'ProvinceBBB',
             'organisationUnitPath': '/CountryAAAA/ProvinceBBB/ZoneBBBBBBB/CentreBBBBB'},
            {'organisationUnitId': 'UnknownAAAA', 'validationRuleId': 'RuleCCCCCCC'},
        ])

    def test_country_province_and_actual_zone_ids_are_distinct_from_dps_target(self):
        raw = self.raw()
        details = build_validation_details(raw)
        result = enrich_detail_geography(details, raw, self.catalog)
        self.assertEqual(result['_zone_id'].tolist(), ['ProvinceAAA', 'ProvinceBBB', ''])
        self.assertEqual(result['_health_zone_id'].tolist(), ['ZoneAAAAAAA', 'ZoneBBBBBBB', ''])
        self.assertEqual(result['_province_id'].tolist(), ['ProvinceAAA', 'ProvinceBBB', ''])
        self.assertEqual(result['_country_id'].tolist(), ['CountryAAAA', 'CountryAAAA', ''])
        self.assertEqual(result['Zone de santé'].tolist(), ['ZS commune (Haut Uele)', 'ZS commune (Bas Uele)', 'Non renseignée'])
        self.assertEqual(result['Province'].tolist(), ['Haut Uele', 'Bas Uele', 'Non renseignée'])
        self.assertNotIn('_country_id', details)

    def test_exact_filters_keep_unknown_rows_only_when_no_filter_is_active(self):
        raw = self.raw()
        enriched = enrich_detail_geography(build_validation_details(raw), raw, self.catalog)
        self.assertEqual(len(filter_detail_geography(enriched)), 3)
        self.assertEqual(len(filter_detail_geography(enriched, country_id='CountryAAAA')), 2)
        self.assertEqual(filter_detail_geography(enriched, province_id='ProvinceAAA')['_ou_id'].tolist(), ['CentreAAAAA'])
        self.assertEqual(filter_detail_geography(enriched, zone_id='ZoneBBBBBBB')['_ou_id'].tolist(), ['CentreBBBBB'])
        self.assertTrue(filter_detail_geography(enriched, country_id='UnknownCountry').empty)
        self.assertTrue(filter_detail_geography(enriched, zone_id='').empty)

    def test_metadata_ancestors_resolve_without_path_and_preserve_duplicate_indexes(self):
        raw = self.raw().iloc[[0, 0]].drop(columns='organisationUnitPath')
        raw['_zone_id'] = ''
        details = build_validation_details(raw)
        details.index = [8, 8]
        metadata = {'CentreAAAAA': {'ancestors': [
            {'id': 'ZoneAAAAAAA', 'level': 3},
        ]}}
        result = enrich_detail_geography(details, raw, self.catalog, metadata)
        self.assertEqual(result.index.tolist(), [8, 8])
        self.assertEqual(result['_health_zone_id'].tolist(), ['ZoneAAAAAAA', 'ZoneAAAAAAA'])
        self.assertEqual(result['_province_id'].tolist(), ['ProvinceAAA', 'ProvinceAAA'])
        self.assertEqual(result['_ou_id'].tolist(), ['CentreAAAAA', 'CentreAAAAA'])

    def test_malformed_path_is_ignored_and_names_do_not_assign_geography(self):
        raw = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA', 'organisationUnitDisplayName': 'ZS commune',
                             'organisationUnitPath': '/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/OtherCentre'}])
        result = enrich_detail_geography(build_validation_details(raw), raw, self.catalog)
        self.assertEqual(result['_country_id'].tolist(), [''])
        self.assertEqual(result['_health_zone_id'].tolist(), [''])

    def test_contradictory_paths_do_not_assign_an_arbitrary_zone_or_province(self):
        raw = self.raw().iloc[:1]
        metadata = {'CentreAAAAA': {'path': '/CountryAAAA/ProvinceBBB/ZoneBBBBBBB/CentreAAAAA'}}
        result = enrich_detail_geography(build_validation_details(raw), raw, self.catalog, metadata)
        self.assertEqual(result['_country_id'].tolist(), ['CountryAAAA'])
        self.assertEqual(result['_province_id'].tolist(), [''])
        self.assertEqual(result['_health_zone_id'].tolist(), [''])

    def test_assigned_single_zone_gives_its_catalog_ancestors_without_a_centre_path(self):
        raw = pd.DataFrame([{'organisationUnitId': 'CentreAAAAA', '_zone_id': 'ZoneAAAAAAA'}])
        result = enrich_detail_geography(build_validation_details(raw), raw, self.catalog)
        self.assertEqual(result['_health_zone_id'].tolist(), ['ZoneAAAAAAA'])
        self.assertEqual(result['_province_id'].tolist(), ['ProvinceAAA'])

    def test_dps_rollup_keeps_the_requested_zone_without_a_centre_path(self):
        raw = pd.DataFrame([{
            'organisationUnitId': 'CentreAAAAA', '_zone_id': 'ProvinceAAA',
            '_analysis_target_id': 'ProvinceAAA', '_request_unit_id': 'ZoneAAAAAAA',
        }])
        result = enrich_detail_geography(build_validation_details(raw), raw, self.catalog)
        self.assertEqual(result['_zone_id'].tolist(), ['ProvinceAAA'])
        self.assertEqual(result['_health_zone_id'].tolist(), ['ZoneAAAAAAA'])
        self.assertEqual(result['_province_id'].tolist(), ['ProvinceAAA'])
        self.assertEqual(result['_country_id'].tolist(), ['CountryAAAA'])

    def test_empty_frames_stay_empty_with_geography_schema(self):
        result = enrich_detail_geography(build_validation_details(pd.DataFrame()), pd.DataFrame(), self.catalog)
        self.assertTrue(result.empty)
        self.assertIn('_health_zone_id', result)
        self.assertIn('Province', result)


class CorrectionStatusTests(unittest.TestCase):
    def test_persistent_corrected_and_unavailable_rows_keep_multiplicity_and_values(self):
        previous = pd.DataFrame([
            {'_zone_id': 'ZoneAAAAAAA', 'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleAAAAAAA', 'leftSideValue': 0},
            {'_zone_id': 'ZoneAAAAAAA', 'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleAAAAAAA', 'leftSideValue': 0},
            {'_zone_id': 'ZoneAAAAAAA', 'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleBBBBBBB', 'leftSideValue': 5},
            {'_zone_id': 'ZoneBBBBBBB', 'organisationUnitId': 'CentreBBBBB', 'validationRuleId': 'RuleCCCCCCC', 'leftSideValue': 8},
        ], index=[4, 4, 3, 2])
        current = previous.iloc[:1]
        result = mark_previous_corrections(previous, current, {'ZoneBBBBBBB': ['HTTP 504']}, signatures)
        self.assertEqual(result[CORRECTION_COL].tolist(), ['Non corrigée', 'Non corrigée', 'Corrigée', 'Non vérifiable'])
        self.assertEqual(result.index.tolist(), [4, 4, 3, 2])
        self.assertEqual(result['leftSideValue'].tolist(), [0, 0, 5, 8])
        self.assertNotIn(CORRECTION_COL, previous)

    def test_successful_empty_current_analysis_marks_previous_rows_corrected(self):
        previous = pd.DataFrame([{'_zone_id': 'ZoneAAAAAAA', 'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleAAAAAAA'}])
        result = mark_previous_corrections(previous, pd.DataFrame(), {}, signatures)
        self.assertEqual(result[CORRECTION_COL].tolist(), ['Corrigée'])

    def test_failed_zone_does_not_hide_correction_states_of_other_zones_in_same_dps(self):
        previous = pd.DataFrame([
            {'_zone_id': 'ProvinceAAA', '_request_unit_id': 'ZoneAAAAAAA',
             'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleAAAAAAA'},
            {'_zone_id': 'ProvinceAAA', '_request_unit_id': 'ZoneAAAAAAA',
             'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleBBBBBBB'},
            {'_zone_id': 'ProvinceAAA', '_request_unit_id': 'ZoneBBBBBBB',
             'organisationUnitId': 'CentreBBBBB', 'validationRuleId': 'RuleCCCCCCC'},
        ])
        result = mark_previous_corrections(
            previous, previous.iloc[:1], {'ZoneBBBBBBB': ['HTTP 504']}, signatures,
        )
        self.assertEqual(result[CORRECTION_COL].tolist(), ['Non corrigée', 'Corrigée', 'Non vérifiable'])
        self.assertEqual(result['_zone_id'].tolist(), ['ProvinceAAA', 'ProvinceAAA', 'ProvinceAAA'])

    def test_missing_target_or_signature_is_not_a_proven_correction(self):
        previous = pd.DataFrame([
            {'_zone_id': '', 'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleAAAAAAA'},
            {'_zone_id': 'ZoneAAAAAAA', 'organisationUnitId': 'CentreAAAAA'},
        ])
        result = mark_previous_corrections(previous, pd.DataFrame(), {}, signatures)
        self.assertEqual(result[CORRECTION_COL].tolist(), ['Non vérifiable', 'Non vérifiable'])


if __name__ == '__main__':
    unittest.main()
