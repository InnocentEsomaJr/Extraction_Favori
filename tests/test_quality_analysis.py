"""Les actions M et M-1 conservent leurs résultats et leurs erreurs par unité."""

import unittest

import pandas as pd

from quality_analysis import (
    analyse_quality_period, build_validation_comparison_rows,
    quality_request_targets, rollup_period_to_scores,
)
from geography import build_geography_catalog


class QualityAnalysisTests(unittest.TestCase):
    def setUp(self):
        self.targets = pd.DataFrame([
            {'Zone_ID': 'ZoneAAAAAAA', 'Zone': 'Aba', 'level': 3},
            {'Zone_ID': 'ZoneBBBBBBB', 'Zone': 'Autre ZS', 'level': 3},
        ])

    def frame(self, rules, zone='ZoneAAAAAAA'):
        return pd.DataFrame([
            {'organisationUnitId': 'CentreAAAAA', 'validationRuleId': rule,
             'organisationUnitPath': f'/CountryAAAA/ProvinceAAA/{zone}/CentreAAAAA',
             'leftSideValue': 10, 'rightSideValue': 1}
            for rule in rules
        ])

    def batch(self, first, second=None, errors=None):
        errors = errors or {}
        return {
            'frames_by_id': {'ZoneAAAAAAA': first, 'ZoneBBBBBBB': second if second is not None else pd.DataFrame()},
            'errors_by_id': errors,
            'succeeded_by_id': {identifier: not bool(errors.get(identifier)) for identifier in ['ZoneAAAAAAA', 'ZoneBBBBBBB']},
        }

    def test_m_only_counts_all_rows_and_keeps_reference_unexecuted(self):
        current, errors, warnings = analyse_quality_period(
            self.batch(self.frame(['RuleAAAAAAA', 'RuleAAAAAAA'])), self.targets
        )
        result = build_validation_comparison_rows(self.targets, current, errors).set_index('Zone_ID')
        self.assertEqual(result['Violations_Actuelles'].tolist(), [2, 0])
        self.assertTrue(result['Violations_Mois_Precedent'].isna().all())
        self.assertTrue(result['Violations_Corrigees'].isna().all())
        self.assertFalse(result['Controle_Corrections_Effectue'].any())
        self.assertEqual(warnings, [])

        self.assertEqual(current['_analysis_target_id'].tolist(), ['ZoneAAAAAAA', 'ZoneAAAAAAA'])
    def test_second_action_adds_m1_without_altering_m_and_preserves_multiplicity(self):
        current, current_errors, _ = analyse_quality_period(self.batch(self.frame(['RuleAAAAAAA'])), self.targets)
        previous, previous_errors, _ = analyse_quality_period(
            self.batch(self.frame(['RuleAAAAAAA', 'RuleBBBBBBB', 'RuleBBBBBBB'])), self.targets
        )
        original = current.copy(deep=True)
        result = build_validation_comparison_rows(self.targets, current, current_errors, previous, previous_errors)
        self.assertEqual(result['Violations_Actuelles'].tolist(), [1, 0])
        self.assertEqual(result['Violations_Mois_Precedent'].tolist(), [3, 0])
        self.assertEqual(result['Violations_Corrigees'].tolist(), [2, 0])
        self.assertTrue(result['Controle_Corrections_Effectue'].all())
        pd.testing.assert_frame_equal(current, original)

    def test_failed_m1_keeps_current_counts_and_other_units_complete(self):
        current, current_errors, _ = analyse_quality_period(self.batch(self.frame(['RuleAAAAAAA'])), self.targets)
        previous, previous_errors, warnings = analyse_quality_period(
            self.batch(pd.DataFrame(), errors={'ZoneAAAAAAA': ['M-1 : HTTP 504']}), self.targets
        )
        result = build_validation_comparison_rows(self.targets, current, current_errors, previous, previous_errors).set_index('Zone_ID')
        self.assertEqual(result['Violations_Actuelles'].tolist(), [1, 0])
        self.assertTrue(pd.isna(result.at['ZoneAAAAAAA', 'Violations_Mois_Precedent']))
        self.assertTrue(pd.isna(result.at['ZoneAAAAAAA', 'Violations_Corrigees']))
        self.assertEqual(result.at['ZoneBBBBBBB', 'Violations_Corrigees'], 0)
        self.assertEqual(warnings, ['Aba : M-1 : HTTP 504'])

    def test_successful_m1_remains_visible_when_m_failed(self):
        current, current_errors, _ = analyse_quality_period(
            self.batch(pd.DataFrame(), errors={'ZoneAAAAAAA': ['M : HTTP 504']}), self.targets
        )
        previous, previous_errors, _ = analyse_quality_period(self.batch(self.frame(['RuleAAAAAAA'])), self.targets)
        result = build_validation_comparison_rows(self.targets, current, current_errors, previous, previous_errors).set_index('Zone_ID')
        self.assertTrue(pd.isna(result.at['ZoneAAAAAAA', 'Violations_Actuelles']))
        self.assertEqual(result.at['ZoneAAAAAAA', 'Violations_Mois_Precedent'], 1)
        self.assertTrue(pd.isna(result.at['ZoneAAAAAAA', 'Violations_Corrigees']))
        self.assertEqual(result.at['ZoneBBBBBBB', 'Violations_Corrigees'], 0)

    def test_unassigned_line_invalidates_only_its_requested_target_and_is_kept(self):
        unknown = self.frame(['RuleAAAAAAA'])
        unknown['organisationUnitPath'] = '/CountryAAAA/OtherProvAA/OtherZoneAA/CentreAAAAA'
        current, errors, warnings = analyse_quality_period(self.batch(unknown), self.targets)
        result = build_validation_comparison_rows(self.targets, current, errors).set_index('Zone_ID')
        self.assertEqual(len(current), 1)
        self.assertEqual(current['_zone_id'].tolist(), [''])
        self.assertTrue(pd.isna(result.at['ZoneAAAAAAA', 'Violations_Actuelles']))
        self.assertEqual(result.at['ZoneBBBBBBB', 'Violations_Actuelles'], 0)
        self.assertTrue(warnings)
        self.assertEqual(current.attrs, {})

    def test_missing_batch_target_is_unavailable_instead_of_false_zero(self):
        batch = self.batch(pd.DataFrame())
        del batch['frames_by_id']['ZoneAAAAAAA']
        current, errors, _ = analyse_quality_period(batch, self.targets)
        result = build_validation_comparison_rows(self.targets, current, errors).set_index('Zone_ID')
        self.assertTrue(pd.isna(result.at['ZoneAAAAAAA', 'Violations_Actuelles']))
        self.assertEqual(result.at['ZoneBBBBBBB', 'Violations_Actuelles'], 0)

    def test_custom_signature_function_preserves_existing_identity_contract(self):
        current, errors, _ = analyse_quality_period(self.batch(self.frame(['RuleAAAAAAA'])), self.targets)
        previous = current.copy(deep=True)
        calls = []

        def signatures(frame):
            calls.append(len(frame))
            return ['stable-key'] * len(frame)

        result = build_validation_comparison_rows(self.targets, current, errors, previous, {}, signatures)
        self.assertEqual(result['Violations_Corrigees'].tolist(), [0, 0])
        self.assertEqual(calls, [1, 1, 0, 0])

    def test_dps_uses_its_own_analysis_target_and_keeps_zeros_for_other_provinces(self):
        targets = pd.DataFrame([
            {'Zone_ID': 'ProvinceAAA', 'Zone': 'DPS A', 'level': 2},
            {'Zone_ID': 'ProvinceBBB', 'Zone': 'DPS B', 'level': 2},
        ])
        batch = {
            'frames_by_id': {'ProvinceAAA': self.frame(['RuleAAAAAAA']), 'ProvinceBBB': pd.DataFrame()},
            'errors_by_id': {}, 'succeeded_by_id': {'ProvinceAAA': True, 'ProvinceBBB': True},
        }
        current, errors, warnings = analyse_quality_period(batch, targets)
        result = build_validation_comparison_rows(targets, current, errors)
        self.assertEqual(current['_zone_id'].tolist(), ['ProvinceAAA'])
        self.assertEqual(result['Violations_Actuelles'].tolist(), [1, 0])
        self.assertEqual(warnings, [])


class QualityScoreRollupTests(unittest.TestCase):
    def setUp(self):
        country = {'id': 'CountryAAAA', 'displayName': 'RDC', 'level': 1}
        first = {'id': 'ProvinceAAA', 'displayName': 'DPS A', 'level': 2}
        second = {'id': 'ProvinceBBB', 'displayName': 'DPS B', 'level': 2}
        self.catalog = build_geography_catalog([
            country, {**first, 'ancestors': [country]}, {**second, 'ancestors': [country]},
            {'id': 'ZoneAAAAAAA', 'displayName': 'Aba', 'level': 3, 'ancestors': [country, first]},
            {'id': 'ZoneBBBBBBB', 'displayName': 'ZS B', 'level': 3, 'ancestors': [country, first]},
            {'id': 'ZoneCCCCCCC', 'displayName': 'ZS C', 'level': 3, 'ancestors': [country, second]},
        ])
        self.scores = pd.DataFrame([
            {'Zone_ID': 'ProvinceAAA', 'Zone': 'DPS A', 'level': 2, 'Profil_Bareme': 'DPS', 'Pays_ID': 'CountryAAAA'},
            {'Zone_ID': 'ProvinceBBB', 'Zone': 'DPS B', 'level': 2, 'Profil_Bareme': 'DPS', 'Pays_ID': 'CountryAAAA'},
        ])

    def test_country_requests_only_accessible_zones_and_preserves_the_score_link(self):
        requests = quality_request_targets(self.scores, self.catalog)
        self.assertEqual(requests['Zone_ID'].tolist(), ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'])
        self.assertEqual(requests['_score_target_id'].tolist(), ['ProvinceAAA', 'ProvinceAAA', 'ProvinceBBB'])
        self.assertTrue(requests['level'].eq(3).all())
        self.assertTrue(requests['Profil_Bareme'].eq('ZS').all())

    def test_two_zones_roll_up_to_a_dps_and_other_province_failure_stays_isolated(self):
        requests = quality_request_targets(self.scores, self.catalog)
        leaf_frame = pd.DataFrame([
            {'_analysis_target_id': 'ZoneAAAAAAA', '_zone_id': 'ZoneAAAAAAA',
             'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleAAAAAAA', 'leftSideValue': 10},
            {'_analysis_target_id': 'ZoneAAAAAAA', '_zone_id': 'ZoneAAAAAAA',
             'organisationUnitId': 'CentreAAAAA', 'validationRuleId': 'RuleBBBBBBB', 'leftSideValue': 20},
        ])
        rolled, errors, warnings = rollup_period_to_scores(
            leaf_frame, {'ZoneAAAAAAA': [], 'ZoneBBBBBBB': [], 'ZoneCCCCCCC': ['HTTP 504']}, requests, self.scores,
        )
        result = build_validation_comparison_rows(self.scores, rolled, errors).set_index('Zone_ID')
        self.assertEqual(result.at['ProvinceAAA', 'Violations_Actuelles'], 2)
        self.assertTrue(pd.isna(result.at['ProvinceBBB', 'Violations_Actuelles']))
        self.assertEqual(rolled['_request_unit_id'].tolist(), ['ZoneAAAAAAA', 'ZoneAAAAAAA'])
        self.assertEqual(rolled['_zone_id'].tolist(), ['ProvinceAAA', 'ProvinceAAA'])
        self.assertEqual(rolled['_zone_name'].tolist(), ['DPS A', 'DPS A'])
        self.assertEqual(rolled['leftSideValue'].tolist(), [10, 20])
        self.assertEqual(errors['ProvinceAAA'], [])
        self.assertEqual(warnings, ['DPS B : ZS C : HTTP 504'])
        self.assertNotIn('_request_unit_id', leaf_frame)

    def test_failure_of_one_leaf_invalidates_its_dps_only(self):
        requests = quality_request_targets(self.scores, self.catalog)
        rolled, errors, _ = rollup_period_to_scores(
            pd.DataFrame(), {'ZoneBBBBBBB': ['M-1 indisponible']}, requests, self.scores,
        )
        result = build_validation_comparison_rows(self.scores, rolled, errors).set_index('Zone_ID')
        self.assertTrue(pd.isna(result.at['ProvinceAAA', 'Violations_Actuelles']))
        self.assertEqual(result.at['ProvinceBBB', 'Violations_Actuelles'], 0)

    def test_dps_without_a_leaf_is_unavailable_and_does_not_create_a_province_request(self):
        scores = pd.concat([self.scores, pd.DataFrame([
            {'Zone_ID': 'ProvinceCCC', 'Zone': 'DPS sans ZS', 'level': 2, 'Profil_Bareme': 'DPS'},
        ])], ignore_index=True)
        requests = quality_request_targets(scores, self.catalog)
        self.assertNotIn('ProvinceCCC', requests['Zone_ID'].tolist())
        rolled, errors, warnings = rollup_period_to_scores(pd.DataFrame(), {}, requests, scores)
        result = build_validation_comparison_rows(scores, rolled, errors).set_index('Zone_ID')
        self.assertEqual(result.loc[['ProvinceAAA', 'ProvinceBBB'], 'Violations_Actuelles'].tolist(), [0, 0])
        self.assertTrue(pd.isna(result.at['ProvinceCCC', 'Violations_Actuelles']))
        self.assertIn('Aucune zone de santé accessible', warnings[0])

    def test_single_selected_zone_keeps_its_direct_request_and_own_zero(self):
        scores = pd.DataFrame([{'Zone_ID': 'ZoneAAAAAAA', 'Zone': 'Aba', 'level': 3, 'Profil_Bareme': 'ZS'}])
        requests = quality_request_targets(scores, self.catalog)
        self.assertEqual(requests['Zone_ID'].tolist(), ['ZoneAAAAAAA'])
        self.assertEqual(requests['_score_target_id'].tolist(), ['ZoneAAAAAAA'])
        rolled, errors, warnings = rollup_period_to_scores(pd.DataFrame(), {}, requests, scores)
        result = build_validation_comparison_rows(scores, rolled, errors)
        self.assertEqual(result['Violations_Actuelles'].tolist(), [0])
        self.assertEqual(warnings, [])


if __name__ == '__main__':
    unittest.main()
