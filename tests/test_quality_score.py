import unittest

import pandas as pd

from quality_score import build_quality_rows


class QualityRowsTests(unittest.TestCase):
    def test_quality_uses_live_reporting_and_joins_validation_only_by_id(self):
        zones = pd.DataFrame([
            {'Zone_ID': 'ZoneAAAAAAA', 'Zone': 'Même zone', 'Province': 'Province A'},
            {'Zone_ID': 'ZoneBBBBBBB', 'Zone': 'Même zone', 'Province': 'Province B'},
        ])
        metrics = pd.DataFrame([
            {'Zone_ID': 'ZoneAAAAAAA', 'Complétude_Globale': 98, 'Promptitude_Globale': 80,
             'Actual_Reports': 98, 'Nombre_Datasets_95': 10},
            {'Zone_ID': 'ZoneBBBBBBB', 'Complétude_Globale': 50, 'Promptitude_Globale': 30,
             'Actual_Reports': 50, 'Nombre_Datasets_95': 4},
        ])
        validation = pd.DataFrame([
            {'Zone_ID': 'ZoneBBBBBBB', 'Zone de santé': 'Même zone', 'Violations_Actuelles': 7,
             'Complétude_Globale': 999, 'Actual_Reports': 999},
            {'Zone_ID': 'OtherZoneAA', 'Violations_Actuelles': 999},
        ])
        result = build_quality_rows(metrics, zones, validation).set_index('Zone_ID')
        self.assertEqual(result.loc['ZoneAAAAAAA', 'Score_Qualite'], 89)
        self.assertEqual(result.loc['ZoneBBBBBBB', 'Score_Qualite'], 40)
        self.assertEqual(result.loc['ZoneBBBBBBB', 'Actual_Reports'], 50)
        self.assertTrue(pd.isna(result.loc['ZoneAAAAAAA', 'Violations_Actuelles']))
        self.assertEqual(result.loc['ZoneBBBBBBB', 'Violations_Actuelles'], 7)
        self.assertNotEqual(result.iloc[0]['Zone de santé'], result.iloc[1]['Zone de santé'])

    def test_missing_inputs_stay_numeric_missing_and_zero_remains_zero(self):
        zones = pd.DataFrame([{'Zone_ID': 'ZoneAAAAAAA', 'Zone': 'A'}])
        missing = build_quality_rows(pd.DataFrame(), zones)
        for column in ['Complétude_Globale', 'Promptitude_Globale', 'Actual_Reports', 'Score_Qualite', 'Violations_Actuelles']:
            self.assertTrue(pd.api.types.is_numeric_dtype(missing[column]))
            self.assertTrue(missing[column].isna().all())
        zeros = build_quality_rows(pd.DataFrame([{
            'Zone_ID': 'ZoneAAAAAAA', 'Complétude_Globale': 0, 'Promptitude_Globale': 0,
            'Actual_Reports': 0, 'Nombre_Datasets_95': 0,
        }]), zones)
        self.assertEqual(zeros.iloc[0]['Score_Qualite'], 0)
        self.assertEqual(zeros.iloc[0]['Actual_Reports'], 0)
        self.assertFalse(zeros.iloc[0]['Controle_Corrections_Effectue'])


if __name__ == '__main__':
    unittest.main()
