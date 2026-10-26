import unittest

import numpy as np
import pandas as pd

from categorisation import (
    COMPONENT_COLS, DATASET_COUNT_COL, NOTE_COLS, QUALITY_NOTE_COL,
    REMAINING_COL, TOTAL_COL, build_baremes_df, calculer_categorisation,
    score_datasets, score_taux, score_violations, style_context_tab5, style_note_20,
)


class ScoreTests(unittest.TestCase):
    def test_rate_thresholds_and_decimals(self):
        source = pd.Series([0, 49.99, 50, 59.99, 60, 79.99, 80, 94.99, 95, 100])
        self.assertEqual(score_taux(source).tolist(), [0, 0, 5, 5, 10, 10, 15, 15, 20, 20])

    def test_invalid_inputs_do_not_become_good_scores(self):
        source = pd.Series([None, np.nan, -1, np.inf, -np.inf, "inconnu", pd.NA])
        for actual in [score_taux(source), score_violations(source), score_datasets(source, "DPS")]:
            self.assertTrue(actual.isna().all())

    def test_dataset_profiles_have_distinct_thresholds(self):
        counts = pd.Series([0, 4, 5, 6, 7, 8, 9, 10, 11, 12])
        self.assertEqual(score_datasets(counts, "DPS").tolist(), [0, 0, 0, 5, 5, 10, 10, 15, 15, 20])
        self.assertEqual(score_datasets(counts, "ZS").tolist(), [0, 0, 5, 10, 10, 15, 15, 20, 20, 20])
        mixed = score_datasets(pd.Series([10, 10, 10], index=[9, 4, 2]), pd.Series(["ZS", None, "DPS"], index=[4, 2, 9]))
        self.assertEqual(mixed.loc[9], 15)
        self.assertEqual(mixed.loc[4], 20)
        self.assertTrue(pd.isna(mixed.loc[2]))

    def test_violation_ratio_boundaries(self):
        ratios = pd.Series([0, 5, 5.01, 10, 10.01, 15, 15.01, 19.99, 20, 20.01])
        self.assertEqual(score_violations(ratios).tolist(), [20, 20, 15, 15, 10, 10, 5, 5, 5, 0])


class CategorisationTests(unittest.TestCase):
    def frame(self, **changes):
        data = {
            "Zone de santé": ["Zone test"], "Complétude_Globale": [95.0],
            "Promptitude_Globale": [95.0], "Score_Qualite": [95.0],
            "Actual_Reports": [100.0], "Violations_Actuelles": [5],
            "Violations_Mois_Precedent": [8], "Violations_Corrigees": [8],
            "Nombre_Datasets_95": [12], "Profil_Bareme": ["DPS"],
            "Erreur_Validation": [""],
        }
        data.update({name: [value] for name, value in changes.items()})
        return pd.DataFrame(data)

    def test_complete_total_uses_five_components_not_historical_quality(self):
        source = self.frame(Score_Qualite=55)
        actual = calculer_categorisation(source)
        self.assertEqual(actual.loc[0, QUALITY_NOTE_COL], 5)
        self.assertEqual(actual.loc[0, COMPONENT_COLS].tolist(), [20, 20, 20, 20, 20])
        self.assertEqual(actual.loc[0, TOTAL_COL], 100)
        self.assertEqual(actual.loc[0, REMAINING_COL], 0)
        self.assertEqual(actual.loc[0, "Statut catégorisation"], "Complète")
        self.assertEqual(source.columns.tolist(), actual.columns[:len(source.columns)].tolist())
        self.assertNotIn(TOTAL_COL, source.columns)

    def test_missing_auto_profile_is_explicit_and_never_guessed(self):
        actual = calculer_categorisation(self.frame(Profil_Bareme=None))
        self.assertTrue(pd.isna(actual.loc[0, COMPONENT_COLS[2]]))
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))
        self.assertIn("barème datasets à préciser", actual.loc[0, "Statut catégorisation"])
        forced = calculer_categorisation(self.frame(Profil_Bareme=None, Nombre_Datasets_95=10), profil="ZS")
        self.assertEqual(forced.loc[0, COMPONENT_COLS[2]], 20)
        self.assertEqual(forced.loc[0, "Barème datasets"], "ZS")

    def test_zero_reports_cannot_produce_a_perfect_violation_score(self):
        for reports in [0, -1, None, np.inf]:
            with self.subTest(reports=reports):
                actual = calculer_categorisation(self.frame(Actual_Reports=reports))
                self.assertTrue(pd.isna(actual.loc[0, "Ratio_Violations_sur_100"]))
                self.assertTrue(pd.isna(actual.loc[0, COMPONENT_COLS[3]]))
                self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))

    def test_corrections_remaining_and_incoherent_counts(self):
        outstanding = calculer_categorisation(self.frame(Violations_Corrigees=5))
        self.assertEqual(outstanding.loc[0, REMAINING_COL], 3)
        self.assertEqual(outstanding.loc[0, COMPONENT_COLS[4]], 0)
        none_previous = calculer_categorisation(self.frame(Violations_Mois_Precedent=0, Violations_Corrigees=0))
        self.assertEqual(none_previous.loc[0, COMPONENT_COLS[4]], 20)
        for corrected in [None, -1, 9]:
            with self.subTest(corrected=corrected):
                invalid = calculer_categorisation(self.frame(Violations_Corrigees=corrected))
                self.assertTrue(pd.isna(invalid.loc[0, REMAINING_COL]))
                self.assertTrue(pd.isna(invalid.loc[0, COMPONENT_COLS[4]]))
                self.assertTrue(pd.isna(invalid.loc[0, TOTAL_COL]))

    def test_failed_analysis_keeps_both_violation_components_unavailable(self):
        actual = calculer_categorisation(self.frame(Erreur_Validation="période indisponible"))
        self.assertTrue(actual.loc[0, COMPONENT_COLS[3:]].isna().all())
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))
        self.assertIn("analyse des violations indisponible", actual.loc[0, "Statut catégorisation"])

    def test_ratio_exactly_twenty_retains_five_points(self):
        actual = calculer_categorisation(self.frame(Violations_Actuelles=20))
        self.assertEqual(actual.loc[0, COMPONENT_COLS[3]], 5)
        self.assertEqual(actual.loc[0, TOTAL_COL], 85)

    def test_unrounded_ratios_determine_the_band(self):
        for ratio, expected in [(5.004, 15), (10.004, 10), (15.004, 5), (20.004, 0)]:
            with self.subTest(ratio=ratio):
                actual = calculer_categorisation(self.frame(Violations_Actuelles=ratio, Actual_Reports=100))
                self.assertAlmostEqual(actual.loc[0, "Ratio_Violations_sur_100"], ratio)
                self.assertEqual(actual.loc[0, COMPONENT_COLS[3]], expected)

    def test_missing_component_produces_no_partial_total_and_empty_is_supported(self):
        actual = calculer_categorisation(self.frame(Nombre_Datasets_95=None))
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))
        self.assertTrue(pd.isna(actual.loc[0, DATASET_COUNT_COL]))
        empty = calculer_categorisation(pd.DataFrame())
        self.assertTrue(empty.empty)
        self.assertIn(TOTAL_COL, empty.columns)

    def test_style_context_and_legend_match_confirmed_rules(self):
        colors = [style_note_20(note) for note in [0, 5, 10, 15, 20]]
        self.assertEqual(len(set(colors)), 5)
        self.assertEqual(style_note_20(np.nan), "")
        context = style_context_tab5()
        self.assertEqual(context["note_cols"], NOTE_COLS)
        self.assertIn(TOTAL_COL, context["int_cols"])
        legend = build_baremes_df()
        self.assertIn("95 % et plus", legend["Condition"].tolist())
        self.assertIn("Plus de 20", legend["Condition"].tolist())
        self.assertIn("Aucune règle restante", legend["Condition"].tolist())


if __name__ == "__main__":
    unittest.main()
