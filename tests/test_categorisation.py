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
            "Violations_Mois_Precedent": [0], "Violations_Corrigees": [0],
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
        outstanding = calculer_categorisation(self.frame(Violations_Mois_Precedent=8, Violations_Corrigees=5))
        self.assertEqual(outstanding.loc[0, REMAINING_COL], 3)
        self.assertEqual(outstanding.loc[0, COMPONENT_COLS[4]], 0)
        none_previous = calculer_categorisation(self.frame(Violations_Mois_Precedent=0, Violations_Corrigees=0))
        self.assertEqual(none_previous.loc[0, COMPONENT_COLS[4]], 20)
        for corrected in [None, -1, 9]:
            with self.subTest(corrected=corrected):
                invalid = calculer_categorisation(self.frame(Violations_Mois_Precedent=8, Violations_Corrigees=corrected))
                self.assertTrue(pd.isna(invalid.loc[0, REMAINING_COL]))
                self.assertEqual(invalid.loc[0, COMPONENT_COLS[4]], 0)
                self.assertEqual(invalid.loc[0, TOTAL_COL], 80)

    def test_failed_current_analysis_keeps_known_previous_note_independent(self):
        actual = calculer_categorisation(self.frame(Erreur_Validation="période indisponible"))
        self.assertTrue(pd.isna(actual.loc[0, COMPONENT_COLS[3]]))
        self.assertEqual(actual.loc[0, COMPONENT_COLS[4]], 20)
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))
        self.assertIn("analyse des violations indisponible", actual.loc[0, "Statut catégorisation"])

    def test_current_analysis_scores_violations_before_optional_correction_control(self):
        actual = calculer_categorisation(self.frame(
            Controle_Corrections_Effectue=False, Erreur_Corrections="",
            Violations_Mois_Precedent=None, Violations_Corrigees=None,
        ))
        self.assertEqual(actual.loc[0, COMPONENT_COLS[3]], 20)
        self.assertTrue(pd.isna(actual.loc[0, COMPONENT_COLS[4]]))
        self.assertTrue(pd.isna(actual.loc[0, REMAINING_COL]))
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))
        self.assertEqual(actual.loc[0, "Statut catégorisation"],
                         "Incomplète : contrôle des corrections non effectué")

    def test_unperformed_control_does_not_reuse_stale_correction_counts(self):
        actual = calculer_categorisation(self.frame(Controle_Corrections_Effectue=False))
        self.assertEqual(actual.loc[0, COMPONENT_COLS[3]], 20)
        self.assertTrue(pd.isna(actual.loc[0, COMPONENT_COLS[4]]))
        self.assertTrue(pd.isna(actual.loc[0, REMAINING_COL]))
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))

    def test_failed_reference_control_does_not_erase_a_valid_current_violation_score(self):
        actual = calculer_categorisation(self.frame(
            Controle_Corrections_Effectue=True, Erreur_Corrections="M-1 : HTTP 403",
            Violations_Mois_Precedent=None, Violations_Corrigees=None,
            Violations_Actuelles=20,
        ))
        self.assertEqual(actual.loc[0, COMPONENT_COLS[3]], 5)
        self.assertTrue(pd.isna(actual.loc[0, COMPONENT_COLS[4]]))
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))
        self.assertEqual(actual.loc[0, "Statut catégorisation"],
                         "Incomplète : contrôle des corrections indisponible")

    def test_successful_reference_control_completes_the_total(self):
        actual = calculer_categorisation(self.frame(
            Controle_Corrections_Effectue=True, Erreur_Corrections="",
        ))
        self.assertEqual(actual.loc[0, COMPONENT_COLS[3:]].tolist(), [20, 20])
        self.assertEqual(actual.loc[0, TOTAL_COL], 100)
        self.assertEqual(actual.loc[0, "Statut catégorisation"], "Complète")

    def test_current_failure_blocks_total_but_not_known_reference_note(self):
        actual = calculer_categorisation(self.frame(
            Controle_Corrections_Effectue=True, Erreur_Corrections="", Erreur_Validation="M : HTTP 403",
        ))
        self.assertTrue(pd.isna(actual.loc[0, COMPONENT_COLS[3]]))
        self.assertEqual(actual.loc[0, COMPONENT_COLS[4]], 20)
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))
        self.assertIn("analyse des violations indisponible", actual.loc[0, "Statut catégorisation"])

    def test_mixed_control_states_stay_independent_per_zone(self):
        source = pd.concat([
            self.frame(Controle_Corrections_Effectue=False, Erreur_Corrections=""),
            self.frame(Controle_Corrections_Effectue=True, Erreur_Corrections="M-1 indisponible"),
            self.frame(Controle_Corrections_Effectue=True, Erreur_Corrections=""),
        ], ignore_index=True)
        actual = calculer_categorisation(source)
        self.assertEqual(actual[COMPONENT_COLS[3]].tolist(), [20, 20, 20])
        self.assertTrue(actual.loc[:1, COMPONENT_COLS[4]].isna().all())
        self.assertEqual(actual.loc[2, TOTAL_COL], 100)

    def test_ratio_exactly_twenty_retains_five_points(self):
        actual = calculer_categorisation(self.frame(Violations_Actuelles=20))
        self.assertEqual(actual.loc[0, COMPONENT_COLS[3]], 5)
        self.assertEqual(actual.loc[0, TOTAL_COL], 85)

    def test_previous_period_is_all_or_nothing_even_when_every_violation_is_corrected(self):
        for previous, corrected in [(1, 0), (1, 1), (2, 2), (8, 8), (8, 0)]:
            with self.subTest(previous=previous, corrected=corrected):
                actual = calculer_categorisation(self.frame(
                    Violations_Mois_Precedent=previous, Violations_Corrigees=corrected,
                    Controle_Corrections_Effectue=True,
                ))
                self.assertEqual(actual.loc[0, COMPONENT_COLS[4]], 0)
                self.assertEqual(actual.loc[0, REMAINING_COL], previous - corrected)
                self.assertEqual(actual.loc[0, TOTAL_COL], 80)

    def test_aba_strict_previous_rule_scores_seventy_five(self):
        actual = calculer_categorisation(self.frame(
            Profil_Bareme="ZS", Complétude_Globale=100, Promptitude_Globale=100,
            Score_Qualite=100, Nombre_Datasets_95=9, Actual_Reports=134,
            Violations_Actuelles=5, Violations_Mois_Precedent=2,
            Violations_Corrigees=2, Controle_Corrections_Effectue=True,
        ))
        self.assertEqual(actual.loc[0, COMPONENT_COLS].tolist(), [20, 20, 15, 20, 0])
        self.assertEqual(actual.loc[0, TOTAL_COL], 75)
        self.assertEqual(actual.loc[0, REMAINING_COL], 0)

    def test_missing_previous_count_is_unavailable_even_if_comparison_claims_no_errors(self):
        actual = calculer_categorisation(self.frame(
            Violations_Mois_Precedent=None, Violations_Corrigees=0,
            Controle_Corrections_Effectue=True,
        ))
        self.assertTrue(pd.isna(actual.loc[0, COMPONENT_COLS[4]]))
        self.assertTrue(pd.isna(actual.loc[0, TOTAL_COL]))
        self.assertIn("violations de la période précédente indisponibles", actual.loc[0, "Statut catégorisation"])

    def test_total_is_an_integer_multiple_of_five_between_zero_and_one_hundred(self):
        frames = []
        for rate in [0, 50, 60, 80, 95]:
            for datasets in [0, 5, 6, 8, 10]:
                for violations in [0, 7, 12, 17, 25]:
                    for previous in [0, 1]:
                        frames.append(self.frame(
                            Profil_Bareme="ZS", Complétude_Globale=rate,
                            Promptitude_Globale=rate, Nombre_Datasets_95=datasets,
                            Violations_Actuelles=violations, Violations_Mois_Precedent=previous,
                            Violations_Corrigees=previous,
                        ))
        totals = calculer_categorisation(pd.concat(frames, ignore_index=True))[TOTAL_COL]
        self.assertTrue(totals.between(0, 100).all())
        self.assertTrue(totals.mod(5).eq(0).all())

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
        self.assertIn("Aucune règle violée en M-1", legend["Condition"].tolist())
        self.assertIn("Au moins une règle violée en M-1, même corrigée en M", legend["Condition"].tolist())


if __name__ == "__main__":
    unittest.main()
