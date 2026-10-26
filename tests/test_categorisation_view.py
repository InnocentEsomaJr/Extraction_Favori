import unittest

import pandas as pd

from categorisation import TOTAL_COL
from categorisation_view import (
    VIEW_SORT_OPTIONS, VIEW_STATUS_OPTIONS, filter_categorisation,
)


class CategorisationViewTests(unittest.TestCase):
    def frame(self):
        return pd.DataFrame({
            "Zone de santé": ["ZS Équateur", "ZS  Bumba", "ZS Bukavu", "ZS zéro", "ZS Kasai"],
            "Barème datasets": ["ZS", "DPS", "ZS", "ZS", "DPS"],
            "Statut catégorisation": ["Complète", "Incomplète", "Complète", "Complète", "Complète"],
            "Province": ["Équateur", "Mongala", "Sud Kivu", "Kasaï", "Kasaï"],
            TOTAL_COL: [90.0, float("nan"), 60.0, 0.0, 60.0],
            "Règles violées (M)": [1, 2, 3, 4, 5],
        }, index=[17, 9, 42, 6, 21])

    def test_search_ignores_case_accents_and_extra_spaces(self):
        frame = self.frame()
        self.assertEqual(filter_categorisation(frame, query="  zs   EQUATEUR ").index.tolist(), [17])
        self.assertEqual(filter_categorisation(frame, query="zs bumba").index.tolist(), [9])
        self.assertEqual(filter_categorisation(frame, query="kasai").index.tolist(), [6, 21])
        self.assertEqual(filter_categorisation(frame, query="dps").index.tolist(), [9, 21])
        self.assertEqual(filter_categorisation(frame, query="incomplete").index.tolist(), [9])

    def test_query_uses_literal_text_not_regular_expressions(self):
        frame = self.frame()
        self.assertTrue(filter_categorisation(frame, query=".*").empty)

    def test_complete_and_incomplete_use_scores_not_status_text(self):
        frame = self.frame()
        frame.loc[17, "Statut catégorisation"] = "Incomplète"
        frame.loc[9, "Statut catégorisation"] = "Complète"
        self.assertEqual(
            filter_categorisation(frame, status="Scores complets").index.tolist(), [17, 42, 6, 21]
        )
        self.assertEqual(filter_categorisation(frame, status="Scores incomplets").index.tolist(), [9])

    def test_zero_is_complete_and_priority_does_not_include_missing_scores(self):
        frame = self.frame()
        frame.loc[17, TOTAL_COL] = 80
        self.assertEqual(
            filter_categorisation(frame, status="Score inférieur à 80").index.tolist(), [42, 6, 21]
        )
        self.assertIn(6, filter_categorisation(frame, status="Scores complets").index)

    def test_score_sorts_are_stable_and_missing_is_last(self):
        frame = self.frame()
        self.assertEqual(
            filter_categorisation(frame, sort_by="Score croissant").index.tolist(), [6, 42, 21, 17, 9]
        )
        self.assertEqual(
            filter_categorisation(frame, sort_by="Score décroissant").index.tolist(), [17, 42, 21, 6, 9]
        )
        self.assertEqual(filter_categorisation(frame).index.tolist(), frame.index.tolist())

    def test_name_sort_ignores_accents_and_places_missing_names_last(self):
        frame = pd.DataFrame({"Zone de santé": ["Équateur", "bumba", None, "equateur"]}, index=[4, 3, 2, 1])
        self.assertEqual(filter_categorisation(frame, sort_by="Nom de zone").index.tolist(), [3, 4, 1, 2])

    def test_combined_filters_preserve_columns_types_and_source(self):
        frame = self.frame()
        before = frame.copy(deep=True)
        actual = filter_categorisation(frame, query="zs", status="Scores complets", sort_by="Score croissant")
        self.assertEqual(actual.index.tolist(), [6, 42, 21, 17])
        self.assertEqual(actual.columns.tolist(), frame.columns.tolist())
        pd.testing.assert_series_equal(actual.dtypes, frame.dtypes)
        actual.loc[6, "Règles violées (M)"] = 99
        pd.testing.assert_frame_equal(frame, before)

    def test_duplicate_indexes_do_not_duplicate_rows_when_sorting(self):
        frame = self.frame()
        frame.index = [1, 1, 2, 2, 2]
        actual = filter_categorisation(frame, sort_by="Score croissant")
        self.assertEqual(len(actual), len(frame))
        self.assertEqual(actual["Règles violées (M)"].tolist(), [4, 3, 5, 1, 2])

    def test_empty_and_missing_text_columns_are_supported(self):
        empty = self.frame().iloc[:0]
        pd.testing.assert_frame_equal(filter_categorisation(empty, query="test", status="Scores incomplets"), empty)
        frame = pd.DataFrame({TOTAL_COL: [0.0, float("nan")]}, index=[5, 6])
        self.assertTrue(filter_categorisation(frame, query="absent").empty)
        self.assertEqual(filter_categorisation(frame, status="Scores incomplets").index.tolist(), [6])
        pd.testing.assert_frame_equal(filter_categorisation(frame, sort_by="Nom de zone"), frame)

    def test_missing_score_column_is_incomplete_and_has_no_priority_rows(self):
        frame = pd.DataFrame({"Zone de santé": ["ZS A", "ZS B"]})
        self.assertTrue(filter_categorisation(frame, status="Scores complets").empty)
        self.assertEqual(len(filter_categorisation(frame, status="Scores incomplets")), 2)
        self.assertTrue(filter_categorisation(frame, status="Score inférieur à 80").empty)
        pd.testing.assert_frame_equal(filter_categorisation(frame, sort_by="Score croissant"), frame)

    def test_option_constants_are_supported_and_invalid_choices_are_explicit(self):
        for status in VIEW_STATUS_OPTIONS:
            for sort_by in VIEW_SORT_OPTIONS:
                filter_categorisation(self.frame(), status=status, sort_by=sort_by)
        with self.assertRaises(ValueError):
            filter_categorisation(self.frame(), status="Inconnu")
        with self.assertRaises(ValueError):
            filter_categorisation(self.frame(), sort_by="Inconnu")


if __name__ == "__main__":
    unittest.main()
