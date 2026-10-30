"""Comparaison des corrections avec champs facultatifs et lignes répétées."""

import ast
from copy import deepcopy
from pathlib import Path
import unittest

import pandas as pd


def signature_functions():
    dashboard = Path(__file__).resolve().parents[1] / "dashbord.py"
    tree = ast.parse(dashboard.read_text(encoding="utf-8-sig"))
    names = {"_extract_identifier", "build_violation_signatures", "build_violation_signature_set"}
    nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names]
    namespace = {"pd": pd}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(dashboard), "exec"), namespace)
    return namespace


class ValidationSignatureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.functions = signature_functions()

    def row(self, **changes):
        row = {
            "organisationUnit": {"id": "CentreAAAAA"},
            "validationRule": {"id": "RuleAAAAAAA"},
            "period": {"id": "202609"}, "leftsideValue": 84, "rightsideValue": 59,
        }
        row.update(changes)
        return row

    def signatures(self, rows):
        return self.functions["build_violation_signatures"](pd.DataFrame(rows))

    def test_missing_identifier_variants_do_not_become_distinct_text_ids(self):
        identifier = self.functions["_extract_identifier"]
        for value in [None, float("nan"), pd.NA, pd.NaT, "", " nan ", "NULL", "None", "<NA>", "NaT"]:
            with self.subTest(value=repr(value)):
                self.assertEqual(identifier(value), "")
        self.assertEqual(identifier({"id": pd.NA, "uid": "AocAAAAAAAA"}), "AocAAAAAAAA")
        self.assertEqual(identifier({"id": float("nan"), "displayName": "Combinaison A"}), "Combinaison A")

    def test_mixed_optional_columns_do_not_report_an_ongoing_violation_as_corrected(self):
        previous = self.signatures([self.row()])
        current = self.signatures([
            self.row(period={"id": "202610"}),
            self.row(
                organisationUnit={"id": "CentreBBBBB"}, validationRule={"id": "RuleBBBBBBB"},
                attributeOptionCombo={"id": "AocAAAAAAAA"}, categoryOptionCombo={"id": "CocAAAAAAAA"},
                dataSet={"id": "DatasetAAAA"},
            ),
        ])
        self.assertIn(previous[0], current)
        self.assertEqual(sum(signature not in set(current) for signature in previous), 0)

    def test_each_previous_row_is_counted_when_a_repeated_violation_is_gone(self):
        rows = [self.row(), self.row(), self.row(validationRule={"id": "RuleBBBBBBB"})]
        previous = self.signatures(rows)
        self.assertEqual(len(previous), 3)
        self.assertEqual(previous[0], previous[1])
        current = self.signatures([self.row(validationRule={"id": "RuleBBBBBBB"})])
        corrected = sum(signature not in set(current) for signature in previous)
        self.assertEqual(corrected, 2)
        self.assertEqual(len(previous) - corrected, 1)

    def test_period_and_values_do_not_change_identity_but_combinations_do(self):
        previous = self.signatures([self.row(attributeOptionCombo={"id": "AocAAAAAAAA"})])[0]
        current = self.signatures([
            self.row(
                period={"id": "202610"}, leftsideValue=20, rightsideValue=18,
                attributeOptionCombo={"id": "AocAAAAAAAA"},
            ),
            self.row(attributeOptionCombo={"id": "AocBBBBBBBB"}),
        ])
        self.assertEqual(previous, current[0])
        self.assertNotEqual(previous, current[1])

    def test_empty_or_unidentified_results_and_source_data_are_safe(self):
        rows = [self.row(), {"organisationUnit": pd.NA, "validationRule": float("nan")}]
        frame = pd.DataFrame(rows)
        original = deepcopy(frame.to_dict("records"))
        self.assertEqual(len(self.functions["build_violation_signatures"](frame)), 1)
        self.assertEqual(frame.iloc[0].to_dict(), original[0])
        self.assertEqual(self.functions["build_violation_signatures"](None), [])
        self.assertEqual(self.functions["build_violation_signature_set"](pd.DataFrame()), set())


if __name__ == "__main__":
    unittest.main()
