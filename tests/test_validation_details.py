import copy
import unittest

import pandas as pd

from validation_details import (
    DISPLAY_COLUMNS, INTERNAL_COLUMNS, build_validation_details, detail_export_frame,
    filter_validation_details, style_importance, validation_detail_identifiers,
)


class ValidationDetailsTests(unittest.TestCase):
    def row(self, **changes):
        row = {
            "organisationUnit": {"id": "OuAAAAAAA"}, "validationRule": {"id": "RuleAAAAA"},
            "leftsideValue": 0, "rightsideValue": 3, "period": "202609",
            "_zone_id": "ZoneAAAAA", "_zone_name": "ZS Équateur",
        }
        row.update(changes)
        return row

    def metadata(self):
        return {
            "RuleAAAAA": {
                "displayName": "Doses administrées", "importance": "HIGH", "operator": "less_than_or_equal_to",
                "instruction": "Vérifier les quantités.",
                "leftSide": {"description": "Doses", "expression": "#{doses}"},
                "rightSide": {"description": "Flacons", "expression": "#{flacons} * 10"},
            },
        }, {"OuAAAAAAA": {"displayName": "Centre de santé A"}}

    def test_nested_api_shape_metadata_and_zero(self):
        rules, units = self.metadata()
        result = build_validation_details([self.row()], rules, units)
        self.assertEqual(result.columns.tolist(), DISPLAY_COLUMNS + INTERNAL_COLUMNS)
        self.assertEqual(result.loc[0, "Unité d’organisation"], "Centre de santé A")
        self.assertEqual(result.loc[0, "Règle de validation"], "Doses administrées")
        self.assertEqual(result.loc[0, "Valeur gauche"], 0)
        self.assertEqual(result.loc[0, "Valeur droite"], 3)
        self.assertEqual(result.loc[0, "Opérateur"], "<=")
        self.assertEqual(result.loc[0, "Importance"], "HIGH")
        self.assertEqual(result.loc[0, "Période"], "Septembre 2026")
        self.assertIn("Vérifier les quantités.", result.loc[0, "Détails"])
        self.assertIn("Gauche : Doses (#{doses})", result.loc[0, "Détails"])

    def test_flat_fresh_analysis_fields_have_priority(self):
        rules, units = self.metadata()
        row = self.row(
            validationRuleId="RuleBBBBB", validationRuleDescription="Règle fraîche",
            organisationUnitId="OuBBBBBBB", organisationUnitDisplayName="Centre frais",
            periodId="202610", periodDisplayName="October 2026", importance="MEDIUM",
            leftSideValue="0", rightSideValue=4.5, operator=">",
            attributeOptionComboId="ComboAAAAA", attributeOptionComboDisplayName="Attribut A",
        )
        result = build_validation_details([row], rules, units)
        self.assertEqual(result.loc[0, "_rule_id"], "RuleBBBBB")
        self.assertEqual(result.loc[0, "_ou_id"], "OuBBBBBBB")
        self.assertEqual(result.loc[0, "Unité d’organisation"], "Centre frais")
        self.assertEqual(result.loc[0, "Règle de validation"], "Règle fraîche")
        self.assertEqual(result.loc[0, "Période"], "Octobre 2026")
        self.assertEqual(result.loc[0, "Valeur gauche"], 0)
        self.assertEqual(result.loc[0, "Valeur droite"], 4.5)
        self.assertEqual(result.loc[0, "Opérateur"], ">")
        self.assertIn("Attribut A", result.loc[0, "Détails"])

    def test_missing_values_are_not_zero_and_metadata_absence_shows_ids(self):
        result = build_validation_details([self.row(leftsideValue=None, rightsideValue=pd.NA)])
        self.assertTrue(pd.isna(result.loc[0, "Valeur gauche"]))
        self.assertTrue(pd.isna(result.loc[0, "Valeur droite"]))
        self.assertEqual(result.loc[0, "Règle de validation"], "RuleAAAAA")
        self.assertEqual(result.loc[0, "Unité d’organisation"], "OuAAAAAAA")
        self.assertEqual(result.loc[0, "Importance"], "Non renseignée")

    def test_month_year_other_period_and_requested_fallback(self):
        rows = [
            self.row(period={"id": "202602", "name": "February 2026"}),
            self.row(period={"id": "2026"}),
            self.row(period={"id": "2026Q3", "name": "Troisième trimestre 2026"}),
            self.row(period=None, _requested_period="202601"),
            self.row(period={"startDate": "2026-09-01"}),
        ]
        result = build_validation_details(rows)
        self.assertEqual(result["Période"].tolist(), [
            "Février 2026", "2026", "Troisième trimestre 2026", "Janvier 2026", "2026-09-01",
        ])

    def test_duplicates_are_removed_without_merging_periods_values_or_combinations(self):
        base = self.row()
        rows = [
            base, copy.deepcopy(base), self.row(period="202610"), self.row(leftsideValue=1),
            self.row(attributeOptionCombo="ComboAAAAA"), self.row(categoryOptionCombo="ComboAAAAA"),
            self.row(dayInPeriod=0), self.row(dayInPeriod=1),
        ]
        result = build_validation_details(rows)
        self.assertEqual(len(result), 7)
        self.assertEqual(result["_period_id"].tolist().count("202609"), 6)

    def test_nested_iso_precedes_database_id_but_flat_period_remains_priority(self):
        result = build_validation_details([
            self.row(period={"id": 12345, "iso": "202609"}),
            self.row(period={"id": 12345, "iso": "202609"}, periodId="202610"),
        ])
        self.assertEqual(result["_period_id"].tolist(), ["202609", "202610"])
        self.assertEqual(result["Période"].tolist(), ["Septembre 2026", "Octobre 2026"])

    def test_result_uid_deduplicates_repeated_observation_but_retains_other_period(self):
        rows = [self.row(id="ResultAAA"), self.row(id="ResultAAA"), self.row(id="ResultAAA", period="202610")]
        result = build_validation_details(rows)
        self.assertEqual(len(result), 2)
        self.assertEqual(result["_result_id"].tolist(), ["ResultAAA", "ResultAAA"])

    def test_homonyms_stay_distinct_by_unit_rule_ids(self):
        rows = [self.row(), self.row(organisationUnit="OuBBBBBBB"), self.row(validationRule="RuleBBBBB")]
        rules = {key: {"displayName": "Même règle"} for key in ["RuleAAAAA", "RuleBBBBB"]}
        units = {key: {"displayName": "Même centre"} for key in ["OuAAAAAAA", "OuBBBBBBB"]}
        self.assertEqual(len(build_validation_details(rows, rules, units)), 3)

    def test_zone_name_can_be_resolved_from_path_and_ancestors(self):
        rows = [
            self.row(_zone_id=None, _zone_name=None, organisationUnitPath="/CountryAA/ProvinceA/ZoneAAAAA/OuAAAAAAA"),
            self.row(_zone_id=None, _zone_name=None, organisationUnit={
                "id": "OuBBBBBBB", "ancestors": [{"id": "ZoneBBBBB", "displayName": "ZS B", "level": 3}],
            }),
            self.row(_zone_id=None, _zone_name=None, organisationUnit={"id": "ZoneCCCCC", "displayName": "ZS C", "level": 3}),
        ]
        result = build_validation_details(rows, org_metadata={"ZoneAAAAA": {"displayName": "ZS A"}})
        self.assertEqual(result["_zone_id"].tolist(), ["ZoneAAAAA", "ZoneBBBBB", "ZoneCCCCC"])
        self.assertEqual(result["Zone de santé"].tolist(), ["ZS A", "ZS B", "ZS C"])

    def test_operator_variants_and_unknown_values_are_preserved(self):
        operators = ["equal_to", "equal", "equals", "==", "=", "not_equal_to", "less_than", "greater_than_or_equal_to", "compulsory_pair", "exclusive_pair", "CUSTOM"]
        rows = [self.row(id=str(index), operator=operator) for index, operator in enumerate(operators)]
        result = build_validation_details(rows)
        self.assertEqual(result["Opérateur"].tolist(), ["==", "==", "==", "==", "=", "!=", "<", ">=", "Paire obligatoire", "Paire exclusive", "CUSTOM"])

    def test_identifiers_helper_handles_flat_nested_and_numeric_periods(self):
        identifiers = validation_detail_identifiers([
            self.row(), self.row(), self.row(validationRuleId="RuleBBBBB", organisationUnitId="OuBBBBBBB"),
        ])
        self.assertEqual(identifiers, {"rule_ids": ["RuleAAAAA", "RuleBBBBB"], "org_ids": ["OuAAAAAAA", "OuBBBBBBB"]})
        result = build_validation_details([self.row(periodId=202609.0)])
        self.assertEqual(result.loc[0, "_period_id"], "202609")

    def test_source_dataframe_and_metadata_are_not_mutated(self):
        rules, units = self.metadata()
        source = pd.DataFrame([self.row()])
        original_source = copy.deepcopy(source.to_dict("records"))
        original_rules = copy.deepcopy(rules)
        original_units = copy.deepcopy(units)
        build_validation_details(source, rules, units)
        self.assertEqual(source.to_dict("records"), original_source)
        self.assertEqual(rules, original_rules)
        self.assertEqual(units, original_units)


class ValidationDetailsViewTests(unittest.TestCase):
    def frame(self):
        return build_validation_details([
            {"validationRuleId": "RuleAAAAA", "validationRuleDescription": "Doses administrées", "organisationUnitId": "OuAAAAAAA",
             "organisationUnitDisplayName": "Centre Équateur", "periodId": "202609", "importance": "HIGH",
             "_zone_id": "ZoneAAAAA", "_zone_name": "ZS Équateur", "leftSideValue": 0, "rightSideValue": 3},
            {"validationRuleId": "RuleBBBBB", "organisationUnitId": "OuBBBBBBB", "periodId": "202610",
             "importance": "LOW", "_zone_id": "ZoneBBBBB", "_zone_name": "ZS B"},
        ])

    def test_combined_filters_use_ids_and_plain_accent_insensitive_text(self):
        frame = self.frame()
        actual = filter_validation_details(frame, query="  centre   EQUATEUR ", importance="high", zone_id="ZoneAAAAA", period_id="202609")
        self.assertEqual(actual["_rule_id"].tolist(), ["RuleAAAAA"])
        self.assertTrue(filter_validation_details(frame, query=".*").empty)
        self.assertTrue(filter_validation_details(frame, zone_id="ZoneBBBBB", period_id="202609").empty)
        self.assertTrue(filter_validation_details(frame, zone_id="").empty)

    def test_filters_and_exports_are_independent_copies(self):
        frame = self.frame()
        before = frame.copy(deep=True)
        actual = filter_validation_details(frame, importance="LOW")
        actual.loc[1, "Règle de validation"] = "Changed"
        exported = detail_export_frame(frame)
        self.assertEqual(exported.columns.tolist(), DISPLAY_COLUMNS)
        self.assertEqual(detail_export_frame(frame, include_ids=True).columns[-2:].tolist(), ["ID règle", "ID unité"])
        pd.testing.assert_frame_equal(frame, before)

    def test_empty_inputs_and_missing_filter_columns_are_safe(self):
        frame = build_validation_details(None)
        self.assertTrue(frame.empty)
        self.assertEqual(frame.columns.tolist(), DISPLAY_COLUMNS + INTERNAL_COLUMNS)
        pd.testing.assert_frame_equal(filter_validation_details(frame, query="anything"), frame)
        self.assertEqual(detail_export_frame(pd.DataFrame()).columns.tolist(), DISPLAY_COLUMNS)
        self.assertTrue(filter_validation_details(pd.DataFrame({"Règle de validation": ["A"]}), zone_id="ZoneA").empty)

    def test_severity_styles_are_static_and_unknown_is_neutral(self):
        self.assertEqual(len(set(style_importance(value) for value in ["HIGH", "MEDIUM", "LOW"])), 3)
        self.assertEqual(style_importance(None), "")
        self.assertEqual(style_importance("Non renseignée"), "")
        self.assertEqual(style_importance("<script>"), "")


if __name__ == "__main__":
    unittest.main()
