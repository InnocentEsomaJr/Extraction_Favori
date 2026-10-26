import unittest

import pandas as pd

from scoped_analytics import analytics_to_reporting_frame, build_analytics_query


class AnalyticsQueryTests(unittest.TestCase):
    def favorite(self):
        return {
            "dataDimensionItems": [
                {"dataDimensionItemType": "REPORTING_RATE", "reportingRate": {
                    "dataSet": {"id": "DatasetAAAA", "displayName": "Consultations"}, "metric": metric,
                }}
                for metric in ["REPORTING_RATE", "REPORTING_RATE_ON_TIME", "ACTUAL_REPORTS", "ACTUAL_REPORTS_ON_TIME", "EXPECTED_REPORTS"]
            ],
            "columns": [{"id": "dx"}], "rows": [{"id": "ou"}],
            "filters": [{"dimension": "pe", "items": [{"id": "202401"}]}],
            "organisationUnits": [{"id": "ProvinceAAA"}],
        }

    def test_selected_scope_and_period_override_favorite_and_metrics_are_canonical(self):
        params, labels = build_analytics_query(self.favorite(), "202501;202502", "LEVEL-3;CountryAAAA")
        self.assertIn(("dimension", "ou:LEVEL-3;CountryAAAA"), params)
        self.assertIn(("filter", "pe:202501;202502"), params)
        self.assertNotIn(("dimension", "pe:202501;202502"), params)
        self.assertNotIn("ProvinceAAA", str(params))
        self.assertNotIn("202401", str(params))
        self.assertEqual(len([key for key, _ in params if key == "dimension"]), 2)
        self.assertEqual(list(labels.values()), [
            "Consultations Reporting rate", "Consultations Reporting rate on time", "Consultations Actual reports",
            "Consultations Actual reports on time", "Consultations Expected reports",
        ])

    def test_other_explicit_filters_and_aggregation_options_are_preserved(self):
        favorite = self.favorite()
        favorite["filters"].append({"dimension": "CategoryAAA", "items": [{"id": "OptionAAAAA"}]})
        favorite.update(aggregationType="SUM", skipRounding=True, completedOnly=False)
        params, _ = build_analytics_query(favorite, ["202501", "202502"], "LEVEL-3;ProvinceAAA")
        self.assertIn(("filter", "CategoryAAA:OptionAAAAA"), params)
        self.assertIn(("aggregationType", "SUM"), params)
        self.assertIn(("skipRounding", "true"), params)
        self.assertIn(("completedOnly", "false"), params)

    def test_fallback_dx_items_and_detail_category_selections(self):
        favorite = {"columns": [{"dimension": "dx", "items": [{"id": "SetAAAAAAA.ACTUAL_REPORTS", "name": "Consultations Actual reports"}]}],
                    "filters": [{"id": "CategoryAAA"}],
                    "categoryDimensions": [{"category": {"id": "CategoryAAA"}, "categoryOptions": [{"id": "OptionAAAAA"}]}]}
        params, labels = build_analytics_query(favorite, "202501", "ZoneAAAAAAA")
        self.assertEqual(labels, {"SetAAAAAAA.ACTUAL_REPORTS": "Consultations Actual reports"})
        self.assertIn(("filter", "CategoryAAA:OptionAAAAA"), params)

    def test_report_metric_id_only_and_generic_indicator(self):
        favorite = {"dataDimensionItems": [
            {"reportingRate": {"id": "SetAAAAAAA.EXPECTED_REPORTS", "displayName": "Consultations Expected reports"}},
            {"indicator": {"id": "IndicatorAA", "displayName": "Indicateur"}},
        ]}
        _, labels = build_analytics_query(favorite, "202501", "ZoneAAAAAAA")
        self.assertEqual(labels["SetAAAAAAA.EXPECTED_REPORTS"], "Consultations Expected reports")
        self.assertEqual(labels["IndicatorAA"], "Indicateur")

    def test_unknown_or_unselected_extra_dimensions_are_not_silently_aggregated(self):
        favorite = self.favorite()
        favorite["rows"].append({"id": "CategoryAAA"})
        with self.assertRaisesRegex(ValueError, "sélection explicite"):
            build_analytics_query(favorite, "202501", "LEVEL-3;ProvinceAAA")
        with self.assertRaises(ValueError):
            build_analytics_query({}, "202501", "ZoneAAAAAAA")
        with self.assertRaises(ValueError):
            build_analytics_query(self.favorite(), "", "ZoneAAAAAAA")

    def test_tracker_attribute_dimensions_are_not_mistaken_for_category_filters(self):
        favorite = self.favorite()
        favorite["attributeDimensions"] = [{"attribute": {"id": "AttributeAA"}, "filter": "EQ:valeur"}]
        with self.assertRaisesRegex(ValueError, "Tracker"):
            build_analytics_query(favorite, "202501", "ZoneAAAAAAA")

    def test_homonymous_datasets_receive_distinct_consistent_metric_prefixes(self):
        favorite = {"dataDimensionItems": [
            {"reportingRate": {"dataSet": {"id": dataset, "name": "Même nom"}, "metric": metric}}
            for dataset in ["DatasetAAAA", "DatasetBBBB"] for metric in ["REPORTING_RATE", "ACTUAL_REPORTS", "EXPECTED_REPORTS"]
        ]}
        _, labels = build_analytics_query(favorite, "202501", "ZoneAAAAAAA")
        self.assertEqual(len(set(labels.values())), 6)
        self.assertEqual(labels["DatasetAAAA.REPORTING_RATE"], "Même nom (DatasetAAAA) Reporting rate")
        self.assertEqual(labels["DatasetAAAA.ACTUAL_REPORTS"], "Même nom (DatasetAAAA) Actual reports")


class AnalyticsFrameTests(unittest.TestCase):
    def response(self):
        return {
            "headers": [{"name": "dx"}, {"name": "ou"}, {"name": "value"}],
            "rows": [["DatasetAAAA.ACTUAL_REPORTS", "ZoneAAAAAAA", "10"],
                     ["DatasetAAAA.ACTUAL_REPORTS", "ZoneBBBBBBB", "20"],
                     ["DatasetAAAA.EXPECTED_REPORTS", "ZoneAAAAAAA", "10"]],
            "metaData": {"items": {"ZoneAAAAAAA": {"name": "Homonyme"}, "ZoneBBBBBBB": {"name": "Homonyme"}, "DatasetAAAA": {"name": "Consultations"}}},
        }

    def test_homonymous_org_units_stay_separate_by_identifier(self):
        frame = analytics_to_reporting_frame(self.response())
        self.assertEqual(frame["Organisation unit"].tolist(), ["Homonyme", "Homonyme"])
        self.assertEqual(frame["Organisation unit ID"].tolist(), ["ZoneAAAAAAA", "ZoneBBBBBBB"])
        self.assertEqual(frame["Consultations Actual reports"].tolist(), [10, 20])
        self.assertTrue(pd.isna(frame.loc[1, "Consultations Expected reports"]))

    def test_requested_columns_order_and_absent_metrics_are_preserved(self):
        labels = {"DatasetAAAA.EXPECTED_REPORTS": "Consultations Expected reports", "DatasetAAAA.ACTUAL_REPORTS": "Consultations Actual reports", "DatasetAAAA.REPORTING_RATE": "Consultations Reporting rate"}
        frame = analytics_to_reporting_frame(self.response(), labels)
        self.assertEqual(frame.columns.tolist(), ["Organisation unit", "Organisation unit ID", *labels.values()])
        self.assertTrue(frame["Consultations Reporting rate"].isna().all())

    def test_metadata_enriches_only_identifier_fallbacks_and_keeps_explicit_names(self):
        response = self.response()
        labels = {"DatasetAAAA.ACTUAL_REPORTS": "DatasetAAAA Actual reports",
                  "DatasetAAAA.EXPECTED_REPORTS": "Libellé choisi Expected reports"}
        frame = analytics_to_reporting_frame(response, labels)
        self.assertEqual(frame.columns.tolist(), ["Organisation unit", "Organisation unit ID",
                                                "Consultations Actual reports", "Libellé choisi Expected reports"])
        response["metaData"]["items"].update({"IndicatorAA": {"name": "Indicateur lisible"}})
        frame = analytics_to_reporting_frame({"rows": [], "metaData": response["metaData"]}, {"IndicatorAA": "IndicatorAA"})
        self.assertIn("Indicateur lisible", frame.columns)

    def test_metadata_enrichment_resolves_homonymous_dataset_names_without_collision(self):
        response = self.response()
        response["rows"].append(["DatasetBBBB.ACTUAL_REPORTS", "ZoneAAAAAAA", "7"])
        response["metaData"]["items"]["DatasetBBBB"] = {"name": "Consultations"}
        labels = {"DatasetAAAA.ACTUAL_REPORTS": "DatasetAAAA Actual reports",
                  "DatasetBBBB.ACTUAL_REPORTS": "DatasetBBBB Actual reports"}
        frame = analytics_to_reporting_frame(response, labels)
        self.assertIn("Consultations (DatasetAAAA) Actual reports", frame.columns)
        self.assertIn("Consultations (DatasetBBBB) Actual reports", frame.columns)
        self.assertEqual(frame.loc[0, "Consultations (DatasetBBBB) Actual reports"], 7)

    def test_duplicate_values_raise_instead_of_averaging(self):
        response = self.response()
        response["rows"].append(response["rows"][0])
        with self.assertRaisesRegex(ValueError, "plusieurs valeurs"):
            analytics_to_reporting_frame(response)

    def test_extra_dimensions_and_invalid_values_are_explicit(self):
        response = self.response()
        response["headers"].append({"name": "pe"})
        with self.assertRaisesRegex(ValueError, "uniquement"):
            analytics_to_reporting_frame(response)
        response = self.response()
        response["rows"][0][2] = "indisponible"
        with self.assertRaisesRegex(ValueError, "numérique"):
            analytics_to_reporting_frame(response)

    def test_empty_response_keeps_identifier_and_requested_columns(self):
        frame = analytics_to_reporting_frame({"rows": []}, {"set.ACTUAL_REPORTS": "DS Actual reports"})
        self.assertTrue(frame.empty)
        self.assertEqual(frame.columns.tolist(), ["Organisation unit", "Organisation unit ID", "DS Actual reports"])


if __name__ == "__main__":
    unittest.main()
