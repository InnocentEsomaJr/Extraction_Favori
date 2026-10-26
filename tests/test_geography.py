import unittest

import pandas as pd

from geography import (
    CATALOG_COLUMNS, build_geography_catalog, labels_for_units,
    resolve_hierarchy, zone_catalog,
)


COUNTRY = {"id": "country-a", "displayName": "Pays A", "level": 1}
PROVINCE_A = {"id": "province-a", "displayName": "Province A", "level": 2}
PROVINCE_B = {"id": "province-b", "displayName": "Province B", "level": 2}


def fixture_units():
    return [
        COUNTRY,
        {**PROVINCE_A, "ancestors": [COUNTRY]},
        {**PROVINCE_B, "ancestors": [COUNTRY]},
        {"id": "zone-a", "displayName": "Zone commune", "level": 3, "ancestors": [PROVINCE_A, COUNTRY]},
        {"id": "zone-b", "displayName": "Zone B", "level": 3, "ancestors": [COUNTRY, PROVINCE_A]},
        {"id": "zone-c", "displayName": "Zone commune", "level": 3, "ancestors": [COUNTRY, PROVINCE_B]},
    ]


class HierarchyTests(unittest.TestCase):
    def test_levels_are_resolved_exactly_regardless_of_ancestor_order(self):
        hierarchy = resolve_hierarchy(fixture_units())
        self.assertEqual(hierarchy["zone-a"], {
            "level": 3, "country_name": "Pays A", "country_id": "country-a",
            "province_name": "Province A", "province_id": "province-a",
            "zone_name": "Zone commune", "zone_id": "zone-a",
        })
        self.assertEqual(hierarchy["province-a"]["zone_id"], "")
        self.assertEqual(hierarchy["province-a"]["zone_name"], "")
        self.assertEqual(hierarchy["country-a"]["province_id"], "")
        self.assertEqual(hierarchy["country-a"]["zone_id"], "")

    def test_missing_province_is_not_replaced_by_country(self):
        hierarchy = resolve_hierarchy([
            {"id": "zone-a", "displayName": "Zone A", "level": 3, "ancestors": [COUNTRY]},
        ])["zone-a"]
        self.assertEqual(hierarchy["country_id"], "country-a")
        self.assertEqual(hierarchy["province_id"], "")
        self.assertEqual(hierarchy["zone_id"], "zone-a")

    def test_level_four_uses_ancestor_zone_but_does_not_invent_it(self):
        zone = {"id": "zone-a", "displayName": "Zone A", "level": 3}
        hierarchy = resolve_hierarchy([
            {"id": "area-a", "displayName": "Aire A", "level": 4, "ancestors": [COUNTRY, PROVINCE_A, zone]},
            {"id": "area-b", "displayName": "Aire B", "level": 4, "ancestors": [COUNTRY, PROVINCE_A]},
        ])
        self.assertEqual(hierarchy["area-a"]["zone_id"], "zone-a")
        self.assertEqual(hierarchy["area-b"]["zone_id"], "")

    def test_invalid_metadata_does_not_create_a_false_identity(self):
        hierarchy = resolve_hierarchy([
            None, {}, {"id": None, "level": 3},
            {"id": "  province-a  ", "displayName": " Province A ", "level": "2", "ancestors": [None, COUNTRY]},
        ])
        self.assertEqual(list(hierarchy), ["province-a"])
        self.assertEqual(hierarchy["province-a"]["level"], 2)
        self.assertEqual(hierarchy["province-a"]["province_name"], "Province A")
        self.assertEqual(hierarchy["province-a"]["zone_id"], "")

    def test_conflicting_ancestors_at_same_level_remain_unresolved(self):
        unit = {"id": "zone-a", "level": 3, "ancestors": [COUNTRY, PROVINCE_A, PROVINCE_B]}
        self.assertEqual(resolve_hierarchy([unit])["zone-a"]["province_id"], "")


class CatalogTests(unittest.TestCase):
    def test_catalog_keeps_units_and_marks_only_resolved_zones(self):
        catalog = build_geography_catalog(fixture_units())
        self.assertEqual(catalog.columns.tolist(), CATALOG_COLUMNS)
        self.assertEqual(len(catalog), 6)
        self.assertEqual(catalog["is_zone_candidate"].tolist(), [False, False, False, True, True, True])
        self.assertEqual(catalog.loc[1, "Zone_ID"], "")

    def test_country_province_and_zone_filters_keep_exact_ids(self):
        catalog = build_geography_catalog(fixture_units())
        self.assertEqual(set(zone_catalog(catalog, country_id="country-a")["Zone_ID"]), {"zone-a", "zone-b", "zone-c"})
        self.assertEqual(set(zone_catalog(catalog, country_id="country-a", province_id="province-a")["Zone_ID"]), {"zone-a", "zone-b"})
        self.assertEqual(zone_catalog(catalog, country_id="country-a", province_id="province-a", zone_id="zone-a")["Zone_ID"].tolist(), ["zone-a"])
        self.assertTrue(zone_catalog(catalog, country_id="country-a", province_id="province-b", zone_id="zone-a").empty)

    def test_only_actual_level_three_units_are_listed_as_zones(self):
        units = fixture_units() + [{
            "id": "area-a", "displayName": "Aire A", "level": 4,
            "ancestors": [COUNTRY, PROVINCE_A, {"id": "zone-a", "displayName": "Zone commune", "level": 3}],
        }]
        catalog = build_geography_catalog(units)
        self.assertTrue(catalog.iloc[-1]["is_zone_candidate"])
        self.assertEqual(len(zone_catalog(catalog)), 3)
        self.assertNotIn("area-a", zone_catalog(catalog)["Organisation unit ID"].tolist())

    def test_unknown_explicit_scope_does_not_expand_to_entire_country(self):
        catalog = build_geography_catalog(fixture_units())
        for identifier in ["", "inconnu", pd.NA]:
            with self.subTest(identifier=identifier):
                self.assertTrue(zone_catalog(catalog, province_id=identifier).empty)

    def test_duplicate_units_and_empty_catalog_are_supported(self):
        catalog = build_geography_catalog(fixture_units() + fixture_units())
        self.assertEqual(len(zone_catalog(catalog)), 3)
        self.assertTrue(build_geography_catalog([]).empty)
        self.assertEqual(zone_catalog(None).columns.tolist(), CATALOG_COLUMNS)


class LabelTests(unittest.TestCase):
    def test_homonymous_zones_are_distinguished_by_province(self):
        catalog = zone_catalog(build_geography_catalog(fixture_units()))
        labels = labels_for_units(catalog, "Zone_ID", "Zone")
        self.assertEqual(labels, {
            "zone-a": "Zone commune (Province A)",
            "zone-b": "Zone B",
            "zone-c": "Zone commune (Province B)",
        })

    def test_same_province_homonyms_fall_back_to_ids(self):
        frame = pd.DataFrame({"Zone_ID": ["id-a", "id-b"], "Zone": ["Même zone", "Même zone"], "Province": ["Province A", "Province A"]})
        self.assertEqual(labels_for_units(frame, "Zone_ID", "Zone"), {"id-a": "Même zone (id-a)", "id-b": "Même zone (id-b)"})

    def test_repeated_ids_do_not_make_unique_names_ambiguous(self):
        frame = pd.DataFrame({"Province_ID": ["id-a", "id-a"], "Province": ["Province A", "Province A"]})
        self.assertEqual(labels_for_units(frame, "Province_ID", "Province"), {"id-a": "Province A"})
        self.assertEqual(labels_for_units(pd.DataFrame(), "ID", "Nom"), {})

    def test_homonym_without_province_uses_its_id(self):
        frame = pd.DataFrame({"Zone_ID": ["id-a", "id-b"], "Zone": ["Même zone", "Même zone"], "Province": ["Province A", ""]})
        self.assertEqual(labels_for_units(frame, "Zone_ID", "Zone"), {"id-a": "Même zone (Province A)", "id-b": "Même zone (id-b)"})


if __name__ == "__main__":
    unittest.main()
