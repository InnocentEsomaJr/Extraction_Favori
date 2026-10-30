"""Intégration de l'onglet de catégorisation, sans DHIS2 ni lecture de secrets."""

import ast
import copy
import io
from pathlib import Path
import unittest
from unittest import mock
from xml.etree import ElementTree as ET
import zipfile

import numpy as np
import pandas as pd
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL, build_baremes_df, style_context_tab5, style_note_20
from validation_api import fetch_validation_analysis


DASHBOARD = Path(__file__).resolve().parents[1] / "dashbord.py"


def dashboard_functions(*names, **overrides):
    tree = ast.parse(DASHBOARD.read_text(encoding="utf-8-sig"))
    functions = []
    for name in names:
        node = copy.deepcopy(next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == name))
        node.decorator_list = []
        functions.append(node)
    namespace = {
        "pd": pd, "np": np, "io": io, "style_note_20": style_note_20,
        "build_baremes_df": build_baremes_df, "fetch_validation_analysis": fetch_validation_analysis,
        **overrides,
    }
    exec(compile(ast.Module(body=functions, type_ignores=[]), str(DASHBOARD), "exec"), namespace)
    return namespace


class ValidationApiTests(unittest.TestCase):
    def test_client_posts_validation_json_using_the_authenticated_session(self):
        session = mock.Mock()
        response = session.request.return_value
        response.status_code = 200
        response.text = "[]"
        response.json.return_value = []
        requests_module = mock.Mock()
        requests_module.Session.return_value = session
        build_client = dashboard_functions(
            "_build_dhis2_client", requests=requests_module,
            Retry=mock.Mock(), HTTPAdapter=mock.Mock(),
        )["_build_dhis2_client"]
        client = build_client("https://fixture.invalid/", "fixture-user", "fixture-password")
        payload = {
            "ou": "ZoneAAAAAAA", "startDate": "2026-09-01", "endDate": "2026-09-30",
            "persist": False, "notification": False,
        }
        self.assertEqual(client.post("dataAnalysis/validationRules", data=payload), [])
        self.assertEqual(session.auth, ("fixture-user", "fixture-password"))
        session.request.assert_called_once_with(
            "POST", "https://fixture.invalid/api/dataAnalysis/validationRules",
            params=None, json=payload, timeout=(10, 90),
        )

    def test_empty_success_is_distinct_from_failed_and_malformed_periods(self):
        class Client:
            def post(self, endpoint, data=None):
                self.assert_endpoint(endpoint)
                period = data["startDate"][:7].replace("-", "")
                if period == "202501":
                    return []
                if period == "202502":
                    return [{"organisationUnitId": "ZoneAAAAAAA", "validationRuleId": "RuleAAAAAAA", "periodId": period, "leftSideValue": 0, "rightSideValue": 1}]
                if period == "202503":
                    raise RuntimeError("fixture indisponible")
                return {"unexpected": []}

            def assert_endpoint(self, endpoint):
                if endpoint != "dataAnalysis/validationRules":
                    raise AssertionError("Appel fixture non prévu")

        namespace = dashboard_functions("get_validation_results", get_dhis2_client=lambda: Client())
        get_results = namespace["get_validation_results"]
        empty = get_results("ZoneAAAAAAA", ["202501"], None)
        self.assertTrue(empty.empty)
        self.assertEqual(empty.attrs["failed_periods"], [])
        available = get_results("ZoneAAAAAAA", ["202502"], "groupA")
        self.assertEqual(len(available), 1)
        self.assertEqual(available.attrs["failed_periods"], [])
        failed = get_results("ZoneAAAAAAA", ["202503", "202504"], "groupA")
        self.assertTrue(failed.empty)
        self.assertEqual(failed.attrs["failed_periods"], ["202503", "202504"])
        malformed = get_results("ZoneAAAAAAA", ["202504"], None)
        self.assertTrue(malformed.empty)
        self.assertEqual(malformed.attrs["failed_periods"], ["202504"])

    def test_bad_identifiers_and_client_failure_are_unavailable(self):
        unavailable = mock.Mock(side_effect=RuntimeError("client indisponible"))
        get_results = dashboard_functions("get_validation_results", get_dhis2_client=unavailable)["get_validation_results"]
        invalid = get_results("USER_ORGUNIT", ["202501"], None)
        self.assertEqual(invalid.attrs["failed_periods"], ["202501"])
        unavailable.assert_not_called()
        failed = get_results("ZoneAAAAAAA", ["202501", "202502"], None)
        self.assertEqual(failed.attrs["failed_periods"], ["202501", "202502"])


class CategorisationExportTests(unittest.TestCase):
    def test_excel_preserves_new_columns_numeric_scores_missing_cells_and_note_colors(self):
        namespace = dashboard_functions("style_taux", "style_score", "_build_styler", "to_excel_dashboard_report")
        frame = pd.DataFrame({
            "Zone de santé": ["ZS A", "ZS B"],
            "Score de qualité": [95.0, 80.0],
            "Score de qualité /20": [20.0, 15.0],
            "Score complétude /20": [20.0, 0.0],
            "Score promptitude /20": [20.0, 5.0],
            "Score datasets /20": [20.0, 10.0],
            "Score violations /20": [20.0, 15.0],
            "Score corrections /20": [20.0, np.nan],
            TOTAL_COL: [100.0, np.nan],
            "Statut catégorisation": ["Complète", "Incomplète"],
        })
        context = style_context_tab5()
        styler = namespace["_build_styler"](frame, **context)
        html = styler.to_html()
        for color in ["#B91C1C", "#FDBA74", "#FDE047", "#86EFAC", "#15803D"]:
            self.assertIn(color, html)
        self.assertNotIn(">nan</td>", html)
        minimal = pd.DataFrame({"Test": [1]})
        report = namespace["to_excel_dashboard_report"](
            df_base=minimal, df_actual_detail=minimal, df_expected_detail=minimal,
            df_final_c=minimal, df_final_p=minimal, df_synth=minimal,
            comments_df=minimal, df_tab5=frame, style_context={"tab5": context},
        )
        x = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with zipfile.ZipFile(io.BytesIO(report)) as archive:
            strings = ["".join(item.itertext()) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
            sheets = ET.fromstring(archive.read("xl/workbook.xml")).find("m:sheets", x)
            self.assertIn("Baremes_Categorisation", [sheet.attrib["name"] for sheet in sheets])
            index = next(i for i, sheet in enumerate(sheets, 1) if sheet.attrib["name"] == "Resultats_Regles")
            sheet = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
            self.assertEqual(sheet.find("m:sheetViews/m:sheetView/m:pane", x).attrib["topLeftCell"], "B2")
            rows = sheet.find("m:sheetData", x)
            headers = [strings[int(cell.find("m:v", x).text)] for cell in rows[0]]
            self.assertEqual(headers, frame.columns.tolist())
            cells = {cell.attrib["r"]: cell for row in rows for cell in row}
            self.assertEqual(cells["I2"].find("m:v", x).text, "100")
            self.assertNotIn("t", cells["I2"].attrib)
            self.assertTrue("I3" not in cells or cells["I3"].find("m:v", x) is None)
            styles = ET.fromstring(archive.read("xl/styles.xml"))
            formats = styles.find("m:cellXfs", x)
            fills = styles.find("m:fills", x)
            colors = set()
            for reference in ["D2", "D3", "E3", "F3", "G3"]:
                style = formats[int(cells[reference].attrib["s"])]
                fill = fills[int(style.attrib["fillId"])]
                colors.add(fill.find("m:patternFill/m:fgColor", x).attrib["rgb"])
            self.assertEqual(len(colors), 5)


FIXTURE_FUNCTIONS = '''
def _read_config_value(key):
    return "https://fixture.invalid"

def get_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    st.session_state["_fixture_data_scope"] = ou_dimension
    data = {"Organisation unit": ["ZS A"], "Organisation unit ID": ["ZoneAAAAAAA"]}
    for number in range(10):
        prefix = f"Dataset {number} "
        data[prefix + "Actual reports"] = [10]
        data[prefix + "Expected reports"] = [10]
        data[prefix + "Actual reports on time"] = [10]
        data[prefix + "Reporting rate"] = [100]
        data[prefix + "Reporting rate on time"] = [100]
    return pd.DataFrame(data)

def get_organisation_catalog(cache_user=None):
    from geography import build_geography_catalog
    country = {"id": "CountryAAAA", "displayName": "Pays test", "level": 1}
    province = {"id": "ProvinceAAA", "displayName": "Province test", "level": 2}
    zone = {"id": "ZoneAAAAAAA", "displayName": "ZS A", "level": 3}
    return build_geography_catalog([
        country,
        {**province, "ancestors": [country]},
        {**zone, "ancestors": [country, province]},
    ])

def get_org_units_hierarchy(org_unit_ids, cache_user=None):
    return {"ZoneAAAAAAA": {"country_name": "Pays test", "country_id": "CountryAAAA",
        "province_name": "Province test", "province_id": "ProvinceAAA",
        "zone_name": "ZS A", "zone_id": "ZoneAAAAAAA", "level": 3}}

def get_validation_groups(cache_user=None):
    return [{"id": "groupAAAAAA", "displayName": "Groupe test"}]

def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            if endpoint != "dataAnalysis/validationRules":
                raise AssertionError("Appel fixture non prévu")
            st.session_state["_fixture_validation_calls"] = st.session_state.get("_fixture_validation_calls", 0) + 1
            targets = list(st.session_state.get("_fixture_validation_targets", []))
            targets.append(data["ou"])
            st.session_state["_fixture_validation_targets"] = targets
            return []
    return Client()

def test_dhis2_credentials(base_url, username, password):
    return True, None
'''


def app_fixture_source(replacements=None):
    """Remplace uniquement les frontières réseau/configuration de l'application."""
    tree = ast.parse(DASHBOARD.read_text(encoding="utf-8-sig"))
    fixture_nodes = {node.name: node for node in ast.parse(FIXTURE_FUNCTIONS).body}
    for name, source in (replacements or {}).items():
        fixture_nodes[name] = ast.parse(source).body[0]

    class Boundaries(ast.NodeTransformer):
        def visit_FunctionDef(self, node):
            return copy.deepcopy(fixture_nodes[node.name]) if node.name in fixture_nodes else node

        def visit_Expr(self, node):
            call = node.value
            if isinstance(call, ast.Call) and isinstance(call.func, ast.Attribute) and isinstance(call.func.value, ast.Name) and call.func.value.id == "st" and call.func.attr == "image":
                return ast.copy_location(ast.Pass(), node)
            return self.generic_visit(node)

    transformed = ast.fix_missing_locations(Boundaries().visit(tree))
    source = ast.unparse(transformed) + '\nst.session_state["_test_export_frame"] = df_tab5_export.copy()\n'
    # AppTest écrit son fichier temporaire avec l'encodage système de Windows.
    return source.encode("ascii", errors="backslashreplace").decode("ascii")


class CategorisationAppTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(widget for widget in getattr(app, kind) if widget.label == label)

    def run_app(self, app):
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self, replacements=None):
        app = AppTest.from_string(app_fixture_source(replacements), default_timeout=30)
        app.session_state["connected"] = True
        app.session_state["dhis2_user"] = "fixture-user"
        app.session_state["dhis2_pass"] = "fixture-password"
        return app

    def test_analysis_profile_preview_and_context_invalidation(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Tout réseau est interdit dans ce test")):
            self.run_app(app)
            self.assertTrue(app.session_state["_test_export_frame"].empty)
            self.widget(app, "button", "Lancer l'analyse des violations").click()
            self.run_app(app)
            frame = app.session_state["_test_export_frame"]
            self.assertEqual(frame[TOTAL_COL].tolist(), [100])
            self.assertEqual(frame["Barème datasets"].tolist(), ["ZS"])
            self.assertEqual(frame.columns[:9].tolist(), [
                "Zone de santé", "Complétude globale (%)", "Promptitude globale (%)", "Reports_Actual",
                "Règles violées (M-1)", "Règles corrigées (M-1 -> M)", "Règles violées (M)",
                "Ratio / 100 rapports", "Score de qualité",
            ])
            calls = app.session_state["_fixture_validation_calls"]
            self.widget(app, "selectbox", "Barème du nombre de datasets à complétude ≥95 % :").select("DPS")
            self.run_app(app)
            self.assertEqual(app.session_state["_test_export_frame"][TOTAL_COL].tolist(), [95])
            self.assertEqual(app.session_state["_fixture_validation_calls"], calls)
            self.widget(app, "button", "Visualiser le rapport").click()
            self.run_app(app)
            self.assertEqual(app.session_state["_test_export_frame"][TOTAL_COL].tolist(), [95])
            self.assertEqual(app.session_state["_fixture_validation_calls"], calls)
            self.widget(app, "selectbox", "Sélectionner le Groupe de Règles de Validation :").select("Groupe test")
            self.run_app(app)
            self.assertTrue(app.session_state["_test_export_frame"].empty)
            self.widget(app, "button", "Lancer l'analyse des violations").click()
            self.run_app(app)
            self.assertEqual(app.session_state["_test_export_frame"][TOTAL_COL].tolist(), [95])
            self.widget(app, "selectbox", "Mois fin :").select("Février")
            self.run_app(app)
            self.assertTrue(app.session_state["_test_export_frame"].empty)

    def test_country_requests_real_zone_level_and_uses_zs_profile(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Tout réseau est interdit dans ce test")):
            self.run_app(app)
            self.assertEqual(app.session_state["_fixture_data_scope"], "LEVEL-3;CountryAAAA")
            self.widget(app, "button", "Lancer l'analyse des violations").click()
            self.run_app(app)
            frame = app.session_state["_test_export_frame"]
            self.assertEqual(frame["Barème datasets"].tolist(), ["ZS"])
            self.assertEqual(frame["Score datasets /20"].tolist(), [20])
            self.assertEqual(frame[TOTAL_COL].tolist(), [100])

    def test_unavailable_catalog_does_not_invent_health_zones(self):
        app = self.new_app({"get_organisation_catalog": '''
def get_organisation_catalog(cache_user=None):
    return pd.DataFrame()
'''})
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Tout réseau est interdit dans ce test")):
            self.run_app(app)
            self.assertFalse(any(button.label == "Lancer l'analyse des violations" for button in app.button))
            self.assertGreater(len(app.error) + len(app.warning) + len(app.info), 0)


if __name__ == "__main__":
    unittest.main()
