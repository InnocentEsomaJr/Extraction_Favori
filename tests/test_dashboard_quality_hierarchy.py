"""Scores groupés pays/DPS/ZS et calcul de l'exemple Aba, sans réseau."""

from datetime import date
import unittest
from unittest import mock

import pandas as pd
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
import test_dashboard_categorisation as dashboard_test


CATALOG = '''
def get_organisation_catalog(cache_user=None):
    from geography import build_geography_catalog
    country = {"id": "CountryAAAA", "displayName": "RDC test", "level": 1}
    pa = {"id": "ProvinceAAA", "displayName": "Haut Uele", "level": 2}
    pb = {"id": "ProvinceBBB", "displayName": "Autre DPS", "level": 2}
    units = [country, {**pa, "ancestors": [country]}, {**pb, "ancestors": [country]}]
    for uid, name, province in [
        ("ZoneAAAAAAA", "Aba", pa),
        ("ZoneBBBBBBB", "ZS sans violation", pa),
        ("ZoneCCCCCCC", "ZS indisponible", pa),
        ("ZoneDDDDDDD", "ZS autre DPS", pb),
    ]:
        units.append({"id": uid, "displayName": name, "level": 3, "ancestors": [country, province]})
    return build_geography_catalog(units)
'''


DATA = '''
def get_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    zones = [
        ("ZoneAAAAAAA", "Aba", "ProvinceAAA"),
        ("ZoneBBBBBBB", "ZS sans violation", "ProvinceAAA"),
        ("ZoneCCCCCCC", "ZS indisponible", "ProvinceAAA"),
        ("ZoneDDDDDDD", "ZS autre DPS", "ProvinceBBB"),
    ]
    if ou_dimension == "LEVEL-2;CountryAAAA":
        selected = [("ProvinceAAA", "Haut Uele", "ProvinceAAA"), ("ProvinceBBB", "Autre DPS", "ProvinceBBB")]
    elif ou_dimension == "LEVEL-3;CountryAAAA":
        selected = zones
    elif ou_dimension and ou_dimension.startswith("LEVEL-3;"):
        selected = [unit for unit in zones if unit[2] == ou_dimension.split(";", 1)[1]]
    else:
        selected = [unit for unit in zones if unit[0] == ou_dimension]
    rows = []
    for uid, name, province in selected:
        row = {"Organisation unit": name, "Organisation unit ID": uid}
        count = 12 if uid.startswith("Province") else 9
        for number in range(count):
            expected = 14 if number == 0 and uid == "ZoneAAAAAAA" else 15
            rate = 60 if str(period).endswith("02") else 100
            prefix = f"Dataset {number} "
            row[prefix + "Actual reports"] = expected * rate / 100
            row[prefix + "Expected reports"] = expected
            row[prefix + "Actual reports on time"] = expected * rate / 100
            row[prefix + "Reporting rate"] = rate
            row[prefix + "Reporting rate on time"] = rate
        rows.append(row)
    return pd.DataFrame(rows)
'''


REPORTING = '''
def get_quality_reporting_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    calls = list(st.session_state.get("_hierarchy_reporting", []))
    calls.append({"period": period, "ou_dimension": ou_dimension})
    st.session_state["_hierarchy_reporting"] = calls
    return get_data(favori_id, period=period, cache_user=cache_user, ou_dimension=ou_dimension)
'''


CLIENT = '''
def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            import requests
            if endpoint != "dataAnalysis/validationRules":
                raise AssertionError(endpoint)
            calls = list(st.session_state.get("_hierarchy_validation", []))
            calls.append(dict(data))
            st.session_state["_hierarchy_validation"] = calls
            previous = data["startDate"] < st.session_state["validation_start_date"].isoformat()
            if (data["ou"] in st.session_state.get("_hierarchy_failed_ids", [])
                or previous and data["ou"] in st.session_state.get("_hierarchy_failed_previous_ids", [])):
                response = requests.Response()
                response.status_code = 504
                raise requests.HTTPError("Fixture : relais trop lent", response=response)
            if data["ou"] not in ["CountryAAAA", "ProvinceAAA", "ZoneAAAAAAA"]:
                return []
            rows = []
            for number in range(2 if previous else 5):
                centre = "CentreAAAAA"
                rows.append({
                    "organisationUnitId": centre,
                    "organisationUnitDisplayName": "Centre Aba",
                    "organisationUnitPath": "/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/" + centre,
                    "validationRuleId": f"Rule{number:07d}",
                    "validationRuleDescription": f"Règle {number}",
                    "periodId": data["startDate"][:7].replace("-", ""),
                    "leftSideValue": 5 + number,
                    "rightSideValue": 1,
                    "operator": "<=", "importance": "MEDIUM",
                })
            return rows
    return Client()
'''


METADATA = '''
def get_validation_detail_metadata(rule_ids, org_ids, cache_user=None):
    return {}, {"CentreAAAAA": {
        "id": "CentreAAAAA", "displayName": "Centre Aba",
        "path": "/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/CentreAAAAA",
    }}, []
'''


class QualityHierarchyAppTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        app.run(timeout=45)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self):
        app = AppTest.from_string(dashboard_test.app_fixture_source({
            'get_organisation_catalog': CATALOG, 'get_data': DATA,
            'get_quality_reporting_data': REPORTING, 'get_dhis2_client': CLIENT,
            'get_validation_detail_metadata': METADATA,
        }), default_timeout=45)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-hierarchy-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        app.session_state['quality_detail_level'] = 'DPS / provinces'
        return app

    def frame(self, app):
        return app.session_state['_test_export_frame'].copy()

    def select_province(self, app):
        self.widget(app, 'selectbox', 'Province de l’analyse').select('ProvinceAAA')
        self.run_app(app)

    def select_country_dps(self, app):
        self.widget(app, 'selectbox', 'Détail du tableau qualité').select('DPS / provinces')
        self.run_app(app)

    def batch(self, app):
        self.widget(app, 'button', 'Calculer tous les scores (M et M-1)').click()
        self.run_app(app)

    def detail_view(self, app):
        frames = [table.value for table in app.dataframe if {'Règle de validation', 'Importance', 'Période'}.issubset(table.value.columns)]
        return frames[0] if frames else pd.DataFrame()

    def test_country_dps_and_explicit_zone_detail_use_the_matching_level(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_country_dps(app)
            frame = self.frame(app)
            self.assertEqual(frame.columns[0], 'Province / DPS')
            self.assertEqual(set(frame['Province / DPS']), {'Haut Uele', 'Autre DPS'})
            self.assertEqual(frame['Barème datasets'].tolist(), ['DPS', 'DPS'])
            self.assertEqual(frame['Nombre de datasets complétude ≥95 %'].tolist(), [12, 12])
            self.assertTrue(frame[TOTAL_COL].isna().all())
            self.assertEqual(app.session_state['_hierarchy_reporting'][-1]['ou_dimension'], 'LEVEL-2;CountryAAAA')
            self.widget(app, 'selectbox', 'Détail du tableau qualité').select('Zones de santé')
            self.run_app(app)
            zones = self.frame(app)
            self.assertEqual(len(zones), 4)
            self.assertEqual(set(zones['Barème datasets']), {'ZS'})
            self.assertNotIn('_hierarchy_validation', app.session_state)

    def test_province_batch_computes_aba_75_keeps_zero_and_isolates_failed_zone(self):
        app = self.new_app()
        app.session_state['_hierarchy_failed_ids'] = ['ZoneCCCCCCC']
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_province(app)
            self.batch(app)
            frame = self.frame(app).set_index('Zone de santé')
            self.assertEqual(set(frame.index), {'Aba', 'ZS sans violation', 'ZS indisponible'})
            aba = frame.loc['Aba']
            self.assertEqual(aba['Reports_Actual'], 134)
            self.assertEqual(aba['Règles violées (M)'], 5)
            self.assertAlmostEqual(aba['Ratio / 100 rapports'], 5 / 134 * 100)
            self.assertEqual(aba['Règles violées (M-1)'], 2)
            self.assertEqual(aba['Règles corrigées (M-1 -> M)'], 0)
            self.assertEqual(aba[[
                'Score complétude /20', 'Score promptitude /20', 'Score datasets /20',
                'Score violations /20', 'Score corrections /20',
            ]].tolist(), [20, 20, 15, 20, 0])
            self.assertEqual(aba[TOTAL_COL], 75)
            self.assertEqual(frame.loc['ZS sans violation', 'Règles violées (M)'], 0)
            self.assertEqual(frame.loc['ZS sans violation', TOTAL_COL], 95)
            self.assertTrue(frame.loc['ZS indisponible', [TOTAL_COL, 'Règles violées (M)']].isna().all())
            self.assertEqual(frame.loc['ZS indisponible', 'Complétude globale (%)'], 100)
            calls = app.session_state['_hierarchy_validation']
            self.assertEqual(set(call['ou'] for call in calls), {'ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'})
            self.assertTrue(all(call['persist'] is False and call['notification'] is False for call in calls))
            ordered = self.frame(app).columns.tolist()
            self.assertLess(ordered.index('Score complétude /20'), ordered.index('Promptitude globale (%)'))
            self.assertLess(ordered.index('Score promptitude /20'), ordered.index('Nombre de datasets complétude ≥95 %'))
            self.assertLess(ordered.index('Score datasets /20'), ordered.index('Ratio / 100 rapports'))
            self.assertLess(ordered.index('Score violations /20'), ordered.index('Règles restantes (M-1 -> M)'))
            self.assertLess(ordered.index('Score corrections /20'), ordered.index(TOTAL_COL))

    def test_selected_zone_and_dates_replace_scores_without_static_results(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_province(app)
            self.widget(app, 'selectbox', 'Zone de santé de l’analyse').select('ZoneAAAAAAA')
            self.run_app(app)
            self.batch(app)
            self.assertEqual(self.frame(app)[TOTAL_COL].tolist(), [75])
            calls = len(app.session_state['_hierarchy_validation'])
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            self.widget(app, 'date_input', 'Date de début').set_value(date(year, 2, 1))
            self.widget(app, 'date_input', 'Date de fin').set_value(date(year, 2, 28))
            self.run_app(app)
            february = self.frame(app)
            self.assertEqual(february['Complétude globale (%)'].tolist(), [60])
            self.assertAlmostEqual(february['Reports_Actual'].iloc[0], 134 * 0.6)
            self.assertTrue(february['Règles violées (M)'].isna().all())
            self.assertTrue(february[TOTAL_COL].isna().all())
            self.assertEqual(len(app.session_state['_hierarchy_validation']), calls)
            self.widget(app, 'selectbox', 'Zone de santé de l’analyse').select('ZoneBBBBBBB')
            self.run_app(app)
            self.batch(app)
            zero = self.frame(app)
            self.assertEqual(zero['Zone de santé'].tolist(), ['ZS sans violation'])
            self.assertEqual(zero['Règles violées (M)'].tolist(), [0])
            self.assertTrue(zero[TOTAL_COL].notna().all())

    def test_country_batch_keeps_successful_dps_when_another_dps_times_out(self):
        app = self.new_app()
        app.session_state['_hierarchy_failed_ids'] = ['ZoneDDDDDDD']
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_country_dps(app)
            self.batch(app)
            frame = self.frame(app).set_index('Province / DPS')
            self.assertEqual(frame.loc['Haut Uele', TOTAL_COL], 80)
            self.assertEqual(frame.loc['Haut Uele', 'Règles violées (M)'], 5)
            self.assertTrue(frame.loc['Autre DPS', [TOTAL_COL, 'Règles violées (M)']].isna().all())
            self.assertEqual(frame.loc['Autre DPS', 'Complétude globale (%)'], 100)
            calls = app.session_state['_hierarchy_validation']
            self.assertEqual(set(call['ou'] for call in calls), {
                'ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC', 'ZoneDDDDDDD',
            })
            self.assertFalse({'CountryAAAA', 'ProvinceAAA', 'ProvinceBBB'}.intersection(call['ou'] for call in calls))

    def test_previous_month_failure_keeps_current_violations_and_other_zone_scores(self):
        app = self.new_app()
        app.session_state['_hierarchy_failed_previous_ids'] = ['ZoneAAAAAAA']
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_province(app)
            self.batch(app)
            frame = self.frame(app).set_index('Zone de santé')
            aba = frame.loc['Aba']
            self.assertEqual(aba['Complétude globale (%)'], 100)
            self.assertEqual(aba['Promptitude globale (%)'], 100)
            self.assertEqual(aba['Reports_Actual'], 134)
            self.assertEqual(aba['Règles violées (M)'], 5)
            self.assertAlmostEqual(aba['Ratio / 100 rapports'], 5 / 134 * 100)
            self.assertEqual(aba['Score violations /20'], 20)
            self.assertTrue(aba[[
                'Règles violées (M-1)', 'Règles corrigées (M-1 -> M)',
                'Règles restantes (M-1 -> M)', 'Score corrections /20', TOTAL_COL,
            ]].isna().all())
            for label in ['ZS sans violation', 'ZS indisponible']:
                self.assertEqual(frame.loc[label, 'Règles violées (M)'], 0)
                self.assertEqual(frame.loc[label, 'Règles violées (M-1)'], 0)
                self.assertEqual(frame.loc[label, 'Score corrections /20'], 20)
                self.assertEqual(frame.loc[label, TOTAL_COL], 95)
            self.assertEqual(len(app.session_state['_hierarchy_validation']), 6)

    def test_previous_button_displays_its_own_rules_and_current_view_switch_never_reloads(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_province(app)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            current_view = self.detail_view(app).copy(deep=True)
            current_saved = app.session_state['categorisation_result']['violations'].copy(deep=True)
            self.assertEqual(len(current_view), 5)
            self.assertEqual(len(app.session_state['_hierarchy_validation']), 3)
            self.assertEqual(self.widget(app, 'selectbox', 'Période des violations affichées').value, 'Période filtrée (M)')
            self.assertTrue(self.frame(app)['Règles violées (M-1)'].isna().all())
            self.assertTrue(self.frame(app)[TOTAL_COL].isna().all())
            self.assertTrue(any(button.label == 'Calculer tous les scores (M et M-1)' for button in app.button))

            self.widget(app, 'button', 'Contrôler les corrections du mois précédent').click()
            self.run_app(app)
            saved = app.session_state['categorisation_result']
            self.assertEqual(len(app.session_state['_hierarchy_validation']), 6)
            self.assertEqual(self.widget(app, 'selectbox', 'Période des violations affichées').value, 'Mois précédent (M-1)')
            self.assertEqual(len(self.detail_view(app)), 2)
            self.assertEqual(len(saved['previous_violations']), 2)
            self.assertNotEqual(set(saved['previous_violations']['_period_id']), set(saved['violations']['_period_id']))
            pd.testing.assert_frame_equal(saved['violations'], current_saved)

            self.widget(app, 'selectbox', 'Période des violations affichées').select('Période filtrée (M)')
            self.run_app(app)
            pd.testing.assert_frame_equal(self.detail_view(app), current_view)
            self.assertEqual(len(app.session_state['_hierarchy_validation']), 6)

            self.widget(app, 'selectbox', 'Période des violations affichées').select('Mois précédent (M-1)')
            self.widget(app, 'text_input', 'Rechercher dans les violations').input('Règle 1')
            self.run_app(app)
            self.assertEqual(self.detail_view(app)['Règle de validation'].tolist(), ['Règle 1'])
            self.assertEqual(len(app.session_state['_hierarchy_validation']), 6)


if __name__ == '__main__':
    unittest.main()
