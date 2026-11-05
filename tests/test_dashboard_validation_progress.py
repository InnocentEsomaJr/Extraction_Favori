"""Les réponses de validation survivent à une interruption au milieu des ZS."""

import unittest
from unittest import mock

import pandas as pd
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
import test_dashboard_categorisation as dashboard_test


CATALOG = '''
def get_organisation_catalog(cache_user=None):
    from geography import build_geography_catalog
    country = {"id": "CountryAAAA", "displayName": "Pays test", "level": 1}
    province = {"id": "ProvinceAAA", "displayName": "Province huit zones", "level": 2}
    units = [country, {**province, "ancestors": [country]}]
    for number in range(1, 9):
        units.append({
            "id": f"Zone{number:07d}", "displayName": f"ZS {number:02d}", "level": 3,
            "ancestors": [country, province],
        })
    return build_geography_catalog(units)
'''


DATA = '''
def get_data(favori_id, period=None, cache_user=None, ou_dimension=None, unit_ids=None):
    selected = [f"Zone{number:07d}" for number in range(1, 9)]
    if ou_dimension and not ou_dimension.startswith("LEVEL-"):
        selected = [uid for uid in selected if uid == ou_dimension]
    rows = []
    for uid in selected:
        row = {"Organisation unit ID": uid, "Organisation unit": f"ZS {int(uid[4:]):02d}"}
        for number in range(9):
            prefix = f"Dataset {number} "
            row[prefix + "Actual reports"] = 15
            row[prefix + "Expected reports"] = 15
            row[prefix + "Actual reports on time"] = 15
            row[prefix + "Reporting rate"] = 100
            row[prefix + "Reporting rate on time"] = 100
        rows.append(row)
    return pd.DataFrame(rows)
'''


REPORTING = '''
def get_quality_reporting_data(favori_id, period=None, cache_user=None, ou_dimension=None, unit_ids=None):
    return get_data(favori_id, period=period, cache_user=cache_user, ou_dimension=ou_dimension, unit_ids=unit_ids)
'''


METADATA = '''
def get_validation_detail_metadata(rule_ids, org_ids, cache_user=None):
    return {}, {}, []
'''


CLIENT = '''
def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            from streamlit.runtime.scriptrunner_utils.exceptions import StopException
            if endpoint != "dataAnalysis/validationRules":
                raise AssertionError(endpoint)
            previous = data["startDate"] < st.session_state["validation_start_date"].isoformat()
            phase = "M-1" if previous else "M"
            calls = list(st.session_state.get("_progress_validation_calls", []))
            prior_calls = [call for call in calls if call["phase"] == phase]
            job = st.session_state.get("validation_job", {})
            snapshot = job.get("batches", {}).get(phase) or {}
            saved_ids = list(snapshot.get("processed_ids", []))
            if prior_calls and prior_calls[-1]["ou"] not in saved_ids:
                failures = list(st.session_state.get("_progress_checkpoint_failures", []))
                failures.append({"phase": phase, "previous": prior_calls[-1]["ou"], "saved": saved_ids})
                st.session_state["_progress_checkpoint_failures"] = failures
            calls.append({**dict(data), "phase": phase})
            st.session_state["_progress_validation_calls"] = calls
            number = int(data["ou"][4:])
            if not previous and number == st.session_state.get("_progress_interrupt_at"):
                st.session_state["_progress_interrupt_at"] = None
                raise StopException()
            if number > 6:
                return []
            rows = []
            for rule in range(2 if previous else 5):
                centre = f"Centre{number:05d}"
                rows.append({
                    "organisationUnitId": centre,
                    "organisationUnitDisplayName": f"Centre ZS {number:02d}",
                    "organisationUnitPath": f"/CountryAAAA/ProvinceAAA/{data['ou']}/{centre}",
                    "validationRuleId": f"Rule{rule:07d}",
                    "validationRuleDescription": f"Règle {rule}",
                    "periodId": data["startDate"][:7].replace("-", ""),
                    "leftSideValue": 20 + rule, "rightSideValue": 10,
                    "operator": "<=", "importance": "MEDIUM",
                })
            return rows
    return Client()
'''


class ValidationProgressAppTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        app.run(timeout=90)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self):
        app = AppTest.from_string(dashboard_test.app_fixture_source({
            'get_organisation_catalog': CATALOG, 'get_data': DATA,
            'get_quality_reporting_data': REPORTING, 'get_dhis2_client': CLIENT,
            'get_validation_detail_metadata': METADATA,
        }), default_timeout=90)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-progress-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        app.session_state['filter_level_main'] = 'Province'
        app.session_state['filter_province_main'] = 'ProvinceAAA'
        return app

    def calls(self, app, phase=None):
        calls = list(app.session_state.filtered_state.get('_progress_validation_calls', []))
        return calls if phase is None else [call for call in calls if call['phase'] == phase]

    def summary(self, app):
        return app.session_state['categorisation_result']['data'].set_index('Zone_ID').copy()

    def assert_success_notice(self, app):
        messages = [str(item.value).casefold() for item in app.success]
        self.assertTrue(any('termin' in message or 'trait' in message for message in messages), messages)

    def test_current_analysis_saves_every_zone_before_the_next_request_and_notifies_completion(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            job = app.session_state['validation_job']
            self.assertEqual(job['status'], 'complete')
            self.assertTrue(job['batches']['M']['finished'])
            self.assertEqual(len(job['batches']['M']['processed_ids']), 8)
            self.assertEqual(job['batches']['M']['pending_ids'], [])
            self.assertEqual(len(self.calls(app, 'M')), 8)
            self.assertEqual(self.calls(app, 'M-1'), [])
            self.assertEqual(app.session_state.filtered_state.get('_progress_checkpoint_failures', []), [])
            frame = self.summary(app)
            self.assertEqual(frame['Violations_Actuelles'].tolist(), [5] * 6 + [0, 0])
            self.assertTrue(frame['Violations_Mois_Precedent'].isna().all())
            self.assert_success_notice(app)

    def test_seventh_zone_interruption_keeps_six_results_and_resume_only_requests_eighth(self):
        app = self.new_app()
        app.session_state['_progress_interrupt_at'] = 7
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            job = app.session_state['validation_job']
            self.assertEqual(job['batches']['M']['processed_ids'], [f'Zone{number:07d}' for number in range(1, 7)])
            self.assertEqual(job['batches']['M']['pending_ids'], ['Zone0000007', 'Zone0000008'])
            self.assertEqual(len(self.calls(app, 'M')), 7)
            saved = self.summary(app)
            self.assertTrue(saved.loc[['Zone0000007', 'Zone0000008'], 'Violations_Actuelles'].isna().all())
            self.assertEqual(saved.loc[[f'Zone{number:07d}' for number in range(1, 7)], 'Violations_Actuelles'].tolist(), [5] * 6)

            self.run_app(app)
            recovered = app.session_state['validation_job']
            self.assertEqual(recovered['status'], 'interrupted')
            self.assertEqual(len(self.calls(app, 'M')), 7)
            notices = list(app.warning) + list(app.info) + list(app.error)
            self.assertTrue(any('interromp' in str(item.value).casefold() for item in notices))
            self.widget(app, 'button', 'Reprendre les zones restantes').click()
            self.run_app(app)
            final = app.session_state['validation_job']
            self.assertEqual(final['status'], 'complete')
            self.assertEqual([call['ou'] for call in self.calls(app, 'M')], [f'Zone{number:07d}' for number in range(1, 9)])
            self.assertEqual(final['batches']['M']['pending_ids'], [])
            self.assertEqual(app.session_state.filtered_state.get('_progress_checkpoint_failures', []), [])
            result = self.summary(app)
            self.assertTrue(pd.isna(result.loc['Zone0000007', 'Violations_Actuelles']))
            self.assertEqual(result.loc['Zone0000008', 'Violations_Actuelles'], 0)
            self.assertEqual(result.loc[[f'Zone{number:07d}' for number in range(1, 7)], 'Violations_Actuelles'].tolist(), [5] * 6)

    def test_calculate_all_continues_from_m_to_reference_and_preserves_m_checkpoints(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.widget(app, 'button', 'Calculer tous les scores (M et M-1)').click()
            self.run_app(app)
            job = app.session_state['validation_job']
            self.assertEqual(job['status'], 'complete')
            self.assertTrue(job['batches']['M']['finished'])
            self.assertTrue(job['batches']['M-1']['finished'])
            self.assertEqual([call['phase'] for call in self.calls(app)], ['M'] * 8 + ['M-1'] * 8)
            self.assertEqual(app.session_state.filtered_state.get('_progress_checkpoint_failures', []), [])
            saved = app.session_state['categorisation_result']
            self.assertTrue(saved['corrections_checked'])
            self.assertEqual(saved['data']['Violations_Actuelles'].tolist(), [5] * 6 + [0, 0])
            self.assertEqual(saved['data']['Violations_Mois_Precedent'].tolist(), [2] * 6 + [0, 0])
            self.assertEqual(saved['data']['Violations_Corrigees'].tolist(), [0] * 8)
            self.assertEqual(app.session_state['_test_export_frame'][TOTAL_COL].tolist(), [75] * 6 + [95, 95])
            self.assert_success_notice(app)

    def test_reference_action_keeps_current_rows_and_does_not_request_m_again(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            current = app.session_state['categorisation_result']['violations'].copy(deep=True)
            self.widget(app, 'button', 'Contrôler les corrections du mois précédent').click()
            self.run_app(app)
            self.assertEqual(len(self.calls(app, 'M')), 8)
            self.assertEqual(len(self.calls(app, 'M-1')), 8)
            pd.testing.assert_frame_equal(app.session_state['categorisation_result']['violations'], current)
            self.assertEqual(app.session_state['validation_job']['status'], 'complete')
            self.assertEqual(len(app.session_state['categorisation_result']['previous_violations']), 12)
            self.assertEqual(self.widget(app, 'selectbox', 'Période des violations affichées').value, 'Mois précédent (M-1)')


if __name__ == '__main__':
    unittest.main()
