"""Un travail interrompu conserve les réponses reçues et évite les POST répétés."""

from datetime import date
import unittest
from unittest import mock

import pandas as pd

from quality_batch import run_quality_batch
from validation_job import (
    INTERRUPTED_MESSAGE, advance_validation_job, create_validation_job,
    recover_validation_job, resume_validation_job,
)


class UserInterruption(BaseException):
    pass


class ValidationJobTests(unittest.TestCase):
    def setUp(self):
        self.targets = pd.DataFrame([
            {'Zone_ID': f'Zone{number:07d}', 'Zone': f'ZS {number}'} for number in range(1, 9)
        ])
        self.periods = ['202609']
        self.previous = ['202608']
        self.bounds = {'start_date': date(2026, 9, 1), 'end_date': date(2026, 9, 30)}

    def advance(self, job, fetch, **kwargs):
        return advance_validation_job(
            job, self.targets, self.periods, self.previous, 'GroupAAAAAA', fetch,
            **self.bounds, **kwargs,
        )

    def test_create_validates_action_and_previous_base_without_network(self):
        job = create_validation_job(('context',), 'M', ('user', 'revision', 'nonce'))
        self.assertEqual(job['phase'], 'M')
        self.assertEqual(job['batches'], {'M': None, 'M-1': None})
        self.assertEqual(job['status'], 'running')
        self.assertFalse(job['auto_continue'])
        self.assertIsNone(job['inflight_id'])
        for action in ['unknown', None, '']:
            with self.subTest(action=action), self.assertRaises(ValueError):
                create_validation_job('context', action, 'cache')
        with self.assertRaises(ValueError):
            create_validation_job('context', 'M-1', 'cache')

    def test_one_advance_stores_before_ui_and_preserves_source_job(self):
        job = create_validation_job('context', 'M', ('user', 'nonce'))
        checkpoints = []
        events = []

        def checkpoint(state):
            checkpoints.append(state)
            events.append(('checkpoint', state['inflight_id']))

        def progress(position, total, identifier, name):
            self.assertEqual(checkpoints[-1]['inflight_id'], identifier)
            events.append(('ui', identifier))

        def fetch(*args, **kwargs):
            self.assertEqual(events[:2], [('checkpoint', 'Zone0000001'), ('ui', 'Zone0000001')])
            self.assertEqual(args, ('Zone0000001', ['202609'], 'GroupAAAAAA'))
            self.assertEqual(kwargs, {'cache_user': ('user', 'nonce'), **self.bounds})
            return pd.DataFrame([{'validationRule': {'id': 'RuleAAAAAAA'}}])

        result = self.advance(job, fetch, on_checkpoint=checkpoint, on_progress=progress)
        self.assertEqual(result['batches']['M']['processed_ids'], ['Zone0000001'])
        self.assertEqual(len(result['batches']['M']['pending_ids']), 7)
        self.assertEqual(result['status'], 'running')
        self.assertIsNone(result['inflight_id'])
        self.assertFalse(result['auto_continue'])
        self.assertIsNone(job['batches']['M'])
        self.assertTrue(any(
            state['batches']['M'] is not None and state['inflight_id'] is None for state in checkpoints
        ))

    def test_interruption_on_seventh_preserves_six_and_resumes_at_eighth_without_retry(self):
        job = create_validation_job('context', 'M', ('user', 'nonce'))
        latest = {}
        attempted = []

        def checkpoint(state):
            latest['job'] = state

        def fetch(identifier, *args, **kwargs):
            attempted.append(identifier)
            if identifier == 'Zone0000007':
                raise UserInterruption()
            return pd.DataFrame()

        for _ in range(6):
            job = self.advance(job, fetch, on_checkpoint=checkpoint)
        with self.assertRaises(UserInterruption):
            self.advance(job, fetch, on_checkpoint=checkpoint)
        saved = latest['job']
        self.assertEqual(len(saved['batches']['M']['processed_ids']), 6)
        self.assertEqual(saved['inflight_id'], 'Zone0000007')
        recovered = recover_validation_job(saved, self.targets, self.periods, self.previous)
        self.assertEqual(recovered['status'], 'interrupted')
        self.assertIsNone(recovered['inflight_id'])
        snapshot = recovered['batches']['M']
        self.assertEqual(snapshot['pending_ids'], ['Zone0000008'])
        self.assertEqual(len(snapshot['processed_ids']), 7)
        self.assertFalse(snapshot['succeeded_by_id']['Zone0000007'])
        self.assertIn(INTERRUPTED_MESSAGE, '; '.join(snapshot['errors_by_id']['Zone0000007']))
        self.assertTrue(pd.isna(snapshot['requests'][-1]['Lignes reçues']))
        self.assertFalse(saved['batches']['M']['finished'])
        resumed = resume_validation_job(recovered)
        finished = self.advance(resumed, fetch)
        self.assertEqual(finished['status'], 'complete')
        self.assertEqual(attempted, self.targets['Zone_ID'].tolist())
        self.assertEqual(finished['batches']['M']['pending_ids'], [])
        self.assertFalse(finished['batches']['M']['complete'])
        self.assertEqual(recovered['status'], 'interrupted')

    def test_both_action_advances_to_previous_after_current_attempts_even_with_failure(self):
        targets = self.targets.iloc[:1]
        job = create_validation_job('context', 'both', 'nonce')
        failed = pd.DataFrame()
        failed.attrs = {'failed_periods': ['202609'], 'errors': {'202609': ['HTTP 504']}}
        fetch = mock.Mock(side_effect=[failed, pd.DataFrame()])
        after_m = advance_validation_job(
            job, targets, self.periods, self.previous, None, fetch, **self.bounds
        )
        self.assertEqual(after_m['phase'], 'M-1')
        self.assertEqual(after_m['status'], 'running')
        self.assertFalse(after_m['batches']['M']['complete'])
        final = advance_validation_job(
            after_m, targets, self.periods, self.previous, None, fetch, **self.bounds
        )
        self.assertEqual(final['status'], 'complete')
        self.assertEqual(fetch.call_count, 2)
        self.assertEqual(fetch.call_args_list[1], mock.call(
            'Zone0000001', ['202608'], None, cache_user='nonce'
        ))
        self.assertTrue(final['batches']['M-1']['complete'])

    def test_previous_action_retains_current_snapshot_and_never_calls_current_again(self):
        current = run_quality_batch(self.targets, self.periods, 'GroupAAAAAA', mock.Mock(return_value=pd.DataFrame()))
        job = create_validation_job('context', 'M-1', 'next-nonce', current_batch=current)
        fetch = mock.Mock(return_value=pd.DataFrame())
        result = self.advance(job, fetch)
        fetch.assert_called_once_with('Zone0000001', ['202608'], 'GroupAAAAAA', cache_user='next-nonce')
        self.assertEqual(result['batches']['M']['processed_ids'], self.targets['Zone_ID'].tolist())
        self.assertEqual(result['batches']['M-1']['processed_ids'], ['Zone0000001'])
        self.assertEqual(current['processed_ids'], self.targets['Zone_ID'].tolist())

    def test_recover_after_first_inflight_builds_diagnostics_without_fetching(self):
        job = create_validation_job('context', 'M', 'nonce')
        latest = {}
        with self.assertRaises(UserInterruption):
            self.advance(
                job, mock.Mock(side_effect=UserInterruption()),
                on_checkpoint=lambda state: latest.update(job=state),
            )
        recovered = recover_validation_job(latest['job'], self.targets, self.periods, self.previous)
        batch = recovered['batches']['M']
        self.assertEqual(batch['processed_ids'], ['Zone0000001'])
        self.assertEqual(batch['requests'][0]['Début'], '2026-09-01')
        self.assertEqual(batch['requests'][0]['Fin'], '2026-09-30')
        self.assertEqual(batch['context']['group_id'], 'GroupAAAAAA')
        self.assertTrue(batch['raw'].empty)

    def test_recover_previous_inflight_uses_reference_month_and_keeps_current(self):
        current = run_quality_batch(self.targets.iloc[:1], self.periods, None, mock.Mock(return_value=pd.DataFrame()))
        job = create_validation_job('context', 'M-1', 'nonce', current)
        latest = {}
        with self.assertRaises(UserInterruption):
            self.advance(job, mock.Mock(side_effect=UserInterruption()), on_checkpoint=lambda state: latest.update(job=state))
        recovered = recover_validation_job(latest['job'], self.targets, self.periods, self.previous)
        request = recovered['batches']['M-1']['requests'][0]
        self.assertEqual((request['Début'], request['Fin']), ('2026-08-01', '2026-08-31'))
        self.assertEqual(recovered['batches']['M']['processed_ids'], ['Zone0000001'])

    def test_auto_continuation_is_not_mistaken_for_an_interruption(self):
        job = create_validation_job('context', 'M', 'nonce')
        job['auto_continue'] = True
        recovered = recover_validation_job(job, self.targets, self.periods, self.previous)
        self.assertEqual(recovered['status'], 'running')
        self.assertTrue(recovered['auto_continue'])
        job['inflight_id'] = 'Zone0000001'
        recovered = recover_validation_job(job, self.targets, self.periods, self.previous)
        self.assertEqual(recovered['status'], 'interrupted')

    def test_complete_or_interrupted_jobs_do_not_fetch_without_resume(self):
        job = create_validation_job('context', 'M', 'nonce')
        fetch = mock.Mock(side_effect=AssertionError('Aucun appel attendu'))
        for status in ['complete', 'interrupted']:
            job['status'] = status
            result = self.advance(job, fetch)
            self.assertEqual(result['status'], status)
        fetch.assert_not_called()

    def test_checkpoints_and_created_current_are_deeply_independent(self):
        frame = pd.DataFrame([{'organisationUnit': {'id': 'CentreAAAAA'}}])
        current = run_quality_batch(self.targets.iloc[:1], self.periods, None, mock.Mock(return_value=frame))
        job = create_validation_job({'ids': ['a']}, 'M-1', ('user', 'nonce'), current)
        job['batches']['M']['frames_by_id']['Zone0000001'].iloc[0]['organisationUnit']['id'] = 'Changed'
        self.assertEqual(current['frames_by_id']['Zone0000001'].iloc[0]['organisationUnit']['id'], 'CentreAAAAA')
        states = []

        def mutate(state):
            states.append(state)
            state['context']['ids'].append('mutated')
            state['cache_user'] = 'wrong-user'

        result = self.advance(job, mock.Mock(return_value=pd.DataFrame()), on_checkpoint=mutate)
        self.assertEqual(result['context'], {'ids': ['a']})
        self.assertEqual(result['cache_user'], ('user', 'nonce'))


if __name__ == '__main__':
    unittest.main()
