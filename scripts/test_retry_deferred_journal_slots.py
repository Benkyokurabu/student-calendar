import datetime as dt
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch, Mock
import urllib.error

import retry_deferred_journal_slots as retry


class FakeGraph:
    def __init__(self, changed=False, locked=False, bad_backup=False, bad_stage=False):
        self.calls = []
        self.meta_reads = 0
        self.changed = changed
        self.locked = locked
        self.bad_backup = bad_backup
        self.bad_stage = bad_stage
        self.production = b'original'
        self.identity = retry.TARGETS[0][1]
        self.path = retry.item_path(self.identity)

    def request(self, path, method='GET', data=None, headers=None):
        self.calls.append((path, method, data, headers))
        if path == self.path and method == 'GET':
            self.meta_reads += 1
            return {'id': self.identity, 'name': retry.TARGETS[0][0],
                    'eTag': 'edited' if self.changed and self.meta_reads >= 3 else 'original-etag',
                    'parentReference': {'id': 'parent'}}
        if path == self.path + '/content':
            if self.locked:
                raise urllib.error.HTTPError(path, 423, 'Locked', {}, None)
            self.production = data
            return {'id': self.identity}
        if path.endswith('/children'):
            return {'id': 'backup-folder'}
        if 'before.xlsx' in path:
            return {'id': 'backup'}
        if 'validated.xlsx' in path:
            return {'id': 'staged'}
        if path == retry.item_path('backup'):
            return {'id': 'backup'}
        if path == retry.item_path('staged'):
            return {'id': 'staged'}
        if '/workbook/worksheets?' in path:
            return {'value': [{'name': '2026-09'}]}
        if '/range(' in path:
            cells = [[''] * 9 for _ in range(12)]
            cells[5][0] = 'クラス'
            if self.bad_stage and path.startswith(retry.item_path('staged') + '/'):
                cells[0][0] = '#REF!'
            return {'text': cells}
        raise AssertionError('Unexpected API call')

    def download(self, meta):
        if meta['id'] == 'backup':
            return b'bad' if self.bad_backup else b'original'
        if meta['id'] == 'staged':
            return b'candidate'
        return self.production


def builder(source, output):
    assert source.read_bytes() == b'original'
    output.write_bytes(b'candidate')
    return {'original_cells_and_caches_preserved': True, 'unrelated_parts_preserved': True,
            'sheets': [{'name': '2026-09'}]}


class RetryTests(unittest.TestCase):
    def run_repair(self, graph, build=builder):
        with tempfile.TemporaryDirectory() as directory, patch.object(
            retry, 'already_expanded', side_effect=lambda b: b == b'candidate'
        ):
            return retry.repair(graph, *retry.TARGETS[0], Path(directory), builder=build)

    def test_exact_time_window(self):
        self.assertFalse(retry.due(retry.START - dt.timedelta(seconds=1)))
        self.assertTrue(retry.due(retry.START))
        self.assertFalse(retry.due(retry.END))

    def test_before_start_and_after_expiry_never_connect(self):
        factory = Mock(side_effect=AssertionError('Cloud accessed early'))
        for now in (retry.START - dt.timedelta(seconds=1), retry.END):
            self.assertEqual(retry.main(now, factory), [])
        factory.assert_not_called()

    def test_success_requires_backup_stage_conditional_put_and_reopen(self):
        graph = FakeGraph()
        self.assertEqual(self.run_repair(graph), 'published_and_verified')
        writes = [(i, c) for i, c in enumerate(graph.calls) if c[0] == graph.path + '/content']
        self.assertEqual(len(writes), 1)
        index, call = writes[0]
        self.assertEqual(call[3]['If-Match'], 'original-etag')
        self.assertTrue(any('/range(' in c[0] for c in graph.calls[:index]))
        self.assertTrue(any('/range(' in c[0] for c in graph.calls[index + 1:]))
        self.assertTrue(any('before.xlsx' in c[0] for c in graph.calls[:index]))

    def test_edit_during_validation_never_writes_production(self):
        graph = FakeGraph(changed=True)
        self.assertEqual(self.run_repair(graph), 'deferred_concurrent_edit')
        self.assertEqual(graph.production, b'original')
        self.assertFalse(any(c[0] == graph.path + '/content' for c in graph.calls))

    def test_lock_never_forces_overwrite(self):
        graph = FakeGraph(locked=True)
        with self.assertRaises(urllib.error.HTTPError):
            self.run_repair(graph)
        self.assertEqual(graph.production, b'original')

    def test_backup_mismatch_stops_before_production(self):
        graph = FakeGraph(bad_backup=True)
        with self.assertRaises(ValueError):
            self.run_repair(graph)
        self.assertEqual(graph.production, b'original')

    def test_excel_stage_error_stops_before_production(self):
        graph = FakeGraph(bad_stage=True)
        with self.assertRaises(ValueError):
            self.run_repair(graph)
        self.assertEqual(graph.production, b'original')

    def test_preservation_failure_creates_no_cloud_files(self):
        graph = FakeGraph()
        with self.assertRaises(ValueError):
            self.run_repair(graph, lambda *args: {'original_cells_and_caches_preserved': False})
        self.assertTrue(all(c[1] == 'GET' for c in graph.calls))

    def test_only_three_authorized_targets_even_when_one_locked(self):
        statuses = [urllib.error.HTTPError('', 423, 'Locked', {}, None), 'published_and_verified', 'published_and_verified']
        with patch.object(retry, 'repair', side_effect=statuses) as repair, patch.dict('os.environ', {}, clear=True):
            results = retry.main(retry.START, lambda: object())
        self.assertEqual(len(results), 3)
        self.assertEqual([(c.args[1], c.args[2]) for c in repair.call_args_list], list(retry.TARGETS))
        self.assertEqual(results[0]['status'], 'deferred_lock_or_concurrent_edit')

    def test_already_expanded_only_verifies_without_writing(self):
        graph = FakeGraph()
        workbook = Mock()
        with tempfile.TemporaryDirectory() as directory, patch.object(
            retry, 'already_expanded', return_value=True
        ), patch.object(retry.openpyxl, 'load_workbook', return_value=workbook), patch.object(
            retry, 'target_sheets', return_value=[Mock(title='2026-09')]
        ):
            status = retry.repair(graph, *retry.TARGETS[0], Path(directory))
        self.assertEqual(status, 'already_expanded_and_verified')
        self.assertTrue(all(c[1] == 'GET' for c in graph.calls))
        workbook.close.assert_called_once()

    def test_unexpected_failure_stops_remaining_targets_without_leaking_exception(self):
        with patch.object(retry, 'repair', side_effect=RuntimeError('private signed URL')) as repair, \
             patch.dict('os.environ', {}, clear=True), patch('builtins.print') as output:
            results = retry.main(retry.START, lambda: object())
        self.assertEqual(repair.call_count, 1)
        self.assertEqual(results[0]['status'], 'stopped_safely')
        self.assertNotIn('private signed URL', str(output.call_args_list))


if __name__ == '__main__':
    unittest.main()
