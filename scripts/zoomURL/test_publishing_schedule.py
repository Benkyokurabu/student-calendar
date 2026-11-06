"""Independent recording publication must use the authoritative current month."""
from datetime import datetime
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

import publish_zoom_recording_url_json as publisher
import recording_publication_filter as policy


class PublishingScheduleTests(unittest.TestCase):
    def test_reads_remote_month_even_when_local_schedule_is_stale(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            def git(*args):
                return subprocess.run(['git', *args], cwd=repo, check=True, capture_output=True)
            git('init')
            git('config', 'user.email', 'test@example.test')
            git('config', 'user.name', 'Isolated test')
            expected = [{'date': '2026-11-06', 'teacher': 'source'}]
            file = repo / 'schedule_2026-11.json'
            file.write_text(json.dumps(expected), 'utf-8')
            git('add', file.name)
            git('commit', '-m', 'authoritative month')
            git('update-ref', 'refs/remotes/origin/main', 'HEAD')
            file.write_text('[{"date":"2026-10-01"}]', 'utf-8')
            self.assertEqual(publisher.load_publishing_schedule(repo, '2026-11'), expected)
            with self.assertRaisesRegex(RuntimeError, '正本予定がありません'):
                publisher.load_publishing_schedule(repo, '2026-12')
            self.assertEqual(json.loads(file.read_text('utf-8'))[0]['date'], '2026-10-01')

    def test_rejects_invalid_or_wrong_month_source(self):
        for value in [[], {}, [{'date': '2026-10-31'}], [None], [{'date': '2026-11-06'}, {'date': '2026-12-01'}]]:
            with self.subTest(value=value), patch.object(publisher, 'run', return_value=subprocess.CompletedProcess([], 0, json.dumps(value), '')):
                with self.assertRaisesRegex(RuntimeError, '月・形式が不正'):
                    publisher.load_publishing_schedule(Path('.'), '2026-11')

    def test_default_month_rollover_and_dry_run_preserve_files_and_git(self):
        for current in [datetime(2026, 11, 30, 23, 59), datetime(2026, 12, 1, 0, 1), datetime(2027, 1, 1)]:
            class Clock(datetime):
                @classmethod
                def now(cls, tz=None):
                    return current.replace(tzinfo=tz)
            month = current.strftime('%Y-%m')
            schedule = [{'date': month + '-01'}]
            payload = {'month': month, 'matched': 0, 'missing': 1, 'entries': {}}
            with tempfile.TemporaryDirectory() as tmp:
                repo = Path(tmp)
                sentinel = repo / 'schedule_latest.json'
                sentinel.write_text('[{"date":"2026-10-01"}]', 'utf-8')
                with patch.object(sys, 'argv', ['publisher', '--dry-run', '--repo', str(repo)]), \
                     patch.object(publisher, 'datetime', Clock), patch.object(publisher, 'PublishLock'), \
                     patch.object(publisher, 'run', return_value=subprocess.CompletedProcess([], 0, '', '')) as git, \
                     patch.object(publisher, 'load_publishing_schedule', return_value=schedule) as source, \
                     patch.object(publisher.url_list, 'make_recording_json', return_value=payload) as generate, \
                     patch.object(publisher.url_list, 'SYSTEM_DIR', repo), \
                     patch.object(policy, 'load_rules', return_value=[]), \
                     patch.object(publisher, 'push_pending_commits', side_effect=AssertionError('dry-run cannot push')), \
                     patch.object(publisher, 'pull_rebase_with_zoom_resolution', side_effect=AssertionError('dry-run cannot rebase')):
                    self.assertEqual(publisher.main(), 0)
                    source.assert_called_once_with(repo, month)
                    generate.assert_called_once_with(month, schedule=schedule)
                    self.assertEqual(git.call_args.args[0], ['git', 'fetch', 'origin', 'main'])
                self.assertEqual([p.name for p in repo.iterdir()], ['schedule_latest.json'])

    def test_missing_current_source_stops_before_zoom_and_publication(self):
        with patch.object(sys, 'argv', ['publisher', '--repo', '.']), patch.object(publisher, 'PublishLock'), \
             patch.object(publisher, 'run', return_value=subprocess.CompletedProcess([], 1, '', 'missing')), \
             patch.object(publisher.url_list, 'make_recording_json', side_effect=AssertionError('must not query Zoom')):
            with self.assertRaisesRegex(RuntimeError, '正本予定がありません'):
                publisher.main()


if __name__ == '__main__':
    unittest.main()
