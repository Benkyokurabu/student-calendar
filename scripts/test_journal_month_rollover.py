"""A new month must not freeze the last lessons displayed in preparation sheets."""
import tempfile
import contextlib
import io
import json
from types import SimpleNamespace
import unittest
from datetime import date, datetime, timezone
from pathlib import Path
from unittest.mock import patch

import ci_extract_and_push as pipeline


class JournalMonthRolloverTests(unittest.TestCase):
    def months(self, today, **kwargs):
        with patch.object(pipeline, 'journal_today', return_value=today):
            return pipeline.get_months_to_process(**kwargs)

    def test_october_corrections_are_retrieved_all_through_november(self):
        for day in [1, 4, 8, 15, 25, 30]:
            with self.subTest(day=day):
                self.assertEqual(self.months(date(2026, 11, day)), ['2026-10', '2026-11'])

    def test_year_and_leap_month_boundaries(self):
        for today, expected in [
            (date(2027, 1, 1), ['2026-12', '2027-01']),
            (date(2028, 3, 1), ['2028-02', '2028-03']),
            (date(2026, 12, 31), ['2026-11', '2026-12']),
        ]:
            with self.subTest(today=today):
                self.assertEqual(self.months(today), expected)

    def test_japan_month_is_used_before_utc_midnight(self):
        with patch.object(pipeline, 'datetime') as clock:
            clock.now.return_value = datetime(2026, 10, 31, 19, 17, tzinfo=timezone.utc).astimezone(pipeline.ZoneInfo('Asia/Tokyo'))
            self.assertEqual(pipeline.get_months_to_process(), ['2026-10', '2026-11'])
            clock.now.assert_called_once_with(pipeline.ZoneInfo('Asia/Tokyo'))

    def test_available_next_month_does_not_replace_previous_month(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / '2026年12月スケジュール.xlsm').touch()
            self.assertEqual(self.months(date(2026, 11, 4), script_dir=root), ['2026-10', '2026-11', '2026-12'])

    def test_explicit_manual_month_remains_limited(self):
        self.assertEqual(pipeline.get_months_to_process('2026-09'), ['2026-09'])

    def test_previous_month_refresh_does_not_add_workbook_maintenance(self):
        for today in [date(2026, 11, 1), date(2026, 11, 4), date(2026, 11, 25), date(2027, 1, 1)]:
            with self.subTest(today=today), patch.object(pipeline, 'journal_today', return_value=today):
                months = pipeline.get_months_to_process()
                self.assertEqual(pipeline.get_months_to_maintain(months), [months[-1]])
                self.assertEqual(len(months), 2)

    def test_existing_month_end_maintenance_window_is_preserved(self):
        for day in [26, 30]:
            with self.subTest(day=day), patch.object(pipeline, 'journal_today', return_value=date(2026, 11, day)):
                months = pipeline.get_months_to_process()
                self.assertEqual(pipeline.get_months_to_maintain(months), months)

    def test_manual_previous_month_maintenance_and_next_month_creation_remain(self):
        with patch.object(pipeline, 'journal_today', return_value=date(2026, 11, 4)):
            self.assertEqual(pipeline.get_months_to_maintain(['2026-10'], '2026-10'), ['2026-10'])
            self.assertEqual(pipeline.get_months_to_maintain(['2026-10', '2026-11', '2026-12']), ['2026-11', '2026-12'])

    def test_pipeline_refreshes_previous_content_without_rewriting_previous_schedule(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            scripts = root / 'scripts'
            scripts.mkdir()
            (scripts / '_cloud_journal').mkdir()
            (root / 'update_journal.py').touch()
            (scripts / 'sync_journal_across_campus.py').touch()
            for month in ['2026-10', '2026-11', '2026-12']:
                (scripts / f'2026年{int(month[-2:])}月スケジュール.xlsm').touch()
                (root / f'journal_{month}.json').write_text(json.dumps({'month': month, 'entries': {}}), encoding='utf-8')
                (root / f'schedule_{month}.json').write_text('[]', encoding='utf-8')
            with patch.object(pipeline, '__file__', str(scripts / 'ci_extract_and_push.py')), \
                 patch.object(pipeline.sys, 'argv', ['ci_extract_and_push.py']), \
                 patch.object(pipeline, 'journal_today', return_value=date(2026, 11, 4)), \
                 patch.object(pipeline, 'repair_hidden_template_validations'), \
                 patch.object(pipeline, 'create_month_sheets', return_value=0) as create, \
                 patch.object(pipeline.subprocess, 'run', return_value=SimpleNamespace(returncode=0, stdout='', stderr='')) as update, \
                 patch.object(pipeline, 'run') as publish, contextlib.redirect_stdout(io.StringIO()):
                pipeline.main()
            self.assertEqual(create.call_args.args[1], ['2026-11', '2026-12'])
            update_months = [Path(call.args[0][3]).name for call in update.call_args_list]
            self.assertEqual(update_months, ['2026年11月スケジュール.xlsm', '2026年12月スケジュール.xlsm'])
            for script in ['sync_journal_across_campus.py', 'extract_journal_to_json.py']:
                calls = [call.args[0] for call in publish.call_args_list if Path(call.args[0][1]).name == script]
                self.assertEqual([args[args.index('--month') + 1] for args in calls], ['2026-10', '2026-11', '2026-12'])


if __name__ == '__main__':
    unittest.main()
