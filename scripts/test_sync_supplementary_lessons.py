import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch
import openpyxl
from openpyxl.styles import PatternFill
from sync_supplementary_lessons import extract_supplements, sync_month, allowed_months
from datetime import datetime, timezone

class SupplementTests(unittest.TestCase):
    def source(self, folder):
        book = openpyxl.Workbook()
        book.remove(book.active)
        for name in ['本校教務部用', '南教室教務部用']:
            sheet = book.create_sheet(name)
            sheet.merge_cells('F9:K9')
            sheet.merge_cells('F10:K10')
            sheet['F9'] = '小学部4:55～6:15'
            sheet['F10'] = '中学部4:45～6:15'
            sheet['A12'] = 1
            sheet['B12'] = '火'
            sheet['Z12'] = 1
            sheet['AA12'] = '火'
            sheet['F12'] = '数学　補講①'
            sheet['H56'] = 'テスト甲'
            sheet['F12'].fill = sheet['H56'].fill = PatternFill('solid', fgColor='FF00FF00')
        path = Path(folder) / '2026年9月スケジュール.xlsm'
        book.save(path)
        return path

    def test_existing_grade_aware_reader_and_legend_are_reused_for_both_campuses(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.source(folder)
            before = path.read_bytes()
            events = extract_supplements(path, '2026-09')
            self.assertEqual(len(events), 2)
            self.assertEqual({e['campus'] for e in events}, {'本校', '南教室'})
            self.assertTrue(all(e['time'] == '16:45〜' and e['teacher'] == 'テスト甲' for e in events))
            self.assertEqual(path.read_bytes(), before)

    def test_uncertain_post_is_not_retried(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.source(folder)
            calls = []
            def send(body):
                calls.append(body)
                raise TimeoutError()
            with self.assertRaises(TimeoutError):
                sync_month(path, '2026-09', dry_run=False, parent_run_id=None, download=path.read_bytes, send=send)
            self.assertEqual(len(calls), 1)

    def test_source_change_prevents_request(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.source(folder)
            with self.assertRaises(RuntimeError):
                sync_month(path, '2026-09', dry_run=False, parent_run_id=None, download=lambda: b'changed', send=lambda _: self.fail('unexpected request'))

    def test_batch_continuation_and_december_rollover(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.source(folder)
            results = iter([dict(created=1, pending=1), dict(created=1, pending=0)])
            result = sync_month(path, '2026-09', dry_run=False, parent_run_id='42', download=path.read_bytes, send=lambda _: next(results))
            self.assertEqual(result['created'], 2)
        self.assertEqual(allowed_months(datetime(2026, 12, 31, tzinfo=timezone.utc)), ['2026-12', '2027-01'])

if __name__ == '__main__':
    unittest.main()
