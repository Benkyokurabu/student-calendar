"""A new month must not freeze the last lessons displayed in preparation sheets."""
import tempfile
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


if __name__ == '__main__':
    unittest.main()
