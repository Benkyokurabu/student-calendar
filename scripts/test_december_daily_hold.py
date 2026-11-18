from datetime import date
from pathlib import Path
import tempfile,unittest
from unittest.mock import patch
from ci_extract_and_push import get_months_to_process,december_creation_ready

class DecemberDailyHoldTests(unittest.TestCase):
    def test_completed_schedule_alone_does_not_release_the_december_hold(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);(root/'2026年12月スケジュール.xlsm').write_bytes(b'fixture')
            for today in (date(2026,11,18),date(2026,11,20)):
                with patch('ci_extract_and_push.journal_today',return_value=today):
                    self.assertNotIn('2026-12',get_months_to_process(script_dir=root))
            (root/'TEMPLATE_MAIN_COMMON.xlsx').write_bytes(b'registered-fixture')
            with patch('ci_extract_and_push.journal_today',return_value=date(2026,11,19)):
                self.assertNotIn('2026-12',get_months_to_process(script_dir=root))
            with patch('ci_extract_and_push.journal_today',return_value=date(2026,11,20)):
                self.assertIn('2026-12',get_months_to_process(script_dir=root))
                self.assertIn('2026-11',get_months_to_process(script_dir=root))
            self.assertTrue(december_creation_ready('2027-01',None,date(2026,12,1)))

    def test_explicit_december_input_cannot_bypass_the_original_hold(self):
        with patch('ci_extract_and_push.journal_today',return_value=date(2026,11,18)):
            with self.assertRaises(ValueError):get_months_to_process('2026-12')

if __name__=='__main__':unittest.main()
