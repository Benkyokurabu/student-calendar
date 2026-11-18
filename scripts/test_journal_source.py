import os
import unittest
from unittest.mock import patch

from journal_source import source_remote, source_file


class JournalSourceTests(unittest.TestCase):
    def test_default_keeps_current_desktop_operation(self):
        with patch.dict(os.environ, {}, clear=True):
            self.assertEqual(source_remote(), "onedrive:デスクトップ/【完成版】授業日誌システム")

    def test_all_future_inputs_select_the_same_bentan_folder(self):
        with patch.dict(os.environ, {"JOURNAL_INPUT_LOCATION": "bentan"}):
            for name in ("2027年1月スケジュール.xlsm", "TEMPLATE_MAIN_COMMON.xlsx"):
                self.assertEqual(source_file(name), "onedrive:デスクトップ/勉たん/業務データ/授業日誌/" + name)

    def test_invalid_settings_and_paths_stop_instead_of_using_desktop(self):
        for profile in ("", "other", "../desktop", "bentan "):
            with patch.dict(os.environ, {"JOURNAL_INPUT_LOCATION": profile}):
                with self.assertRaises(ValueError):
                    source_remote()
        for name in ("", "..", "../schedule.xlsm", "sub/schedule.xlsm", "sub\\schedule.xlsm"):
            with self.assertRaises(ValueError):
                source_file(name, "bentan")


if __name__ == "__main__":
    unittest.main()
