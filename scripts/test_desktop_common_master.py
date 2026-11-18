from pathlib import Path
import contextlib,io,unittest
from unittest.mock import patch
import openpyxl
import export_by_grade_subject as generator
import extract_and_push as desktop
import test_journal_common_master as common_tests


class DesktopMasterTests(unittest.TestCase):
    fixture=common_tests.CommonMasterTests.fixture
    annual=common_tests.CommonMasterTests.annual
    events=common_tests.CommonMasterTests.events

    def test_desktop_new_month_uses_common_master_and_previous_annual_counter(self):
        for supplement in (False,True):
            with self.subTest(supplement=supplement):
                root,common,_=self.fixture()
                scripts=root/'scripts';scripts.mkdir()
                (scripts/common.name).write_bytes(common.read_bytes())
                for name in ('TEMPLATE_MAIN.xlsx','TEMPLATE_X.xlsx'):
                    (scripts/name).write_bytes((common_tests.SCRIPTS/name).read_bytes())
                journal=root/'journal';journal.mkdir()
                name='南教室数学補講_2026.xlsx' if supplement else '南教室中２数学_2026.xlsx'
                annual=journal/name
                book,_,_=self.annual();book.save(annual)
                schedule=root/'schedule.xlsx';source=openpyxl.Workbook();source.save(schedule);source.close()
                events=[event._replace(text='数学補講①') for event in self.events()] if supplement else self.events()
                with patch.object(generator,'pick_schedule_in_same_folder',return_value=(2026,12,schedule)), \
                     patch.object(generator,'choose_target_sheets',return_value=[('南教室','教務',object())]), \
                     patch.object(generator,'collect_events',return_value=events),contextlib.redirect_stdout(io.StringIO()):
                    self.assertEqual(desktop.create_month_sheets(scripts,['2026-12'],journal),1)
                saved=openpyxl.load_workbook(annual)
                try:
                    sheet=saved['2026-12']
                    self.assertEqual(sheet['D6'].value,'共通レイアウト')
                    self.assertEqual(sheet['F2'].value,42)
                    self.assertEqual(sheet['HA3'].value,'数学補講' if supplement else '数学')
                    self.assertEqual(saved['2026-11']['D7'].value,'保存済みの授業内容')
                finally:saved.close()


if __name__=='__main__':unittest.main()
