from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch,Mock
from io import BytesIO
from zipfile import ZipFile
import unittest
import bentan_monthly_create as monthly

def workbook(month):
    content=BytesIO()
    with ZipFile(content,'w') as z:z.writestr('xl/workbook.xml',f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="{month}"/></sheets></workbook>')
    return content.getvalue()

class MonthlyCreationTests(unittest.TestCase):
    def setUp(self):
        self.folder=TemporaryDirectory();self.addCleanup(self.folder.cleanup);self.root=Path(self.folder.name)
        self.item={'key':'a'*64,'name':'南教室中１数学_2026.xlsx','driveId':'drive','itemId':'item'}
        self.entry={'driveId':'drive','itemId':'item','eTag':'current','sha256':monthly.digest(workbook('2026-11'))}
        self.client=Mock();self.client.metadata.return_value={'eTag':'current'}
        self.info={'filename':self.item['name']}
    def run_one(self,action='create',opened=None,inputs=None):
        return monthly.prepare_and_save(self.item,self.info,'2026-12',self.root,self.root,action,'request',set() if opened is None else opened,inputs or [],lambda:self.client)
    def generated(self,source,target,*args):
        self.assertEqual(source.read_bytes(),workbook('2026-11'))
        target.write_bytes(workbook('2026-12'))
        return {'first':{'S':{'counter':45},'A':{'counter':44},'B':{'counter':44}}}
    def test_local_open_is_skipped_before_cloud_access(self):
        result=self.run_one(opened={self.item['name'].casefold()});self.assertEqual(result['status'],'open');self.client.metadata.assert_not_called()
    def test_existing_month_never_generates_or_writes(self):
        with patch.object(monthly,'get_file',return_value=(workbook('2026-12'),self.entry)),patch.object(monthly,'generate_month_workbook') as generate,patch.object(monthly,'replace_existing_if_unchanged') as put:
            self.assertEqual(self.run_one()['status'],'exists');generate.assert_not_called();put.assert_not_called()
    def test_preview_never_backs_up_or_puts(self):
        with patch.object(monthly,'get_file',return_value=(workbook('2026-11'),self.entry)),patch.object(monthly,'generate_month_workbook',side_effect=self.generated),patch.object(monthly,'backup_file') as backup,patch.object(monthly,'replace_existing_if_unchanged') as put:
            result=self.run_one('preview');self.assertEqual(result['status'],'planned');self.assertTrue(result['adjusted']);backup.assert_not_called();put.assert_not_called()
    def test_fresh_version_is_used_for_the_conditional_write(self):
        entry={**self.entry,'eTag':'human-saved-current-version'}
        with patch.object(monthly,'get_file',return_value=(workbook('2026-11'),entry)),patch.object(monthly,'generate_month_workbook',side_effect=self.generated),patch.object(monthly,'backup_file') as backup,patch.object(monthly,'replace_existing_if_unchanged') as put:
            self.assertEqual(self.run_one()['status'],'created');self.assertEqual(put.call_args.args[3],'human-saved-current-version');backup.assert_called_once()
    def test_locked_put_leaves_only_that_file_pending(self):
        with patch.object(monthly,'get_file',return_value=(workbook('2026-11'),self.entry)),patch.object(monthly,'generate_month_workbook',side_effect=self.generated),patch.object(monthly,'backup_file'),patch.object(monthly,'replace_existing_if_unchanged',side_effect=monthly.GraphError(423)) as put:
            self.assertEqual(self.run_one()['status'],'open');self.assertEqual(put.call_count,1)
    def test_lost_put_is_reconciled_without_repeating_it(self):
        with patch.object(monthly,'get_file',side_effect=[(workbook('2026-11'),self.entry),(workbook('2026-12'),self.entry)]),patch.object(monthly,'generate_month_workbook',side_effect=self.generated),patch.object(monthly,'backup_file'),patch.object(monthly,'replace_existing_if_unchanged',side_effect=monthly.UploadSafetyError('lost')) as put:
            self.assertEqual(self.run_one()['status'],'created');self.assertEqual(put.call_count,1)
    def test_unreadable_manual_counter_requires_review(self):
        with patch.object(monthly,'get_file',return_value=(workbook('2026-11'),self.entry)),patch.object(monthly,'generate_month_workbook',side_effect=ValueError('counter')),patch.object(monthly,'replace_existing_if_unchanged') as put:
            self.assertEqual(self.run_one()['status'],'review');put.assert_not_called()
    def test_changed_timetable_blocks_the_write(self):
        self.client.metadata.return_value={'eTag':'new'}
        with patch.object(monthly,'get_file',return_value=(workbook('2026-11'),self.entry)),patch.object(monthly,'generate_month_workbook',side_effect=self.generated),patch.object(monthly,'backup_file') as backup,patch.object(monthly,'replace_existing_if_unchanged') as put:
            result=self.run_one(inputs=[self.entry]);self.assertEqual(result['reason'],'input_changed');backup.assert_not_called();put.assert_not_called()
    def test_scope_cannot_contain_paths_or_duplicate_keys(self):
        for value in ['["../workbook.xlsx"]','["'+('a'*64)+'","'+('a'*64)+'"]']:
            with self.assertRaises(ValueError):monthly.parse_keys(value)
    def test_january_and_february_use_the_existing_march_school_year(self):
        self.assertEqual(monthly.journal_workbook_year(2027,1),2026)
        self.assertEqual(monthly.journal_workbook_year(2027,2),2026)
        self.assertEqual(monthly.journal_workbook_year(2027,3),2027)

if __name__=='__main__':unittest.main()
