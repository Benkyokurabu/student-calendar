import copy
from datetime import datetime
from pathlib import Path
import tempfile
import unittest
import openpyxl
from bentan_journal_operation import (JST, validate_operation, workbook_sheets,
    assert_only_new_month, source_digest, merge_for_publication, validate_master, has_month_sheet)

IDENTITY = '12345678-1234-4234-8234-123456789abc'

class Boundaries(unittest.TestCase):
    def test_current_next_month_and_existing_december_hold(self):
        now = datetime(2026, 11, 5, tzinfo=JST)
        validate_operation('preview', '2026-12', IDENTITY, now)
        for action, month, identity in [('create','2026-12',IDENTITY), ('publish','2026-12',IDENTITY),
                                        ('create','2026-10',IDENTITY), ('preview','2026-11',';bad'),
                                        ('legacy','2026-11',IDENTITY)]:
            with self.assertRaises(ValueError):
                validate_operation(action, month, identity, now)
        validate_operation('create', '2026-12', IDENTITY, datetime(2026,11,20,tzinfo=JST))
        with self.assertRaises(ValueError):
            validate_master('create', '2026-12', False)
        validate_master('create', '2026-12', True)
        validate_master('preview', '2026-12', False)

    def test_old_sheets_and_user_input_must_survive_new_month_creation(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'journal.xlsx'
            book = openpyxl.Workbook()
            book.active.title='2026-10'
            book.active['A1']='入力済み本文'
            book.active['B2']='=1+2'
            book.active.merge_cells('C3:D4')
            book.save(path)
            self.assertTrue(has_month_sheet(path,'2026-10'))
            self.assertFalse(has_month_sheet(path,'2026-11'))
            before=workbook_sheets(path)
            book.create_sheet('2026-11')
            book.save(path)
            self.assertTrue(has_month_sheet(path,'2026-11'))
            assert_only_new_month(before,workbook_sheets(path),'2026-11')
            book['2026-10']['A1']='changed'
            book.save(path)
            with self.assertRaises(ValueError):
                assert_only_new_month(before,workbook_sheets(path),'2026-11')

    def test_snapshot_changes_when_any_original_version_or_public_baseline_changes(self):
        manifest={'entries':{'a':{'itemId':'1','eTag':'v1','sha256':'old'}}}
        baseline=source_digest(manifest,b'schedule',b'publicSchedule',b'journal')
        changed=copy.deepcopy(manifest);changed['entries']['a']['eTag']='v2'
        self.assertNotEqual(baseline,source_digest(changed,b'schedule',b'publicSchedule',b'journal'))
        self.assertNotEqual(baseline,source_digest(manifest,b'edited',b'publicSchedule',b'journal'))
        self.assertNotEqual(baseline,source_digest(manifest,b'schedule',b'publicSchedule',b'newer'))
        self.assertNotEqual(source_digest(manifest,b'schedule',b'publicSchedule',b'journal',b'master-v1'),
                            source_digest(manifest,b'schedule',b'publicSchedule',b'journal',b'master-v2'))

    def test_publication_preserves_existing_content_recording_policy_and_future_entries(self):
        key='2026-11-05|time|hon|j2mathS|1'
        future='2026-11-09|time|hon|j2mathS|1'
        old={'entries':{key:{'content':'保存済み','homework':[{'text':'宿題'}],'recordingUrl':''},future:{'content':'未来の入力','recordingUrl':''}}}
        new={'entries':{key:{'content':'','homework':[],'recordingUrl':'https://zoom.invalid/private','prevEntry':{'recordingUrl':'https://zoom.invalid/private'}},future:{'content':'別の入力','recordingUrl':'https://zoom.invalid/private'}}}
        merged=merge_for_publication(old,new,'2026-11-05')
        self.assertEqual(merged['entries'][key]['content'],'保存済み')
        self.assertEqual(merged['entries'][key]['homework'],old['entries'][key]['homework'])
        self.assertEqual(merged['entries'][key]['recordingUrl'],'')
        self.assertEqual(merged['entries'][key]['prevEntry']['recordingUrl'],'')
        self.assertEqual(merged['entries'][future],old['entries'][future])
        with self.assertRaises(ValueError):
            merge_for_publication(old,{'entries':{}},'2026-11-05')

if __name__=='__main__':
    unittest.main()
