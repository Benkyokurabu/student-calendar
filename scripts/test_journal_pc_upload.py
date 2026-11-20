import json
import ast
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import download_journal_from_cloud as download
import journal_pc_upload as pc
from safe_journal_upload import GraphError, UploadSafetyError
from test_safe_journal_upload import FakeGraph


class PcUploadTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        (self.root/'book.xlsx').write_bytes(b'original')
        self.graph = FakeGraph()
        self.rows = [{'Path':'book.xlsx','ID':'drive#book'}]
        for target, value in [('graph_client',self.graph), ('mapped_listing',self.rows),
                              ('probe_parents',{'drive':'scratch'})]:
            mock = patch.object(pc,target,return_value=value)
            mock.start(); self.addCleanup(mock.stop)
        pc.freeze_download('rclone','onedrive:test',self.root)

    def apply(self):
        return pc.guarded_upload('rclone','onedrive:test',self.root)

    def change(self):
        (self.root/'book.xlsx').write_bytes(b'maintenance')

    def test_download_snapshot_preserves_shared_identity(self):
        data=json.loads((self.root/pc.MANIFEST).read_text(encoding='utf-8'))
        self.assertEqual(data['entries']['book.xlsx']['driveId'],'drive')
        self.assertEqual(data['localRoot'],str(self.root.resolve()))

    def test_teacher_edit_after_download_blocks_write_and_publication(self):
        self.change(); self.graph.edit('book',b'teacher saved')
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(self.graph.puts,[])
        self.assertEqual(self.graph.items['book']['bytes'],b'teacher saved')

    def test_teacher_edit_during_put_wins(self):
        self.change()
        self.graph.before_original_put=lambda:self.graph.edit('book',b'teacher wins')
        with self.assertRaises(GraphError) as e:self.apply()
        self.assertEqual(e.exception.status,412)
        self.assertEqual(self.graph.items['book']['bytes'],b'teacher wins')

    def test_successful_save_is_verified_and_cannot_be_replayed(self):
        self.change(); self.assertEqual(self.apply()['uploaded'],1)
        self.assertEqual(self.graph.items['book']['bytes'],b'maintenance')
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(len([x for x in self.graph.puts if x[0]=='book']),1)

    def test_unchanged_originals_do_not_need_a_probe_or_any_put(self):
        with patch.object(pc,'probe_parents',side_effect=AssertionError('must not probe')):
            self.assertEqual(self.apply()['uploaded'],0)
        self.assertEqual(self.graph.puts,[])

    def test_lost_response_is_not_retried(self):
        self.change(); self.graph.original_put_error='lost'
        with self.assertRaises(UploadSafetyError):self.apply()
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(len([x for x in self.graph.puts if x[0]=='book']),1)

    def test_locked_workbook_fails_without_fallback(self):
        self.change(); self.graph.original_put_error=423
        with self.assertRaises(GraphError):self.apply()
        self.assertEqual(self.graph.items['book']['bytes'],b'original')

    def test_missing_snapshot_never_uploads_old_cache(self):
        (self.root/pc.MANIFEST).unlink();self.change()
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(self.graph.puts,[])

    def test_snapshot_from_another_folder_is_rejected(self):
        target=self.root/pc.MANIFEST
        data=json.loads(target.read_text(encoding='utf-8'));data['localRoot']='another-folder'
        target.write_text(json.dumps(data),encoding='utf-8')
        with self.assertRaises(UploadSafetyError):self.apply()

    def test_replaced_session_cannot_upload_another_process_copy(self):
        target=self.root/pc.MANIFEST
        data=json.loads(target.read_text(encoding='utf-8'));data['session']='another-session'
        target.write_text(json.dumps(data),encoding='utf-8')
        with self.assertRaises(UploadSafetyError):self.apply()

    def test_same_directory_cannot_be_downloaded_and_uploaded_together(self):
        with pc.operation_lock(self.root):
            with self.assertRaises(UploadSafetyError):
                with pc.operation_lock(self.root):self.fail('must not acquire twice')
        self.assertFalse((self.root/'.journal-operation.lock').exists())

    def test_default_downloads_have_separate_working_copies(self):
        with patch.object(download,'_download_journal_to',side_effect=lambda root,exclude:root):
            with patch.object(pc.tempfile,'mkdtemp',side_effect=[str(self.root/'a'),str(self.root/'b')]):
                a=download.download_journal();b=download.download_journal()
        self.assertNotEqual(a,b)

    def test_failed_download_invalidates_old_snapshot(self):
        with patch.object(download,'_download_journal_to',side_effect=RuntimeError('download failed')):
            with self.assertRaises(RuntimeError):download.download_journal(self.root)
        self.assertFalse((self.root/pc.MANIFEST).exists())

    def test_public_upload_propagates_conflict(self):
        with patch.object(download,'find_rclone',return_value='rclone'), \
             patch.object(download,'validate_remote_tree_names'), \
             patch.object(download,'guarded_upload',side_effect=UploadSafetyError('edited')):
            with self.assertRaises(UploadSafetyError):download.upload_journal(self.root)


class WindowsPathTests(unittest.TestCase):
    def test_encoded_colon_resolves_only_to_the_downloaded_workbook(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);(root/'本校\u201b：中２').mkdir()
            (root/'本校\u201b：中２'/'book.xlsx').write_bytes(b'original')
            rows=[{'Path':'本校:中２/book.xlsx','ID':'drive#book'}]
            # Call the real mapper; no fullwidth/raw-colon folder is created.
            with patch.object(pc,'rclone_json',return_value=rows):
                self.assertEqual(pc.mapped_listing('rclone','remote',root)[0]['Path'],
                                 '本校\u201b：中２/book.xlsx')

    def test_remote_aliases_and_duplicate_ids_are_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);(root/'book.xlsx').write_bytes(b'original')
            rows=[{'Path':'book.xlsx','ID':'drive#book'},
                  {'Path':'other.xlsx','ID':'drive#book'}]
            with patch.object(pc,'rclone_json',return_value=rows):
                with self.assertRaises(UploadSafetyError):pc.mapped_listing('rclone','remote',root)

    def test_no_registered_probe_never_creates_one(self):
        with patch.dict(pc.os.environ,{},clear=True),patch.object(pc,'rclone_json',return_value=[]):
            with self.assertRaises(UploadSafetyError):pc.probe_parents('rclone','remote',{'book':{'driveId':'drive'}})


class PublicationStopsTests(unittest.TestCase):
    def test_code_push_cannot_start_original_workbook_job(self):
        workflow=(Path(__file__).resolve().parents[1]/'.github/workflows/update-journal.yml').read_text(encoding='utf-8')
        update=workflow.split('  update:',1)[1].split('    runs-on:',1)[0]
        self.assertIn("if: github.event_name != 'push'",update)

    def statement(self, name):
        source=Path(__file__).with_name('extract_and_push.py').read_text(encoding='utf-8')
        tree=ast.parse(source)
        for node in ast.walk(tree):
            if not isinstance(node,ast.Try):continue
            if any(isinstance(x,ast.Call) and isinstance(x.func,ast.Name) and x.func.id==name
                   for body in node.body for x in ast.walk(body)):
                return node
        self.fail('pipeline stage missing')

    def test_failed_original_save_cannot_continue_to_publication(self):
        from unittest.mock import Mock
        failure=UploadSafetyError('teacher edited')
        stage=self.statement('upload_journal')
        publication=Mock()
        code=ast.Module(body=[stage,ast.Expr(value=ast.Call(func=ast.Name(id='publish',ctx=ast.Load()),args=[],keywords=[]))],type_ignores=[])
        ast.fix_missing_locations(code)
        with self.assertRaises(UploadSafetyError):
            exec(compile(code,'pipeline-upload','exec'),{'upload_journal':Mock(side_effect=failure),
                 'cloud_journal_dir':Path('isolated'),'publish':publication})
        publication.assert_not_called()

    def test_failed_original_download_cannot_fall_back_to_stale_files(self):
        from unittest.mock import Mock
        import sys,types
        stage=self.statement('download_journal');publication=Mock()
        module=types.ModuleType('download_journal_from_cloud')
        module.download_journal=Mock(side_effect=RuntimeError('connection lost'))
        module.upload_journal=Mock()
        code=ast.Module(body=[stage,ast.Expr(value=ast.Call(func=ast.Name(id='publish',ctx=ast.Load()),args=[],keywords=[]))],type_ignores=[])
        ast.fix_missing_locations(code)
        with patch.dict(sys.modules,{'download_journal_from_cloud':module}):
            with self.assertRaises(RuntimeError):exec(compile(code,'pipeline-download','exec'),{'publish':publication})
        publication.assert_not_called()


if __name__=='__main__':unittest.main()
