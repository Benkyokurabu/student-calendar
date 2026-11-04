import copy
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from safe_journal_upload import (GraphError, UploadSafetyError, apply_manifest,
    capture_manifest, digest, main, read_original, replace_existing_if_unchanged,
    verify_listing_unchanged, verify_manifest)


class FakeGraph:
    def __init__(self):
        self.items = {'book': self.file('book','book.xlsx',b'original'),
                      'scratch': {'id':'scratch','name':'.bentan-cas-probe-test','folder':{}}}
        self.puts=[]; self.deletes=[]; self.ignore_conditions=False
        self.before_original_put=None; self.after_original_put=None
        self.metadata_hook=None; self.download_hook=None
        self.original_put_error=None; self.metadata_error=None

    def file(self,item,name,data,version=1):
        return {'id':item,'name':name,'file':{},'parentReference':{'driveId':'drive'},
                'eTag':f'"{item}-{version}"','size':len(data),'bytes':data,'version':version}

    def metadata(self,drive,item):
        if self.metadata_error: raise GraphError(self.metadata_error)
        if item not in self.items: raise GraphError(404)
        value=copy.deepcopy(self.items[item]); value.pop('bytes',None)
        if self.metadata_hook: self.metadata_hook(drive,item)
        return value

    def download(self,drive,item):
        if self.download_hook: self.download_hook(drive,item)
        return self.items[item]['bytes']

    def edit(self,item,data):
        old=self.items[item];self.items[item]=self.file(item,old['name'],data,old['version']+1)

    def put_if_match(self,drive,item,content,etag):
        self.puts.append((item,etag))
        if item=='book' and self.before_original_put: self.before_original_put()
        if item=='book' and self.original_put_error:
            if self.original_put_error=='lost':
                self.edit(item,content)
                raise UploadSafetyError('Transport failed; uncertain write')
            raise GraphError(self.original_put_error)
        if not self.ignore_conditions and self.items[item]['eTag']!=etag:raise GraphError(412)
        self.edit(item,content);result=self.metadata(drive,item)
        if item=='book' and self.after_original_put:self.after_original_put()
        return result

    @staticmethod
    def route(drive,item):return '/drives/'+drive+'/items/'+item

    def request(self,method,route,body=None,headers=None,binary=False):
        if method=='PUT' and ':/' in route:
            name=route.split(':/',1)[1].removesuffix(':/content')
            self.items['probe']=self.file('probe',name,body)
            return self.metadata('drive','probe'),201
        if method=='DELETE':
            item=route.rsplit('/',1)[-1]
            if self.items[item]['eTag']!=headers['If-Match']:raise GraphError(412)
            self.deletes.append(item);del self.items[item];return {},204
        raise AssertionError((method,route))


class SafeUploadTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.root=Path(self.tmp.name);(self.root/'book.xlsx').write_bytes(b'original')
        self.graph=FakeGraph()
        self.listing=[{'Path':'book.xlsx','ID':'drive#book','Size':8}]
        self.manifest=capture_manifest(self.graph,self.root,self.listing,'onedrive:test')

    def change_local(self): (self.root/'book.xlsx').write_bytes(b'maintenance')

    def apply(self):return apply_manifest(self.graph,self.root,self.manifest,{'drive':'scratch'})

    def test_capture_matches_original_bytes_and_version_not_mtime(self):
        entry=self.manifest['entries']['book.xlsx']
        self.assertEqual(entry['sha256'],digest(b'original'))
        self.assertEqual(entry['eTag'],'"book-1"')
        (self.root/'book.xlsx').write_bytes(b'different')
        with self.assertRaises(UploadSafetyError):capture_manifest(self.graph,self.root,self.listing,'onedrive:test')

    def test_edit_during_original_download_is_rejected(self):
        self.graph.download_hook=lambda d,i:self.graph.edit(i,b'new teacher data')
        with self.assertRaises(UploadSafetyError):read_original(self.graph,'drive','book')

    def test_snapshot_coverage_and_traversal_fail_closed(self):
        for path in ['../book.xlsx','missing.xlsx']:
            with self.subTest(path=path),self.assertRaises(UploadSafetyError):
                capture_manifest(self.graph,self.root,[{'Path':path,'ID':'drive#book'}],'onedrive:test')
        (self.root/'new.xlsx').write_bytes(b'new')
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(self.graph.puts,[])

    def test_shared_drive_identity_is_not_replaced_by_configured_root(self):
        with self.assertRaises(UploadSafetyError):
            capture_manifest(self.graph,self.root,[{'Path':'book.xlsx','ID':'book'}],'onedrive:test')
        self.graph.items['book']['parentReference']['driveId']='another-shared-drive'
        with self.assertRaises(UploadSafetyError):
            capture_manifest(self.graph,self.root,self.listing,'onedrive:test')
        self.assertEqual(self.graph.puts,[])

    def test_capture_rejects_original_renamed_since_listing_even_with_same_bytes(self):
        self.graph.items['book']['name']='renamed.xlsx'
        with self.assertRaises(UploadSafetyError):
            capture_manifest(self.graph,self.root,self.listing,'onedrive:test')
        self.assertEqual(self.graph.puts,[])

    def test_capture_rejects_duplicate_paths_case_collisions_and_item_aliases(self):
        cases = [self.listing+self.listing,
                 self.listing+[{'Path':'BOOK.xlsx','ID':'drive#another-item'}],
                 self.listing+[{'Path':'moved/book.xlsx','ID':'DRIVE#book'}]]
        for listing in cases:
            with self.subTest(listing=listing),self.assertRaises(UploadSafetyError):
                capture_manifest(self.graph,self.root,listing,'onedrive:test')

    def test_final_listing_requires_exact_path_to_identity_mapping(self):
        verify_listing_unchanged(self.manifest,copy.deepcopy(self.listing))
        verify_listing_unchanged(self.manifest,[{'Path':'book.xlsx','ID':'DRIVE#book'}])
        cases = [[],
                 [{'Path':'renamed.xlsx','ID':'drive#book'}],
                 [{'Path':'moved/book.xlsx','ID':'drive#book'}],
                 [{'Path':'book.xlsx','ID':'drive#replacement'}],
                 self.listing+[{'Path':'new.xlsx','ID':'drive#new'}],
                 self.listing+self.listing]
        for listing in cases:
            with self.subTest(listing=listing),self.assertRaises(UploadSafetyError):
                verify_listing_unchanged(self.manifest,listing)

    def test_cli_capture_relists_before_saving_and_rejects_move_or_rename(self):
        output=self.root/'manifest.json'
        argv=['safe_journal_upload.py','capture','--remote','onedrive:test',
              '--local-dir',str(self.root),'--manifest',str(output)]
        for final in [self.listing,
                      [{'Path':'renamed.xlsx','ID':'drive#book'}],
                      [{'Path':'moved/book.xlsx','ID':'drive#book'}]]:
            output.unlink(missing_ok=True)
            with self.subTest(final=final),patch('sys.argv',argv),\
                 patch('safe_journal_upload.GraphClient.from_rclone',return_value=self.graph),\
                 patch('safe_journal_upload.remote_listing',side_effect=[self.listing,final]) as listing:
                if final==self.listing:
                    self.assertEqual(main(),0)
                    self.assertEqual(json.loads(output.read_text(encoding='utf-8')),self.manifest)
                else:
                    with self.assertRaises(UploadSafetyError):main()
                    self.assertFalse(output.exists())
                self.assertEqual(listing.call_count,2)
        self.assertEqual(self.graph.puts,[])

    def test_unchanged_workbooks_do_not_upload_or_need_a_probe(self):
        result=apply_manifest(self.graph,self.root,self.manifest,{})
        self.assertEqual(result,{'checked':1,'uploaded':0,'unchanged':1})
        self.assertEqual(self.graph.puts,[])

    def test_validation_exercises_guard_without_any_original_write(self):
        result = verify_manifest(self.graph,self.root,self.manifest,{'drive':'scratch'})
        self.assertTrue(result['guardVerified'])
        self.assertEqual(result['uploaded'],0)
        self.assertEqual(self.graph.items['book']['bytes'],b'original')
        self.assertFalse(any(item=='book' for item,etag in self.graph.puts))
        self.assertEqual(self.graph.deletes,['probe'])

    def test_validation_refuses_modified_local_or_remote_content(self):
        self.change_local()
        with self.assertRaises(UploadSafetyError):
            verify_manifest(self.graph,self.root,self.manifest,{'drive':'scratch'})
        self.assertEqual(self.graph.puts,[])
        (self.root/'book.xlsx').write_bytes(b'original')
        self.graph.edit('book',b'new teacher entry')
        with self.assertRaises(UploadSafetyError):
            verify_manifest(self.graph,self.root,self.manifest,{'drive':'scratch'})
        self.assertEqual(self.graph.puts,[])

    def test_successful_upload_uses_frozen_etag_and_checks_result(self):
        self.change_local();result=self.apply()
        self.assertEqual(result['uploaded'],1)
        self.assertEqual(self.graph.items['book']['bytes'],b'maintenance')
        self.assertEqual([call for call in self.graph.puts if call[0]=='book'],[('book','"book-1"')])
        self.assertEqual(self.graph.deletes,['probe'])

    def test_remote_edit_during_maintenance_blocks_all_uploads(self):
        self.change_local();self.graph.edit('book',b'teacher edited')
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(self.graph.puts,[])
        self.assertEqual(self.graph.items['book']['bytes'],b'teacher edited')

    def test_edit_between_precheck_and_put_is_atomically_rejected(self):
        self.change_local()
        self.graph.before_original_put=lambda:self.graph.edit('book',b'teacher wins race')
        with self.assertRaises(GraphError) as raised:self.apply()
        self.assertEqual(raised.exception.status,412)
        self.assertEqual(self.graph.items['book']['bytes'],b'teacher wins race')

    def test_a_locked_workbook_is_not_reported_as_success(self):
        self.change_local();self.graph.original_put_error=423
        with self.assertRaises(GraphError) as raised:self.apply()
        self.assertEqual(raised.exception.status,423)
        self.assertEqual(self.graph.items['book']['bytes'],b'original')

    def test_authentication_failure_does_not_write(self):
        self.change_local();self.graph.metadata_error=401
        with self.assertRaises(GraphError):self.apply()
        self.assertEqual(self.graph.puts,[])

    def test_unsupported_conditional_put_never_touches_original(self):
        self.change_local();self.graph.ignore_conditions=True
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertFalse(any(item=='book' for item,etag in self.graph.puts))
        self.assertEqual(self.graph.items['book']['bytes'],b'original')

    def test_missing_probe_registration_cannot_enable_writes(self):
        self.change_local()
        with self.assertRaises(UploadSafetyError):apply_manifest(self.graph,self.root,self.manifest,{})
        self.assertEqual(self.graph.puts,[])

    def test_probe_cannot_use_an_existing_journal_folder(self):
        self.change_local();self.graph.items['scratch']['name']='授業日誌'
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(self.graph.puts,[])

    def test_newer_native_edit_after_put_is_preserved_and_blocks_publication(self):
        self.change_local()
        self.graph.after_original_put=lambda:self.graph.edit('book',b'newer native edit')
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(self.graph.items['book']['bytes'],b'newer native edit')

    def test_lost_put_response_is_not_retried_or_rolled_back(self):
        self.change_local();self.graph.original_put_error='lost'
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(len([i for i,e in self.graph.puts if i=='book']),1)
        self.assertEqual(self.graph.items['book']['bytes'],b'maintenance')

    def test_original_content_mismatch_is_checked_even_with_same_etag(self):
        self.graph.items['book']['bytes']=b'mutated without version'
        self.graph.items['book']['size']=len(b'mutated without version')
        with self.assertRaises(UploadSafetyError):
            replace_existing_if_unchanged(self.graph,'drive','book','"book-1"',digest(b'original'),b'maintenance')
        self.assertEqual(self.graph.puts,[])

    def test_edit_of_untouched_original_before_publication_stops_publication(self):
        calls=[]
        def edit_after_first_read(drive,item):
            calls.append(item)
            if len(calls)==1:self.graph.edit(item,b'teacher edit during final checks')
        self.graph.metadata_hook=edit_after_first_read
        with self.assertRaises(UploadSafetyError):self.apply()
        self.assertEqual(self.graph.puts,[])
        self.assertEqual(self.graph.items['book']['bytes'],b'teacher edit during final checks')

    def test_workflow_requires_snapshot_before_maintenance_and_blocks_on_conflict(self):
        workflow=(Path(__file__).resolve().parents[1]/'.github/workflows/update-journal.yml').read_text(encoding='utf-8')
        capture=workflow.index('python scripts/safe_journal_upload.py capture')
        pipeline=workflow.index('python ci_extract_and_push.py')
        apply=workflow.index('python scripts/safe_journal_upload.py apply')
        publication=workflow.index('- name: Commit and push JSON files')
        self.assertLess(capture,pipeline);self.assertLess(pipeline,apply);self.assertLess(apply,publication)
        upload_step=workflow[workflow.rfind('- name:',0,apply):publication]
        self.assertNotIn('continue-on-error',upload_step)
        self.assertNotIn('--ignore-errors',workflow)
        self.assertIn('vars.JOURNAL_UPLOAD_PROBE_PARENTS',upload_step)
        for step in ['Copy schedule Excel to scripts folder','Download registered common master from OneDrive',
                     'Run pipeline','Upload only unchanged original journal versions with atomic If-Match',
                     'Commit and push JSON files']:
            block=workflow.split('- name: '+step,1)[1].split('- name:',1)[0]
            self.assertIn("if: ${{ inputs.validate_only != true }}",block)
        independent=(Path(__file__).resolve().parents[1]/'.github/workflows/test-student-calendar.yml').read_text(encoding='utf-8')
        self.assertIn('test_safe_journal_upload.py',independent)
        self.assertIn('test_retry_deferred_journal_slots.py',independent)


if __name__=='__main__':unittest.main()
