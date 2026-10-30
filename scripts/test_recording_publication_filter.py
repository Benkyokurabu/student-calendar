import unittest
import hashlib
import copy
from unittest.mock import patch
import io
import json
import tempfile
from pathlib import Path
from datetime import datetime, timezone, timedelta
from types import SimpleNamespace
from recording_publication_filter import apply_payload, preserve_recording_fields, publish_text_with_existing_recordings, refresh_approved_snapshots

class PublicationFilterTests(unittest.TestCase):
    def test_journal_only_publication_never_contacts_the_policy_service(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / 'journal_2026-10.json'
            path.write_text(json.dumps({'entries': {'lesson': {'content': 'latest', 'recordingUrl': 'new'}}}), 'utf-8')
            baseline = {'entries': {'lesson': {'content': 'old', 'recordingUrl': 'existing'}}}
            with patch('recording_publication_filter.urlopen') as network, patch('recording_publication_filter.subprocess.run', return_value=SimpleNamespace(returncode=0, stdout=json.dumps(baseline).encode())):
                publish_text_with_existing_recordings(root)
                network.assert_not_called()
            result = json.loads(path.read_text('utf-8'))['entries']['lesson']
            self.assertEqual(result, {'content': 'latest', 'recordingUrl': 'existing'})

    def test_verified_snapshots_refresh_automatically_and_exclude_restrictions(self):
        now = datetime(2026, 10, 30, tzinfo=timezone.utc)
        keys = [f'2026-10-29|18:00|hon|hon_j{i}_S_math|1' for i in range(1, 8)]
        urls = [f'https://zoom.us/rec/play/video{i}' for i in range(7)]
        entries = {key: {'url': url} for key, url in zip(keys, urls)}
        entries[keys[1]]['hidden'] = True
        entries[keys[4]]['url'] = 'https://zoom.us.evil.test/rec/play/wrong'
        rules = [{'key': keys[2], 'status': 'hidden', 'eventKeys': [keys[2]]},
                 {'key': 'test', 'status': 'hidden', 'eventKeys': [], 'match': {'date': '2026-10-29', 'campus': 'hon', 'group': 'hon_j4_S_math'}},
                 {'key': 'alias', 'status': 'hidden', 'eventKeys': [], 'urlHashes': [hashlib.sha256(urls[5].encode()).hexdigest()]}]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'zoom_recording_urls_2026-10.json').write_text(json.dumps({'entries': entries}), 'utf-8')
            (root / 'recording_release.json').write_text(json.dumps({'rules': [{'eventKeys': [keys[6]], 'status': 'blocked'}]}), 'utf-8')
            self.assertEqual(refresh_approved_snapshots(root, rules, now), ['recording_emergency_2026-10.json'])
            snapshot = json.loads((root / 'recording_emergency_2026-10.json').read_text('utf-8'))
            self.assertEqual(snapshot['entries'], {keys[0]: urls[0]})
            self.assertEqual(datetime.fromisoformat(snapshot['expiresAt']), now + timedelta(days=7))
            self.assertEqual(refresh_approved_snapshots(root, rules, now + timedelta(hours=1)), [])
            self.assertEqual(len(refresh_approved_snapshots(root, rules, now + timedelta(days=1))), 1)
            rules.append({'key': keys[0], 'status': 'hidden', 'eventKeys': [keys[0]]})
            refresh_approved_snapshots(root, rules, now + timedelta(days=1, hours=1))
            self.assertEqual(json.loads((root / 'recording_emergency_2026-10.json').read_text('utf-8'))['entries'], {})

    def test_outage_publishes_text_but_only_existing_recording_urls(self):
        old = {'entries': {'same': {'content': 'old', 'recordingUrl': 'https://zoom.us/old'},
                           'hidden': {'recordingUrl': '', 'hidden': True, 'recordingPublicationKey': 'hidden'}}}
        new = {'entries': {'same': {'content': 'new', 'homework': ['new homework'], 'recordingUrl': 'https://zoom.us/new'},
                           'hidden': {'content': 'new test', 'recordingUrl': 'https://zoom.us/private'},
                           'new': {'content': 'new lesson', 'recordingUrl': 'https://zoom.us/unverified'}}}
        preserve_recording_fields(new, old)
        self.assertEqual(new['entries']['same']['content'], 'new')
        self.assertEqual(new['entries']['same']['homework'], ['new homework'])
        self.assertEqual(new['entries']['same']['recordingUrl'], 'https://zoom.us/old')
        self.assertEqual(new['entries']['hidden']['recordingUrl'], '')
        self.assertTrue(new['entries']['hidden']['hidden'])
        self.assertEqual(new['entries']['new']['recordingUrl'], '')

    def test_outage_does_not_reuse_a_previous_lesson_url_when_identity_changed(self):
        old = {'prevEntry': {'date': '2026-10-20', 'recordingUrl': 'https://zoom.us/old'}}
        new = {'prevEntry': {'date': '2026-10-27', 'content': 'latest', 'recordingUrl': 'https://zoom.us/new'}}
        preserve_recording_fields(new, old)
        self.assertEqual(new['prevEntry']['content'], 'latest')
        self.assertEqual(new['prevEntry']['recordingUrl'], '')

    def test_cloud_publisher_uses_signed_workflow_identity(self):
        key='2026-10-01|18:00～19:00|hon|hon_j1_S_math|2'
        payload={'entries':{key:{'url':'https://example.test/private'}}}
        rule={'key':'test:row','eventKeys':[],'status':'hidden','url':'','match':{'date':'2026-10-01','campus':'hon','group':'hon_j1_S_math'}}
        env={'GITHUB_ACTIONS':'true','ACTIONS_ID_TOKEN_REQUEST_URL':'https://example.test/oidc?request=1','ACTIONS_ID_TOKEN_REQUEST_TOKEN':'test-only'}
        with patch.dict('os.environ',env),patch('recording_publication_filter.urlopen',side_effect=[io.BytesIO(b'{"value":"signed-test-identity"}'),io.BytesIO(b'{"captured":1}')]) as requests,patch('recording_publication_filter.load_rules',return_value=[rule]):
            apply_payload(payload,[rule])
            self.assertEqual(requests.call_args_list[1].args[0].get_header('Authorization'),'Bearer signed-test-identity')
        self.assertEqual(payload['entries'][key]['url'],'')
    def test_partial_capture_stops_before_stripping_original_url(self):
        key='2026-10-01|18:00～19:00|hon|hon_j1_S_math|2'
        payload={'entries':{key:{'url':'https://example.test/private'}}}
        rule={'key':'test:row','eventKeys':[],'status':'hidden','url':'','match':{'date':'2026-10-01','campus':'hon','group':'hon_j1_S_math'}}
        with patch.dict('os.environ',{'GITHUB_TOKEN':'test-only','GITHUB_ACTIONS':'false'}),patch('recording_publication_filter.urlopen',return_value=io.BytesIO(b'{"captured":0}')):
            self.assertRaises(ValueError,apply_payload,payload,[rule])
        self.assertEqual(payload['entries'][key]['url'],'https://example.test/private')

    def test_more_than_100_test_recordings_are_captured_in_batches(self):
        rules=[];entries={}
        for day in range(1,102):
            key=f'2026-10-01|time{day}|hon|hon_j1_S_math|2'
            entries[key]={'url':f'https://example.test/video{day}'}
        rule={'key':'test:row','eventKeys':[],'status':'hidden','url':'','match':{'date':'2026-10-01','campus':'hon','group':'hon_j1_S_math'}}
        with patch.dict('os.environ',{'GITHUB_TOKEN':'test-only','GITHUB_ACTIONS':'false'}),patch('recording_publication_filter.urlopen',side_effect=[io.BytesIO(b'{"captured":100}'),io.BytesIO(b'{"captured":1}')]) as capture,patch('recording_publication_filter.load_rules',return_value=[rule]):
            apply_payload({'entries':entries},[rule]);self.assertEqual(capture.call_count,2)
        self.assertTrue(all(not entry['url'] for entry in entries.values()))

    def test_new_test_recording_is_stored_privately_before_first_publication(self):
        key='2026-10-01|18:00～19:00|hon|hon_j1_S_math|2'
        url='https://example.test/new-test'
        automatic={'key':'test:row','eventKeys':[],'status':'hidden','url':'','match':{'date':'2026-10-01','campus':'hon','group':'hon_j1_S_math'}}
        stored={'key':key,'eventKeys':[key],'status':'hidden','url':'','urlHashes':[hashlib.sha256(url.encode()).hexdigest()]}
        payload={'entries':{key:{'url':url},'alias':{'url':url},'next-day':{'url':'https://example.test/ordinary'}}}
        with patch.dict('os.environ',{'GITHUB_TOKEN':'test-only','GITHUB_ACTIONS':'false'}),patch('recording_publication_filter.urlopen',return_value=io.BytesIO(b'{"captured":1}')) as capture,patch('recording_publication_filter.load_rules',return_value=[stored,automatic]):
            apply_payload(payload,[automatic])
            self.assertEqual(capture.call_count,1)
            self.assertEqual(capture.call_args.args[0].method,'POST')
        self.assertNotIn(url,str(payload))
        self.assertEqual(payload['entries'][key]['recordingPublicationKey'],key)
        self.assertEqual(payload['entries']['next-day']['url'],'https://example.test/ordinary')

    def test_capture_failure_stops_without_stripping_restore_url(self):
        key='2026-10-01|18:00～19:00|hon|hon_j1_S_math|2'
        payload={'entries':{key:{'url':'https://example.test/private'}}}
        rule={'key':'test:row','eventKeys':[],'status':'hidden','url':'','match':{'date':'2026-10-01','campus':'hon','group':'hon_j1_S_math'}}
        with patch.dict('os.environ',{'GITHUB_TOKEN':'test-only','GITHUB_ACTIONS':'false'}),patch('recording_publication_filter.urlopen',side_effect=OSError('unavailable')):
            self.assertRaises(OSError,apply_payload,payload,[rule])
        self.assertEqual(payload['entries'][key]['url'],'https://example.test/private')

    def test_hide_restore_and_aliases_without_touching_other_lessons(self):
        url = 'https://example.test/private'
        key = '2026-10-01|6:35～8:05|hon|hon_j1_S_math|2'
        rules = [{'key': key, 'eventKeys': [key], 'status': 'hidden', 'url': '', 'urlHashes': [hashlib.sha256(url.encode()).hexdigest()]}]
        payload = {'entries': {key: {'url': url, 'content': '単元テスト'}, 'next': {'recordingUrl': 'https://example.test/next', 'prevEntry': {'recordingUrl': url}}, 'other': {'url': 'https://example.test/other'}}}
        apply_payload(payload, rules)
        self.assertNotIn(url, str(payload))
        self.assertEqual(payload['entries'][key]['content'], '単元テスト')
        self.assertEqual(payload['entries']['next']['recordingUrl'], 'https://example.test/next')
        self.assertEqual(payload['entries']['other']['url'], 'https://example.test/other')
        again = copy.deepcopy(payload)
        apply_payload(payload, rules)
        self.assertEqual(payload, again)
        rules[0].update(status='public', url=url)
        apply_payload(payload, rules)
        self.assertEqual(payload['entries'][key]['url'], url)
        self.assertEqual(payload['entries']['next']['prevEntry']['recordingUrl'], url)
        self.assertNotIn('hidden', payload['entries'][key])

    def test_newer_public_urls_are_preserved(self):
        key='lesson'
        payload={'entries': {key:{'url':'https://example.test/current'}}}
        apply_payload(payload,[{'key':key,'eventKeys':[key],'status':'public','url':'https://example.test/old'}])
        self.assertEqual(payload['entries'][key]['url'],'https://example.test/current')

if __name__ == '__main__':
    unittest.main()
