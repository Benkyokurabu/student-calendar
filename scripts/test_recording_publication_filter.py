import unittest
import hashlib
import copy
from unittest.mock import patch
import io
from recording_publication_filter import apply_payload

class PublicationFilterTests(unittest.TestCase):
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
