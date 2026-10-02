import unittest
import hashlib
import copy
from recording_publication_filter import apply_payload

class PublicationFilterTests(unittest.TestCase):
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
