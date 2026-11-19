import os
import json
from pathlib import Path
import unittest
from unittest.mock import patch
from bentan_endpoints import checked_origin, endpoint, public_read_headers


class EndpointTests(unittest.TestCase):
    def test_default_uses_the_reviewed_shipping_manifest(self):
        manifest = json.loads((Path(__file__).resolve().parent.parent / 'bentan-endpoints.json').read_text('utf-8'))
        self.assertIn(manifest['apiBaseUrl'], ('https://line-check-system.vercel.app', 'https://bentang10ops20261119.shares.zrok.io'))
        with patch.dict(os.environ, {'BENTAN_API_BASE_URL': ''}):
            self.assertEqual(endpoint('/api/recordings/publication'), manifest['apiBaseUrl'] + '/api/recordings/publication')

    def test_all_existing_routes_switch_together(self):
        with patch.dict(os.environ, {'BENTAN_API_BASE_URL': 'https://g10.example.invalid/'}):
            for route in ('/api/recordings/publication', '/api/recordings/capture', '/api/journal-management/callback', '/api/journal-management/supplements'):
                self.assertEqual(endpoint(route), 'https://g10.example.invalid' + route)

    def test_invalid_origin_never_receives_publisher_credentials(self):
        for origin in ('http://example.invalid', 'https://user:pass@example.invalid', 'https://example.invalid/path', 'https://example.invalid?x=1', 'https://example.invalid/#x', 'https://example.invalid\n', 'https://example.invalid:bad'):
            with self.subTest(origin=origin), self.assertRaises(ValueError):
                checked_origin(origin)

    def test_arbitrary_routes_are_rejected(self):
        with self.assertRaises(ValueError):
            endpoint('/rest/v1/students')

    def test_zrok_header_only_for_the_approved_configured_public_origin(self):
        target='https://bentang10ops20261119.shares.zrok.io'
        with patch.dict(os.environ, {'BENTAN_API_BASE_URL':target}):
            self.assertEqual(public_read_headers(target+'/api/recordings/publication'),{'skip_zrok_interstitial':'true'})
            for url in ('https://other.shares.zrok.io/api/recordings/publication',target+'/api/recordings/capture','http://bentang10ops20261119.shares.zrok.io/api/recordings/publication'):
                self.assertEqual(public_read_headers(url),{})
        with patch.dict(os.environ, {'BENTAN_API_BASE_URL':''}):
            self.assertEqual(public_read_headers('https://line-check-system.vercel.app/api/recordings/publication'),{})


if __name__ == '__main__':
    unittest.main()
