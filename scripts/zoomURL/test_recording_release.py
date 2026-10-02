import unittest
from datetime import datetime, timezone
from unittest.mock import patch
import recording_release as r

class ReleaseTests(unittest.TestCase):
    def test_japan_time_and_boundary(self):
        rule = {'mode': 'scheduled', 'releaseAt': '2026-10-10T22:00', 'originalShare': 'internally'}
        self.assertEqual(r.desired_share(rule, datetime(2026, 10, 10, 12, 59, tzinfo=timezone.utc)), 'none')
        self.assertEqual(r.desired_share(rule, datetime(2026, 10, 10, 13, 0, tzinfo=timezone.utc)), 'internally')

    def test_private_has_no_expiry(self):
        self.assertEqual(r.desired_share({'mode': 'private'}, datetime.now(timezone.utc)), 'none')

    def test_uuid_encoding(self):
        self.assertTrue(r.settings_url('/a//b=').endswith('%252Fa%252F%252Fb%253D/recordings/settings'))
        self.assertTrue(r.settings_url('a/b=').endswith('a%2Fb%3D/recordings/settings'))

    def test_failed_patch_never_marks_released(self):
        rule = {'uuid': 'abc', 'mode': 'public', 'originalShare': 'publicly', 'status': 'blocked'}
        with patch.object(r, 'setting', side_effect=[{'share_recording': 'none'}, RuntimeError('failed')]):
            with self.assertRaises(RuntimeError):
                r.apply_rule(object(), rule, datetime.now(timezone.utc))
        self.assertEqual(rule['status'], 'blocked')

    def test_verified_release(self):
        rule = {'uuid': 'abc', 'mode': 'public', 'originalShare': 'publicly', 'status': 'blocked'}
        with patch.object(r, 'setting', side_effect=[{'share_recording': 'none'}, {'share_recording': 'publicly'}]) as setting:
            r.apply_rule(object(), rule, datetime.now(timezone.utc))
        self.assertEqual(rule['status'], 'released')
        self.assertEqual(setting.call_args.args[2], 'publicly')

if __name__ == '__main__':
    unittest.main()
