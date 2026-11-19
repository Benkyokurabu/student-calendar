"""Real local Git proof that legacy main pushes cannot alter G10 publication."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from publication_branch import publication_branch, require_publication_branch, G10_BRANCH

sys.path.insert(0, str(Path(__file__).parent / 'zoomURL'))
import publish_zoom_recording_url_json as publisher


class PublicationBranchTests(unittest.TestCase):
    def test_legacy_default_and_invalid_targets(self):
        with patch.dict(os.environ, {}, clear=True):
            self.assertEqual(publication_branch(), 'main')
        for value in ('other', 'refs/heads/main', 'main:other', '../g10'):
            with patch.dict(os.environ, {'BENTAN_PUBLICATION_BRANCH': value}):
                with self.assertRaises(ValueError):
                    publication_branch()

    def test_independent_publication_and_arbitrary_legacy_pending_commit(self):
        with tempfile.TemporaryDirectory(prefix='bentan-branch-proof-') as temp:
            root=Path(temp)
            def git(cwd, *args):
                return subprocess.run(['git', *args], cwd=cwd, check=True,
                                      capture_output=True, text=True).stdout.strip()
            git(root, 'init', '--bare', 'remote.git')
            git(root, 'init', '--initial-branch=main', 'seed')
            seed=root/'seed'
            git(seed, 'config', 'user.name', 'Isolated fixture')
            git(seed, 'config', 'user.email', 'fixture@example.invalid')
            git(seed, 'remote', 'add', 'origin', str(root/'remote.git'))
            schedule=seed/'schedule_2026-11.json'
            schedule.write_text(json.dumps([{'date':'2026-11-20','source':'legacy'}]))
            (seed/'calendar.html').write_text('legacy static fixture')
            git(seed, 'add', '.')
            git(seed, 'commit', '-m', 'legacy fixture')
            git(seed, 'push', '-u', 'origin', 'main')
            legacy_head=git(seed, 'rev-parse', 'HEAD')
            git(seed, 'switch', '-c', G10_BRANCH)
            schedule.write_text(json.dumps([{'date':'2026-11-20','source':'g10'}]))
            (seed/'calendar.html').write_text('reviewed G10 static fixture')
            git(seed, 'add', '.')
            git(seed, 'commit', '-m', 'reviewed publication fixture')
            git(seed, 'push', '-u', 'origin', G10_BRANCH)
            git(root, 'clone', '--branch', G10_BRANCH, str(root/'remote.git'), 'publication')
            public=root/'publication'
            git(public, 'config', 'user.name', 'Isolated publisher')
            git(public, 'config', 'user.email', 'publisher@example.invalid')
            with patch.dict(os.environ, {'BENTAN_PUBLICATION_BRANCH':G10_BRANCH}):
                require_publication_branch(public)
                self.assertEqual(publisher.load_publishing_schedule(public, '2026-11')[0]['source'], 'g10')
                (public/'zoom_recording_urls_2026-11.json').write_text('{"fixture":true}')
                git(public, 'add', 'zoom_recording_urls_2026-11.json')
                git(public, 'commit', '-m', 'fixture publication output')
                self.assertTrue(publisher.push_pending_commits(public, '{}', '2026-11'))
            self.assertEqual(git(root/'remote.git', 'rev-parse', 'main'), legacy_head)
            public_tree=git(root/'remote.git', 'rev-parse', G10_BRANCH+'^{tree}')
            git(seed, 'switch', 'main')
            with patch.dict(os.environ, {'BENTAN_PUBLICATION_BRANCH':G10_BRANCH}):
                with self.assertRaises(ValueError):
                    require_publication_branch(seed)
            # A legacy pending commit may contain more than its normal JSON.
            (seed/'calendar.html').write_text('old pending control file')
            (seed/'zoom_recording_urls_2026-11.json').write_text('{"old":true}')
            git(seed, 'add', '.')
            git(seed, 'commit', '-m', 'arbitrary legacy pending main commit')
            git(seed, 'push', 'origin', 'main')
            self.assertEqual(git(root/'remote.git', 'rev-parse', G10_BRANCH+'^{tree}'), public_tree)


if __name__=='__main__':
    unittest.main()
