import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import publish_zoom_recording_url_json as publisher
import recording_publication_filter as policy

class HistoricalPublicationTests(unittest.TestCase):
    def test_previous_month_does_not_replace_current_latest(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp); output = root / "output"; repo = root / "repo"
            output.mkdir(); repo.mkdir()
            sentinel = json.dumps({"month": "2026-10", "entries": {"current": {"url": "https://example.test/current"}}})
            (output / "zoom_recording_urls_latest.json").write_text(sentinel, "utf-8")
            (repo / "zoom_recording_urls_latest.json").write_text(sentinel, "utf-8")
            payload = {"month": "2026-09", "matched": 1, "missing": 0, "entries": {"lesson": {"date": "2026-09-21", "url": "https://example.test/old"}}}
            def git(args, **kwargs):
                code = 1 if args[:3] == ["git", "diff", "--cached"] else 0
                return subprocess.CompletedProcess(args, code, "", "")
            with patch.object(sys, "argv", ["publisher", "--month", "2026-09"]), patch.object(publisher.url_list, "SYSTEM_DIR", output), patch.object(publisher.url_list.z, "determine_latest_schedule_month", return_value="2026-10"), patch.object(publisher.url_list, "make_recording_json", return_value=payload), patch.object(publisher.url_list, "repo_dir", return_value=repo), patch.object(publisher, "PublishLock"), patch.object(publisher, "push_pending_commits", return_value=False), patch.object(publisher, "pull_rebase_with_zoom_resolution"), patch.object(publisher, "run", side_effect=git) as calls, patch.object(policy, "load_rules", return_value=[]), patch.object(policy, "apply_payload", side_effect=lambda data, rules: data):
                self.assertEqual(publisher.main(), 0)
            self.assertEqual((output / "zoom_recording_urls_latest.json").read_text("utf-8"), sentinel)
            self.assertEqual((repo / "zoom_recording_urls_latest.json").read_text("utf-8"), sentinel)
            self.assertTrue((repo / "zoom_recording_urls_2026-09.json").exists())
            added = next(call.args[0] for call in calls.call_args_list if call.args[0][:2] == ["git", "add"])
            self.assertNotIn("zoom_recording_urls_latest.json", added)

if __name__ == "__main__":
    unittest.main()
