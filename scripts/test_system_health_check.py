from datetime import datetime, timedelta
import hashlib
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import system_health_check as health

class HealthMonitoringTests(unittest.TestCase):
    def test_reads_new_runtime_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            file = Path(tmp) / "log"
            file.write_text("[2026-10-05T19:14:16.123456] END exit=2\n", "utf-8")
            stamp, code = health.last_log_exit(file)
            self.assertEqual(code, 2)
            self.assertEqual(stamp.hour, 19)
            self.assertFalse(health.zoom_failure_needs_alert(stamp, code, stamp + timedelta(seconds=1)))
            self.assertTrue(health.zoom_failure_needs_alert(stamp, code, stamp - timedelta(seconds=1)))

    def test_keeps_reading_legacy_log_format(self):
        with tempfile.TemporaryDirectory() as tmp:
            file = Path(tmp) / "log"
            file.write_text("[2026/10/05 18:15:28.79] END exit=2\n", "utf-8")
            self.assertEqual(health.last_log_exit(file)[1], 2)

    def test_prefers_installed_runtime_log(self):
        with tempfile.TemporaryDirectory() as tmp:
            file = Path(tmp) / "BenkyoClub/zoom-publisher/logs/zoom_recording_json.log"
            file.parent.mkdir(parents=True)
            file.write_text("", "utf-8")
            with patch.dict(os.environ, {"LOCALAPPDATA": tmp}):
                self.assertEqual(health.zoom_log_path(), file)

    def test_detects_deleted_or_changed_program(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            code = root / "publisher.py"
            code.write_bytes(b"verified")
            manifest = {"files": {"publisher.py": hashlib.sha256(code.read_bytes()).hexdigest()}}
            (root / "runtime_integrity.json").write_text(json.dumps(manifest), "utf-8")
            self.assertFalse(health.runtime_integrity_errors(root))
            code.write_bytes(b"unexpected change")
            self.assertEqual(len(health.runtime_integrity_errors(root)), 1)
            code.unlink()
            self.assertEqual(len(health.runtime_integrity_errors(root)), 1)

    def test_journal_month_and_staleness_are_checked(self):
        now = datetime(2026, 10, 5, 19, tzinfo=health.JST)
        payload = {"month": "2026-10", "entries": {"lesson": {}}, "generatedAt": now.isoformat()}
        self.assertFalse(health.validate_journal_health(payload, "2026-10", now))
        self.assertTrue(health.validate_journal_health(payload, "2026-09", now))
        payload["generatedAt"] = (now - timedelta(hours=27)).isoformat()
        self.assertTrue(health.validate_journal_health(payload, "2026-10", now))

if __name__ == "__main__":
    unittest.main()
