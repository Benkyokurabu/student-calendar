#!/usr/bin/env python3
import json
import os
import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import check_journal_freshness as freshness
import check_journal_publish_age as publish_age


class FreshnessTests(unittest.TestCase):
    def test_backup_paths_are_ignored_at_any_depth(self):
        self.assertTrue(freshness.is_backup_path("_backup/old.xlsx"))
        self.assertTrue(freshness.is_backup_path("campus/class/_backup/old.xlsx"))
        self.assertTrue(freshness.is_backup_path("campus/退避2026/old.xlsx"))
        self.assertFalse(freshness.is_backup_path("campus/class/current.xlsx"))

    def test_old_but_matching_file_warns_without_blocking(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            local = root / "class.xlsx"
            local.write_bytes(b"xlsx")
            old = datetime.now(timezone.utc) - timedelta(hours=72)
            os.utime(local, (old.timestamp(), old.timestamp()))
            warning = root / "warning.txt"
            listing = [{"Path": "class.xlsx", "Size": 4, "ModTime": old.isoformat()}]
            argv = ["check", "--remote", "remote:path", "--local-dir", str(root),
                    "--warning-file", str(warning)]
            with patch.object(freshness, "remote_listing", return_value=listing), patch.object(sys, "argv", argv):
                self.assertEqual(freshness.main(), 0)
            self.assertIn("old remote workbook", warning.read_text(encoding="utf-8"))

    def test_size_mismatch_blocks(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            local = root / "class.xlsx"
            local.write_bytes(b"xlsx")
            now = datetime.now(timezone.utc)
            os.utime(local, (now.timestamp(), now.timestamp()))
            listing = [{"Path": "class.xlsx", "Size": 99, "ModTime": now.isoformat()}]
            argv = ["check", "--remote", "remote:path", "--local-dir", str(root)]
            with patch.object(freshness, "remote_listing", return_value=listing), patch.object(sys, "argv", argv):
                self.assertEqual(freshness.main(), 2)


class PublishAgeTests(unittest.TestCase):
    def test_old_publish_creates_warning(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            journal = root / "journal.json"
            generated = datetime.now(timezone.utc) - timedelta(hours=30)
            journal.write_text(json.dumps({"generatedAt": generated.isoformat()}), encoding="utf-8")
            warning = root / "warning.txt"
            argv = ["check", "--journal", str(journal), "--warning-file", str(warning)]
            with patch.object(sys, "argv", argv):
                self.assertEqual(publish_age.main(), 0)
            self.assertTrue(warning.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
