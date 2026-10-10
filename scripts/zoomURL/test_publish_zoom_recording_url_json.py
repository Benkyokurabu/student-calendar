#!/usr/bin/env python3
import json
import tempfile
import unittest
from pathlib import Path

import publish_zoom_recording_url_json as publisher
from recording_publication_filter import apply_payload


class ExistingPayloadTests(unittest.TestCase):
    def test_loads_valid_existing_payload(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "zoom.json"
            expected = {"month": "2026-08", "entries": {"lesson": {"url": "https://example.test"}}}
            path.write_text(json.dumps(expected), encoding="utf-8")
            self.assertEqual(expected, publisher.load_existing_payload(path))

    def test_invalid_json_is_treated_as_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "zoom.json"
            path.write_text("not-json", encoding="utf-8")
            self.assertIsNone(publisher.load_existing_payload(path))

    def test_existing_entries_cannot_silently_disappear(self):
        existing = {
            "month": "2026-08",
            "entries": {"lesson": {"date": "2026-08-18", "url": "https://example.test/old"}},
        }
        generated = {"month": "2026-08", "matched": 0, "entries": {}}
        with self.assertRaisesRegex(RuntimeError, "would remove"):
            publisher.validate_payload(generated, "2026-08", existing)

    def test_hidden_entry_keeps_empty_url_without_exposing_recording(self):
        key = "2026-10-01|lesson|hon|group|2"
        hidden = {
            "date": "2026-10-01", "url": "", "hidden": True,
            "recordingPublicationKey": key,
        }
        payload = {"month": "2026-10", "matched": 1, "entries": {key: hidden}}
        publisher.validate_payload(payload, "2026-10")
        for invalid in ({**hidden, "hidden": False}, {**hidden, "recordingPublicationKey": "other"}):
            with self.assertRaisesRegex(RuntimeError, "invalid entries"):
                publisher.validate_payload({**payload, "entries": {key: invalid}}, "2026-10")

    def test_merge_preserves_old_url_and_adds_new_entry(self):
        existing = {
            "month": "2026-08",
            "matched": 1,
            "missing": 1,
            "entries": {"old": {"date": "2026-08-03", "url": "https://example.test/stable"}},
        }
        generated = {
            "month": "2026-08",
            "matched": 2,
            "missing": 1,
            "entries": {
                "old": {"date": "2026-08-03", "url": "https://example.test/rotated"},
                "new": {"date": "2026-08-24", "url": "https://example.test/new"},
            },
        }
        merged = publisher.merge_preserving_published_entries(generated, existing)
        self.assertEqual("https://example.test/stable", merged["entries"]["old"]["url"])
        self.assertIn("new", merged["entries"])
        self.assertEqual(2, merged["matched"])
        self.assertEqual(1, merged["missing"])

    def test_public_rule_changes_hidden_entry_without_mutating_existing(self):
        key = "2026-10-09|lesson|minami|minami_e6_A_eng|2"
        existing = {
            "month": "2026-10", "matched": 1, "missing": 0,
            "entries": {key: {
                "date": "2026-10-09", "url": "", "hidden": True,
                "recordingPublicationKey": key,
            }},
        }
        generated = {"month": "2026-10", "matched": 0, "missing": 1, "entries": {}}
        rules = [{
            "key": key, "eventKeys": [key], "status": "public",
            "url": "https://example.test/recording", "urlHashes": [],
        }]

        published = apply_payload(
            publisher.merge_preserving_published_entries(generated, existing), rules,
        )

        self.assertEqual("", existing["entries"][key]["url"])
        self.assertTrue(existing["entries"][key]["hidden"])
        self.assertEqual("https://example.test/recording", published["entries"][key]["url"])
        self.assertNotIn("hidden", published["entries"][key])
        self.assertNotEqual(
            publisher.comparable_payload(existing), publisher.comparable_payload(published),
        )


if __name__ == "__main__":
    unittest.main()
