import json
import tempfile
import unittest
from pathlib import Path

from ci_extract_and_push import merge_entry_keys, require_valid_publication
from validate_journal_publish import validate


class ScheduleMergeTests(unittest.TestCase):
    def setUp(self):
        self.event = dict(date="2026-09-08", time="8:25～9:55", campus="minami",
                          groupKey="minami_j3_B_math", room="1")
        self.key = "|".join(self.event.values())
        self.entry = dict(teacher="teacher", sessionNumber="1", monthNum="9",
                          weekNum="1", content="equations", homework=["p.12"])

    def check_entries(self, entries):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            candidate = root / "journal.json"
            schedule = root / "schedule.json"
            candidate.write_text(json.dumps(dict(month="2026-09", entries=entries)), encoding="utf-8")
            schedule.write_text(json.dumps([self.event]), encoding="utf-8")
            return validate(candidate, schedule)

    def test_renamed_old_slots_do_not_block_new_content(self):
        old = {self.key: dict(self.entry, content=""), "retired-special-key": self.entry}
        new = {self.key: self.entry}
        merged = {key: new[key] for key in merge_entry_keys(old, new)}
        self.assertEqual(self.check_entries(merged), [])
        self.assertEqual(merged[self.key]["homework"], ["p.12"])
        self.assertEqual(len(old), 2)

    def test_missing_extraction_is_not_masked_by_old_entries(self):
        keys = merge_entry_keys({self.key: self.entry}, {})
        self.assertEqual(keys, [])
        self.assertTrue(self.check_entries({}))

    def test_unexpected_new_slot_is_still_rejected(self):
        new = {self.key: self.entry, "unexpected": self.entry}
        self.assertEqual(set(merge_entry_keys({}, new)), set(new))
        self.assertTrue(self.check_entries(new))

    def test_failed_validation_stops_publication(self):
        with self.assertRaises(RuntimeError):
            require_valid_publication(True)
        require_valid_publication(False)


if __name__ == "__main__":
    unittest.main()
