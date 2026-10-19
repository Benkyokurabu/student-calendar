import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from build_teacher_calendar import teacher_payload, build


class TeacherCalendarTests(unittest.TestCase):
    def row(self, month):
        return {"date": month + "-02", "time": "5:00～6:00", "campus": "hon", "room": "1", "groupKey": "hon_j1_S_math", "teacher": "工藤", "label": "中1S 数学"}

    def test_same_canonical_event_and_labels(self):
        payload = teacher_payload([self.row("2026-11")], "2026-11")
        self.assertEqual(payload["items"][0]["room"], "1")
        self.assertEqual(payload["items"][0]["campusLabel"], "本校")
        self.assertEqual(payload["teachers"], ["工藤"])

    def test_wrong_month_empty_duplicate_stop(self):
        for rows in ([], [self.row("2026-10")], [self.row("2026-11")] * 2):
            with self.assertRaises(ValueError):
                teacher_payload(rows, "2026-11")

    def test_next_month_does_not_replace_current_latest(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            for month in ("2026-10", "2026-11"):
                (root / f"schedule_{month}.json").write_text(json.dumps([self.row(month)]), "utf-8")
            build(root, datetime(2026, 10, 19))
            latest = json.loads((root / "teacher_schedule_latest.json").read_text("utf-8"))
            self.assertEqual(latest["month"], 10)
            self.assertTrue((root / "teacher_schedule_2026-11.json").exists())


if __name__ == "__main__":
    unittest.main()
