import unittest
from datetime import datetime, timedelta

import zoom_recording_urls as z


def lesson(start="18:35"):
    return {"date": "2026-08-08", "time": f"{start}～20:05"}


def recording(offset_minutes, duration_minutes=90, topic="授業"):
    lesson_start, _ = z.parse_lesson_window(lesson())
    start = lesson_start + timedelta(minutes=offset_minutes)
    return z.RecordingCandidate(
        meeting_id="1",
        start_time=start,
        end_time=start + timedelta(minutes=duration_minutes),
        topic=topic,
        url="https://example.test/recording",
        raw={},
    )


class LessonOnlyRecordingTests(unittest.TestCase):
    def match(self, candidate):
        return z.match_recording(lesson(), [candidate], 30, 30)

    def test_accepts_full_lesson_started_twenty_minutes_early(self):
        self.assertIsNotNone(self.match(recording(-20, 110)))

    def test_rejects_interview_topic(self):
        self.assertIsNone(self.match(recording(0, 90, "保護者面談")))

    def test_rejects_short_recording(self):
        self.assertIsNone(self.match(recording(0, 25)))

    def test_rejects_too_early_start(self):
        self.assertIsNone(self.match(recording(-31, 121)))

    def test_rejects_too_late_start(self):
        self.assertIsNone(self.match(recording(21, 90)))

    def test_rejects_insufficient_lesson_overlap(self):
        self.assertIsNone(self.match(recording(15, 10)))

    def test_rejects_unknown_end(self):
        candidate = recording(0, 90)
        candidate.end_time = None
        self.assertIsNone(self.match(candidate))


if __name__ == "__main__":
    unittest.main()
