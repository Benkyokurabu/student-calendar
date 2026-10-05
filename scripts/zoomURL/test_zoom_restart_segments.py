from datetime import datetime, timedelta
import unittest
import zoom_recording_urls as zoom

class RestartSegmentTests(unittest.TestCase):
    def recordings(self):
        start = datetime(2026, 9, 21, 20, 25, 4, tzinfo=zoom.JST)
        first = zoom.RecordingCandidate("meeting", start, datetime(2026, 9, 21, 21, 31, 23, tzinfo=zoom.JST), "本校 第1教室", "https://example.test/one", {"uuid": "same-occurrence"})
        second = zoom.RecordingCandidate("meeting", datetime(2026, 9, 21, 21, 32, 25, tzinfo=zoom.JST), datetime(2026, 9, 21, 21, 54, 42, tzinfo=zoom.JST), "本校 第1教室", "https://example.test/two", {"uuid": "same-occurrence"})
        return first, second

    def test_same_occurrence_sequential_restart_is_resolved(self):
        first, second = self.recordings()
        self.assertIs(zoom.first_restart_segment([second, first]), first)
        event = {"date": "2026-09-21", "time": "8:25～9:55", "campus": "hon", "room": "1"}
        self.assertIs(zoom.match_recording(event, [first, second], 30, 75), first)

    def test_distinct_occurrences_remain_ambiguous(self):
        first, second = self.recordings()
        second.raw = {"uuid": "other-occurrence"}
        self.assertIsNone(zoom.first_restart_segment([first, second]))

    def test_missing_occurrence_identity_remains_ambiguous(self):
        first, second = self.recordings()
        first.raw = {}; second.raw = {}
        self.assertIsNone(zoom.first_restart_segment([first, second]))

    def test_overlapping_variants_remain_ambiguous(self):
        first, second = self.recordings()
        second.start_time = first.start_time + timedelta(minutes=5)
        self.assertIsNone(zoom.first_restart_segment([first, second]))

if __name__ == "__main__":
    unittest.main()
