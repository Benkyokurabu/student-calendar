import unittest
from datetime import datetime, timedelta
from unittest.mock import patch

import zoom_recording_urls as z
import zoom_recording_url_list as url_list


def lesson(start="18:35", campus="hon", room="1"):
    return {"date": "2026-08-08", "time": f"{start}～20:05", "campus": campus, "room": room}


def recording(offset_minutes, duration_minutes=90, topic="本校 第1教室"):
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
        self.assertIsNone(self.match(recording(0, 90, "本校 第1教室 保護者面談")))

    def test_rejects_short_recording(self):
        self.assertIsNone(self.match(recording(0, 25)))

    def test_rejects_too_early_start(self):
        self.assertIsNone(self.match(recording(-31, 121)))

    def test_rejects_too_late_start(self):
        self.assertIsNone(self.match(recording(31, 90)))

    def test_accepts_fifty_eight_minute_partial_lesson_recording(self):
        self.assertIsNotNone(self.match(recording(24, 58)))

    def test_rejects_recording_from_other_campus_room(self):
        self.assertIsNone(self.match(recording(0, 90, "南校 第2教室")))

    def test_rejects_multiple_plausible_recordings(self):
        self.assertIsNone(z.match_recording(lesson(), [recording(0), recording(1)], 30, 30))

    def test_supplement_is_not_a_relevant_event(self):
        events = [{
            **lesson(),
            "label": "数学 補講①",
            "groupKey": "hon_special_数学 補講①",
        }]
        self.assertEqual([], z.relevant_events(events, "2026-08", {"hon": {"1": "123"}}))

    def test_overlapping_lessons_in_same_physical_slot_are_excluded(self):
        events = [
            {**lesson(), "groupKey": "hon_j3_A_eng"},
            {**lesson(), "groupKey": "hon_j3_B_eng"},
        ]
        self.assertEqual([], z.relevant_events(events, "2026-08", {"hon": {"1": "123"}}))

    def test_flatten_uses_actual_video_file_times(self):
        payload = {
            "start_time": "2026-08-08T09:30:00Z",
            "duration": 20,
            "topic": "本校 第1教室",
            "recording_files": [{
                "recording_type": "active_speaker",
                "status": "completed",
                "recording_start": "2026-08-08T09:59:00Z",
                "recording_end": "2026-08-08T10:57:00Z",
                "play_url": "https://example.test/file",
            }],
        }
        candidate = z.flatten_recordings("1", payload)[0]
        self.assertEqual("18:59", candidate.start_time.strftime("%H:%M"))
        self.assertEqual("19:57", candidate.end_time.strftime("%H:%M"))

    @patch.object(url_list.z, "ZoomClient")
    @patch.object(url_list.z, "fetch_recordings_for_events")
    @patch.object(url_list.z, "load_schedule")
    @patch.object(url_list.z, "load_meeting_ids")
    def test_does_not_copy_recording_between_campuses(
        self, load_meeting_ids, load_schedule, fetch_recordings, zoom_client
    ):
        common = {
            "date": "2026-08-08",
            "time": "6:35～8:05",
            "grade": "j3",
            "class": "B",
            "subject": "eng",
            "faceToFace": False,
        }
        hon = {**common, "campus": "hon", "room": "1", "groupKey": "hon_j3_B_eng"}
        minami = {**common, "campus": "minami", "room": "2", "groupKey": "minami_j3_B_eng"}
        load_meeting_ids.return_value = {"hon": {"1": "111"}, "minami": {"2": "222"}}
        load_schedule.return_value = [hon, minami]
        fetch_recordings.return_value = {"111": [], "222": [recording(0, 90, "南校 第2教室")]}

        payload = url_list.make_recording_json("2026-08")

        self.assertNotIn(url_list.event_key(hon), payload["entries"])
        self.assertIn(url_list.event_key(minami), payload["entries"])

    def test_rejects_insufficient_lesson_overlap(self):
        self.assertIsNone(self.match(recording(15, 10)))

    def test_rejects_unknown_end(self):
        candidate = recording(0, 90)
        candidate.end_time = None
        self.assertIsNone(self.match(candidate))


if __name__ == "__main__":
    unittest.main()
