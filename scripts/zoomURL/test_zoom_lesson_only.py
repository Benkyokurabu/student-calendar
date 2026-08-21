import json
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import patch

import zoom_recording_urls as z
import zoom_recording_url_list as url_list


REPO_ROOT = Path(__file__).resolve().parents[2]


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

    def test_flatten_keeps_recording_restart_segments_separate(self):
        payload = {
            "start_time": "2026-08-11T05:55:32Z",
            "duration": 163,
            "topic": "本校 第1教室",
            "recording_files": [
                {
                    "id": "audio-2",
                    "recording_type": "audio_only",
                    "status": "completed",
                    "recording_start": "2026-08-11T07:50:26Z",
                    "recording_end": "2026-08-11T09:12:49Z",
                    "play_url": "https://example.test/audio-2",
                },
                {
                    "id": "video-1",
                    "recording_type": "shared_screen_with_speaker_view",
                    "status": "completed",
                    "recording_start": "2026-08-11T06:01:12Z",
                    "recording_end": "2026-08-11T07:22:06Z",
                    "play_url": "https://example.test/video-1",
                },
                {
                    "id": "video-2",
                    "recording_type": "shared_screen_with_speaker_view",
                    "status": "completed",
                    "recording_start": "2026-08-11T07:50:26Z",
                    "recording_end": "2026-08-11T09:12:49Z",
                    "play_url": "https://example.test/video-2",
                },
                {
                    "id": "audio-1",
                    "recording_type": "audio_only",
                    "status": "completed",
                    "recording_start": "2026-08-11T06:01:12Z",
                    "recording_end": "2026-08-11T07:22:06Z",
                    "play_url": "https://example.test/audio-1",
                },
            ],
        }

        candidates = z.flatten_recordings("1", payload)

        self.assertEqual(2, len(candidates))
        self.assertEqual(["15:01", "16:50"], [c.start_time.strftime("%H:%M") for c in candidates])
        self.assertEqual(
            ["https://example.test/video-1", "https://example.test/video-2"],
            [c.url for c in candidates],
        )

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

    def test_schedule_pages_require_same_time_for_group_fallback(self):
        unsafe_patterns = {
            "lesson_prep.html": 'key.startsWith(ev.date + "|") && key.includes("|" + gk + "|")',
            "calendar.html": 'key.startsWith(`${it.date}|`) && key.includes(`|${it.groupKey}|`)',
        }
        for filename, unsafe_pattern in unsafe_patterns.items():
            page = (REPO_ROOT / filename).read_text(encoding="utf-8")
            self.assertNotIn(unsafe_pattern, page, filename)
            self.assertIn("eventKeyMatchesLessonGroup", page, filename)
            self.assertIn("sameLessonTime(parts[1]", page, filename)


class ScheduleMonthSelectionTests(unittest.TestCase):
    class FixedDateTime(datetime):
        @classmethod
        def now(cls, tz=None):
            return cls(2026, 8, 21, 17, 0, tzinfo=tz)

    def test_current_month_wins_when_latest_schedule_is_next_month(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            scripts = root / "scripts"
            scripts.mkdir()
            (root / "schedule_2026-08.json").write_text("[]", encoding="utf-8")
            (root / "schedule_latest.json").write_text(
                json.dumps([{"date": "2026-09-07"}]), encoding="utf-8"
            )
            with patch.object(z, "SYSTEM_DIR", scripts), patch.object(z, "datetime", self.FixedDateTime):
                self.assertEqual("2026-08", z.determine_latest_schedule_month())

    def test_missing_requested_month_never_falls_back_to_latest(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            scripts = root / "scripts"
            scripts.mkdir()
            (root / "schedule_latest.json").write_text(
                json.dumps([{"date": "2026-09-07"}]), encoding="utf-8"
            )
            with patch.object(z, "SYSTEM_DIR", scripts):
                with self.assertRaises(FileNotFoundError):
                    z.load_schedule("2026-08")


if __name__ == "__main__":
    unittest.main()
