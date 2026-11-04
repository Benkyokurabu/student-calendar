import unittest
from unittest.mock import patch
import zoom_recording_urls as zoom
import zoom_recording_url_list as generator


class RecordingCompletionTests(unittest.TestCase):
    def meeting(self, files=None, include_files=True):
        value={'start_time':'2026-11-04T07:55:00Z','duration':80,'topic':'南校 第1教室',
               'share_url':'https://example.test/legacy'}
        if include_files:
            value['recording_files']=files
        return value

    def video(self, **extra):
        return {'status':'completed','file_type':'MP4','recording_type':'shared_screen_with_speaker_view',
                'play_url':'https://example.test/video','recording_start':'2026-11-04T07:55:00Z',
                'recording_end':'2026-11-04T09:15:00Z',**extra}

    def test_completed_playable_video_proves_completion(self):
        candidates=zoom.flatten_recordings('123',self.meeting([self.video()]))
        self.assertEqual(len(candidates),1)
        self.assertTrue(candidates[0].recording_completed)
        self.assertEqual(candidates[0].url,'https://example.test/video')

    def test_processing_or_unplayable_video_does_not_fall_back_to_share_link(self):
        for files in [[self.video(status='processing')],[self.video(play_url='')],[],None]:
            with self.subTest(files=files):
                meeting=self.meeting(files)
                self.assertEqual(zoom.flatten_recordings('123',meeting),[])
                self.assertEqual(zoom.recording_url(meeting),'')

    def test_audio_transcript_or_unknown_file_is_not_a_completed_video(self):
        for extra in [{'file_type':'M4A','recording_type':'audio_only'},
                      {'file_type':'MP3','recording_type':'audio_only'},
                      {'file_type':'TXT','recording_type':'audio_transcript'},
                      {'file_type':'','recording_type':'unknown'}]:
            with self.subTest(extra=extra):
                self.assertEqual(zoom.flatten_recordings('123',self.meeting([self.video(**extra)])),[])

    def test_old_share_link_remains_usable_without_claiming_completed_video(self):
        candidates=zoom.flatten_recordings('123',self.meeting(include_files=False))
        self.assertEqual(len(candidates),1)
        self.assertFalse(candidates[0].recording_completed)
        self.assertEqual(candidates[0].url,'https://example.test/legacy')

    def test_old_known_video_type_without_file_type_still_proves_video(self):
        video=self.video();video.pop('file_type')
        self.assertTrue(zoom.flatten_recordings('123',self.meeting([video]))[0].recording_completed)

    def test_independent_generator_sets_evidence_without_opening_or_writing_excel(self):
        event={'date':'2026-11-04','time':'4:55～6:15','campus':'minami','room':'1',
               'grade':'e6','class':'A','subject':'arith','groupKey':'minami_e6_A_arith','teacher':'確認用職員'}
        candidate=zoom.flatten_recordings('123',self.meeting([self.video()]))[0]
        with patch.object(zoom,'ZoomClient'),patch.object(zoom,'load_meeting_ids',return_value={'minami':{'1':'123'}}), \
             patch.object(zoom,'load_schedule',return_value=[event]), \
             patch.object(zoom,'fetch_recordings_for_events',return_value={'123':[candidate]}), \
             patch.object(zoom,'WorkbookWriter',side_effect=AssertionError('Excel writer must not run')), \
             patch.object(zoom.openpyxl,'load_workbook',side_effect=AssertionError('Excel workbook must not be opened')):
            payload=generator.make_recording_json('2026-11')
        self.assertTrue(payload['entries'][generator.event_key(event)]['recordingCompleted'])


if __name__=='__main__':
    unittest.main()
