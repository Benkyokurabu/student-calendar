import unittest

from sync_journal_across_campus import compare_block_fields


class CompareBlockFieldsTests(unittest.TestCase):
    def test_partially_filled_destination_receives_only_missing_fields(self):
        hon = {
            "content": "現代文", "page": "12-15", "homework1": "漢字",
            "homework2": "", "report": "", "recordingUrl": "https://example.test/video",
        }
        minami = {
            "content": "", "page": "", "homework1": "", "homework2": "",
            "report": "", "recordingUrl": "https://example.test/video",
        }

        to_hon, to_minami, conflicts = compare_block_fields(hon, minami)

        self.assertEqual({}, to_hon)
        self.assertEqual({"content": "現代文", "page": "12-15", "homework1": "漢字"}, to_minami)
        self.assertEqual([], conflicts)

    def test_complementary_fields_are_copied_in_both_directions(self):
        hon = {"content": "現代文", "page": "", "homework1": "", "homework2": "", "report": "", "recordingUrl": ""}
        minami = {"content": "", "page": "20", "homework1": "", "homework2": "", "report": "", "recordingUrl": ""}

        to_hon, to_minami, conflicts = compare_block_fields(hon, minami)

        self.assertEqual({"page": "20"}, to_hon)
        self.assertEqual({"content": "現代文"}, to_minami)
        self.assertEqual([], conflicts)

    def test_different_nonempty_values_are_never_overwritten(self):
        hon = {"content": "本校の記入", "page": "", "homework1": "", "homework2": "", "report": "", "recordingUrl": ""}
        minami = {"content": "南の記入", "page": "", "homework1": "", "homework2": "", "report": "", "recordingUrl": ""}

        to_hon, to_minami, conflicts = compare_block_fields(hon, minami)

        self.assertEqual({}, to_hon)
        self.assertEqual({}, to_minami)
        self.assertEqual(["content"], conflicts)


if __name__ == "__main__":
    unittest.main()
