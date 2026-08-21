#!/usr/bin/env python3
import unittest

from openpyxl import Workbook

import export_schedule_json as export


class SpecialLessonRoomTests(unittest.TestCase):
    def test_supplement_sequence_marker_does_not_override_room_header(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.cell(row=11, column=7, value="③")

        # The label may be "数学 補講④", but the physical room is the header ③.
        self.assertEqual("3", export.special_event_room(sheet, 11, 7))


if __name__ == "__main__":
    unittest.main()
