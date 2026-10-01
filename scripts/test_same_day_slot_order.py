import unittest
from types import SimpleNamespace

from export_by_grade_subject import _build_merged_slots


def lesson(day, time, *, special):
    return SimpleNamespace(day=str(day), time=time, special=special)


class SameDaySlotOrderTests(unittest.TestCase):
    def test_early_special_precedes_later_regular(self):
        early = lesson(23, "4:45～6:15", special=True)
        late = lesson(23, "6:35～8:05", special=False)
        slots = _build_merged_slots({"B": [early, late]}, ["S", "A", "B"])
        self.assertEqual([slot["B"] for slot in slots], [early, late])
        self.assertEqual([slot["_special"] for slot in slots], [True, False])

    def test_later_special_follows_earlier_regular(self):
        early = lesson(23, "4:45～6:15", special=False)
        late = lesson(23, "6:35～8:05", special=True)
        slots = _build_merged_slots({"B": [late, early]}, ["S", "A", "B"])
        self.assertEqual([slot["B"] for slot in slots], [early, late])

    def test_other_class_regular_slot_is_kept(self):
        b_early = lesson(23, "4:45～6:15", special=True)
        b_late = lesson(23, "6:35～8:05", special=False)
        s_regular = lesson(23, "6:35～8:05", special=False)
        slots = _build_merged_slots(
            {"S": [s_regular], "B": [b_early, b_late]}, ["S", "A", "B"]
        )
        self.assertIsNone(slots[0]["S"])
        self.assertIs(slots[1]["S"], s_regular)

    def test_ambiguous_same_day_time_stops(self):
        special = lesson(23, "", special=True)
        regular = lesson(23, "6:35～8:05", special=False)
        with self.assertRaisesRegex(ValueError, "時刻順を判定できません"):
            _build_merged_slots({"B": [special, regular]}, ["B"])


if __name__ == "__main__":
    unittest.main()
