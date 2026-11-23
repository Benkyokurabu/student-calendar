import unittest
from types import SimpleNamespace
import openpyxl
from export_by_grade_subject import _build_merged_slots, _GRAY_PLACEHOLDER, patch_counter_formulas
from journal_class_alignment import align_new_month, previous_class_sessions

class ClassAlignmentTests(unittest.TestCase):
    def fixture(self,ahead='S',gap=1,special=False):
        wb=openpyxl.Workbook();prev=wb.active;prev.title='2026-11'
        prev['F2']=43;prev['HD2']=43
        for index in range(1+gap):
            left=2+10*index
            for klass,top in [('S',6),('A',26),('B',46)]:
                if index and klass!=ahead:continue
                prev.cell(top+1,left,klass);prev.cell(top+5,left,25+index)
        patch_counter_formulas(prev)
        if special:
            left=2+10*(1+gap);prev.cell(2,left+4,'特');prev.cell(7,left,'特');prev.cell(11,left,30)
        out=wb.create_sheet('2026-12');out['F2']=100;out['HD2']=100
        classes={k:[SimpleNamespace(special=False,day=2,time='18:35'),SimpleNamespace(special=False,day=9,time='18:35')] for k in ['S','A','B']}
        return wb,prev,out,classes

    def test_s_leads_and_old_month_remains(self):
        wb,prev,out,classes=self.fixture()
        before={(r,c):v.value for (r,c),v in prev._cells.items()}
        offsets=align_new_month(out,12,classes)
        self.assertEqual(offsets,{'S':1,'A':0,'B':0});self.assertEqual(out['F2'].value,44)
        rows=_build_merged_slots(classes,['S','A','B'],offsets)
        self.assertIs(rows[0]['S'],_GRAY_PLACEHOLDER);self.assertEqual(rows[1]['S'].day,2)
        self.assertEqual(before,{(r,c):v.value for (r,c),v in prev._cells.items()})

    def test_a_leads(self):
        wb,prev,out,classes=self.fixture('A')
        self.assertEqual(align_new_month(out,12,classes),{'S':0,'A':1,'B':0})

    def test_multiple_sessions_and_special_not_counted(self):
        wb,prev,out,classes=self.fixture('S',gap=2,special=True)
        self.assertEqual(previous_class_sessions(prev),{'S':45,'A':43,'B':43})
        self.assertEqual(align_new_month(out,12,classes),{'S':2,'A':0,'B':0})

    def test_conflicting_marker_rejected(self):
        wb,prev,out,classes=self.fixture()
        with self.assertRaises(ValueError):align_new_month(out,12,classes,{'A':1})

    def test_missing_class_rejected(self):
        wb,prev,out,classes=self.fixture()
        prev['B31']=None
        with self.assertRaises(ValueError):align_new_month(out,12,classes)

if __name__=='__main__':unittest.main()
