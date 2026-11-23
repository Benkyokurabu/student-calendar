from pathlib import Path
from tempfile import TemporaryDirectory,gettempdir
from zipfile import ZipFile
from copy import copy
import shutil,unittest
import openpyxl
import export_by_grade_subject as ex
from journal_slots import ensure_lesson_capacity,slot_columns
from journal_class_alignment import class_session_entries
from journal_month_workbook import generate_month_workbook

SCRIPTS=Path(__file__).resolve().parent
class PreservingMonthWriterTests(unittest.TestCase):
    def fixture(self,previous='2026-11',month='2026-12'):
        temp=TemporaryDirectory();self.addCleanup(temp.cleanup);root=Path(temp.name)
        templates=root/'templates';templates.mkdir()
        for name in ('TEMPLATE_MAIN.xlsx','TEMPLATE_X.xlsx'):shutil.copy2(SCRIPTS/name,templates/name)
        book=openpyxl.load_workbook(templates/'TEMPLATE_MAIN.xlsx');hidden=book.active
        hidden.title='__TEMPLATE_MAIN__';ensure_lesson_capacity(hidden);hidden.sheet_state='hidden'
        prev=book.copy_worksheet(hidden);prev.title=previous;prev.sheet_state='visible'
        prev['F2']=43;prev['HD2']=43
        for left in slot_columns(prev):
            for row in (11,31,51):prev.cell(row,left).value=None
        for klass,top in [('S',6),('A',26),('B',46)]:prev.cell(top+1,2,klass);prev.cell(top+5,2,25)
        prev['L7']='S';prev['L11']=30;prev['HD10']='Existing human note'
        ex.patch_counter_formulas(prev)
        original=root/'before.xlsx';book.save(original);book.close()
        number=int(month[-2:]);year=month[:4]
        info={'filename':f'南教室中１数学_{year}.xlsx','grade':'中１','subject':'数学','supplementary':False,'classes':{}}
        for klass,day in [('S',7),('A',2),('B',2)]:info['classes'][klass]=[list(ex.Event(number,day,'月','18:35～20:05','①','1',klass,'数','1'+klass+'数','担当',False,False,10,3))]
        return root,templates,original,info
    def test_real_month_generation_retains_all_old_parts_and_human_counters(self):
        root,templates,original,info=self.fixture();output=root/'after.xlsx'
        result=generate_month_workbook(original,output,'2026-12',info,templates,root/'work')
        self.assertEqual(result['first']['S']['counter'],45);self.assertEqual(result['first']['A']['counter'],44);self.assertEqual(result['first']['B']['counter'],44)
        with ZipFile(original) as a,ZipFile(output) as b:
            for name in a.namelist():
                if name not in {'xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml','xl/styles.xml'}:self.assertEqual(a.read(name),b.read(name),name)
        # Save a synthetic, pupil-free workbook for native Excel acceptance on Windows.
        shutil.copy2(output,Path(gettempdir())/'bentan-month-library-native-test.xlsx')
        repeated=root/'repeat.xlsx';self.assertTrue(generate_month_workbook(output,repeated,'2026-12',info,templates,root/'repeat-work')['existing']);self.assertFalse(repeated.exists())
    def test_year_boundary_and_latest_human_override_are_read(self):
        root,templates,original,info=self.fixture('2026-12','2027-01')
        book=openpyxl.load_workbook(original);book['2026-12']['P2']=47;book.save(original);book.close()
        result=generate_month_workbook(original,root/'after.xlsx','2027-01',info,templates,root/'work')
        self.assertEqual(result['first']['S']['counter'],48);self.assertEqual(result['first']['A']['counter'],44)
        self.assertEqual(result['first']['S']['column'],42)

if __name__=='__main__':unittest.main()
