import unittest
from worksheet_order import NS,ORDER,assert_worksheet_order,insert_data_validations
class WorksheetOrderTest(unittest.TestCase):
 def wrap(self,body):return '<worksheet xmlns="'+NS+'">'+body+'</worksheet>'
 def test_previous_corruption_is_rejected(self):
  with self.assertRaises(ValueError):assert_worksheet_order(self.wrap('<sheetData/><hyperlinks/><dataValidations/><pageMargins/>'))
 def test_validation_precedes_every_later_element(self):
  for name in ORDER[ORDER.index('dataValidations')+1:]:
   with self.subTest(name=name):
    result=insert_data_validations(self.wrap('<sheetData/><'+name+'/>'),'<dataValidations/>')
    self.assertLess(result.index('<dataValidations'),result.index('<'+name))
 def test_prior_elements_and_no_later_element(self):
  result=insert_data_validations(self.wrap('<sheetData/><mergeCells/><conditionalFormatting/>'),'<dataValidations/>')
  self.assertIn('<conditionalFormatting/><dataValidations/></worksheet>',result)
 def test_duplicate_rejected(self):
  with self.assertRaises(ValueError):insert_data_validations(self.wrap('<sheetData/><dataValidations/>'),'<dataValidations/>')
if __name__=='__main__':unittest.main()
