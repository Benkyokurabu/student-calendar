"""Validate CT_Worksheet order before publishing a surgical OOXML patch."""
import re
import xml.etree.ElementTree as ET
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ORDER='sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr sheetProtection protectedRanges scenarios autoFilter sortState dataConsolidate customSheetViews mergeCells phoneticPr conditionalFormatting dataValidations hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF picture oleObjects controls webPublishItems tableParts extLst'.split()
def assert_worksheet_order(raw):
    root=ET.fromstring(raw)
    seen=[]
    for child in root:
        if child.tag.startswith('{'+NS+'}'):
            name=child.tag.split('}')[-1]
            if name not in ORDER:raise ValueError('Unknown worksheet child: '+name)
            seen.append((ORDER.index(name),name))
    if [i for i,_ in seen]!=sorted(i for i,_ in seen):
        raise ValueError('Invalid worksheet child order: '+' '.join(n for _,n in seen))
def insert_data_validations(raw,dvtext):
    assert_worksheet_order(raw)
    if ET.fromstring(raw).find('{'+NS+'}dataValidations') is not None:
        raise ValueError('dataValidations already exists')
    later=ORDER[ORDER.index('dataValidations')+1:]
    match=re.search(r'<(?:'+ '|'.join(later)+r')\b',raw)
    position=match.start() if match else raw.index('</worksheet>')
    result=raw[:position]+dvtext+raw[position:]
    assert_worksheet_order(result)
    return result
