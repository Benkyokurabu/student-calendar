"""Transplant only authored validation/fill changes, preserving original ZIP parts and cell XML."""
from worksheet_order import insert_data_validations,assert_worksheet_order
import sys,pathlib,tempfile,zipfile,xml.etree.ElementTree as E,json,re,copy,hashlib
plan_file=pathlib.Path(sys.argv[1]).resolve()
root=plan_file.parent
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';N={'s':NS};E.register_namespace('',NS)
def tag(t):return '{'+NS+'}'+t
reports=[]
for plan_index,plan in enumerate(json.loads(plan_file.read_text(encoding='utf8'))):
 if len(sys.argv)>2 and plan_index!=int(sys.argv[2]):continue
 with zipfile.ZipFile(plan['input']) as original,zipfile.ZipFile(plan['artifact']) as artifact:
  patches={}; style_raw=original.read('xl/styles.xml').decode(); style=E.fromstring(style_raw);xfs=style.find('s:cellXfs',N);mapping={}
  for i,sp in enumerate(plan['sheets'],1):
   if not sp['rules'] and not sp['fills']:continue
   part=f'xl/worksheets/sheet{i}.xml';raw=original.read(part).decode();s=E.fromstring(raw)
   a=E.fromstring(artifact.read(part)); authored={d.get('sqref'):d for d in a.find('s:dataValidations',N)}
   dv=s.find('s:dataValidations',N)
   if dv is None:dv=E.Element(tag('dataValidations'))
   grouped={}
   for request in sp['rules']:
    d=authored[request['range']];formula=d.find('s:formula1',N).text
    assert formula==request['formula1'],(formula,request)
    if formula not in grouped:
     nd=copy.deepcopy(d);nd.set('sqref','');nd.set('showDropDown','0');nd.set('allowBlank','1');grouped[formula]=nd
    nd=grouped[formula];nd.set('sqref',(nd.get('sqref')+' '+request['range']).strip())
   for d in grouped.values():dv.append(d)
   dv.set('count',str(len(dv)));dvtext=E.tostring(dv,encoding='unicode')
   if re.search(r'<dataValidations\b',raw):raw=re.sub(r'<dataValidations\b[^>]*(?:/>|>.*?</dataValidations>)',lambda m:dvtext,raw,flags=re.S)
   else:raw=insert_data_validations(raw,dvtext)
   assert_worksheet_order(raw)
   cells={c.get('r'):c for c in s.findall('s:sheetData/s:row/s:c',N)}
   for coord in sp['fills']:
    if coord not in cells:continue
    oldid=int(cells[coord].get('s','0'))
    if oldid not in mapping:
     xf=copy.deepcopy(xfs[oldid]);xf.set('fillId','0')
     matches=[j for j,x in enumerate(xfs) if E.tostring(x)==E.tostring(xf)]
     if matches:mapping[oldid]=matches[0]
     else:mapping[oldid]=len(xfs);xfs.append(xf)
    pat=r'<c\b(?=[^>]*\br="'+re.escape(coord)+r'")[^>]*>'
    def change(m):
     t=m.group();t=re.sub(r'\bs="\d+"',f's="{mapping[oldid]}"',t) if ' s=' in t else t[:-1]+f' s="{mapping[oldid]}">';return t
    raw,n=re.subn(pat,change,raw);assert n==1,(coord,n)
   patches[part]=raw.encode()
   # All worksheet data and structure must be identical when only allowed controls/fills are stripped.
   after=E.fromstring(raw)
   for node in [s,after]:
    d=node.find('s:dataValidations',N)
    if d is not None:node.remove(d)
    for c in node.findall('s:sheetData/s:row/s:c',N):
     if c.get('r') in sp['fills']:c.attrib.pop('s',None)
   assert E.tostring(s)==E.tostring(after),(plan['input'],sp['name'],'unrelated sheet change')
  if mapping:
   xfs.set('count',str(len(xfs)))
   newstyles=re.sub(r'<cellXfs\b[^>]*>.*?</cellXfs>',lambda m:E.tostring(xfs,encoding='unicode'),style_raw,flags=re.S)
   patches['xl/styles.xml']=newstyles.encode()
  with zipfile.ZipFile(plan['output'],'w',zipfile.ZIP_DEFLATED) as out:
   for item in original.infolist():out.writestr(copy.copy(item),patches.get(item.filename,original.read(item.filename)))
  with zipfile.ZipFile(plan['output']) as out:
   assert original.namelist()==out.namelist()
   assert all(original.read(n)==out.read(n) for n in original.namelist() if n not in patches)
  reports.append({'input':plan['input'],'output':plan['output'],'changed_parts':list(patches),'sha256':hashlib.sha256(pathlib.Path(plan['output']).read_bytes()).hexdigest(),'unchanged_content_and_formula_xml':True})
(root/'preservation.json').write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(reports,ensure_ascii=False,indent=2))


