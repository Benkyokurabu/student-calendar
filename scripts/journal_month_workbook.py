"""Reusable monthly journal writer. Never replace any existing month."""
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as ET
import os,sys,json,posixpath,uuid,copy
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
REL='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'


def absolute(part,target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.dirname(part)+'/'+target)

def xml(element):return ET.tostring(element,encoding='utf-8',xml_declaration=True)

def require(condition,reason):
    if not condition:raise ValueError(str(reason))

def generate_month_workbook(source: Path, output: Path, month: str, info: dict, templates: Path, scratch: Path):
    """Add one absent month. Existing sheet parts and style definitions are preserved."""
    from datetime import date, timedelta
    import re
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])",month):raise ValueError('Invalid month')
    year,number=map(int,month.split('-'))
    previous=(date(year,number,1)-timedelta(days=1)).strftime('%Y-%m')
    relative=info['filename']
    MONTH=month
    import openpyxl
    import export_by_grade_subject as ex
    from journal_class_alignment import previous_class_sessions, class_session_entries
    from journal_slots import slot_columns
    from journal_input_controls import assert_template_controls
    data=source.read_bytes();name=info["filename"]
    scratch.mkdir(parents=True,exist_ok=True)
    with ZipFile(BytesIO(data)) as archive:parts={p:archive.read(p) for p in archive.namelist()}
    originalbook=ET.fromstring(parts['xl/workbook.xml']);sheets=originalbook.find('m:sheets',NS)
    if any(s.get('name')==MONTH for s in sheets):return {'path':relative,'existing':True}
    is_x='X' in name;hidden='__TEMPLATE_X__' if is_x else '__TEMPLATE_MAIN__'
    wanted={hidden,previous}
    refs={r.get('Id'):r.get('Target') for r in ET.fromstring(parts['xl/_rels/workbook.xml.rels'])}
    thinbook=copy.deepcopy(originalbook);thinsheets=thinbook.find('m:sheets',NS)
    for s in list(thinsheets):
        if s.get('name') not in wanted:thinsheets.remove(s)
    require(any(s.get('name')==hidden for s in thinsheets),(name,'hidden template missing'))
    thinparts=dict(parts);thinparts['xl/workbook.xml']=xml(thinbook)
    for s in thinsheets:
        if s.get('name')!=previous:continue
        part=absolute('xl/workbook.xml',refs[s.get(RID)])
        full=ET.fromstring(parts[part]);sparse=ET.Element(full.tag)
        sd=ET.SubElement(sparse,'{'+NS['m']+'}sheetData')
        for row in full.find('m:sheetData',NS):
            if int(row.get('r')) in {2,3,4,5,6,7,10,11,27,30,31,47,50,51}:sd.append(copy.deepcopy(row))
        thinparts[part]=xml(sparse)
    thin=scratch/'thin.xlsx'
    with ZipFile(thin,'w',ZIP_DEFLATED) as z:
        for p,b in thinparts.items():z.writestr(p,b)
    wb=openpyxl.load_workbook(thin)
    campus='南教室' if name.startswith('南教室') else '本校'
    classes={k:[ex.Event(*values) for values in entries] for k,entries in info['classes'].items()}
    template=templates/('TEMPLATE_X.xlsx' if is_x else 'TEMPLATE_MAIN.xlsx')
    prior=previous_class_sessions(wb[previous]) if previous in wb else {}
    tmpl=ex.ensure_hidden_template_sheet(wb,template,hidden)
    sheet=ex.create_month_sheet(wb,tmpl,year,number,template_path=template)
    ex.set_header_cells(sheet,campus,info['grade'],info['subject'],number,wb=wb,year=year)
    if is_x:ex.fill_sheet_x(sheet,number,classes['X'],teacher_blank=False)
    else:ex.fill_sheet_main(sheet,number,classes,teacher_blank=info['supplementary'])
    require(len(slot_columns(sheet))==21,'Unexpected lesson capacity')
    assert_template_controls(sheet)
    # Continue each class from the actual previous month, including manual overrides.
    first={k:entries[0] for k,entries in class_session_entries(sheet).items() if entries}
    if not is_x and len([k for k in first if k in prior])>=2:
        if any(first[k]['counter']!=prior[k]+1 for k in first if k in prior):
            raise ValueError('Class annual session continuation was not preserved')
    new=scratch/'generated.xlsx';wb.save(new);wb.close()
    with ZipFile(new) as z:generated={p:z.read(p) for p in z.namelist()}
    from openpyxl.styles.stylesheet import Stylesheet
    oldstyles=Stylesheet.from_tree(ET.fromstring(parts['xl/styles.xml']))
    newstyles=Stylesheet.from_tree(ET.fromstring(generated['xl/styles.xml']))
    for attr in ('fonts','fills','borders','dxfs','tableStyles'):
        before=getattr(oldstyles,attr);after=getattr(newstyles,attr)
        if isinstance(before,list):require(before==after[:len(before)],(name,'style prefix',attr))
        else:require(before==after,(name,'style property',attr))
    for attr,sub in [('cellXfs','xf'),('cellStyleXfs','xf'),('numFmts','numFmt')]:
        before=getattr(getattr(oldstyles,attr),sub);after=getattr(getattr(newstyles,attr),sub)
        if sub=='xf':
            for item in before+after:
                item.pivotButton=bool(item.pivotButton);item.quotePrefix=bool(item.quotePrefix)
        require(before==after[:len(before)],(name,'style prefix',attr))
    require(sorted(oldstyles.cellStyles.cellStyle,key=lambda s:s.name)==sorted(newstyles.cellStyles.cellStyle,key=lambda s:s.name),'Named styles changed')
    # Retain original style definitions, appending only genuinely new indexed styles.
    mergedstyles=ET.fromstring(parts['xl/styles.xml']);generatedstyles=ET.fromstring(generated['xl/styles.xml'])
    for local in ('fonts','fills','borders','cellXfs','cellStyleXfs','numFmts','dxfs'):
        oldnode=mergedstyles.find('m:'+local,NS);newnode=generatedstyles.find('m:'+local,NS)
        if newnode is None:continue
        if oldnode is None:
            if len(newnode):mergedstyles.append(copy.deepcopy(newnode))
            continue
        for item in list(newnode)[len(oldnode):]:oldnode.append(copy.deepcopy(item))
        if 'count' in oldnode.attrib:oldnode.set('count',str(len(oldnode)))
    genbook=ET.fromstring(generated['xl/workbook.xml'])
    gensheet=next(s for s in genbook.find('m:sheets',NS) if s.get('name')==MONTH)
    genrels={r.get('Id'):r for r in ET.fromstring(generated['xl/_rels/workbook.xml.rels'])}
    genpart=absolute('xl/workbook.xml',genrels[gensheet.get(RID)].get('Target'))
    require(not any(c.get('t')=='s' for c in ET.fromstring(generated[genpart]).findall('.//m:c',NS)),(name,'new shared strings unsupported'))
    originalrels=ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    relids={r.get('Id') for r in originalrels};newid='rIdBentanMonth'+month.replace('-','')
    require(newid not in relids,'Month relationship already exists')
    newpart='xl/worksheets/bentan_month_'+month.replace('-','')+'.xml';require(newpart not in parts,'Month part already exists')
    added=copy.deepcopy(gensheet);added.set(RID,newid);added.set('sheetId',str(max(int(s.get('sheetId')) for s in sheets)+1));sheets.append(added)
    ET.SubElement(originalrels,'{'+REL+'}Relationship',{'Id':newid,'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet','Target':'/'+newpart})
    types=ET.fromstring(parts['[Content_Types].xml']);gentypes=ET.fromstring(generated['[Content_Types].xml'])
    types_by_part={t.get('PartName','').lstrip('/'):t.get('ContentType') for t in gentypes if t.tag.endswith('Override')}
    existing_defaults={t.get('Extension') for t in types if t.tag.endswith('Default')}
    for t in gentypes:
        if t.tag.endswith('Default') and t.get('Extension') not in existing_defaults:types.append(copy.deepcopy(t))
    additions={};visited={genpart:newpart}
    def import_part(oldpart,target):
        additions[target]=generated[oldpart]
        if oldpart in types_by_part:ET.SubElement(types,'{'+CT+'}Override',{'PartName':'/'+target,'ContentType':types_by_part[oldpart]})
        oldrel=posixpath.dirname(oldpart)+'/_rels/'+posixpath.basename(oldpart)+'.rels'
        if oldrel not in generated:return
        trel=ET.fromstring(generated[oldrel])
        for relation in trel:
            if relation.get('TargetMode')=='External':continue
            part=absolute(oldpart,relation.get('Target'))
            if part not in visited:
                dest='xl/bentan_month_'+month.replace('-','')+'/'+uuid.uuid4().hex+'_'+posixpath.basename(part);visited[part]=dest;import_part(part,dest)
            relation.set('Target','/'+visited[part])
        additions[posixpath.dirname(target)+'/_rels/'+posixpath.basename(target)+'.rels']=xml(trel)
    import_part(genpart,newpart)
    names=originalbook.find('m:definedNames',NS)
    gennames=genbook.find('m:definedNames',NS)
    genindex=list(genbook.find('m:sheets',NS)).index(gensheet)
    if gennames is not None:
        for definition in gennames:
            if definition.get('localSheetId')==str(genindex):
                if names is None:
                    names=ET.Element('{'+NS['m']+'}definedNames')
                    after=originalbook.find('m:calcPr',NS)
                    if after is None:after=originalbook.find('m:extLst',NS)
                    if after is None:originalbook.append(names)
                    else:originalbook.insert(list(originalbook).index(after),names)
                item=copy.deepcopy(definition);item.set('localSheetId',str(len(sheets)-1));names.append(item)
    calc=originalbook.find('m:calcPr',NS)
    if calc is not None:calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
    output.parent.mkdir(parents=True,exist_ok=True)
    replacements={'xl/workbook.xml':xml(originalbook),'xl/_rels/workbook.xml.rels':xml(originalrels),
        '[Content_Types].xml':xml(types),'xl/styles.xml':xml(mergedstyles)}
    with ZipFile(output,'w',ZIP_DEFLATED) as z:
        for p,b in parts.items():z.writestr(p,replacements.get(p,b))
        for p,b in additions.items():require(p not in parts,'Package part collision');z.writestr(p,b)
    with ZipFile(output) as z:
        for p,b in parts.items():
            if p not in replacements:require(z.read(p)==b,(name,'historic XML changed',p))
    result={'path':relative,'existing':False,'prior':prior,'first':first,'oldPartsUnchanged':len(parts)-4,
        'output':str(output),'sha256':__import__('hashlib').sha256(output.read_bytes()).hexdigest()}
    (scratch/'result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return result
