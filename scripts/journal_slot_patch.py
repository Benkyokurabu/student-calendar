"""Append journal forms while retaining original OOXML cells, caches and unrelated parts."""
from copy import deepcopy,copy
from pathlib import Path
import re,zipfile,xml.etree.ElementTree as E,json
import openpyxl
from openpyxl.utils import coordinate_to_tuple,column_index_from_string
from journal_slots import ensure_lesson_capacity
from journal_input_controls import lesson_columns
from worksheet_order import assert_worksheet_order,ORDER
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';N={'s':NS};E.register_namespace('',NS)
def xml(e):return E.tostring(e,encoding='unicode')
def tag(name):return '{'+NS+'}'+name
def replace_section(raw,name,element):
    pattern=r'<'+name+r'\b[^>]*(?:/>|>.*?</'+name+r'>)'
    text=xml(element)
    if re.search(pattern,raw,re.S):return re.sub(pattern,lambda m:text,raw,count=1,flags=re.S)
    later=ORDER[ORDER.index(name)+1:]
    match=re.search(r'<(?:'+'|'.join(later)+r')\b',raw)
    p=match.start() if match else raw.index('</worksheet>')
    return raw[:p]+text+raw[p:]
def sheet_map(z):
    wb=E.fromstring(z.read('xl/workbook.xml'))
    rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    out={}
    for s in wb.find('s:sheets',N):
        t=rels[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
        out[s.get('name')]=t.lstrip('/') if t.startswith('/') else 'xl/'+t
    return out
def extend_file(source,output,*,canonical=False,month='2026-09'):
    source=Path(source);output=Path(output)
    w=openpyxl.load_workbook(source)
    targets=[s for s in w if canonical or 'TEMPLATE' in s.title or (re.fullmatch(r'\d{4}-\d{2}',s.title) and s.title>=month)]
    plans=[]
    base=openpyxl.load_workbook(source)
    donor=next((s for s in base if 'TEMPLATE' in s.title),base.worksheets[0])
    for s in targets:
        old=len(lesson_columns(s));before_f={c.coordinate:c.value for row in s for c in row if c.data_type=='f'}
        ensure_lesson_capacity(s,donor if not canonical else None)
        changed_formulas={a:(v,s[a].value) for a,v in before_f.items() if s[a].value!=v}
        for a,(before,after) in changed_formulas.items():
            assert re.sub(r'\$FG\$([23])\b',r'$HD$\1',before)==after,(s.title,a)
        plans.append({'name':s.title,'old':old,'changed_formulas':changed_formulas})
    candidate=output.with_suffix('.candidate.xlsx');w.save(candidate)
    # openpyxl is used by the existing generator; never use its whole-file rewrite in production.
    for attr in ['_fonts','_fills','_borders','_alignments','_protections','_number_formats','_cell_styles']:
        old_table=getattr(base,attr);new_table=getattr(w,attr)
        assert list(new_table[:len(old_table)])==list(old_table),(source.name,attr,'existing style changed')
    base.close();w.close()
    with zipfile.ZipFile(source) as original,zipfile.ZipFile(candidate) as authored:
        original_map=sheet_map(original);new_map=sheet_map(authored);patches={}
        for plan in plans:
            part=original_map[plan['name']];raw=original.read(part).decode();original_sheet=E.fromstring(raw)
            new=E.fromstring(authored.read(new_map[plan['name']]))
            old_end=plan['old']*10
            old_cells={c.get('r'):c for c in original_sheet.findall('s:sheetData/s:row/s:c',N)}
            new_cells={c.get('r'):c for c in new.findall('s:sheetData/s:row/s:c',N)}
            def allowed(a):
                row,col=coordinate_to_tuple(a)
                return old_end<col<=210 or a in ['HD2','HD3']
            for a,old_cell in old_cells.items():
                if allowed(a) and a not in ['FG2','FG3','HD2','HD3'] and a in new_cells:
                    if old_cell.find('s:f',N) is not None or ''.join(old_cell.itertext()).strip():
                        kept=deepcopy(old_cell)
                        if new_cells[a].get('s') is not None:kept.set('s',new_cells[a].get('s'))
                        if a in plan['changed_formulas']:kept.find('s:f',N).text=plan['changed_formulas'][a][1][1:]
                        new_cells[a].clear();new_cells[a].attrib.update(kept.attrib);new_cells[a][:]=list(kept)
            # Keep old cell XML verbatim except the necessary equivalent helper references.
            rows={int(r.get('r')):r for r in new.findall('s:sheetData/s:row',N)}
            data_match=re.search(r'<sheetData\b[^>]*>(.*?)</sheetData>',raw,re.S);assert data_match
            prior_rows={int(re.search(r'\br="(\d+)"',m.group()).group(1)):m.group() for m in re.finditer(r'<row\b[^>]*?(?:/>|>.*?</row>)',data_match.group(1),re.S)}
            row_text=[]
            for row_n in sorted(set(prior_rows)|set(rows)):
                old_row=prior_rows.get(row_n)
                if old_row is None:
                    nr=deepcopy(rows[row_n]);nr[:]=[c for c in nr if c.tag==tag('c') and allowed(c.get('r'))]
                    if len(nr):row_text.append(xml(nr))
                    continue
                cells={}
                for match in re.finditer(r'<c\b[^>]*?(?:/>|>.*?</c>)',old_row,re.S):
                    text=match.group();a=re.search(r'\br="([A-Z]+\d+)"',text).group(1)
                    if allowed(a):
                        if a in new_cells:cells[a]=xml(new_cells[a])
                    else:
                        if a in plan['changed_formulas']:
                            old_f=E.fromstring('<root xmlns="'+NS+'">'+text+'</root>')[0].find('s:f',N)
                            assert old_f is not None
                            old_f.text=plan['changed_formulas'][a][1][1:]
                            text=re.sub(r'<f\b[^>]*>.*?</f>',lambda m:xml(old_f),text,count=1,flags=re.S)
                        cells[a]=text
                for c in rows.get(row_n,[]):
                    if c.tag==tag('c') and allowed(c.get('r')):cells[c.get('r')]=xml(c)
                start=re.match(r'<row\b[^>]*>',old_row).group().replace('/>','>')
                row_text.append(start+''.join(cells[a] for a in sorted(cells,key=lambda a:coordinate_to_tuple(a)[1]))+'</row>')
            data='<sheetData>'+''.join(row_text)+'</sheetData>'
            raw=raw[:data_match.start()]+data+raw[data_match.end():]
            # Existing column dimensions remain unchanged; only append dimensions for new columns.
            cols=deepcopy(original_sheet.find('s:cols',N))
            if cols is None:cols=E.Element(tag('cols'))
            kept=[deepcopy(c) for c in original_sheet.find('s:cols',N)] if original_sheet.find('s:cols',N) is not None else []
            for c in kept:
                if int(c.get('min'))<=old_end<int(c.get('max')):c.set('max',str(old_end))
            cols[:]=[c for c in kept if int(c.get('max'))<=old_end or int(c.get('min'))>212]
            for c in new.find('s:cols',N):
                lo,hi=int(c.get('min')),int(c.get('max'))
                if hi>old_end and lo<=212:
                    d=deepcopy(c);d.set('min',str(max(lo,old_end+1)));cols.append(d)
            raw=replace_section(raw,'cols',cols)
            for name in ['dimension','mergeCells','dataValidations']:
                el=new.find('s:'+name,N)
                if el is not None:raw=replace_section(raw,name,el)
            assert_worksheet_order(raw)
            after=E.fromstring(raw);after_cells={c.get('r'):c for c in after.findall('s:sheetData/s:row/s:c',N)}
            after_rows=after.findall('s:sheetData/s:row',N)
            assert len({r.get('r') for r in after_rows})==len(after_rows),'duplicate rows'
            for row in after_rows:
                assert all(coordinate_to_tuple(c.get('r'))[0]==int(row.get('r')) for c in row if c.tag==tag('c')),'cell is inside the wrong row'
            assert len(after_cells)==sum(len(row) for row in after_rows),'duplicate cells or non-cell row children'
            for a,old in old_cells.items():
                if allowed(a):
                    if a not in ['FG2','FG3','HD2','HD3'] and (old.find('s:f',N) is not None or ''.join(old.itertext()).strip()):
                        expected=deepcopy(old);actual=deepcopy(after_cells[a]);expected.attrib.pop('s',None);actual.attrib.pop('s',None)
                        if a in plan['changed_formulas']:expected.find('s:f',N).text=plan['changed_formulas'][a][1][1:]
                        assert xml(expected)==xml(actual),(source.name,plan['name'],a,'existing value outside frame changed')
                else:
                    expected=deepcopy(old)
                    if a in plan['changed_formulas']:expected.find('s:f',N).text=plan['changed_formulas'][a][1][1:]
                    assert xml(expected)==xml(after_cells[a]),(source.name,plan['name'],a,'old cell changed')
            original_merges={c.get('ref') for c in original_sheet.find('s:mergeCells',N)}
            assert original_merges <= {c.get('ref') for c in after.find('s:mergeCells',N)}
            patches[part]=raw.encode()
        style_raw=original.read('xl/styles.xml').decode();old_style=E.fromstring(style_raw);new_style=E.fromstring(authored.read('xl/styles.xml'))
        for name in ['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs']:
            old=old_style.find('s:'+name,N);new=new_style.find('s:'+name,N)
            if new is None:continue
            n=len(old) if old is not None else 0
            if len(new)<=n:continue
            if old is None:raise ValueError('New style collection needs explicit insertion: '+name)
            out=deepcopy(old)
            for child in list(new)[n:]:out.append(deepcopy(child))
            out.set('count',str(len(out)))
            pattern=r'<'+name+r'\b[^>]*(?:/>|>.*?</'+name+r'>)'
            style_raw=re.sub(pattern,lambda m:xml(out),style_raw,count=1,flags=re.S)
        if style_raw.encode()!=original.read('xl/styles.xml'):patches['xl/styles.xml']=style_raw.encode()
        # Retain defined names, updating only print areas for the targeted sheets.
        workbook_raw=original.read('xl/workbook.xml').decode()
        wb_new=E.fromstring(authored.read('xl/workbook.xml'))
        new_names=wb_new.find('s:definedNames',N)
        if new_names is not None:
            target_ids={str(list(original_map).index(p['name'])) for p in plans}
            desired={(n.get('name'),n.get('localSheetId')):n.text for n in new_names if n.get('name')=='_xlnm.Print_Area' and n.get('localSheetId') in target_ids}
            def name_change(m):
                node=E.fromstring(m.group());key=(node.get('name'),node.get('localSheetId'))
                if key in desired:node.text=desired[key];return xml(node)
                return m.group()
            workbook_raw=re.sub(r'<definedName\b[^>]*>.*?</definedName>',name_change,workbook_raw,flags=re.S)
        if workbook_raw.encode()!=original.read('xl/workbook.xml'):patches['xl/workbook.xml']=workbook_raw.encode()
        with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as out:
            for item in original.infolist():out.writestr(copy(item),patches.get(item.filename,original.read(item.filename)))
        with zipfile.ZipFile(output) as out:
            assert original.namelist()==out.namelist()
            assert all(original.read(n)==out.read(n) for n in original.namelist() if n not in patches)
        return {'source':str(source),'output':str(output),'sheets':plans,'changed_parts':list(patches),'original_cells_and_caches_preserved':True,'unrelated_parts_preserved':True}
