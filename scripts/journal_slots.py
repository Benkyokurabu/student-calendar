"""Journal lesson capacity and counter storage, shared by generation and extraction."""
from copy import copy
import re
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter,range_boundaries
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.dimensions import ColumnDimension
from journal_input_controls import ensure_controls,is_unused_gray,lesson_columns

MIN_LESSON_SLOTS=21
COUNTER_COLUMN=212  # HD: outside the 21st lesson (GT:HB).

def slot_columns(ws):
    headers=lesson_columns(ws)
    if headers:return headers
    # Sparse legacy/test sheets may contain dates without the printed row-6 labels.
    return list(range(2,min(max(ws.max_column,2),211)+1,10))

def read_slot_columns(ws):
    """Also read legacy entries to the right of the printed frames; exclude helper HD."""
    return list(range(2,min(max(ws.max_column,2),211)+1,10))

def counter_address(ws,row):
    return f'HD{row}' if len(lesson_columns(ws))>=17 or ws.cell(row,COUNTER_COLUMN).value is not None else f'FG{row}'

def counter_value(ws,row):
    return ws[counter_address(ws,row)].value

def chain_formula(previous,base):
    chain=f'IF({base}="","",{base})'
    if not previous:return '='+chain
    for cell in previous:chain=f'IF(ISNUMBER({cell}),{cell}+1,{chain})'
    prev=previous[-1]
    return f'=IF({prev}="","",IF({prev}="特",{chain},{prev}+1))'

def ensure_lesson_capacity(ws,source=None,minimum=MIN_LESSON_SLOTS):
    """Append clean forms; do not move or overwrite existing lesson cells."""
    headers=lesson_columns(ws)
    if not headers or headers!=list(range(2,2+10*len(headers),10)):
        raise ValueError(f'{ws.title}: lesson headers are missing or not contiguous')
    source=source if source is not None else ws
    old=len(headers);minimum=max(old,minimum)
    if minimum>21:raise ValueError('Capacity above 21 requires moving the counter storage first')
    # Never overwrite an existing cell used for something else at the new helper location.
    helper={}
    for row,first in [(2,'F2'),(3,'G3')]:
        legacy=ws.cell(row,163)
        value=legacy.value if old<17 else None
        current=ws.cell(row,COUNTER_COLUMN).value
        if current is not None and not isinstance(current,(int,float)):
            raise ValueError(f'{ws.title}!HD{row}: unexpected existing data')
        helper[row]=current if current is not None else value
        if helper[row] is None and isinstance(ws[first].value,(int,float)):
            helper[row]=ws[first].value
    # First-slot source must contain labels only, not a teacher's recorded lesson.
    source_cells=[]
    for row in source.iter_rows(min_col=2,max_col=11,max_row=source.max_row):
        for c in row:
            if not isinstance(c,MergedCell):source_cells.append((c.row,c.column,c.value,copy(c._style)))
    source_merges=[CellRange(str(r)) for r in source.merged_cells.ranges if r.min_col>=2 and r.max_col<=11]
    for i in range(old,minimum):
        shift=i*10
        prior_values={c.coordinate:c.value for row in ws.iter_rows(min_col=2+shift,max_col=11+shift,max_row=source.max_row) for c in row if c.value is not None and c.coordinate not in ['FG2','FG3']}
        for r,c,v,style in source_cells:
            target=ws.cell(r,c+shift)
            prior=prior_values.get(target.coordinate)
            target._style=copy(style)
            if is_unused_gray(target.fill):
                fill=copy(target.fill);fill.patternType=None;target.fill=fill
            # Only static layout labels/class names are copied. User-entry fields stay blank.
            keep=(r<=6 or (c==2 and ((r-7)%20==0 or (r-6)%20 in [0,3,7,10,11,12,14,15])) or
                  (c==3 and (r-6)%20 in [4,5,6]) or (c==4 and (r-6)%20 in [0,2,3,5]))
            target.value=v if keep else None
            if isinstance(target.value,str) and target.value.startswith('='):
                target.value=Translator(target.value,origin=f'{get_column_letter(c)}{r}').translate_formula(target.coordinate)
            if prior is not None:
                if keep and r>=6 and prior!=target.value and not (c==2 and (r-7)%20==0):
                    raise ValueError(f'{ws.title}!{target.coordinate}: added area already has content conflicting with a layout label')
                target.value=prior
        for c in range(2,12):
            origin=next((d for d in source.column_dimensions.values() if (d.min or 0)<=c<=(d.max or 0)),None)
            if origin is None:origin=ColumnDimension(ws,index=get_column_letter(c),width=source.sheet_format.defaultColWidth or 13)
            dim=copy(origin);dim.index=get_column_letter(c+shift);dim.min=c+shift;dim.max=c+shift
            ws.column_dimensions[dim.index]=dim
        for merged in source_merges:
            added=CellRange(str(merged));added.shift(col_shift=shift)
            for row in ws.iter_rows(min_row=added.min_row,max_row=added.max_row,min_col=added.min_col,max_col=added.max_col):
                for c in row:
                    if (c.row,c.column)!=(added.min_row,added.min_col) and c.coordinate in prior_values:
                        raise ValueError(f'{ws.title}!{c.coordinate}: existing content would be hidden by a new merge')
            ws.merge_cells(str(added))
        left=2+shift
        for r,offset,formula in [(2,0,'=$B$2'),(2,7,'=$I$2'),(3,7,'=$I$3'),(3,3,'=$E$3')]:
            c=ws.cell(r,left+offset)
            if c.coordinate not in prior_values:c.value=formula
        for r,offset,base in [(2,4,'$HD$2'),(3,5,'$HD$3')]:
            c=ws.cell(r,left+offset)
            if c.coordinate not in prior_values:c.value=chain_formula([f'{get_column_letter(2+offset+10*j)}{r}' for j in range(i)],base)
    # Move legacy helpers before the new 17th header occupies FG2:FG3.
    for row,v in helper.items():ws.cell(row,COUNTER_COLUMN,v)
    ws.column_dimensions['HD'].hidden=True
    for i in range(minimum):
        for row,col in [(2,6+10*i),(3,7+10*i)]:
            c=ws.cell(row,col)
            if c.data_type=='f':c.value=re.sub(r'\$FG\$([23])\b',r'$HD$\1',c.value)
    ensure_controls(ws,source)
    if ws.print_area:
        areas=[]
        for area in ws.print_area.ranges:
            a=CellRange(str(area))
            if a.max_col>=old*10:a.max_col=max(a.max_col,minimum*10)
            areas.append(str(a))
        ws.print_area=areas
    return minimum-old
