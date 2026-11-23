"""Align classes by annual session, rather than their position within a month."""
import re
from datetime import date,timedelta

def previous_class_sessions(sheet):
    from export_by_grade_subject import _chain_formula
    from openpyxl.utils import get_column_letter
    def read(row,column):
        cell=sheet._cells.get((row,column))
        return cell.value if cell is not None else None
    headers=sorted(c for (r,c),cell in sheet._cells.items() if r==6 and cell.value=='クラス')
    columns=headers or list(range(2,min(max(sheet.max_column,2),211)+1,10))
    base='HD2' if len(headers)>=17 or read(2,212) is not None else 'FG2'
    counters=[];addresses=[];last={}
    for left in columns:
        address=get_column_letter(left+4)+'2';value=read(2,left+4)
        if isinstance(value,str) and value.startswith('='):
            expected=_chain_formula(addresses,'$'+base.replace('2','$2'))
            if value!=expected:
                raise ValueError(f'{sheet.title}!{address}: unknown annual counter formula')
            previous=counters[-1] if counters else None
            if previous in (None,''):value=''
            elif previous=='特':
                numeric=next((x for x in reversed(counters) if isinstance(x,(int,float))),None)
                value=numeric+1 if numeric is not None else read(2,212 if base=='HD2' else 163)
            elif isinstance(previous,(int,float)):value=previous+1
            else:raise ValueError('Cannot read previous class annual session')
        counters.append(value);addresses.append(address)
        if not isinstance(value,(int,float)):continue
        for klass,top in [('S',6),('A',26),('B',46)]:
            day=read(top+5,left)
            label=read(top+1,left)
            if day not in (None,'') and label!='特':last[klass]=int(value)
    return last

def align_new_month(sheet,month,classes,explicit=None):
    from export_by_grade_subject import patch_counter_formulas
    from journal_slots import counter_address
    explicit=dict(explicit or {})
    match=re.fullmatch(r'(\d{4})-(\d{2})',sheet.title)
    if not match:return explicit
    year,number=map(int,match.groups())
    if number!=month:raise ValueError('Sheet and selected month differ')
    previous=(date(year,month,1)-timedelta(days=1)).strftime('%Y-%m')
    if previous not in sheet.parent.sheetnames:return explicit
    active=[k for k in ('S','A','B') if any(not getattr(e,'special',False) for e in classes.get(k,[]))]
    if len(active)<2:return explicit
    prior=previous_class_sessions(sheet.parent[previous])
    if not all(k in prior for k in active):
        raise ValueError('Cannot infer every active class session from the previous month')
    starts={k:prior[k]+1 for k in active}
    start=min(starts.values());offsets={k:n-start for k,n in starts.items()}
    if explicit and any(explicit.get(k,0)!=offsets.get(k,0) for k in active):
        raise ValueError('Carryover markers disagree with the actual previous class sessions')
    sheet['F2']=start;sheet[counter_address(sheet,2)]=start
    patch_counter_formulas(sheet)
    return offsets
