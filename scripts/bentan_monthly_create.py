"""Scoped monthly creation with independent, version-guarded files and resumable results."""
from pathlib import Path,PurePosixPath
from concurrent.futures import ThreadPoolExecutor,as_completed
from datetime import datetime
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import hashlib,json,os,re,subprocess,tempfile
from journal_source import source_remote
from safe_journal_upload import (GraphClient,GraphError,UploadSafetyError,read_original,digest,
    remote_listing,listing_identities,replace_existing_if_unchanged,probe_conditional_put)
from cloud_backup import backup_file
from journal_month_workbook import generate_month_workbook

JOURNAL_REMOTE='onedrive:●勉強クラブ共有/09　授業日誌'
RETRYABLE={'open','changed','temporary'}

def packed(value):return json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()
def file_key(drive,item,relative):return digest((drive.casefold()+'\0'+item+'\0'+relative).encode())
def journal_workbook_year(year,month):return year if month>=3 else year-1

def has_month(data,month):
    with ZipFile(BytesIO(data)) as z:root=ET.fromstring(z.read('xl/workbook.xml'))
    return any(n.get('name')==month for n in root.iter() if n.tag.endswith('}sheet'))

def parse_keys(raw):
    if raw in (None,''):return None
    keys=json.loads(raw)
    if not isinstance(keys,list) or len(keys)>1000 or len(set(keys))!=len(keys) or any(not isinstance(k,str) or not re.fullmatch('[a-f0-9]{64}',k) for k in keys):raise ValueError('Invalid monthly file scope')
    return sorted(keys)

def collect_targets(schedule,year,month):
    import openpyxl,export_by_grade_subject as ex
    from export_schedule_json import convert_events_to_schedule_json,collect_special_events_from_sheet
    book=openpyxl.load_workbook(schedule,data_only=True,keep_vba=True)
    file_year=journal_workbook_year(year,month)
    result={};lessons=0;campuses=set()
    try:
        for campus,_,sheet in ex.choose_target_sheets(book):
            campuses.add(campus);events=ex.collect_events(sheet,month,campus,year,month)
            lessons+=len(convert_events_to_schedule_json(events,campus,year,month))
            lessons+=len(collect_special_events_from_sheet(sheet,campus,year,month))
            for e in events:
                supplementary=next((s for s in ('英語補講','数学補講') if s in e.text),None)
                if supplementary:grade='';subject=supplementary;klass='S';name=campus+subject+f'_{file_year}.xlsx'
                else:
                    grade=ex.GRADE_LABEL[e.grade];subject=ex._display_subj_name(e.grade,e.subj);klass=e.klass
                    name=campus+grade+('X' if klass=='X' else '')+subject+f'_{file_year}.xlsx'
                info=result.setdefault(name,{'filename':name,'grade':grade,'subject':subject,'supplementary':bool(supplementary),'classes':{}})
                info['classes'].setdefault(klass,[]).append(list(e))
    finally:
        book.close()
        if book.vba_archive is not None:book.vba_archive.close()
    if campuses!={'本校','南教室'} or not lessons or not result:raise ValueError('Both campuses and a nonempty timetable are required')
    return result,lessons

def open_excel_names():
    if os.name!='nt':return set()
    import csv
    result=subprocess.run(['tasklist','/FI','IMAGENAME eq EXCEL.EXE','/FO','CSV','/NH'],capture_output=True,encoding='utf-8',errors='replace',check=True)
    count=sum(1 for r in csv.reader(result.stdout.splitlines()) if r and r[0].upper()=='EXCEL.EXE')
    if not count:return set()
    if count>1:return None
    try:
        import win32com.client
        app=win32com.client.GetActiveObject('Excel.Application')
        return {str(w.Name).casefold() for w in app.Workbooks}
    except Exception:return None

def get_file(client,identity):
    drive,item=identity;content,meta=read_original(client,drive,item)
    return content,{'driveId':drive,'itemId':item,'eTag':meta['eTag'],'sha256':digest(content),'size':len(content),
        'name':meta.get('name'),'parentId':meta.get('parentReference',{}).get('id'),'parentPath':meta.get('parentReference',{}).get('path')}

def prepare_and_save(item,info,month,templates,work,action,request_id,opened,input_versions,client_factory):
    key=item['key'];result={'key':key,'name':item['name'],'status':'review'}
    if action=='create' and (opened is None or item['name'].casefold() in opened):return {**result,'status':'open'}
    client=client_factory();payload=None;write_attempted=False
    try:
        original,entry=get_file(client,(item['driveId'],item['itemId']))
        if entry.get('name') and entry['name']!=item['name'] or item.get('parentId') and entry.get('parentId')!=item['parentId'] or item.get('parentPath') and entry.get('parentPath')!=item['parentPath']:return {**result,'status':'changed'}
        if has_month(original,month):return {**result,'status':'exists'}
        meta=client.metadata(entry['driveId'],entry['itemId'])
        if meta.get('publication',{}).get('level')=='checkout':return {**result,'status':'open'}
        source=work/(key+'-before.xlsx');target=work/(key+'-after.xlsx');source.write_bytes(original)
        generated=generate_month_workbook(source,target,month,info,templates,work/key)
        starts={k:v['counter'] for k,v in generated['first'].items()}
        result['starts']=starts
        result['adjusted']=len(set(starts.values()))>1
        if action=='preview':return {**result,'status':'planned'}
        # The schedule and selected registered master must still be this accepted input.
        for identity in input_versions:
            current=client.metadata(identity['driveId'],identity['itemId'])
            if current['eTag']!=identity['eTag'] or current.get('name')!=identity.get('name'):return {**result,'status':'review','reason':'input_changed'}
        backup_file(source,category='bentan-journal-create',timestamp=request_id,
            relative_path=month+'/'+key+'/'+item['name'])
        payload=target.read_bytes()
        write_attempted=True
        replace_existing_if_unchanged(client,entry['driveId'],entry['itemId'],entry['eTag'],entry['sha256'],payload)
        return {**result,'status':'created'}
    except GraphError as error:
        if error.status==423:return {**result,'status':'open'}
        if error.status in (409,412):return {**result,'status':'changed'}
        if error.status in (429,500,502,503,504):return {**result,'status':'temporary'}
        return {**result,'reason':'cloud_access'}
    except UploadSafetyError:
        if not write_attempted:return {**result,'status':'temporary'}
        # An uncertain PUT can already have committed. Reconcile bytes, never repeat it.
        try:
            received,_=get_file(client,(item['driveId'],item['itemId']))
            if payload is not None and digest(received)==digest(payload):return {**result,'status':'created'}
            return {**result,'status':'changed'}
        except UploadSafetyError:return {**result,'reason':'uncertain'}
    except (ValueError,KeyError):return {**result,'reason':'counter_or_template'}
    except (subprocess.SubprocessError,RuntimeError,OSError):return {**result,'status':'temporary'}

def run_month_creation(scripts,repo,operation,month,request_id,notify):
    """Preview and create use the same explicit scope; already-created sheets never change."""
    if operation not in ('preview','create'):raise ValueError('Unsupported monthly operation')
    year,number=map(int,month.split('-'))
    scope=parse_keys(os.environ.get('BENTAN_TARGET_KEYS'))
    resume=os.environ.get('BENTAN_RESUME_ONLY')=='true'
    if resume and (operation!='create' or scope is None):raise ValueError('A retry requires an explicit accepted file scope')
    client=GraphClient.from_rclone('onedrive')
    source_rows=json.loads(subprocess.run(['rclone','lsjson',source_remote(),'--files-only','--max-depth','1'],check=True,capture_output=True,text=True,encoding='utf-8').stdout)
    schedules=[r for r in source_rows if re.fullmatch(rf'{year}年0?{number}月スケジュール\.xlsm',r.get('Name',''))]
    if len(schedules)!=1:raise ValueError('Exactly one accepted monthly timetable is required')
    template_names=['TEMPLATE_MAIN.xlsx','TEMPLATE_X.xlsx']
    registered=any(r.get('Name')=='TEMPLATE_MAIN_COMMON.xlsx' for r in source_rows)
    if registered:template_names.append('TEMPLATE_MAIN_COMMON.xlsx')
    from safe_journal_upload import split_rclone_id
    from bentan_journal_operation import source_digest
    with tempfile.TemporaryDirectory(prefix='bentan-month-') as directory:
        work=Path(directory);templates=work/'templates';templates.mkdir()
        inputs={};input_versions=[]
        for row in schedules+[next(r for r in source_rows if r.get('Name')==name) for name in template_names]:
            content,entry=get_file(client,split_rclone_id(row['ID']));inputs[row['Name']]=content;input_versions.append(entry)
            (templates/row['Name']).write_bytes(content)
        schedule=templates/schedules[0]['Name']
        # collect_events may emit diagnostic text; keep it in the isolated working folder.
        previous_cwd=Path.cwd()
        try:os.chdir(work);targets,lessons=collect_targets(schedule,year,number)
        finally:os.chdir(previous_cwd)
        identities=listing_identities(remote_listing(JOURNAL_REMOTE))
        items=[];by_name={};manifest={'entries':{}}
        for relative,(drive,item_id) in identities.items():
            name=PurePosixPath(relative).name
            if name in by_name:raise ValueError('Workbook names are ambiguous')
            item={'key':file_key(drive,item_id,relative),'name':name,'driveId':drive,'itemId':item_id,'path':relative};items.append(item);by_name[name]=item
        factory=lambda:GraphClient.from_rclone('onedrive')
        def capture(item):
            content,entry=get_file(factory(),(item['driveId'],item['itemId']))
            return item,content,entry
        with ThreadPoolExecutor(max_workers=4) as pool:
            for item,content,entry in pool.map(capture,items):
                if entry['name']!=item['name']:raise ValueError('An original was renamed while reading')
                manifest['entries'][item['path']]={k:entry[k] for k in ('driveId','itemId','eTag','sha256','size')}
                item.update(exists=has_month(content,month),parentId=entry['parentId'],parentPath=entry['parentPath'])
        input_digest=digest(packed({'profile':source_remote(),'month':month,'files':{n:digest(b) for n,b in inputs.items()}}))
        creation_digest=digest(packed({'inputDigest':input_digest,'targets':sorted((name,by_name[name]['key']) for name in targets if name in by_name)}))
        frozen=source_digest(manifest,inputs[schedules[0]['Name']],(repo/f'schedule_{month}.json').read_bytes() if (repo/f'schedule_{month}.json').exists() else b'',(repo/f'journal_{month}.json').read_bytes() if (repo/f'journal_{month}.json').exists() else b'',inputs.get('TEMPLATE_MAIN_COMMON.xlsx'))
        if operation=='create':
            expected=os.environ.get('BENTAN_EXPECTED_DIGEST')
            if resume:
                if expected!=input_digest:raise ValueError('The accepted timetable or template changed; a new preview is required')
            elif expected!=creation_digest:raise ValueError('The accepted monthly scope changed; a new preview is required')
        chosen=[by_name[name] for name in targets if name in by_name and (not by_name[name]['exists'] if scope is None else by_name[name]['key'] in scope)]
        planned_keys=sorted(item['key'] for item in chosen) if scope is None else scope
        missing_names=[name for name in targets if name not in by_name]
        chosen_by_key={item['key']:item for item in chosen};results=[]
        for key in planned_keys:
            if key not in chosen_by_key:results.append({'key':key,'name':'原本未確認.xlsx','status':'review','reason':'source_missing'})
        if operation=='create' and chosen:
            from journal_pc_upload import probe_parents
            parents=probe_parents('rclone',JOURNAL_REMOTE,{i['key']:i for i in chosen})
            for drive,parent in parents.items():probe_conditional_put(client,drive,parent)
        opened=open_excel_names()
        with ThreadPoolExecutor(max_workers=3) as pool:
            futures=[pool.submit(prepare_and_save,item,targets[item['name']],month,templates,work,operation,request_id,opened,input_versions,factory) for item in chosen]
            for future in as_completed(futures):results.append(future.result())
        summary={'sourceDigest':frozen,'creationDigest':creation_digest,'inputDigest':input_digest,'monthlyVersion':1,
            'lessons':lessons,'changedWorkbooks':len(planned_keys),'uploaded':sum(r['status']=='created' for r in results),
            'published':False,'publicationVerified':False,'masterRegistered':registered,
            'templateMode':'registered' if registered else 'current','targetKeys':planned_keys,'results':results,
            'missingWorkbooks':missing_names,'alreadyExisting':sum(item['exists'] for item in items if item['name'] in targets)}
        notify(request_id,operation,month,summary)
        return summary
