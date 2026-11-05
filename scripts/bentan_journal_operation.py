"""Additive, manual journal operations. The desktop and daily pipeline stay intact."""
from __future__ import annotations
import copy
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone, timedelta
from urllib.request import Request, urlopen
from xml.etree.ElementTree import tostring

JST = timezone(timedelta(hours=9))
REMOTE = 'onedrive:●勉強クラブ共有/09　授業日誌'
CALLBACK = 'https://line-check-system.vercel.app/api/journal-management/callback'
PUBLIC = 'https://benkyokurabu.github.io/student-calendar'

def digest(data):
    return hashlib.sha256(data).hexdigest()

def encoded(data):
    return json.dumps(data, ensure_ascii=False, sort_keys=True, default=str).encode()

def validate_operation(operation, month, request_id, now=None):
    now = now or datetime.now(JST)
    current = now.strftime('%Y-%m')
    following = (now.replace(day=28) + timedelta(days=4)).strftime('%Y-%m')
    if operation not in ('preview', 'create', 'publish') or month not in (current, following):
        raise ValueError('Only the current or next month can be selected')
    if not re.fullmatch(r'[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}', request_id):
        raise ValueError('Invalid operation identity')
    if operation == 'publish' and month != current:
        raise ValueError('Only the current month can be published')
    if operation == 'create' and month == '2026-12' and now < datetime(2026, 11, 20, tzinfo=JST):
        raise ValueError('December generation remains on hold until the late-November schedule is ready')

def command(args, cwd=None):
    # Captured output stays on the runner. Avoid disclosing workbook contents in logs.
    result = subprocess.run(args, cwd=cwd, capture_output=True, timeout=1800)
    if result.returncode:
        raise RuntimeError('Operation command failed: ' + Path(args[0]).name)
    return result.stdout

def sheet_fingerprint(sheet):
    return digest(encoded({
        'state': sheet.sheet_state,
        'cells': [(r, c, cell.value, cell.data_type, str(cell._style), cell.number_format,
                   str(cell.hyperlink), str(cell.comment)) for (r, c), cell in sorted(sheet._cells.items())],
        'merged': sorted(str(item) for item in sheet.merged_cells.ranges),
        'rows': [(k, v.height, v.hidden, v.outlineLevel, v.collapsed, str(v._style)) for k, v in sorted(sheet.row_dimensions.items())],
        'columns': [(k, v.width, v.hidden, v.min, v.max, v.outlineLevel, str(v._style)) for k, v in sorted(sheet.column_dimensions.items())],
        'validations': tostring(sheet.data_validations.to_tree()).decode(),
        'conditional': [(str(k), [tostring(r.to_tree()).decode() for r in rules]) for k, rules in sheet.conditional_formatting._cf_rules.items()],
        'print_area': str(sheet.print_area), 'print_rows': sheet.print_title_rows, 'print_cols': sheet.print_title_cols,
        'page': tostring(sheet.page_setup.to_tree()).decode(), 'margins': tostring(sheet.page_margins.to_tree()).decode(),
        'headers': tostring(sheet.HeaderFooter.to_tree()).decode(),
        'protection': tostring(sheet.protection.to_tree()).decode(),
        'format': tostring(sheet.sheet_format.to_tree()).decode(),
        'properties': tostring(sheet.sheet_properties.to_tree()).decode(),
        'view': tostring(sheet.views.to_tree()).decode(),
        'filter': tostring(sheet.auto_filter.to_tree()).decode(),
        'print_options': tostring(sheet.print_options.to_tree()).decode(),
    }))

def workbook_sheets(path):
    import openpyxl
    book = openpyxl.load_workbook(path)
    try:
        return {sheet.title: sheet_fingerprint(sheet) for sheet in book}
    finally:
        book.close()

def assert_only_new_month(before, after, month):
    if any(after.get(name) != value for name, value in before.items()):
        raise ValueError('An existing worksheet changed; original uploads are blocked')
    if set(after) - set(before) - {month}:
        raise ValueError('An unselected worksheet was added')

def source_digest(manifest, schedule_bytes, public_schedule, public_journal):
    return digest(encoded({'entries': manifest['entries'], 'schedule': digest(schedule_bytes),
                           'publicSchedule': digest(public_schedule), 'publicJournal': digest(public_journal)}))

def merge_for_publication(baseline, extracted, today):
    result = copy.deepcopy(extracted)
    old = baseline.get('entries', {})
    if set(old) - set(extracted.get('entries', {})):
        raise ValueError('Existing published lesson identities would be removed')
    fields = ('content', 'page', 'report', 'teacher', 'absence', 'curriculumProgress', 'note', 'sessionNumber', 'monthNum', 'weekNum', 'homework')
    def protect_recordings(value, previous):
        if not isinstance(value, dict):
            return value
        for key in list(value):
            if key in ('recordingUrl', 'recordingUrls', 'zoomUrl', 'zoomURL'):
                value[key] = copy.deepcopy(previous.get(key, '' if key != 'recordingUrls' else []))
            elif isinstance(value[key], dict):
                value[key] = protect_recordings(value[key], previous.get(key, {}) if isinstance(previous.get(key), dict) else {})
        return value
    for key, candidate in result['entries'].items():
        previous = old.get(key, {})
        if key[:10] > today:
            result['entries'][key] = copy.deepcopy(previous) if previous else protect_recordings(copy.deepcopy(candidate), {})
            continue
        merged = {**copy.deepcopy(previous), **copy.deepcopy(candidate)}
        for field in fields:
            if not candidate.get(field) and previous.get(field):
                merged[field] = copy.deepcopy(previous[field])
        result['entries'][key] = protect_recordings(merged, previous)
    return result

def notify(request_id, action, month, summary):
    request = Request(os.environ['ACTIONS_ID_TOKEN_REQUEST_URL'] + '&audience=bentan-journal-management',
                      headers={'Authorization': 'Bearer ' + os.environ['ACTIONS_ID_TOKEN_REQUEST_TOKEN']})
    with urlopen(request, timeout=30) as response:
        identity = json.load(response)['value']
    request = Request(CALLBACK, data=encoded({'id': request_id, 'action': action, 'month': month, 'summary': summary}),
                      headers={'Authorization': 'Bearer ' + identity, 'Content-Type': 'application/json'}, method='POST')
    with urlopen(request, timeout=60) as response:
        if not json.load(response).get('saved'):
            raise RuntimeError('The operation result was not saved')

def main():
    operation, month, request_id = (os.environ[key] for key in ('BENTAN_OPERATION', 'BENTAN_MONTH', 'BENTAN_REQUEST_ID'))
    validate_operation(operation, month, request_id)
    scripts = Path(__file__).resolve().parent
    repo = scripts.parent
    os.chdir(scripts)
    work = scripts / '_bentan_operation'
    work.mkdir()
    print('[BENTAN] Checking the selected schedule original', flush=True)
    listing = command(['rclone', 'lsf', 'onedrive:デスクトップ/【完成版】授業日誌システム', '--files-only', '--max-depth', '1']).decode('utf-8')
    year, number = map(int, month.split('-'))
    names = [name for name in listing.splitlines() if re.fullmatch(rf'{year}年0?{number}月スケジュール\.xlsm', name)]
    if len(names) != 1:
        raise ValueError('Exactly one schedule original is required')
    schedule = scripts / names[0]
    command(['rclone', 'copyto', 'onedrive:デスクトップ/【完成版】授業日誌システム/' + names[0], str(schedule)])
    schedule_bytes = schedule.read_bytes()
    cloud = scripts / '_cloud_journal'
    print('[BENTAN] Downloading shared journals without modifying originals', flush=True)
    command(['rclone', 'copy', REMOTE, str(cloud), '--exclude', '**/_backup*/**', '--exclude', '**/退避/**', '--include', '*.xlsx', '--transfers', '8'])
    from safe_journal_upload import GraphClient, capture_manifest, remote_listing, verify_listing_unchanged, verify_versions, apply_manifest
    from export_schedule_json import export_month_schedule_json
    from build_journal_map import build_slots
    from cloud_backup import backup_file
    client = GraphClient.from_rclone('onedrive')
    print('[BENTAN] Freezing original file identities and versions', flush=True)
    manifest = capture_manifest(client, cloud, remote_listing(REMOTE), REMOTE)
    verify_listing_unchanged(manifest, remote_listing(REMOTE))
    public_schedule = repo / f'schedule_{month}.json'
    public_journal = repo / f'journal_{month}.json'
    frozen = source_digest(manifest, schedule_bytes, public_schedule.read_bytes() if public_schedule.exists() else b'', public_journal.read_bytes() if public_journal.exists() else b'')
    if operation != 'preview' and os.environ.get('BENTAN_EXPECTED_DIGEST') != frozen:
        raise ValueError('The originals or public data changed after preview; confirm again')
    output = export_month_schedule_json(str(schedule))
    rows = json.loads(output.read_text('utf-8'))
    if not rows or {row.get('campus') for row in rows} != {'hon', 'minami'} or any(not row.get('date', '').startswith(month + '-') for row in rows):
        raise ValueError('Both campuses and the selected month must be fully covered')
    originals = work / 'originals'
    changed = []
    if operation in ('preview', 'create'):
        for relative in manifest['entries']:
            destination = originals / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(cloud / relative, destination)
        # An unregistered editing master is never downloaded or registered here.
        command(['rclone', 'copy', 'onedrive:デスクトップ/【完成版】授業日誌システム', str(scripts), '--include', '/TEMPLATE_MAIN_COMMON.xlsx', '--max-depth', '1'])
        from ci_extract_and_push import create_month_sheets
        print('[BENTAN] Checking the selected month on isolated copies', flush=True)
        create_month_sheets(scripts, [month], cloud)
        for relative, entry in manifest['entries'].items():
            if digest((cloud / relative).read_bytes()) != entry['sha256']:
                # Exact byte equality is stronger and much cheaper for unchanged books.
                # Only a changed copy needs every old worksheet compared semantically.
                assert_only_new_month(workbook_sheets(originals / relative), workbook_sheets(cloud / relative), month)
                changed.append(relative)
    # Saving an Excel original after this read cannot silently publish a stale snapshot.
    schedule_check = work / 'schedule-recheck.xlsm'
    command(['rclone', 'copyto', 'onedrive:デスクトップ/【完成版】授業日誌システム/' + names[0], str(schedule_check)])
    if schedule_check.read_bytes() != schedule_bytes:
        raise ValueError('The schedule changed during confirmation')
    print('[BENTAN] Rechecking the schedule and every original version', flush=True)
    verify_listing_unchanged(manifest, remote_listing(REMOTE))
    verify_versions(client, manifest['entries'])
    summary = {'sourceDigest': frozen, 'lessons': len(rows), 'changedWorkbooks': len(changed), 'uploaded': 0, 'published': False, 'publicationVerified': False}
    if operation == 'create':
        # Back up the exact frozen original bytes before any original can be replaced.
        for relative in changed:
            backup_file(originals / relative, category='bentan-journal-create', timestamp=request_id, relative_path=relative)
        result = apply_manifest(client, cloud, manifest, json.loads(os.environ.get('JOURNAL_UPLOAD_PROBE_PARENTS') or '{}'))
        summary['uploaded'] = result['uploaded']
    elif operation == 'publish':
        if not public_schedule.exists() or json.loads(public_schedule.read_text('utf-8')) != rows:
            raise ValueError('The public timetable differs from the original; confirm the existing timetable update first')
        candidate = work / 'candidate'
        candidate.mkdir()
        shutil.copy2(public_schedule, candidate / public_schedule.name)
        mapping = {'month': month, 'slots': build_slots(year, number, schedule)}
        (candidate / f'journal_map_{month}.json').write_bytes(encoded(mapping))
        command([sys.executable, str(scripts / 'extract_journal_to_json.py'), '--month', month, '--journal-dir', str(cloud), '--repo-dir', str(candidate)])
        baseline = json.loads(public_journal.read_text('utf-8')) if public_journal.exists() else {'entries': {}}
        fresh = json.loads((candidate / public_journal.name).read_text('utf-8'))
        merged = merge_for_publication(baseline, fresh, datetime.now(JST).date().isoformat())
        if merged['entries'] == baseline.get('entries'):
            payload = public_journal.read_bytes()
        else:
            payload = json.dumps(merged, ensure_ascii=False, indent=2).encode()
        target = candidate / public_journal.name
        target.write_bytes(payload)
        from validate_journal_publish import validate
        if validate(target, public_schedule, public_journal if public_journal.exists() else None):
            raise ValueError('Publication validation failed')
        verify_versions(client, manifest['entries'])
        public_journal.write_bytes(payload)
        (repo / 'journal_latest.json').write_bytes(payload)
        command(['git', 'add', public_journal.name, 'journal_latest.json'], repo)
        if command(['git', 'diff', '--cached', '--name-only'], repo).strip():
            command(['git', '-c', 'user.name=github-actions[bot]', '-c', 'user.email=github-actions[bot]@users.noreply.github.com', 'commit', '-m', 'Publish journal from Bentan ' + request_id], repo)
            # Never rebase stale generated data over another publisher's changes.
            command(['git', 'push', 'origin', 'HEAD:main'], repo)
        summary['published'] = True
        for attempt in range(18):
            try:
                request = Request(f'{PUBLIC}/{public_journal.name}?bentan={request_id}-{attempt}', headers={'Cache-Control': 'no-cache'})
                with urlopen(request, timeout=20) as response:
                    if digest(response.read()) == digest(payload):
                        summary['publicationVerified'] = True
                        break
            except Exception:
                pass
            time.sleep(10)
        if not summary['publicationVerified']:
            raise RuntimeError('The commit was saved, but the public output has not been verified')
    print('[BENTAN] Saving the verified operation result to Bentan', flush=True)
    notify(request_id, operation, month, summary)
    print(json.dumps({'operation': operation, 'month': month, **summary}))

if __name__ == '__main__':
    main()
