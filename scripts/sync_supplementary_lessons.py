"""Independent post-publication supplement sync; never edits journal originals."""
from __future__ import annotations
import argparse
from datetime import date, datetime, timezone, timedelta
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile
import time
import unicodedata
from urllib.parse import urlencode
from urllib.request import Request, urlopen

ENDPOINT = 'https://line-check-system.vercel.app/api/journal-management/supplements'
REMOTE = 'onedrive:デスクトップ/【完成版】授業日誌システム'
JST = timezone(timedelta(hours=9))

def allowed_months(now=None):
    today = (now or datetime.now(JST)).date()
    following = today.replace(day=28) + timedelta(days=4)
    return [today.strftime('%Y-%m'), following.strftime('%Y-%m')]

def extract_supplements(path, month):
    import openpyxl
    import export_by_grade_subject as legacy
    book = openpyxl.load_workbook(path, data_only=True)
    events = {}
    try:
        targets = legacy.choose_target_sheets(book)
        if {campus for campus, _, _ in targets} != {'本校', '南教室'}:
            raise ValueError('Both campus sheets are required')
        for campus, _, sheet in targets:
            legend = legacy.detect_teacher_legend(sheet)
            for row in range(2, 55):
                day, _ = legacy.parse_day_week_from_row(sheet, row, allow_prev_row=True)
                if not day:
                    continue
                for col in range(3, 31):
                    text = re.sub(r'\s+', '', legacy.merged_value_text(sheet, row, col))
                    if not re.fullmatch(r'(数学|英語)補講\d{1,2}', unicodedata.normalize('NFKC', text)):
                        continue
                    # Use the existing grade-aware reader, not the elementary header.
                    band = legacy.find_timeband_above(sheet, row, col, grade='1')
                    match = re.search(r'(\d{1,2})[:：](\d{2})', band)
                    if not match:
                        raise ValueError('A supplement time cannot be read')
                    hour, minute = map(int, match.groups())
                    if hour < 10:
                        hour += 12
                    if not (0 <= hour <= 23 and 0 <= minute <= 59):
                        raise ValueError('Invalid supplement start')
                    event = dict(campus=campus, date=date(int(month[:4]), int(month[5:]), day).isoformat(),
                                 name=text, time=f'{hour:02d}:{minute:02d}〜',
                                 teacher=legacy.resolve_teacher_strict(sheet, legend, row, col))
                    key = campus, unicodedata.normalize('NFKC', text)
                    if key in events and events[key] != event:
                        raise ValueError('Ambiguous supplement occurrence')
                    events[key] = event
        return sorted(events.values(), key=lambda event: (event['date'], event['campus'], event['name']))
    finally:
        book.close()

def command(args):
    result = subprocess.run(args, capture_output=True, timeout=90)
    if result.returncode:
        raise RuntimeError('Cloud source retrieval failed')
    return result.stdout

def oidc_token():
    url = os.environ['ACTIONS_ID_TOKEN_REQUEST_URL']
    if not url.startswith('https://'):
        raise ValueError('Invalid identity endpoint')
    url += ('&' if '?' in url else '?') + urlencode({'audience': 'bentan-supplement-sync'})
    request = Request(url, headers={'Authorization': 'Bearer ' + os.environ['ACTIONS_ID_TOKEN_REQUEST_TOKEN']})
    with urlopen(request, timeout=20) as response:
        return json.load(response)['value']

def sync_month(path, month, *, dry_run, parent_run_id, download, send):
    frozen = path.read_bytes()
    source = hashlib.sha256(frozen).hexdigest()
    events = extract_supplements(path, month)
    total = {'created': 0, 'updated': 0}
    for batch in range(100):
        # Re-fetch the cloud original before each bounded batch of Notion writes.
        if download() != frozen:
            raise RuntimeError('The source changed; supplement writes stopped')
        result = send(dict(month=month, events=events, dryRun=dry_run, parentRunId=parent_run_id, source=source))
        for key in total:
            total[key] += result.get(key, 0)
        if result.get('skipped') or not result.get('pending') or dry_run:
            return {**result, **total, 'month': month, 'dryRun': dry_run}
        if result.get('busy'):
            time.sleep(10)
    raise RuntimeError('Supplement batches remain pending')

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--month', default='')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()
    months = allowed_months()
    if args.month:
        if args.month not in months:
            raise ValueError('Only current/next month is allowed')
        months = [args.month]
    listing = json.loads(command(['rclone', 'lsjson', REMOTE, '--files-only', '--max-depth', '1']))
    parent = os.environ.get('PARENT_RUN_ID') or None
    summaries = []
    with tempfile.TemporaryDirectory(prefix='bentan-supplements-') as temporary:
        work = Path(temporary)
        for month in months:
            names = [item['Name'] for item in listing if re.fullmatch(
                rf'{month[:4]}年0?{int(month[5:])}月スケジュール\.xlsm', unicodedata.normalize('NFKC', item['Name']))]
            if not names:
                continue
            if len(names) != 1 or any('/' in name or '\\' in name for name in names):
                raise ValueError('Ambiguous schedule source')
            path = work / f'{month}.xlsm'
            command(['rclone', 'copyto', REMOTE + '/' + names[0], str(path)])
            def download():
                check = work / f'{month}-recheck.xlsm'
                command(['rclone', 'copyto', REMOTE + '/' + names[0], str(check)])
                return check.read_bytes()
            def send(body):
                request = Request(ENDPOINT, data=json.dumps(body, ensure_ascii=False).encode('utf-8'),
                                  headers={'Authorization': 'Bearer ' + oidc_token(), 'Content-Type': 'application/json'}, method='POST')
                with urlopen(request, timeout=55) as response:
                    return json.load(response)
            summaries.append(sync_month(path, month, dry_run=args.dry_run, parent_run_id=parent, download=download, send=send))
    summary = json.dumps({'supplements': summaries}, ensure_ascii=False)
    print(summary)
    if os.environ.get('GITHUB_STEP_SUMMARY'):
        with open(os.environ['GITHUB_STEP_SUMMARY'], 'a', encoding='utf-8') as output:
            output.write('### 補講予定の作成・照合\n\n```json\n' + summary + '\n```\n')

if __name__ == '__main__':
    main()
