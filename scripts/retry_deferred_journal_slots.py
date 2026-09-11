"""Bounded cloud retry for the three workbooks locked during the September migration.

Runs inside the existing journal writer, before its fresh download. No local sync
files or saved candidates are used. Do not extend the date window automatically.
"""
import datetime as dt
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import uuid

import openpyxl
from journal_input_controls import lesson_columns
from journal_slot_patch import extend_file

START = dt.datetime(2026, 9, 11, 19, 17, tzinfo=dt.timezone.utc)
END = dt.datetime(2026, 9, 12, 23, 0, tzinfo=dt.timezone.utc)
DRIVE = '933B1AEEA35AFA41'
TARGETS = (
    ('南教室中１英語_2026.xlsx', DRIVE + '!s0978c01028e5457a8a61391e162c3602'),
    ('南教室小６英語_2026.xlsx', DRIVE + '!sa87c7c324add4e198dbd979ea7176227'),
    ('南教室中２英語_2026.xlsx', DRIVE + '!s44c232f76a734e8999a36fd397c76888'),
)
XLSX = {'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}


def due(now):
    return START <= now < END


def item_path(item):
    return '/drives/' + DRIVE + '/items/' + urllib.parse.quote(item, safe='')


class Graph:
    def __init__(self):
        self.rclone = shutil.which('rclone')
        if not self.rclone:
            raise RuntimeError('rclone unavailable')
        self.refresh()

    def refresh(self):
        # Capture output: credentials and remote listings must never reach CI logs.
        subprocess.run([self.rclone, 'lsf', 'onedrive:', '--max-depth', '1'],
                       capture_output=True, check=True, timeout=120)
        config = json.loads(subprocess.run([self.rclone, 'config', 'dump'],
                            capture_output=True, check=True, timeout=30).stdout)
        self.token = json.loads(config['onedrive']['token'])['access_token']

    def request(self, path, method='GET', data=None, headers=None):
        for attempt in range(2):
            request = urllib.request.Request('https://graph.microsoft.com/v1.0' + path,
                data=data, method=method,
                headers={'Authorization': 'Bearer ' + self.token, **(headers or {})})
            try:
                with urllib.request.urlopen(request, timeout=90) as response:
                    return json.loads(response.read())
            except urllib.error.HTTPError as error:
                if error.code != 401 or attempt:
                    raise
                self.refresh()

    def download(self, metadata):
        # This is a signed URL, not a Graph endpoint; never add a bearer header.
        with urllib.request.urlopen(metadata['@microsoft.graph.downloadUrl'], timeout=90) as response:
            return response.read()


def target_sheets(workbook):
    return [s for s in workbook if 'TEMPLATE' in s.title or
            (re.fullmatch(r'\d{4}-\d{2}', s.title) and s.title >= '2026-09')]


def already_expanded(data):
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    try:
        sheets = target_sheets(workbook)
        return bool(sheets) and all(len(lesson_columns(s)) >= 21 for s in sheets)
    finally:
        workbook.close()


def check_excel(graph, path, names):
    sheets = graph.request(path + '/workbook/worksheets?$select=id,name')['value']
    if not set(names) <= {s['name'] for s in sheets}:
        raise ValueError('Excel worksheet mismatch')
    for name in names:
        quoted = urllib.parse.quote(name.replace("'", "''"), safe='')
        cells = graph.request(path + f"/workbook/worksheets('{quoted}')/range(address='GT1:HB12')?$select=text")['text']
        if cells[5][0] != 'クラス' or any(
            str(v).startswith(('#REF!', '#VALUE!', '#NAME?', '#DIV/0!', '#NUM!'))
            for row in cells for v in row
        ):
            raise ValueError('Excel last-slot validation failed')


def repair(graph, name, identity, directory, builder=extend_file):
    path = item_path(identity)
    metadata = graph.request(path)
    if metadata['id'] != identity or metadata['name'] != name:
        raise ValueError('Target identity changed')
    original = graph.download(metadata)
    if graph.request(path)['eTag'] != metadata['eTag']:
        return 'deferred_concurrent_edit'
    if already_expanded(original):
        workbook = openpyxl.load_workbook(io.BytesIO(original))
        try:
            check_excel(graph, path, [s.title for s in target_sheets(workbook)])
        finally:
            workbook.close()
        return 'already_expanded_and_verified'
    source = directory / 'original.xlsx'
    candidate = directory / 'expanded.xlsx'
    source.write_bytes(original)
    report = builder(source, candidate)
    if not (report['original_cells_and_caches_preserved'] and report['unrelated_parts_preserved']):
        raise ValueError('Preservation checks failed')
    names = [s['name'] for s in report['sheets']]
    if not names or not already_expanded(candidate.read_bytes()):
        raise ValueError('Capacity check failed')
    data = candidate.read_bytes()
    # A new private OneDrive backup folder per attempt; never replace old backups.
    folder = graph.request(item_path(metadata['parentReference']['id']) + '/children', 'POST',
        json.dumps({'name': '_backup_deferred21_' + uuid.uuid4().hex, 'folder': {}}).encode(),
        {'Content-Type': 'application/json'})
    folder_path = item_path(folder['id'])
    backup = graph.request(folder_path + ':/before.xlsx:/content', 'PUT', original, XLSX)
    if graph.download(graph.request(item_path(backup['id']))) != original:
        raise ValueError('Backup verification failed')
    staged = graph.request(folder_path + ':/validated.xlsx:/content', 'PUT', data, XLSX)
    staged_path = item_path(staged['id'])
    if graph.download(graph.request(staged_path)) != data:
        raise ValueError('Staging byte mismatch')
    check_excel(graph, staged_path, names)
    # Last-minute concurrency check plus atomic If-Match; no unconditional overwrite.
    current = graph.request(path)
    if current['eTag'] != metadata['eTag'] or graph.download(current) != original:
        return 'deferred_concurrent_edit'
    result = graph.request(path + '/content', 'PUT', data, {**XLSX, 'If-Match': metadata['eTag']})
    if result['id'] != identity:
        raise ValueError('Published identity mismatch')
    if graph.download(graph.request(path)) != data:
        raise ValueError('Production byte mismatch')
    check_excel(graph, path, names)
    return 'published_and_verified'


def main(now=None, graph_factory=Graph):
    now = now or dt.datetime.now(dt.timezone.utc)
    if not due(now):
        print('Deferred migration: outside the authorized retry window; no cloud access.')
        return []
    results = []
    try:
        graph = graph_factory()
        for name, identity in TARGETS:
            try:
                with tempfile.TemporaryDirectory(prefix='journal-deferred-') as directory:
                    status = repair(graph, name, identity, Path(directory))
            except urllib.error.HTTPError as error:
                if error.code not in (412, 423):
                    raise
                status = 'deferred_lock_or_concurrent_edit'
            results.append({'file': name, 'status': status})
    except Exception as error:
        # Never print exception URLs, request bodies, workbook data or credentials.
        results.append({'status': 'stopped_safely', 'error_type': type(error).__name__})
    print(json.dumps(results, ensure_ascii=False))
    summary = os.environ.get('GITHUB_STEP_SUMMARY')
    if summary:
        with open(summary, 'a', encoding='utf-8') as stream:
            stream.write('\n### Deferred 21-slot migration\n\n')
            for result in results:
                stream.write('- ' + result.get('file', 'Migration') + ': ' + result['status'] + '\n')
    # Existing journal publication continues even when a deferred file is locked.
    return results


if __name__ == '__main__':
    main()
