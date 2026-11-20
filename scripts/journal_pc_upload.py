"""PC journal snapshots use the same conditional-write guard as cloud jobs."""
import json
import os
import subprocess
import tempfile
import uuid
from contextlib import contextmanager
from pathlib import Path

from safe_journal_upload import (GraphClient, UploadSafetyError, apply_manifest,
    capture_manifest, digest, listing_identities, local_file,
    verify_listing_unchanged, workbook_paths)

MANIFEST = '.journal-original-manifest.json'
SESSIONS = {}


def isolated_download_dir():
    return Path(tempfile.mkdtemp(prefix='benkyo_club_cloud_journal_'))


@contextmanager
def operation_lock(root):
    path = root / '.journal-operation.lock'
    try:
        fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    except FileExistsError:
        raise UploadSafetyError('Another journal operation owns this folder; use a new download folder') from None
    try:
        os.write(fd, str(os.getpid()).encode('ascii'))
        os.close(fd)
        yield
    finally:
        path.unlink(missing_ok=True)


def rclone_json(executable, *args):
    result = subprocess.run([executable, *args], check=True, capture_output=True,
                            text=True, encoding='utf-8', timeout=120)
    return json.loads(result.stdout)


def graph_client(executable, remote):
    name = remote.split(':', 1)[0]
    def token():
        config = rclone_json(executable, 'config', 'dump')[name]
        return json.loads(config['token'])['access_token']
    def refresh():
        rclone_json(executable, 'about', name + ':', '--json')
        return token()
    return GraphClient(token(), refresh=refresh)


def logical_path(value):
    return tuple(part.replace('\u201b\uff1a', ':').replace('\uff1a', ':')
                 for part in value.replace('\\', '/').split('/'))


def mapped_listing(executable, remote, root):
    rows = rclone_json(executable, 'lsjson', remote, '--recursive', '--files-only',
                      '--include', '*.xlsx', '--exclude', '**/_backup*/**', '--exclude', '**/退避/**')
    # Validate the actual remote paths and immutable identities before mapping
    # rclone's Windows colon encoding to the corresponding downloaded paths.
    identities = listing_identities(rows)
    local = {}
    for relative in workbook_paths(root):
        key = logical_path(relative)
        if key in local:
            raise UploadSafetyError('Ambiguous Windows workbook path')
        local[key] = relative
    mapped, seen = [], set()
    for relative in identities:
        key = logical_path(relative)
        if key not in local or key in seen:
            raise UploadSafetyError('Remote workbook does not have one exact downloaded counterpart')
        seen.add(key)
        drive, item = identities[relative]
        mapped.append({'Path': local[key], 'ID': drive + '#' + item})
    return mapped


def freeze_download(executable, remote, root):
    client = graph_client(executable, remote)
    manifest = capture_manifest(client, root, mapped_listing(executable, remote, root), remote)
    verify_listing_unchanged(manifest, mapped_listing(executable, remote, root))
    manifest['localRoot'] = str(root.resolve())
    manifest['session'] = uuid.uuid4().hex
    target = root / MANIFEST
    pending = root / (MANIFEST + '.tmp')
    pending.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    pending.replace(target)
    SESSIONS[str(root.resolve())] = manifest['session']
    return manifest


def probe_parents(executable, remote, entries):
    configured = os.environ.get('JOURNAL_UPLOAD_PROBE_PARENTS')
    if configured:
        return json.loads(configured)
    # Reuse only an already registered isolated probe; never create a new
    # folder in the shared originals or use a normal workbook directory.
    rows = rclone_json(executable, 'lsjson', remote, '--dirs-only', '--max-depth', '1')
    candidates = {}
    for row in rows:
        if row.get('IsDir') and row.get('Name', '').startswith('.bentan-cas-probe-'):
            value = row.get('ID', '')
            if '#' not in value:
                raise UploadSafetyError('Probe folder has no shared-drive identity')
            drive, item = value.split('#', 1)
            if drive in candidates:
                raise UploadSafetyError('Ambiguous registered probe folders')
            candidates[drive] = item
    required = {entry['driveId'] for entry in entries.values()}
    if not required.issubset(candidates):
        raise UploadSafetyError('No registered conditional-write probe; journal upload blocked')
    return {drive: candidates[drive] for drive in required}


def guarded_upload(executable, remote, root):
    target = root / MANIFEST
    if not target.is_file():
        raise UploadSafetyError('No frozen original snapshot; download journals again before uploading')
    manifest = json.loads(target.read_text(encoding='utf-8'))
    if manifest.get('remote') != remote or manifest.get('localRoot') != str(root.resolve()):
        raise UploadSafetyError('Snapshot belongs to another source or folder')
    session = SESSIONS.get(str(root.resolve()))
    if session is not None and session != manifest.get('session'):
        raise UploadSafetyError('Another download superseded this processing session')
    verify_listing_unchanged(manifest, mapped_listing(executable, remote, root))
    changed = any(digest(local_file(root, relative).read_bytes()) != entry['sha256']
                  for relative, entry in manifest['entries'].items())
    parents = probe_parents(executable, remote, manifest['entries']) if changed else {}
    # Consume before any PUT. A lost response must never lead to a blind retry.
    consumed = root / (MANIFEST + '.consumed')
    target.replace(consumed)
    result = apply_manifest(graph_client(executable, remote), root, manifest, parents)
    verify_listing_unchanged(manifest, mapped_listing(executable, remote, root))
    return result
