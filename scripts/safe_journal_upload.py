#!/usr/bin/env python3
"""Compare-and-swap journal uploads. Never use local save time as authority.

The read manifest is made before workbook maintenance. Source byte identity and
immutable DriveItem IDs are checked; every write carries the original eTag.
The workflow must stop publication on any conflict or uncertain write result.
"""
import argparse
import hashlib
import json
import os
import subprocess
import sys
import uuid
from pathlib import Path, PurePosixPath
from urllib.parse import quote
from urllib.request import Request, build_opener, HTTPRedirectHandler, urlopen
from urllib.error import HTTPError, URLError

SCHEMA = 1
MAX_BYTES = 250_000_000 - 1
PROBE_PREFIX = '.bentan-cas-probe-'


class UploadSafetyError(RuntimeError):
    pass


class GraphError(UploadSafetyError):
    def __init__(self, status):
        self.status = status
        super().__init__(f'Graph request failed (HTTP {status}); no unguarded retry')


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


def digest(data):
    return hashlib.sha256(data).hexdigest()


class GraphClient:
    def __init__(self, token, refresh=None):
        self.token = token
        self.refresh = refresh
        self.opener = build_opener(NoRedirect())

    @classmethod
    def from_rclone(cls, remote='onedrive'):
        def config():
            result = subprocess.run(['rclone', 'config', 'dump'], check=True,
                                    capture_output=True, text=True, encoding='utf-8')
            data = json.loads(result.stdout)[remote]
            return json.loads(data['token'])['access_token']

        def refresh():
            # Let the existing authenticated rclone refresh its own token.
            subprocess.run(['rclone', 'about', remote + ':', '--json'], check=True,
                           capture_output=True, timeout=60)
            return config()
        return cls(config(), refresh=refresh)

    def request(self, method, route, body=None, headers=None, binary=False):
        if not route.startswith('/drives/'):
            raise UploadSafetyError('Only scoped DriveItem requests are allowed')
        if isinstance(body, dict):
            body = json.dumps(body).encode('utf-8')
            content_type = 'application/json'
        else:
            content_type = 'application/octet-stream'
        request_headers = {'Authorization': 'Bearer ' + self.token,
                           'Content-Type': content_type, **(headers or {})}
        request = Request('https://graph.microsoft.com/v1.0' + route, data=body,
                          method=method, headers=request_headers)
        try:
            with self.opener.open(request, timeout=60) as response:
                status, data = response.status, response.read()
        except HTTPError as error:
            if method == 'GET' and error.code in (301, 302, 303, 307, 308):
                # Signed content URLs must never receive the OAuth header.
                location = error.headers.get('Location', '')
                if not location.startswith('https://'):
                    raise UploadSafetyError('Unsafe content redirect') from None
                try:
                    with urlopen(location, timeout=60) as response:
                        status, data = response.status, response.read()
                except (HTTPError, URLError, TimeoutError, OSError):
                    raise UploadSafetyError('Content download failed; signed URL omitted') from None
            elif error.code == 401 and self.refresh is not None:
                refresh, self.refresh = self.refresh, None
                self.token = refresh()
                return self.request(method, route, body,
                                    {'Content-Type': content_type, **(headers or {})}, binary)
            else:
                raise GraphError(error.code) from None
        except (URLError, TimeoutError, OSError):
            # A lost PUT response may already have committed. Never retry it.
            raise UploadSafetyError('Transport failed; publication blocked pending reconciliation') from None
        return data if binary else (json.loads(data) if data else {}), status

    @staticmethod
    def route(drive_id, item_id):
        return '/drives/' + quote(drive_id, safe='') + '/items/' + quote(item_id, safe='')

    def metadata(self, drive_id, item_id):
        return self.request('GET', self.route(drive_id, item_id))[0]

    def download(self, drive_id, item_id):
        return self.request('GET', self.route(drive_id, item_id) + '/content', binary=True)[0]

    def put_if_match(self, drive_id, item_id, content, etag):
        if not etag or etag == '*' or not content or len(content) > MAX_BYTES:
            raise UploadSafetyError('Missing frozen eTag or unsupported content size')
        return self.request('PUT', self.route(drive_id, item_id) + '/content', content,
                            {'If-Match': etag})[0]


def check_identity(meta, drive_id, item_id):
    if meta.get('id') != item_id or meta.get('parentReference', {}).get('driveId', '').casefold() != drive_id.casefold():
        raise UploadSafetyError('DriveItem identity mismatch')
    if 'file' not in meta or not meta.get('eTag'):
        raise UploadSafetyError('Missing file identity/version')


def read_original(client, drive_id, item_id):
    """Read bytes under a stable version; suitable for targeted restoration."""
    before = client.metadata(drive_id, item_id)
    check_identity(before, drive_id, item_id)
    data = client.download(drive_id, item_id)
    after = client.metadata(drive_id, item_id)
    check_identity(after, drive_id, item_id)
    if before['eTag'] != after['eTag'] or len(data) != after.get('size'):
        raise UploadSafetyError('Original changed while being read')
    return data, after


def replace_existing_if_unchanged(client, drive_id, item_id, original_etag,
                                  original_sha256, payload):
    """Replace only the exact original version. No blind retry or fallback."""
    current, meta = read_original(client, drive_id, item_id)
    if meta['eTag'] != original_etag or digest(current) != original_sha256:
        raise UploadSafetyError('Original was edited; guarded upload refused')
    if digest(payload) == original_sha256:
        return meta
    updated = client.put_if_match(drive_id, item_id, payload, original_etag)
    check_identity(updated, drive_id, item_id)
    received, verified = read_original(client, drive_id, item_id)
    if verified['eTag'] != updated['eTag'] or digest(received) != digest(payload):
        raise UploadSafetyError('Write result was superseded or uncertain; publication blocked')
    return verified


def local_file(root, relative):
    pure = PurePosixPath(relative.replace('\\', '/'))
    if pure.is_absolute() or '..' in pure.parts or not pure.parts:
        raise UploadSafetyError('Invalid relative workbook path')
    target = root.joinpath(*pure.parts)
    if target.is_symlink() or not target.resolve().is_relative_to(root.resolve()):
        raise UploadSafetyError('Workbook escapes local workspace')
    return target


def workbook_paths(root):
    return sorted(p.relative_to(root).as_posix() for p in root.rglob('*.xlsx')
                  if not any(part.startswith(('_backup', '退避')) for part in p.relative_to(root).parts[:-1]))


def split_rclone_id(value):
    if '#' not in value:
        raise UploadSafetyError('rclone must supply driveID#itemID, including shared drive identity')
    drive, item = value.split('#', 1)
    if not drive or not item:
        raise UploadSafetyError('Invalid rclone DriveItem identity')
    return drive, item


def listing_identities(listing):
    """Reject ambiguous paths and aliases before accepting a remote snapshot."""
    identities, seen_paths, seen_items = {}, set(), set()
    for listed in listing:
        relative = listed.get('Path', '')
        if not relative.lower().endswith('.xlsx') or any(part.startswith(('_backup', '退避')) for part in PurePosixPath(relative).parts[:-1]):
            continue
        path = PurePosixPath(relative)
        if path.is_absolute() or '..' in path.parts or path.as_posix() != relative or '\\' in relative:
            raise UploadSafetyError('Invalid relative workbook path in remote listing')
        drive, item = split_rclone_id(listed.get('ID', ''))
        identity = (drive.casefold(), item)
        if relative.casefold() in seen_paths or identity in seen_items:
            raise UploadSafetyError('Duplicate workbook path or DriveItem identity in remote listing')
        seen_paths.add(relative.casefold())
        seen_items.add(identity)
        identities[relative] = (drive, item)
    return identities


def verify_listing_unchanged(manifest, listing):
    expected = {relative: (entry['driveId'].casefold(), entry['itemId'])
                for relative, entry in manifest['entries'].items()}
    actual = {relative: (drive.casefold(), item)
              for relative, (drive, item) in listing_identities(listing).items()}
    if actual != expected:
        raise UploadSafetyError('Remote workbook paths or identities changed during capture')


def capture_manifest(client, root, listing, remote):
    entries = {}
    for relative, (drive, item) in listing_identities(listing).items():
        target = local_file(root, relative)
        if not target.is_file():
            raise UploadSafetyError('Downloaded workbook is missing: ' + relative)
        original, meta = read_original(client, drive, item)
        if meta.get('name') != PurePosixPath(relative).name:
            raise UploadSafetyError('Original was renamed since remote listing: ' + relative)
        if digest(target.read_bytes()) != digest(original):
            raise UploadSafetyError('Downloaded workbook does not match frozen remote bytes: ' + relative)
        entries[relative] = {'driveId': drive, 'itemId': item, 'eTag': meta['eTag'],
                             'sha256': digest(original), 'size': len(original)}
    if not entries or set(entries) != set(workbook_paths(root)):
        raise UploadSafetyError('Manifest must cover every downloaded workbook exactly')
    return {'schema': SCHEMA, 'remote': remote, 'entries': entries}


def verify_versions(client, entries, expected_etags=None):
    for relative, entry in entries.items():
        meta = client.metadata(entry['driveId'], entry['itemId'])
        check_identity(meta, entry['driveId'], entry['itemId'])
        if meta['eTag'] != (expected_etags or {}).get(relative, entry['eTag']):
            raise UploadSafetyError('Remote edited since download: ' + relative)


def probe_conditional_put(client, drive_id, parent_id):
    """Test only a newly created UUID file in a dedicated registered scratch."""
    parent = client.metadata(drive_id, parent_id)
    if 'folder' not in parent or not parent.get('name', '').startswith(PROBE_PREFIX):
        raise UploadSafetyError('Conditional PUT probe requires registered isolated scratch folder')
    name = uuid.uuid4().hex + '.txt'
    initial, current = b'BENTAN-CAS-INITIAL', b'BENTAN-CAS-CURRENT'
    route = client.route(drive_id, parent_id) + ':/' + name + ':/content'
    item, _ = client.request('PUT', route, initial)
    item_id, old_etag = item['id'], item['eTag']
    check_identity(item, drive_id, item_id)
    try:
        try:
            client.put_if_match(drive_id, item_id, b'WRONG-MUST-NOT-PUBLISH', '"wrong-' + uuid.uuid4().hex + '"')
        except GraphError as error:
            if error.status != 412:
                raise
        else:
            raise UploadSafetyError('Conditional PUT is ignored; original uploads disabled')
        if client.download(drive_id, item_id) != initial:
            raise UploadSafetyError('Rejected conditional PUT changed bytes')
        updated = client.put_if_match(drive_id, item_id, current, old_etag)
        if client.download(drive_id, item_id) != current:
            raise UploadSafetyError('Current conditional PUT not verified')
        try:
            client.put_if_match(drive_id, item_id, b'STALE-MUST-NOT-PUBLISH', old_etag)
        except GraphError as error:
            if error.status != 412:
                raise
        else:
            raise UploadSafetyError('Stale eTag write succeeded; original uploads disabled')
        if client.download(drive_id, item_id) != current:
            raise UploadSafetyError('Stale conditional PUT changed bytes')
        # Delete only this newly created ID, with current bytes and eTag checked.
        meta = client.metadata(drive_id, item_id)
        if meta.get('name') != name or meta['eTag'] != updated['eTag']:
            raise UploadSafetyError('Probe was edited; leave it intact for review')
        client.request('DELETE', client.route(drive_id, item_id), headers={'If-Match': meta['eTag']})
    except Exception:
        # Retain unexpected probe state as evidence, never sweep a folder.
        raise


def verify_manifest(client, root, manifest, probe_parents):
    """Exercise the live guard without replacing any original workbook."""
    if manifest.get('schema') != SCHEMA or not manifest.get('entries'):
        raise UploadSafetyError('Missing original download manifest')
    entries = manifest['entries']
    if set(workbook_paths(root)) != set(entries):
        raise UploadSafetyError('Workbook coverage changed')
    verify_versions(client, entries)
    for relative, entry in entries.items():
        if digest(local_file(root, relative).read_bytes()) != entry['sha256']:
            raise UploadSafetyError('Validation requires unchanged downloaded originals')
    for drive in sorted({entry['driveId'] for entry in entries.values()}):
        parent = probe_parents.get(drive)
        if not parent:
            raise UploadSafetyError('No verified scratch folder registered')
        probe_conditional_put(client, drive, parent)
    verify_versions(client, entries)
    return {'checked': len(entries), 'uploaded': 0, 'guardVerified': True}


def apply_manifest(client, root, manifest, probe_parents):
    if manifest.get('schema') != SCHEMA or not manifest.get('entries'):
        raise UploadSafetyError('Missing original download manifest')
    entries = manifest['entries']
    if set(workbook_paths(root)) != set(entries):
        raise UploadSafetyError('New/deleted workbook has no frozen original; publication stopped')
    verify_versions(client, entries)
    changed = [relative for relative, entry in entries.items()
               if digest(local_file(root, relative).read_bytes()) != entry['sha256']]
    for drive in sorted({entries[p]['driveId'] for p in changed}):
        parent = probe_parents.get(drive)
        if not parent:
            raise UploadSafetyError('No verified scratch folder registered; original uploads disabled')
        probe_conditional_put(client, drive, parent)
    etags = {}
    for relative in changed:
        entry = entries[relative]
        meta = replace_existing_if_unchanged(client, entry['driveId'], entry['itemId'],
                    entry['eTag'], entry['sha256'], local_file(root, relative).read_bytes())
        etags[relative] = meta['eTag']
        print('[SAFE_UPLOAD] Updated unchanged original: ' + relative, flush=True)
    # Even an untouched workbook edited during maintenance makes its extracted
    # JSON stale. Never publish a mixed old/new extraction silently.
    verify_versions(client, entries, etags)
    return {'checked': len(entries), 'uploaded': len(changed), 'unchanged': len(entries)-len(changed)}


def remote_listing(remote):
    result = subprocess.run(['rclone','lsjson',remote,'--recursive','--files-only',
                             '--include','*.xlsx','--exclude','**/_backup*/**','--exclude','**/退避/**'],
                            check=True,capture_output=True,text=True,encoding='utf-8')
    return json.loads(result.stdout)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('operation', choices=('capture','apply','verify'))
    parser.add_argument('--remote', required=True)
    parser.add_argument('--local-dir', required=True, type=Path)
    parser.add_argument('--manifest', required=True, type=Path)
    args = parser.parse_args()
    remote_name = args.remote.split(':',1)[0]
    client = GraphClient.from_rclone(remote_name)
    if args.operation == 'capture':
        manifest = capture_manifest(client,args.local_dir,remote_listing(args.remote),args.remote)
        # A stable item can move/rename between listing and its byte read. Freeze
        # only after a second complete listing confirms the path-to-ID mapping.
        verify_listing_unchanged(manifest, remote_listing(args.remote))
        args.manifest.parent.mkdir(parents=True,exist_ok=True)
        args.manifest.write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
        print(f'[SAFE_UPLOAD] Frozen {len(manifest["entries"])} original workbook versions',flush=True)
    else:
        manifest = json.loads(args.manifest.read_text(encoding='utf-8'))
        if manifest.get('remote') != args.remote:
            raise UploadSafetyError('Manifest belongs to another source')
        parents = json.loads(os.environ.get('JOURNAL_UPLOAD_PROBE_PARENTS') or '{}')
        operation = verify_manifest if args.operation == 'verify' else apply_manifest
        print(json.dumps(operation(client,args.local_dir,manifest,parents)),flush=True)
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except (UploadSafetyError, subprocess.CalledProcessError, ValueError, KeyError) as error:
        # Do not print captured command output or credentials.
        print('[SAFE_UPLOAD][BLOCKED] ' + (str(error) if isinstance(error,UploadSafetyError) else type(error).__name__), file=sys.stderr, flush=True)
        sys.exit(2)
