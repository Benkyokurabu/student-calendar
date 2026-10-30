"""Apply Bentan's server-decided publication policy before publishing JSON."""
import argparse
import hashlib
import json
from pathlib import Path
from urllib.request import urlopen
from urllib.request import Request
import os
import subprocess
from datetime import datetime, timezone, timedelta
import re
from urllib.parse import urlparse

POLICY_URL = 'https://line-check-system.vercel.app/api/recordings/publication'

def load_rules():
    with urlopen(POLICY_URL, timeout=30) as response:
        payload = json.load(response)
    rules = payload.get('rules')
    if not isinstance(rules, list) or any(r.get('status') not in ('hidden', 'public') for r in rules):
        raise ValueError('Invalid recording policy; stop publishing rather than exposing hidden URLs.')
    return rules

def apply_payload(payload, rules):
    automatic = [rule for rule in rules if rule.get('match')]
    def matched(key):
        parts = str(key or '').split('|')
        if len(parts) != 5:
            return None
        return next((r for r in automatic if r['match']['date'] == parts[0] and r['match']['campus'] == parts[2] and r['match']['group'] == parts[3]), None)
    # Preserve original URLs in private storage BEFORE stripping a newly generated test recording.
    # Failure stops publication; an unchecked recording is never uploaded as a fallback.
    candidates = {}
    def collect(value, key=None):
        if isinstance(value, list):
            for child in value:
                collect(child)
        elif isinstance(value, dict):
            if matched(key):
                url = value.get('url') or value.get('recordingUrl')
                if isinstance(url, str) and url.startswith('https://'):
                    candidates[key] = {'key': key, 'url': url, 'lesson': {k: v for k, v in value.items() if k in ('date', 'time', 'campus', 'room', 'label', 'grade', 'class', 'subject')}}
            for child_key, child in value.items():
                collect(child, child_key if '|' in child_key else None)
    collect(payload)
    if candidates:
        token = os.environ.get('GITHUB_TOKEN') or os.environ.get('GH_TOKEN')
        # Installation GITHUB_TOKENs do not expose repository.permissions.push.
        # A signed, short-lived Actions identity proves the allowed workflow.
        if os.environ.get('GITHUB_ACTIONS') == 'true':
            oidc_url = os.environ.get('ACTIONS_ID_TOKEN_REQUEST_URL')
            oidc_request_token = os.environ.get('ACTIONS_ID_TOKEN_REQUEST_TOKEN')
            if not oidc_url or not oidc_request_token:
                raise ValueError('GitHub Actions publisher identity unavailable; stop publication.')
            request = Request(oidc_url + '&audience=bentan-recording-capture', headers={'Authorization': 'Bearer ' + oidc_request_token})
            with urlopen(request, timeout=30) as response:
                token = json.load(response).get('value')
            if not token:
                raise ValueError('GitHub Actions publisher identity missing; stop publication.')
        elif not token:
            credential = subprocess.run(['git', 'credential', 'fill'], input='protocol=https\nhost=github.com\n\n', capture_output=True, text=True, timeout=30)
            token = next((line.split('=', 1)[1] for line in credential.stdout.splitlines() if line.startswith('password=')), None)
        if not token:
            raise ValueError('Calendar publisher credentials unavailable; test recording publication stopped.')
        recordings = list(candidates.values())
        for start in range(0, len(recordings), 100):
            batch = recordings[start:start + 100]
            request = Request(POLICY_URL.replace('/publication', '/capture'), data=json.dumps({'recordings': batch}).encode(), headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'}, method='POST')
            with urlopen(request, timeout=60) as response:
                result = json.load(response)
            if result.get('captured') != len(batch):
                raise ValueError('Original test recording storage was not confirmed; stop publication.')
        rules[:] = load_rules()
        automatic = [rule for rule in rules if rule.get('match')]
    by_key = {key: rule for rule in rules for key in rule['eventKeys']}
    by_hash = {digest: rule for rule in rules for digest in rule.get('urlHashes', [])}

    def visit(value, event_key=None):
        if isinstance(value, list):
            for child in value:
                visit(child)
        elif isinstance(value, dict):
            rule = by_key.get(value.get('recordingPublicationKey') or event_key) or matched(event_key)
            fields = [field for field in ('url', 'recordingUrl') if field in value]
            if rule is None:
                for field in fields:
                    url = value.get(field)
                    if isinstance(url, str) and url:
                        rule = by_hash.get(hashlib.sha256(url.encode()).hexdigest())
                        if rule:
                            break
            if rule and fields:
                if rule['status'] == 'hidden':
                    for field in fields:
                        value[field] = ''
                    value['hidden'] = True
                    value['recordingPublicationKey'] = rule['key']
                else:
                    for field in fields:
                        if not value[field] or value.get('hidden'):
                            value[field] = rule['url']
                    value.pop('hidden', None)
                    value.pop('recordingPublicationKey', None)
            for key, child in list(value.items()):
                visit(child, key if '|' in key else None)
    visit(payload)
    return payload

def filter_repository(root, rules):
    changed = []
    # Historical aliases and prevEntry references must not leak the hidden URL.
    files = set(root.glob('zoom_recording_*.json')) | set(root.glob('recording_overrides_*.json')) | set(root.glob('journal_*.json'))
    for path in sorted(files):
        payload = json.loads(path.read_text('utf-8'))
        before = json.dumps(payload, sort_keys=True, ensure_ascii=False)
        apply_payload(payload, rules)
        if json.dumps(payload, sort_keys=True, ensure_ascii=False) != before:
            if isinstance(payload, dict):
                payload['generatedAt'] = datetime.now(timezone.utc).isoformat()
            path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + '\n', 'utf-8')
            changed.append(path.name)
    return changed


def preserve_recording_fields(value, baseline):
    """Publish journal text during an outage without publishing any new URL."""
    if isinstance(value, list):
        old = baseline if isinstance(baseline, list) else []
        for index, child in enumerate(value):
            preserve_recording_fields(child, old[index] if index < len(old) else None)
    elif isinstance(value, dict):
        old = baseline if isinstance(baseline, dict) else {}
        identity = ('date', 'time', 'campus', 'groupKey', 'room')
        if any(value.get(field) != old.get(field) for field in identity):
            old = {}
        for field in ('url', 'recordingUrl'):
            if field in value:
                value[field] = old.get(field, '')
        for field in ('hidden', 'recordingPublicationKey'):
            if field in old:
                value[field] = old[field]
            else:
                value.pop(field, None)
        for key, child in value.items():
            preserve_recording_fields(child, old.get(key))
    return value


def publish_text_with_existing_recordings(root):
    changed = []
    for path in sorted(root.glob('journal_*.json')):
        # The committed public JSON is the authority during an outage. Never
        # obtain a fallback URL from the newly extracted Excel or an empty policy.
        result = subprocess.run(['git', 'show', f'HEAD:{path.name}'], cwd=root,
                                capture_output=True, check=False)
        baseline = json.loads(result.stdout) if result.returncode == 0 else {}
        payload = json.loads(path.read_text('utf-8'))
        before = json.dumps(payload, sort_keys=True, ensure_ascii=False)
        preserve_recording_fields(payload, baseline)
        if json.dumps(payload, sort_keys=True, ensure_ascii=False) != before:
            path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + '\n', 'utf-8')
            changed.append(path.name)
    return changed


def refresh_approved_snapshots(root, rules, now=None):
    """Refresh outage access only after a successful authoritative policy read."""
    now = now or datetime.now(timezone.utc)
    changed = []
    hidden_keys = {key for rule in rules if rule['status'] != 'public' for key in rule['eventKeys']}
    hidden_hashes = {digest for rule in rules if rule['status'] != 'public' for digest in rule.get('urlHashes', [])}
    release_path = root / 'recording_release.json'
    release = json.loads(release_path.read_text('utf-8')) if release_path.exists() else {}
    hidden_keys.update(key for rule in release.get('rules', []) if rule.get('status') != 'released' for key in rule['eventKeys'])
    for source in sorted(root.glob('zoom_recording_urls_*.json')):
        match = re.fullmatch(r'zoom_recording_urls_(20\d{2}-\d{2})\.json', source.name)
        if not match:
            continue
        month = match[1]
        payload = json.loads(source.read_text('utf-8'))
        entries = {}
        for key, row in payload.get('entries', {}).items():
            if not isinstance(row, dict) or row.get('hidden') or key in hidden_keys:
                continue
            parts = key.split('|')
            if len(parts) != 5 or not parts[0].startswith(month + '-'):
                continue
            if any(rule['status'] != 'public' and rule.get('match') == {'date':parts[0], 'campus':parts[2], 'group':parts[3]} for rule in rules):
                continue
            url = row.get('url', '')
            parsed = urlparse(url)
            if (parsed.scheme != 'https' or not parsed.hostname
                    or not (parsed.hostname == 'zoom.us' or parsed.hostname.endswith('.zoom.us'))
                    or parsed.username or parsed.password or not re.match(r'^/rec/(play|share)/', parsed.path)
                    or hashlib.sha256(url.encode()).hexdigest() in hidden_hashes):
                continue
            entries[key] = url
        target = root / f'recording_emergency_{month}.json'
        previous = json.loads(target.read_text('utf-8')) if target.exists() else {}
        try:
            renew_at = datetime.fromisoformat(previous.get('generatedAt', previous.get('approvedAt', '')).replace('Z', '+00:00'))
            expires_at = datetime.fromisoformat(previous.get('expiresAt', '').replace('Z', '+00:00'))
        except ValueError:
            renew_at = expires_at = now - timedelta(days=1)
        if (previous.get('entries') == entries and now - renew_at < timedelta(days=1)
                and expires_at - now > timedelta(days=2)):
            continue
        snapshot = {'version': 1, 'month': month, 'approvedAt': now.isoformat(),
                    'generatedAt': now.isoformat(), 'expiresAt': (now + timedelta(days=7)).isoformat(),
                    'sourceGeneratedAt': payload.get('generatedAt'), 'entries': entries}
        target.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2) + '\n', 'utf-8')
        changed.append(target.name)
    return changed

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parent.parent)
    parser.add_argument('--preserve-recordings', action='store_true',
                        help='Publish journal text independently of the policy service, retaining only committed recording URLs and restrictions.')
    parser.add_argument('--preserve-recordings-on-outage', action='store_true',
                        help='Publish journal text with only previously committed recording URLs if the policy service is unavailable.')
    args = parser.parse_args()
    if args.preserve_recordings:
        print('Recording preservation:', len(publish_text_with_existing_recordings(args.root)), 'JSON files updated; no policy service required')
        raise SystemExit(0)
    try:
        rules = load_rules()
    except OSError:
        if not args.preserve_recordings_on_outage:
            raise
        print('Recording policy unavailable: journal text can update; all recording URLs and restrictions stay at their committed values.')
        print('Recording preservation:', len(publish_text_with_existing_recordings(args.root)), 'JSON files updated')
    else:
        print('Recording publication filter:', len(filter_repository(args.root, rules)), 'JSON files updated')
        print('Approved outage snapshots:', len(refresh_approved_snapshots(args.root, rules)), 'JSON files updated')
