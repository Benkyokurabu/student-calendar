"""Apply Bentan's server-decided publication policy before publishing JSON."""
import argparse
import hashlib
import json
from pathlib import Path
from urllib.request import urlopen
from urllib.request import Request
import os
import subprocess
from datetime import datetime, timezone

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

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parent.parent)
    args = parser.parse_args()
    print('Recording publication filter:', len(filter_repository(args.root, load_rules())), 'JSON files updated')
