"""Apply Bentan's server-decided publication policy before publishing JSON."""
import argparse
import hashlib
import json
from pathlib import Path
from urllib.request import urlopen
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
    by_key = {key: rule for rule in rules for key in rule['eventKeys']}
    by_hash = {digest: rule for rule in rules for digest in rule.get('urlHashes', [])}

    def visit(value, event_key=None):
        if isinstance(value, list):
            for child in value:
                visit(child)
        elif isinstance(value, dict):
            rule = by_key.get(value.get('recordingPublicationKey') or event_key)
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
