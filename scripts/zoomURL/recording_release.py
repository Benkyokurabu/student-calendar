"""Server-side release control. Each rule targets a recording occurrence UUID."""
import argparse
import json
import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import quote
import zoom_recording_urls as z

ROOT = Path(__file__).resolve().parents[2]
STATE = ROOT / 'recording_release.json'

def parse_release(value):
    if not value:
        return None
    result = datetime.fromisoformat(value)
    if result.tzinfo is None:
        result = result.replace(tzinfo=z.JST)
    return result.astimezone(timezone.utc)

def settings_url(uuid):
    # Zoom requires double encoding for UUIDs starting with / or containing //.
    encoded = quote(uuid, safe='')
    if uuid.startswith('/') or '//' in uuid:
        encoded = quote(encoded, safe='')
    return f'https://api.zoom.us/v2/meetings/{encoded}/recordings/settings'

def setting(client, uuid, share=None):
    headers = {'Authorization': f'Bearer {client.access_token()}'}
    if share is not None:
        headers['Content-Type'] = 'application/json'
        z.http_json('PATCH', settings_url(uuid), headers=headers,
                    data=json.dumps({'share_recording': share}).encode())
    result = z.http_json('GET', settings_url(uuid), headers=headers)
    if share is not None and result.get('share_recording') != share:
        raise RuntimeError('Zoom共有設定の反映を確認できません。成功扱いにしません。')
    return result

def desired_share(rule, now):
    if rule['mode'] == 'private':
        return 'none'
    if rule['mode'] == 'scheduled' and now < parse_release(rule['releaseAt']):
        return 'none'
    return rule['originalShare']

def resolve_recording(client, event_key):
    month = event_key[:7]
    payload = json.loads((ROOT / f'zoom_recording_urls_{month}.json').read_text('utf-8'))
    selected = payload['entries'].get(event_key)
    if not selected:
        raise ValueError('指定授業の録画がまだ見つかりません。録画取得後に設定してください。')
    start = z.parse_zoom_time(selected['recordingStart'])
    meetings = client.list_account_recordings(start.date() - timedelta(days=1), start.date() + timedelta(days=1))
    matches = []
    for meeting in meetings:
        if str(meeting.get('id')) != str(selected['meetingId']):
            continue
        candidates = z.flatten_recordings(str(meeting['id']), {'meetings': [meeting]})
        if any(r.start_time == start for r in candidates):
            matches.append(meeting)
    if len(matches) != 1 or not matches[0].get('uuid'):
        raise ValueError('録画を一意に特定できません。別の授業の共有設定は変更しません。')
    meeting = matches[0]
    urls = {r.url for r in z.flatten_recordings(str(meeting['id']), {'meetings': [meeting]})}
    keys = {event_key}
    # Include shared campus copies and restart segments belonging to this occurrence.
    for path in ROOT.glob('zoom_recording_urls_*.json'):
        entries = json.loads(path.read_text('utf-8')).get('entries', {})
        keys.update(k for k, rec in entries.items() if rec.get('url') in urls)
    return meeting['uuid'], sorted(keys)

def apply_rule(client, rule, now):
    desired = desired_share(rule, now)
    current = setting(client, rule['uuid'])
    if current.get('share_recording') != desired:
        setting(client, rule['uuid'], desired)
    rule['status'] = 'blocked' if desired == 'none' else 'released'

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--event-key', default=os.environ.get('RELEASE_EVENT_KEY', ''))
    parser.add_argument('--mode', choices=['scheduled', 'private', 'public'], default=os.environ.get('RELEASE_MODE') or None)
    parser.add_argument('--release-at', default=os.environ.get('RELEASE_AT', ''))
    args = parser.parse_args()
    state = json.loads(STATE.read_text('utf-8')) if STATE.exists() else {'rules': []}
    before = json.dumps(state, sort_keys=True)
    now = datetime.now(timezone.utc)
    if args.mode == 'scheduled':
        release = parse_release(args.release_at)
        if release is None or release <= now:
            raise ValueError('公開予約日時には未来の日本時間を指定してください。')
    if args.mode and not args.event_key:
        raise ValueError('授業キーが必要です。')
    client = z.ZoomClient()
    if args.event_key:
        if not args.mode:
            raise ValueError('公開方法が必要です。')
        uuid, keys = resolve_recording(client, args.event_key)
        rule = next((r for r in state['rules'] if r['uuid'] == uuid), None)
        if rule is None:
            original = setting(client, uuid).get('share_recording')
            if original not in ('publicly', 'internally', 'none'):
                raise ValueError('元のZoom共有設定を取得できません。')
            rule = {'uuid': uuid, 'originalShare': original, 'eventKeys': keys}
            state['rules'].append(rule)
        rule['status'] = 'pending'
        rule.update(mode=args.mode, releaseAt=args.release_at if args.mode == 'scheduled' else '')
        rule['eventKeys'] = sorted(set(rule['eventKeys']) | set(keys))
        # Explicit public overrides an originally private recording, retaining other settings.
        if args.mode == 'public':
            rule['originalShare'] = 'publicly'
        if args.mode == 'scheduled' and rule['originalShare'] == 'none':
            rule['originalShare'] = 'publicly'
    if not state['rules']:
        return
    failures = []
    try:
        for rule in state['rules']:
            try:
                apply_rule(client, rule, now)
            except Exception as error:
                rule['status'] = 'pending'
                failures.append(str(error))
    finally:
        # Persist requested rules even after an API failure, so retries retain the
        # original sharing level and a pending rule never bypasses the page gate.
        if json.dumps(state, sort_keys=True) != before:
            state['generatedAt'] = datetime.now(z.JST).isoformat()
            STATE.write_text(json.dumps(state, ensure_ascii=False, indent=2) + '\n', 'utf-8')
    if failures:
        raise RuntimeError('公開設定を確認できない録画があります。次回に再試行します。\n' + '\n'.join(failures))
    summary = os.environ.get('GITHUB_STEP_SUMMARY')
    if summary:
        with open(summary, 'a', encoding='utf-8') as out:
            out.write('録画公開設定をZoomから再取得して確認しました。\n\n')
            for rule in state['rules']:
                out.write(f"- {rule['eventKeys'][0]}: {rule['status']} / {rule['releaseAt'] or rule['mode']}\n")

if __name__ == '__main__':
    main()
