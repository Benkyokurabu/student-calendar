#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate static Zoom recording URL JSON for lesson_prep.html.

This file contains only matched recording URLs and lesson metadata. It does not
contain Zoom API credentials.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

SCRIPT_DIR = Path(__file__).resolve().parent
SYSTEM_DIR = SCRIPT_DIR.parent

sys.path.insert(0, str(SCRIPT_DIR))
import zoom_recording_urls as z

def event_key(ev: dict) -> str:
    lesson_time = str(ev.get("time") or "").replace("~", "～").strip()
    return "|".join([
        str(ev.get("date") or ""),
        lesson_time,
        str(ev.get("campus") or ""),
        str(ev.get("groupKey") or ""),
        str(ev.get("room") or ""),
    ])


def online_lesson_key(ev: dict) -> str:
    """Return a campus-independent key for one online lesson."""
    lesson_time = str(ev.get("time") or "").replace("~", "～").strip()
    return "|".join([
        str(ev.get("date") or ""),
        lesson_time,
        str(ev.get("grade") or ""),
        str(ev.get("class") or ""),
        str(ev.get("subject") or ""),
    ])


def share_same_meeting_online_lessons(
    events: list[dict],
    local_matches: Dict[str, tuple[dict, z.RecordingCandidate]],
    meeting_ids: Dict[str, Dict[str, str]],
) -> Dict[str, tuple[dict, z.RecordingCandidate]]:
    """Share one verified recording across the two campuses of the same online lesson."""
    resolved = dict(local_matches)
    by_lesson: Dict[str, list[dict]] = {}
    for ev in events:
        by_lesson.setdefault(online_lesson_key(ev), []).append(ev)

    for pair in by_lesson.values():
        if len(pair) != 2 or {ev.get("campus") for ev in pair} != {"hon", "minami"}:
            continue
        if any(ev.get("faceToFace") or ev.get("special") for ev in pair):
            continue
        teachers = {str(ev.get("teacher") or "").strip() for ev in pair}
        if len(teachers) != 1 or not next(iter(teachers)):
            continue
        ids = {z.clean_meeting_id(z.meeting_id_for_event(ev, meeting_ids) or "") for ev in pair}
        if len(ids) != 1 or not next(iter(ids)):
            continue
        sources = [resolved[event_key(ev)] for ev in pair if event_key(ev) in resolved]
        if len(sources) != 1:
            continue
        source_ev, rec = sources[0]
        if z.clean_meeting_id(rec.meeting_id) != next(iter(ids)):
            continue
        target = next(ev for ev in pair if event_key(ev) != event_key(source_ev))
        resolved[event_key(target)] = (source_ev, rec)
    return resolved


def recording_distance_seconds(ev: dict, rec: z.RecordingCandidate) -> float:
    window = z.parse_lesson_window(ev)
    if window is None:
        return float("inf")
    return abs((rec.start_time - window[0]).total_seconds())

def repo_dir() -> Optional[Path]:
    candidates = [
        SYSTEM_DIR.parent,
        SYSTEM_DIR.parent / "生徒スケジュール表",
        Path.home() / "OneDrive" / "デスクトップ" / "生徒スケジュール表",
    ]
    for p in candidates:
        if (p / ".git").exists():
            return p
    return None


def make_recording_json(month: str) -> Dict[str, Any]:
    meeting_ids = z.load_meeting_ids()
    events = z.relevant_events(z.load_schedule(month), month, meeting_ids)
    client = z.ZoomClient()
    recordings_by_id = z.fetch_recordings_for_events(client, events, meeting_ids, month)

    local_matches: Dict[str, tuple[dict, z.RecordingCandidate]] = {}
    for ev in events:
        meeting_id = z.meeting_id_for_event(ev, meeting_ids)
        rec = z.match_recording(
            ev,
            recordings_by_id.get(z.clean_meeting_id(meeting_id or ""), []),
            30,
            30,
        )
        if rec is not None:
            local_matches[event_key(ev)] = (ev, rec)

    resolved_matches = share_same_meeting_online_lessons(events, local_matches, meeting_ids)

    entries: Dict[str, dict] = {}
    matched = 0
    missing = 0
    for ev in events:
        meeting_id = z.meeting_id_for_event(ev, meeting_ids)
        resolved = resolved_matches.get(event_key(ev))
        source_ev, rec = resolved if resolved else (ev, None)
        key = event_key(ev)
        if rec is None:
            missing += 1
            continue
        matched += 1
        entries[key] = {
            "url": rec.url,
            "recordingStart": rec.start_time.isoformat(),
            "meetingId": z.clean_meeting_id(meeting_id or ""),
            "topic": rec.topic,
            "date": ev.get("date", ""),
            "time": ev.get("time", ""),
            "campus": ev.get("campus", ""),
            "room": ev.get("room", ""),
            "grade": ev.get("grade", ""),
            "class": ev.get("class", ""),
            "subject": ev.get("subject", ""),
            "label": ev.get("label") or ev.get("groupKey") or "",
            "teacher": ev.get("teacher", ""),
            "onlineLessonKey": online_lesson_key(ev),
            "recordingCampus": source_ev.get("campus", ""),
            "recordingRoom": source_ev.get("room", ""),
            **({"sharedFromEventKey": event_key(source_ev)} if source_ev is not ev else {}),
            **z.recording_match_audit(source_ev, rec),
        }

    return {
        "month": month,
        "generatedAt": datetime.now(z.JST).isoformat(),
        "matched": matched,
        "missing": missing,
        "entries": entries,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Generate zoom_recording_urls_YYYY-MM.json for lesson_prep.html.")
    ap.add_argument("--month", help="Target month, e.g. 2026-08. Defaults to latest schedule month.")
    ap.add_argument("--copy-repo", action="store_true", help="Copy generated JSON to the student-calendar repo if found.")
    args = ap.parse_args()

    month = args.month or z.determine_latest_schedule_month()
    payload = make_recording_json(month)
    out = SYSTEM_DIR / f"zoom_recording_urls_{month}.json"
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    latest = SYSTEM_DIR / "zoom_recording_urls_latest.json"
    latest.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[write] {out.name} matched={payload['matched']} missing={payload['missing']}")

    if args.copy_repo:
        repo = repo_dir()
        if repo is None:
            print("[WARN] student-calendar repo was not found.")
        else:
            for src in (out, latest):
                dst = repo / src.name
                shutil.copy2(src, dst)
                print(f"[copy] {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
