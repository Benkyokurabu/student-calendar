#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate Zoom recording URL JSON and publish it to GitHub Pages.

This does not write to lesson journal Excel files. It only updates the static
JSON files used by lesson_prep.html and pushes those files to GitHub.
"""

from __future__ import annotations

import argparse
import atexit
import json
import shutil
import subprocess
import sys
from pathlib import Path

import zoom_recording_url_list as url_list


SCRIPT_DIR = Path(__file__).resolve().parent
SYSTEM_DIR = SCRIPT_DIR.parent
if str(SYSTEM_DIR) not in sys.path:
    sys.path.insert(0, str(SYSTEM_DIR))

from publish_lock import PublishLock

def comparable_payload(payload: dict) -> dict:
    data = dict(payload)
    data.pop("generatedAt", None)
    return data


def load_existing_payload(path: Path) -> dict | None:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def run(args: list[str], cwd: Path, *, check: bool = True) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        args,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if check and result.returncode != 0:
        if result.stdout:
            print(result.stdout.strip())
        if result.stderr:
            print(result.stderr.strip())
        raise SystemExit(result.returncode)
    return result


def main() -> int:
    ap = argparse.ArgumentParser(description="Publish Zoom recording URL JSON for lesson_prep.html.")
    ap.add_argument("--month", help="Target month, e.g. 2026-08. Defaults to latest schedule month.")
    ap.add_argument("--dry-run", action="store_true", help="Fetch Zoom data and report whether JSON would change, but do not write, commit, or push.")
    args = ap.parse_args()

    publish_lock = PublishLock(
        SYSTEM_DIR / "logs" / "student_calendar_publish.lock",
        purpose="Zoom録画URL公開",
    )
    publish_lock.acquire()
    atexit.register(publish_lock.release)

    month = args.month or url_list.z.determine_latest_schedule_month()
    print(f"[publish] target month: {month}")

    payload = url_list.make_recording_json(month)
    out = url_list.SYSTEM_DIR / f"zoom_recording_urls_{month}.json"
    latest = url_list.SYSTEM_DIR / "zoom_recording_urls_latest.json"

    repo = url_list.repo_dir()
    if repo is None:
        print("[ERROR] student-calendar repo was not found.")
        return 1

    existing = load_existing_payload(repo / out.name) or load_existing_payload(out)
    if existing and comparable_payload(existing) == comparable_payload(payload):
        print(f"[publish] no recording URL changes matched={payload['matched']} missing={payload['missing']}")
        return 0

    if args.dry_run:
        print(f"[dry-run] recording URL JSON would change matched={payload['matched']} missing={payload['missing']}")
        return 0

    text = json.dumps(payload, ensure_ascii=False, indent=2)
    out.write_text(text, encoding="utf-8")
    latest.write_text(text, encoding="utf-8")
    print(f"[write] {out.name} matched={payload['matched']} missing={payload['missing']}")

    for src in (out, latest):
        dst = repo / src.name
        shutil.copy2(src, dst)
        print(f"[copy] {dst.name}")

    run(["git", "pull", "--rebase", "origin", "main"], cwd=repo, check=False)
    run(["git", "add", f"zoom_recording_urls_{month}.json", "zoom_recording_urls_latest.json"], cwd=repo)

    diff = run(["git", "diff", "--cached", "--quiet"], cwd=repo, check=False)
    if diff.returncode == 0:
        print("[publish] no changes")
        return 0

    run(["git", "commit", "-m", f"Update Zoom recording URLs {month}"], cwd=repo)
    push = run(["git", "push"], cwd=repo, check=False)
    if push.returncode != 0:
        print("[WARN] initial push failed; retrying after rebase")
        if push.stderr:
            print(push.stderr.strip())
        run(["git", "pull", "--rebase", "origin", "main"], cwd=repo)
        run(["git", "push"], cwd=repo)
    print("[publish] pushed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
