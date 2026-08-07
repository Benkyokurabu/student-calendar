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
import time
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


def run(args: list[str], cwd: Path, *, check: bool = True, timeout: int = 60) -> subprocess.CompletedProcess[str]:
    print(f"[run] {' '.join(args)}", flush=True)
    try:
        result = subprocess.run(
            args,
            cwd=str(cwd),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        print(f"[ERROR] command timed out after {timeout}s: {' '.join(args)}", flush=True)
        if check:
            raise SystemExit(124)
        return subprocess.CompletedProcess(args, 124, "", "timeout")
    if check and result.returncode != 0:
        if result.stdout:
            print(result.stdout.strip())
        if result.stderr:
            print(result.stderr.strip())
        raise SystemExit(result.returncode)
    return result


def push_pending_commits(repo: Path) -> bool:
    """Recover a commit left locally after a previous push interruption."""
    fetched = run(["git", "fetch", "origin", "main"], cwd=repo, check=False, timeout=45)
    if fetched.returncode != 0:
        return False
    pending = run(["git", "rev-list", "--count", "origin/main..HEAD"], cwd=repo, check=False, timeout=20)
    if pending.returncode != 0 or not pending.stdout.strip() or int(pending.stdout.strip()) == 0:
        return False
    print(f"[publish] recovering {pending.stdout.strip()} unpushed commit(s)", flush=True)
    for attempt in range(1, 4):
        pushed = run(["git", "push", "origin", "main"], cwd=repo, check=False, timeout=60)
        if pushed.returncode == 0:
            return True
        if attempt < 3:
            run(["git", "pull", "--rebase", "origin", "main"], cwd=repo, timeout=60)
            time.sleep(5 * attempt)
    raise RuntimeError("未pushコミットのpushに3回失敗しました")


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

    pending_pushed = push_pending_commits(repo)

    existing = load_existing_payload(repo / out.name) or load_existing_payload(out)
    if existing and comparable_payload(existing) == comparable_payload(payload):
        if pending_pushed:
            print("[publish] pending commit recovered; no new payload changes")
        else:
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

    run(["git", "pull", "--rebase", "origin", "main"], cwd=repo, check=False, timeout=60)
    run(["git", "add", f"zoom_recording_urls_{month}.json", "zoom_recording_urls_latest.json"], cwd=repo, timeout=20)

    diff = run(["git", "diff", "--cached", "--quiet"], cwd=repo, check=False, timeout=20)
    if diff.returncode == 0:
        print("[publish] no changes")
        return 0

    run(["git", "commit", "-m", f"Update Zoom recording URLs {month}"], cwd=repo, timeout=60)
    push = run(["git", "push", "origin", "main"], cwd=repo, check=False, timeout=60)
    if push.returncode != 0:
        print("[WARN] initial push failed; retrying after rebase")
        for attempt in range(2, 4):
            run(["git", "pull", "--rebase", "origin", "main"], cwd=repo, timeout=60)
            time.sleep(5 * (attempt - 1))
            push = run(["git", "push", "origin", "main"], cwd=repo, check=False, timeout=60)
            if push.returncode == 0:
                break
    if push.returncode != 0:
        raise RuntimeError("Zoom URLコミットのpushに3回失敗しました")
    print("[publish] pushed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
