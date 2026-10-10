#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate Zoom recording URL JSON and publish it to GitHub Pages.

This does not write to lesson journal Excel files. It only updates the static
JSON files used by lesson_prep.html and pushes those files to GitHub.
"""

from __future__ import annotations

import argparse
import atexit
import copy
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


def merge_preserving_published_entries(generated: dict, existing: dict | None) -> dict:
    """Add newly discovered recordings without rotating or dropping published URLs."""
    if not existing or existing.get("month") != generated.get("month"):
        return generated
    existing_entries = existing.get("entries")
    generated_entries = generated.get("entries")
    if not isinstance(existing_entries, dict) or not isinstance(generated_entries, dict):
        return generated

    merged = dict(generated)
    # A Zoom API response can stop returning an older recurring-meeting instance,
    # and play URLs can rotate between calls. Published URLs remain authoritative.
    # The publication filter edits entry dictionaries in place. Keep the
    # comparison source intact so a changed rule cannot look unchanged.
    entries = copy.deepcopy({**generated_entries, **existing_entries})
    merged["entries"] = entries
    merged["matched"] = len(entries)
    total = int(generated.get("matched", 0)) + int(generated.get("missing", 0))
    merged["missing"] = max(0, total - len(entries))
    return merged


def load_existing_payload(path: Path) -> dict | None:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def validate_payload(payload: dict, month: str, existing: dict | None = None) -> None:
    entries = payload.get("entries")
    if payload.get("month") != month or not isinstance(entries, dict):
        raise RuntimeError("generated Zoom payload has an invalid month or entries object")
    if payload.get("matched") != len(entries):
        raise RuntimeError("generated Zoom payload matched count does not equal entry count")
    invalid = [
        key for key, value in entries.items()
        if not isinstance(value, dict)
        or not str(value.get("date", "")).startswith(month + "-")
        or not (str(value.get("url", "")).startswith("https://") or (value.get("hidden") is True and value.get("url") == "" and value.get("recordingPublicationKey") == key))
    ]
    if invalid:
        raise RuntimeError(f"generated Zoom payload contains invalid entries: {invalid[:5]}")
    if existing and existing.get("month") == month and isinstance(existing.get("entries"), dict):
        lost = sorted(set(existing["entries"]) - set(entries))
        if lost:
            raise RuntimeError(
                f"Zoom API result regressed and would remove {len(lost)} published entries; "
                f"publication stopped. examples={lost[:3]}"
            )
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


def rebase_in_progress(repo: Path) -> bool:
    git_dir = repo / ".git"
    return (git_dir / "rebase-merge").exists() or (git_dir / "rebase-apply").exists()


def abort_rebase(repo: Path) -> None:
    """Never leave the shared publishing clone poisoned after a failed rebase."""
    if rebase_in_progress(repo):
        aborted = run(["git", "rebase", "--abort"], cwd=repo, check=False, timeout=30)
        if aborted.returncode != 0:
            print("[ERROR] failed to abort rebase; manual repair is required", flush=True)


def pull_rebase_with_zoom_resolution(repo: Path, payload_text: str, month: str) -> None:
    """Rebase and resolve only the two generated Zoom JSON files.

    GitHub Actions also updates these generated files, so a push race can produce
    a content conflict.  The payload fetched in this run is authoritative.  Any
    conflict involving another file is deliberately left untouched and aborted.
    """
    allowed = {
        f"zoom_recording_urls_{month}.json",
        "zoom_recording_urls_latest.json",
    }
    pulled = run(["git", "pull", "--rebase", "origin", "main"], cwd=repo, check=False, timeout=60)
    if pulled.returncode == 0:
        return

    try:
        while rebase_in_progress(repo):
            conflicts = run(
                ["git", "diff", "--name-only", "--diff-filter=U"],
                cwd=repo,
                check=False,
                timeout=20,
            )
            names = {line.strip().replace("\\", "/") for line in conflicts.stdout.splitlines() if line.strip()}
            if not names or not names.issubset(allowed):
                raise RuntimeError(f"rebase conflict outside generated Zoom JSON: {sorted(names)}")

            print(f"[recover] resolving generated Zoom JSON conflict: {sorted(names)}", flush=True)
            for name in names:
                (repo / name).write_text(payload_text, encoding="utf-8")
            run(["git", "add", *sorted(names)], cwd=repo, timeout=20)
            committed = run(
                ["git", "commit", "--no-edit"],
                cwd=repo,
                check=False,
                timeout=60,
            )
            if committed.returncode != 0:
                if committed.stdout:
                    print(committed.stdout.strip(), flush=True)
                if committed.stderr:
                    print(committed.stderr.strip(), flush=True)
                raise RuntimeError("failed to commit resolved Zoom JSON conflict")
            continued = run(
                ["git", "rebase", "--continue"],
                cwd=repo,
                check=False,
                timeout=60,
            )
            if continued.returncode == 0:
                return
            if continued.stdout:
                print(continued.stdout.strip(), flush=True)
            if continued.stderr:
                print(continued.stderr.strip(), flush=True)
        raise RuntimeError("git pull --rebase failed without a recoverable rebase state")
    except BaseException:
        abort_rebase(repo)
        raise


def push_pending_commits(repo: Path, payload_text: str, month: str) -> bool:
    """Push an earlier local commit before deciding the current payload is unchanged."""
    fetch = run(["git", "fetch", "origin", "main"], cwd=repo, check=False, timeout=45)
    if fetch.returncode != 0:
        print("[WARN] could not fetch origin/main; continuing with normal publish", flush=True)
        return False
    pending = run(
        ["git", "rev-list", "--count", "origin/main..HEAD"],
        cwd=repo,
        check=False,
        timeout=20,
    )
    if pending.returncode != 0 or not pending.stdout.strip() or int(pending.stdout.strip()) == 0:
        return False

    print(f"[publish] {pending.stdout.strip()} unpushed commit(s) detected", flush=True)
    for attempt in range(1, 4):
        print(f"[publish] push attempt {attempt}/3", flush=True)
        pushed = run(["git", "push", "origin", "main"], cwd=repo, check=False, timeout=60)
        if pushed.returncode == 0:
            print("[publish] pending commit(s) pushed", flush=True)
            return True
        if attempt < 3:
            pull_rebase_with_zoom_resolution(repo, payload_text, month)
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

    from recording_publication_filter import apply_payload, load_rules
    publication_rules = load_rules()
    payload = url_list.make_recording_json(month)
    out = url_list.SYSTEM_DIR / f"zoom_recording_urls_{month}.json"
    latest = url_list.SYSTEM_DIR / "zoom_recording_urls_latest.json"
    update_latest = month == url_list.z.determine_latest_schedule_month()
    outputs = (out, latest) if update_latest else (out,)

    repo = url_list.repo_dir()
    if repo is None:
        print("[ERROR] student-calendar repo was not found.")
        return 1

    existing = load_existing_payload(repo / out.name) or load_existing_payload(out)
    payload = apply_payload(merge_preserving_published_entries(payload, existing), publication_rules)
    validate_payload(payload, month, existing)
    payload_text = json.dumps(payload, ensure_ascii=False, indent=2)

    pending_pushed = push_pending_commits(repo, payload_text, month)

    # Start each new publication from the latest remote state.  Do this before
    # copying generated files so pull --rebase never sees our new unstaged data.
    pull_rebase_with_zoom_resolution(repo, payload_text, month)

    existing = load_existing_payload(repo / out.name) or load_existing_payload(out)
    payload = apply_payload(merge_preserving_published_entries(payload, existing), publication_rules)
    validate_payload(payload, month, existing)
    payload_text = json.dumps(payload, ensure_ascii=False, indent=2)
    if existing and comparable_payload(existing) == comparable_payload(payload):
        if pending_pushed:
            print("[publish] pending commit was recovered; no new payload changes")
        else:
            print(f"[publish] no recording URL changes matched={payload['matched']} missing={payload['missing']}")
        return 0

    if args.dry_run:
        print(f"[dry-run] recording URL JSON would change matched={payload['matched']} missing={payload['missing']}")
        return 0

    out.write_text(payload_text, encoding="utf-8")
    if update_latest:
        latest.write_text(payload_text, encoding="utf-8")
    print(f"[write] {out.name} matched={payload['matched']} missing={payload['missing']}")

    for src in outputs:
        dst = repo / src.name
        shutil.copy2(src, dst)
        print(f"[copy] {dst.name}")

    run(["git", "add", *[file.name for file in outputs]], cwd=repo)

    diff = run(["git", "diff", "--cached", "--quiet"], cwd=repo, check=False, timeout=20)
    if diff.returncode == 0:
        print("[publish] no changes")
        return 0

    run(["git", "commit", "-m", f"Update Zoom recording URLs {month}"], cwd=repo, timeout=60)
    pushed = run(["git", "push", "origin", "main"], cwd=repo, check=False, timeout=60)
    if pushed.returncode != 0:
        print("[WARN] initial push failed; retrying after rebase", flush=True)
        for attempt in range(2, 4):
            pull_rebase_with_zoom_resolution(repo, payload_text, month)
            time.sleep(5 * (attempt - 1))
            pushed = run(["git", "push", "origin", "main"], cwd=repo, check=False, timeout=60)
            if pushed.returncode == 0:
                break
    if pushed.returncode != 0:
        raise RuntimeError("Zoom URLコミットのpushに3回失敗しました")
    print("[publish] pushed", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
