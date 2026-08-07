#!/usr/bin/env python3
"""Validate that the journal files downloaded by rclone match the remote listing."""

import argparse
import json
import subprocess
import sys
from datetime import date, datetime, timezone
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--remote", required=True)
    parser.add_argument("--local-dir", type=Path, required=True)
    parser.add_argument("--max-age-hours", type=float, default=48)
    return parser.parse_args()


def remote_listing(remote):
    result = subprocess.run(
        [
            "rclone", "lsjson", remote, "--recursive", "--files-only",
            "--include", "*.xlsx", "--exclude", "_backup/**", "--exclude", "退避/**",
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    return json.loads(result.stdout or "[]")


def parse_time(value):
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def has_current_month_sheet(path):
    """Return whether an xlsx contains the current month's journal sheet."""
    try:
        import openpyxl
        workbook = openpyxl.load_workbook(path, read_only=True)
        current = date.today().strftime("%Y-%m")
        found = current in workbook.sheetnames
        workbook.close()
        return found
    except Exception as exc:
        print(f"[FRESHNESS][WARN] Could not inspect month sheets in {path.name}: {exc}")
        return False


def main():
    args = parse_args()
    entries = remote_listing(args.remote)
    now = datetime.now(timezone.utc)
    problems = []
    checked = 0

    for item in entries:
        remote_path = item.get("Path", "")
        if not remote_path.lower().endswith(".xlsx"):
            continue
        local_path = args.local_dir / Path(remote_path)
        if not local_path.is_file():
            # rclone may flatten a single remote directory; fall back to basename.
            matches = list(args.local_dir.rglob(Path(remote_path).name))
            local_path = matches[0] if len(matches) == 1 else local_path

        remote_size = item.get("Size")
        local_size = local_path.stat().st_size if local_path.is_file() else None
        remote_mtime = parse_time(item["ModTime"]) if item.get("ModTime") else None
        local_mtime = datetime.fromtimestamp(local_path.stat().st_mtime, timezone.utc) if local_path.is_file() else None
        age_hours = (now - remote_mtime).total_seconds() / 3600 if remote_mtime else None
        checked += 1

        print(
            f"[FRESHNESS] {remote_path} | remote_size={remote_size} local_size={local_size} "
            f"remote_mtime={remote_mtime.isoformat() if remote_mtime else 'unknown'} "
            f"local_mtime={local_mtime.isoformat() if local_mtime else 'missing'}"
        )

        if not local_path.is_file():
            problems.append(f"missing local file: {remote_path}")
            continue
        if remote_size is not None and local_size != remote_size:
            problems.append(f"size mismatch: {remote_path} remote={remote_size} local={local_size}")
        if remote_mtime and local_mtime:
            delta = abs((local_mtime - remote_mtime).total_seconds())
            if delta > 300:
                problems.append(f"mtime mismatch: {remote_path} delta={delta:.0f}s")
            if age_hours > args.max_age_hours:
                print(f"[FRESHNESS][WARN] remote file is {age_hours:.1f} hours old: {remote_path}")
                if has_current_month_sheet(local_path):
                    problems.append(
                        f"stale current-month workbook: {remote_path} age={age_hours:.1f}h"
                    )

    if checked == 0:
        print("[FRESHNESS][ERROR] No journal xlsx files found in remote listing.")
        return 1
    if problems:
        print("[FRESHNESS][ERROR] Downloaded journal files do not match the remote listing:")
        for problem in problems:
            print(f"  - {problem}")
        return 2
    print(f"[FRESHNESS] OK: checked {checked} journal xlsx files")
    return 0


if __name__ == "__main__":
    sys.exit(main())
