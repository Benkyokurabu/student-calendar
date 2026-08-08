#!/usr/bin/env python3
"""Validate that the journal files downloaded by rclone match the remote listing."""

import argparse
import json
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--remote", required=True)
    parser.add_argument("--local-dir", type=Path, required=True)
    parser.add_argument("--max-age-hours", type=float, default=48)
    parser.add_argument("--warning-file", type=Path)
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
        encoding="utf-8",
        errors="replace",
    )
    return json.loads(result.stdout or "[]")


def parse_time(value):
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def is_backup_path(value: str) -> bool:
    parts = value.replace("\\", "/").split("/")[:-1]
    return any(part.startswith("_backup") or part.startswith("退避") for part in parts)


def main():
    args = parse_args()
    entries = remote_listing(args.remote)
    now = datetime.now(timezone.utc)
    integrity_problems = []
    age_warnings = []
    checked = 0

    for item in entries:
        remote_path = item.get("Path", "")
        if not remote_path.lower().endswith(".xlsx"):
            continue
        if is_backup_path(remote_path):
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
            integrity_problems.append(f"missing local file: {remote_path}")
            continue
        if remote_size is not None and local_size != remote_size:
            integrity_problems.append(f"size mismatch: {remote_path} remote={remote_size} local={local_size}")
        if remote_mtime and local_mtime:
            delta = abs((local_mtime - remote_mtime).total_seconds())
            if delta > 300:
                integrity_problems.append(f"mtime mismatch: {remote_path} delta={delta:.0f}s")
            if age_hours > args.max_age_hours:
                print(f"[FRESHNESS][WARN] remote file is {age_hours:.1f} hours old: {remote_path}")
                # 更新のない教科・クラスは正常に48時間を超える。年齢だけで公開を止めず、
                # Issue用の警告として残す。remote/local不一致だけを再取得対象にする。
                age_warnings.append(f"old remote workbook: {remote_path} age={age_hours:.1f}h")

    if checked == 0:
        print("[FRESHNESS][ERROR] No journal xlsx files found in remote listing.")
        return 1
    if args.warning_file:
        args.warning_file.parent.mkdir(parents=True, exist_ok=True)
        args.warning_file.write_text("\n".join(age_warnings), encoding="utf-8")
    if integrity_problems:
        print("[FRESHNESS][ERROR] Downloaded journal files do not match the remote listing:")
        for problem in integrity_problems:
            print(f"  - {problem}")
        return 2
    print(f"[FRESHNESS] OK: checked {checked} journal xlsx files; warnings={len(age_warnings)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
