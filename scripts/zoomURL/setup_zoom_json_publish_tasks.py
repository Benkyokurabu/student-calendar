#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Register scheduled tasks that publish Zoom recording URL JSON."""

from __future__ import annotations

import subprocess
from datetime import datetime, timedelta
from pathlib import Path

TASK_PREFIX = "BenkyoZoomRecordingURLJson"
LESSON_END_TIMES = ["16:25", "18:15", "20:05", "21:55"]
OFFSETS_MINUTES = [5, 10, 15, 20, 25, 30]


def run_times() -> list[tuple[int, int]]:
    base_day = datetime(2000, 1, 1)
    times: list[tuple[int, int]] = []
    for end_time in LESSON_END_TIMES:
        hour, minute = map(int, end_time.split(":"))
        base = base_day.replace(hour=hour, minute=minute)
        for offset in OFFSETS_MINUTES:
            target = base + timedelta(minutes=offset)
            times.append((target.hour, target.minute))
    return times


def create_task(hour: int, minute: int) -> bool:
    script_dir = Path(__file__).resolve().parent
    launcher_path = script_dir / "scheduled_zoom_recording_url_json_publish_hidden.vbs"
    if not launcher_path.exists():
        print(f"[ERROR] not found: {launcher_path}")
        return False

    time_str = f"{hour:02d}:{minute:02d}"
    task_name = f"{TASK_PREFIX}_{hour:02d}{minute:02d}"
    args = [
        "schtasks", "/create",
        "/tn", task_name,
        "/tr", f'wscript.exe "{launcher_path}"',
        "/sc", "DAILY",
        "/st", time_str,
        "/rl", "LIMITED",
        "/f",
    ]
    result = subprocess.run(args, capture_output=True, text=True, encoding="cp932", errors="replace")
    if result.returncode == 0:
        print(f"[OK] {time_str} {task_name}")
        return True
    print(f"[NG] {time_str} {task_name}: {result.stderr.strip() or result.stdout.strip()}")
    return False


def main() -> int:
    times = run_times()
    print("Zoom録画URL JSON公開タスク登録")
    print("実行時刻: " + " / ".join(f"{hour:02d}:{minute:02d}" for hour, minute in times))
    registered = sum(create_task(hour, minute) for hour, minute in times)
    if registered != len(times):
        print(f"[WARN] {registered}/{len(times)} 件のみ登録されました。")
        return 1
    print("[DONE] 全タスクを登録しました。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
