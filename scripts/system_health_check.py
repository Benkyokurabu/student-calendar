#!/usr/bin/env python3
"""End-to-end health checks for the journal/calendar publishing system."""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import re
import subprocess
import sys
import time
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from student_calendar_repo import resolve as resolve_repo


SYSTEM_DIR = Path(__file__).resolve().parent
LOG_DIR = SYSTEM_DIR / "logs"
STATUS_PATH = LOG_DIR / "system_health_latest.json"
ALERT_PATH = SYSTEM_DIR / "【要確認】授業日誌システム異常.txt"
PUBLIC_BASE = "https://benkyokurabu.github.io/student-calendar"
JST = timezone(timedelta(hours=9))


def show_windows_alert(message: str) -> None:
    """Show an error dialog without starting a console process."""
    message_box = ctypes.windll.user32.MessageBoxW
    message_box(None, message, "授業日誌システム異常", 0x00000010 | 0x00001000)


def fetch_json(name: str) -> dict[str, Any] | list[Any]:
    url = f"{PUBLIC_BASE}/{name}?health={int(time.time())}"
    request = urllib.request.Request(url, headers={"Cache-Control": "no-cache"})
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8-sig"))


def parse_datetime(value: str) -> datetime | None:
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return parsed if parsed.tzinfo else parsed.replace(tzinfo=JST)
    except (TypeError, ValueError):
        return None


TIME_RE = re.compile(r"(\d{1,2}):(\d{2})")


def lesson_end(event: dict[str, Any]) -> datetime | None:
    raw = str(event.get("time", "")).translate(str.maketrans("０１２３４５６７８９：", "0123456789:"))
    times = TIME_RE.findall(raw)
    try:
        day = datetime.strptime(str(event.get("date", ""))[:10], "%Y-%m-%d").date()
        hour, minute = map(int, times[1])
    except (ValueError, IndexError):
        return None
    if 1 <= hour <= 11:
        hour += 12
    return datetime(day.year, day.month, day.day, hour, minute, tzinfo=JST)


def event_key(event: dict[str, Any]) -> str:
    lesson_time = str(event.get("time") or "").replace("~", "～").strip()
    return "|".join([
        str(event.get("date") or ""), lesson_time,
        str(event.get("campus") or ""), str(event.get("groupKey") or ""),
        str(event.get("room") or ""),
    ])


def git(repo: Path, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", *args], cwd=repo, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=30,
    )


def last_log_exit(path: Path) -> tuple[datetime | None, int | None]:
    if not path.exists():
        return None, None
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    for line in reversed(lines):
        if "] END exit=" not in line:
            continue
        try:
            stamp = line.split("]", 1)[0].lstrip("[")
            code = int(line.rsplit("=", 1)[1].strip())
            return datetime.strptime(stamp, "%Y/%m/%d %H:%M:%S.%f").replace(tzinfo=JST), code
        except (ValueError, IndexError):
            continue
    return None, None


def zoom_failure_needs_alert(
    log_time: datetime | None, exit_code: int | None, public_generated_at: datetime | None
) -> bool:
    """A later successful public update supersedes a local task failure."""
    if exit_code in (None, 0):
        return False
    return not (log_time and public_generated_at and public_generated_at > log_time)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--notify", action="store_true", help="Show a one-time Windows message for a new error set.")
    args = parser.parse_args()
    now = datetime.now(JST)
    errors: list[str] = []
    warnings: list[str] = []
    details: dict[str, Any] = {"checkedAt": now.isoformat()}
    public_zoom_generated_at: datetime | None = None

    try:
        repo = resolve_repo(required=True)
        details["repo"] = str(repo)
        if "OneDrive" in str(repo):
            warnings.append("公開用GitリポジトリがOneDrive内にあります")
        if (repo / ".git" / "rebase-merge").exists() or (repo / ".git" / "rebase-apply").exists():
            errors.append("公開用Gitリポジトリに未完了rebaseがあります")
        status = git(repo, "status", "--porcelain", "--untracked-files=no")
        if status.returncode != 0:
            errors.append("公開用Gitリポジトリの状態を読み取れません")
        elif status.stdout.strip():
            errors.append("公開用Gitリポジトリに未保存の追跡対象変更があります")
        ahead = git(repo, "rev-list", "--count", "origin/main..HEAD")
        if ahead.returncode == 0:
            details["unpushedCommits"] = int(ahead.stdout.strip())
            if int(ahead.stdout.strip()) > 0:
                errors.append(f"未pushコミットが{ahead.stdout.strip()}件あります")
    except Exception as exc:
        errors.append(f"Git確認失敗: {exc}")

    try:
        month = now.strftime("%Y-%m")
        zoom = fetch_json(f"zoom_recording_urls_{month}.json")
        if not isinstance(zoom, dict):
            raise ValueError("Zoom JSONの形式が不正です")
        generated = parse_datetime(str(zoom.get("generatedAt", "")))
        public_zoom_generated_at = generated
        today_entries = [v for v in (zoom.get("entries") or {}).values() if v.get("date") == now.date().isoformat()]
        details["publicZoomGeneratedAt"] = generated.isoformat() if generated else None
        details["todayZoomEntries"] = len(today_entries)
        if not generated:
            errors.append("公開Zoom JSONに有効な生成時刻がありません")

        schedule = fetch_json(f"schedule_{month}.json")
        if not isinstance(schedule, list) or not schedule:
            raise ValueError("当月の公開スケジュールが空または不正です")
        invalid_schedule_dates = [
            str(event.get("date", "")) for event in schedule
            if not str(event.get("date", "")).startswith(month + "-")
        ]
        if invalid_schedule_dates:
            errors.append(f"当月スケジュールに別月の日付があります: {invalid_schedule_dates[:3]}")
        details["publicScheduleEvents"] = len(schedule)

        # 授業日誌JSONは現在の運用では使用していないため、監視対象外とする。
        details["journalMonitoring"] = "ignored: journal JSON is not used in current operations"

        meeting_ids = json.loads((SYSTEM_DIR / "zoomURL" / "zoom_meeting_ids.json").read_text(encoding="utf-8-sig"))
        published_keys = set((zoom.get("entries") or {}).keys())
        expected: list[dict[str, Any]] = []
        for event in schedule if isinstance(schedule, list) else []:
            if event.get("date") != now.date().isoformat() or bool(event.get("faceToFace")):
                continue
            campus, room = str(event.get("campus", "")), str(event.get("room", ""))
            if not (meeting_ids.get(campus) or {}).get(room):
                continue
            end = lesson_end(event)
            if end and now >= end + timedelta(minutes=60):
                expected.append(event)
        missing = [event_key(event) for event in expected if event_key(event) not in published_keys]
        details["endedOnlineLessonsExpected"] = len(expected)
        details["endedOnlineLessonsMissing"] = len(missing)
        if missing:
            errors.append(f"終了後60分を過ぎたオンライン授業のZoom URLが{len(missing)}件ありません: {missing[:3]}")
    except Exception as exc:
        errors.append(f"公開カレンダーデータ確認失敗: {exc}")

    zoom_log_time, zoom_exit = last_log_exit(SYSTEM_DIR / "zoomURL" / "logs" / "zoom_recording_json.log")
    details["zoomTaskLastEnd"] = zoom_log_time.isoformat() if zoom_log_time else None
    details["zoomTaskLastExit"] = zoom_exit
    details["zoomTaskFailureSupersededByPublic"] = bool(
        zoom_exit not in (None, 0)
        and zoom_log_time
        and public_zoom_generated_at
        and public_zoom_generated_at > zoom_log_time
    )
    if zoom_failure_needs_alert(zoom_log_time, zoom_exit, public_zoom_generated_at):
        errors.append(f"Zoom公開タスクの直近終了コードが{zoom_exit}です")

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    previous_signature = ""
    if STATUS_PATH.exists():
        try:
            previous_signature = str(json.loads(STATUS_PATH.read_text(encoding="utf-8")).get("errorSignature", ""))
        except Exception:
            pass
    signature = "\n".join(sorted(errors))
    report = {
        "ok": not errors,
        "errors": errors,
        "warnings": warnings,
        "details": details,
        "errorSignature": signature,
    }
    STATUS_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    if errors:
        message = "授業日誌システムで異常を検出しました。\n\n" + "\n".join(f"・{item}" for item in errors)
        ALERT_PATH.write_text(message + f"\n\n確認時刻: {now:%Y-%m-%d %H:%M:%S}\n", encoding="utf-8")
        if args.notify and signature != previous_signature:
            show_windows_alert(message)
        print(message)
        return 1

    if ALERT_PATH.exists():
        ALERT_PATH.unlink()
    print(f"[OK] system healthy at {now:%Y-%m-%d %H:%M:%S}; today Zoom entries={details.get('todayZoomEntries', 0)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
