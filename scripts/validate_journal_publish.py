#!/usr/bin/env python3
"""授業日誌JSONを公開する前に、明らかな欠損・取り違えを検出する。"""

from __future__ import annotations

import argparse
import json
from pathlib import Path


DATA_FIELDS = ("content", "page", "report", "absence", "note")


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def has_journal_data(entry: dict) -> bool:
    return any(entry.get(field) for field in DATA_FIELDS) or bool(entry.get("homework"))


def validate(candidate_path: Path, schedule_path: Path, baseline_path: Path | None = None) -> list[str]:
    errors: list[str] = []
    candidate = load_json(candidate_path)
    schedule = load_json(schedule_path)

    month = str(candidate.get("month", ""))
    entries = candidate.get("entries")
    if not month or not isinstance(entries, dict):
        return ["month または entries がありません"]

    expected_keys = {
        "|".join([
            str(ev.get("date", "")),
            str(ev.get("time", "")).replace("~", "～"),
            str(ev.get("campus", "")),
            str(ev.get("groupKey", "")),
            str(ev.get("room", "")),
        ])
        for ev in schedule
        if str(ev.get("date", "")).startswith(month + "-")
    }
    actual_keys = set(entries)
    missing = expected_keys - actual_keys
    extra = actual_keys - expected_keys
    if missing:
        errors.append(f"スケジュール枠が {len(missing)} 件欠落")
    if extra:
        errors.append(f"スケジュール外の枠が {len(extra)} 件混入")
    if len(actual_keys) != len(expected_keys):
        errors.append(f"枠数不一致: expected={len(expected_keys)} actual={len(actual_keys)}")

    malformed = 0
    for entry in entries.values():
        if not isinstance(entry, dict):
            malformed += 1
            continue
        if not all(field in entry for field in ("teacher", "sessionNumber", "monthNum", "weekNum")):
            malformed += 1
    if malformed:
        errors.append(f"必須項目が不足した枠: {malformed} 件")

    if baseline_path and baseline_path.exists():
        baseline = load_json(baseline_path)
        if baseline.get("month") == month and isinstance(baseline.get("entries"), dict):
            baseline_filled = sum(has_journal_data(v) for v in baseline["entries"].values() if isinstance(v, dict))
            candidate_filled = sum(has_journal_data(v) for v in entries.values() if isinstance(v, dict))
            if baseline_filled >= 5 and candidate_filled < baseline_filled * 0.7:
                errors.append(
                    f"入力済み件数が急減: baseline={baseline_filled} candidate={candidate_filled}"
                )

    return errors


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate", required=True, type=Path)
    parser.add_argument("--schedule", required=True, type=Path)
    parser.add_argument("--baseline", type=Path)
    args = parser.parse_args()

    errors = validate(args.candidate, args.schedule, args.baseline)
    if errors:
        print("[BLOCK] 授業日誌JSONを公開できません")
        for error in errors:
            print(f"  - {error}")
        raise SystemExit(1)
    print("[OK] 公開前検査に合格しました")


if __name__ == "__main__":
    main()
