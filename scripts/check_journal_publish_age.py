#!/usr/bin/env python3
"""前回の日誌公開時刻を検査し、定期実行漏れを警告ファイルへ残す。"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path


def parse_generated_at(value: str) -> datetime:
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        parsed = parsed.astimezone()
    return parsed.astimezone(timezone.utc)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--journal", type=Path, required=True)
    parser.add_argument("--max-age-hours", type=float, default=26)
    parser.add_argument("--warning-file", type=Path, required=True)
    args = parser.parse_args()

    warnings = []
    try:
        payload = json.loads(args.journal.read_text(encoding="utf-8"))
        generated = parse_generated_at(str(payload.get("generatedAt") or ""))
        age = (datetime.now(timezone.utc) - generated).total_seconds() / 3600
        print(f"[PUBLISH_AGE] generatedAt={generated.isoformat()} age={age:.1f}h")
        if age > args.max_age_hours:
            warnings.append(
                f"前回の日誌JSON生成から {age:.1f} 時間経過しています（上限 {args.max_age_hours:.1f} 時間）"
            )
    except Exception as exc:
        warnings.append(f"前回の日誌JSON生成時刻を確認できません: {exc}")

    args.warning_file.parent.mkdir(parents=True, exist_ok=True)
    args.warning_file.write_text("\n".join(warnings), encoding="utf-8")
    if warnings:
        for warning in warnings:
            print(f"[PUBLISH_AGE][WARN] {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
