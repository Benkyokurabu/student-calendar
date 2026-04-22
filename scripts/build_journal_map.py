# -*- coding: utf-8 -*-
"""
最小版 journal_map 生成

目的
- extract_journal_to_json.py が使う
  journal_map_latest.json / journal_map_YYYY-MM.json
  を作る

方針
- export_by_grade_subject.py の collect_events と同じ並びを使う
- export_schedule_json.py の groupKey 変換を使う
- つまり「授業日誌Excelを作る順番」と同じ順番で map を作る

使い方
python build_journal_map.py
python build_journal_map.py --prefer 4
python build_journal_map.py --schedule "2026年4月スケジュール.xlsm"
"""

from __future__ import annotations

import argparse
import json
import re
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

import openpyxl

# 既存の「本体」と「schedule変換」をそのまま使う
import export_by_grade_subject as legacy
import export_schedule_json as sched


def make_entry_key(ev: dict) -> str:
    return "|".join(
        [
            str(ev.get("date", "")),
            str(ev.get("time", "")).replace("~", "～").strip(),
            str(ev.get("campus", "")),
            str(ev.get("groupKey", "")),
            str(ev.get("room", "")),
        ]
    )


def resolve_schedule_path(arg_schedule: str | None, prefer: str | None):
    if arg_schedule:
        p = Path(arg_schedule)
        if not p.exists():
            raise FileNotFoundError(f"スケジュールが見つかりません: {p}")
        m = re.match(r"^(\d{4})年0?(\d{1,2})月スケジュール\.xlsm$", p.name)
        if not m:
            raise ValueError("スケジュール名は 'YYYY年M月スケジュール.xlsm' 形式にしてください。")
        return int(m.group(1)), int(m.group(2)), p

    prefer_months = [int(x) for x in re.split(r"[,\s]+", prefer.strip()) if x] if prefer else None
    return legacy.pick_schedule_in_same_folder(prefer_months)


def build_slots(year: int, month: int, schedule_path: Path):
    wb_s = openpyxl.load_workbook(schedule_path, data_only=True, keep_vba=True)
    targets = legacy.choose_target_sheets(wb_s)
    if not targets:
        raise RuntimeError("本校/南教室の教務部用シートが見つかりません。")

    slots = OrderedDict()

    for campus_label, _sheetname, sh in targets:
        # 授業日誌本体と同じイベント抽出順
        events = legacy.collect_events(sh, month, campus_label, year, month)

        # schedule JSON と同じ groupKey / campus / subject コードへ変換
        rows = sched.convert_events_to_schedule_json(events, campus_label, year, month)

        for ev in rows:
            gk = str(ev.get("groupKey", "")).strip()
            if not gk:
                continue

            key = make_entry_key(ev)

            if gk not in slots:
                slots[gk] = []

            # 同じ key は重複追加しない
            if key not in slots[gk]:
                slots[gk].append(key)

    wb_s.close()
    return slots


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--schedule", help="対象スケジュール.xlsm のフルパス")
    ap.add_argument("--prefer", help="優先月（例: 4 または 4,3）")
    args = ap.parse_args()

    year, month, schedule_path = resolve_schedule_path(args.schedule, args.prefer)
    target_month = f"{year:04d}-{month:02d}"

    print(f"[INFO] スケジュール: {schedule_path} ({target_month})")

    slots = build_slots(year, month, schedule_path)

    out = {
        "month": target_month,
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "slots": slots,
    }

    latest_path = Path("journal_map_latest.json")
    month_path = Path(f"journal_map_{target_month}.json")

    text = json.dumps(out, ensure_ascii=False, indent=2)
    latest_path.write_text(text, encoding="utf-8")
    month_path.write_text(text, encoding="utf-8")

    print(f"[OK] 出力: {latest_path.resolve()}")
    print(f"[OK] 出力: {month_path.resolve()}")
    print(f"[INFO] month: {target_month}")
    print(f"[INFO] group count: {len(slots)}")


if __name__ == "__main__":
    main()