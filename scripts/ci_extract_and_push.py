#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Actions (CI) 用パイプライン。
rclone でダウンロード済みのファイルを使って JSON を生成する。
git commit/push はワークフロー側で行うため、ここでは行わない。
"""

import calendar
import json
import subprocess
import sys
import os
from datetime import date
from pathlib import Path
import shutil


def get_months_to_process(month_arg=None):
    if month_arg:
        return [month_arg]
    today = date.today()
    last_day = calendar.monthrange(today.year, today.month)[1]
    current_month = f"{today.year}-{today.month:02d}"
    if today.day > last_day - 5:
        if today.month == 1:
            prev_month = f"{today.year - 1}-12"
        else:
            prev_month = f"{today.year}-{today.month - 1:02d}"
        return [prev_month, current_month]
    return [current_month]


def run(args, **kwargs):
    print(f"  > {' '.join(str(a) for a in args)}")
    result = subprocess.run(args, **kwargs)
    if result.returncode != 0:
        print(f"[ERROR] コマンド失敗 (exit {result.returncode})")
        sys.exit(1)
    return result


def main():
    script_dir = Path(__file__).parent.resolve()
    repo_dir = script_dir.parent  # リポのルート
    month_arg = sys.argv[1] if len(sys.argv) > 1 else None
    months = get_months_to_process(month_arg)

    journal_dir = script_dir / "_cloud_journal"
    if not journal_dir.exists():
        print("[ERROR] _cloud_journal フォルダがありません。ワークフローでダウンロードしてください。")
        sys.exit(1)

    print("=" * 48)
    print("  CI: スケジュール＋授業日誌 JSON 生成")
    print(f"  対象月: {', '.join(months)}")
    print(f"  日誌フォルダ: {journal_dir}")
    print("=" * 48)
    print()

    # --- 1. スケジュールJSON再生成 ---
    print("[1/3] スケジュールExcelから JSON を生成中...")
    export_script = script_dir / "export_schedule_json.py"
    if export_script.exists():
        run([sys.executable, str(export_script)], cwd=str(script_dir))
        for f in script_dir.glob("schedule_*.json"):
            dst = repo_dir / f.name
            shutil.copy2(f, dst)
            print(f"  → {dst.name}")
    else:
        print("[SKIP] export_schedule_json.py が見つかりません")
    print()

    # --- 2. 日誌キャンパス間コピー + JSON抽出 ---
    # 日誌JSON退避（マージ用）
    old_jsons = {}
    for m in months:
        repo_journal = repo_dir / f"journal_{m}.json"
        if repo_journal.exists():
            old_jsons[m] = json.loads(repo_journal.read_text(encoding="utf-8"))

    for m in months:
        print(f"--- {m} の処理 ---")

        # 日誌キャンパス間コピー
        print(f"[2/3] オンラインペア授業の日誌をキャンパス間でコピー中... ({m})")
        sync_script = script_dir / "sync_journal_across_campus.py"
        if sync_script.exists():
            run([sys.executable, str(sync_script), "--month", m,
                 "--journal-dir", str(journal_dir)])
        print()

        # 授業日誌JSON抽出
        print(f"[3/3] 授業日誌Excelから JSON を抽出中... ({m})")
        run([sys.executable, str(script_dir / "extract_journal_to_json.py"),
             "--month", m, "--journal-dir", str(journal_dir)])
        print()

    # --- マージ ---
    FIELDS = ["content", "page", "report", "recordingUrl",
              "teacher", "absence", "curriculumProgress", "note",
              "sessionNumber", "monthNum", "weekNum"]

    def entry_has_data(entry):
        if not isinstance(entry, dict):
            return False
        return any(entry.get(f) for f in FIELDS) or bool(entry.get("homework"))

    def merge_entry(old_entry, new_entry):
        merged = {}
        for f in FIELDS:
            new_val = new_entry.get(f, "")
            old_val = old_entry.get(f, "")
            merged[f] = new_val if new_val else old_val
        new_hw = new_entry.get("homework", [])
        old_hw = old_entry.get("homework", [])
        merged["homework"] = new_hw if new_hw else old_hw
        return merged

    print("[MERGE] 日誌データのマージ中...")
    for m in months:
        if m not in old_jsons:
            print(f"  {m}: 前回データなし（初回生成）— そのまま採用")
            continue
        repo_journal = repo_dir / f"journal_{m}.json"
        if not repo_journal.exists():
            continue
        try:
            new_data = json.loads(repo_journal.read_text(encoding="utf-8"))
            old_data = old_jsons[m]
            new_entries = new_data.get("entries", {})
            old_entries = old_data.get("entries", {})

            merged_entries = {}
            all_keys = set(list(old_entries.keys()) + list(new_entries.keys()))
            kept_count = 0
            updated_count = 0

            for key in all_keys:
                old_val = old_entries.get(key, {})
                new_val = new_entries.get(key, {})
                if not isinstance(old_val, dict):
                    old_val = {}
                if not isinstance(new_val, dict):
                    new_val = {}

                if entry_has_data(new_val) and entry_has_data(old_val):
                    merged_entries[key] = merge_entry(old_val, new_val)
                    updated_count += 1
                elif entry_has_data(new_val):
                    merged_entries[key] = new_val
                    updated_count += 1
                elif entry_has_data(old_val):
                    merged_entries[key] = old_val
                    kept_count += 1
                else:
                    merged_entries[key] = new_val if key in new_entries else old_val

            if kept_count > 0:
                print(f"  {m}: {updated_count}件更新, {kept_count}件前回データ維持")
            else:
                print(f"  {m}: OK（{updated_count}件更新）")

            new_data["entries"] = merged_entries
            repo_journal.write_text(
                json.dumps(new_data, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            latest = repo_dir / "journal_latest.json"
            latest.write_text(
                json.dumps(new_data, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception as e:
            print(f"[WARNING] マージ中にエラー: {e}")
    print()

    print("=" * 48)
    print("  CI パイプライン完了（JSON生成済み）")
    print("=" * 48)


if __name__ == "__main__":
    main()
