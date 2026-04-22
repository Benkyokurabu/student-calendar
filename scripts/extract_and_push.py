#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
スケジュールJSON再生成 + 日誌キャンパス間コピー + 授業日誌JSON抽出 → GitHub push を一括実行するスクリプト

使い方:
    python extract_and_push.py              （最新月を自動判定）
    python extract_and_push.py 2026-04      （月を指定）
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
    """処理対象の月リストを返す。
    - month_arg 指定時はその月のみ
    - 未指定時: 今月を基準に、月末5日間なら前月も含める
    """
    if month_arg:
        return [month_arg]

    today = date.today()
    last_day = calendar.monthrange(today.year, today.month)[1]
    current_month = f"{today.year}-{today.month:02d}"

    if today.day > last_day - 5:
        # 月末5日間: 前月も処理
        if today.month == 1:
            prev_month = f"{today.year - 1}-12"
        else:
            prev_month = f"{today.year}-{today.month - 1:02d}"
        return [prev_month, current_month]

    return [current_month]


def get_repo_dir() -> Path:
    candidates = [
        Path.home() / "OneDrive" / "デスクトップ" / "生徒スケジュール表",
    ]
    for c in candidates:
        if c.exists():
            return c
    raise FileNotFoundError("生徒スケジュール表フォルダが見つかりません。")


def get_journal_dir() -> Path:
    candidates = [
        Path.home() / "OneDrive" / "●勉強クラブ共有" / "09　授業日誌",
        Path(r"C:\Users\kudok\OneDrive\●勉強クラブ共有\09　授業日誌"),
    ]
    for c in candidates:
        if c.exists():
            return c
    raise FileNotFoundError("09　授業日誌フォルダが見つかりません。")


SKIP_DIRS = {"_backup", "退避", "__pycache__"}


def find_workbook_in_journal(journal_dir: Path, filename: str):
    p = journal_dir / filename
    if p.exists():
        return p
    for candidate in journal_dir.rglob(filename):
        if candidate.is_file():
            if SKIP_DIRS & {part for part in candidate.relative_to(journal_dir).parts}:
                continue
            return candidate
    return None


def create_month_sheets(script_dir: Path, months: list, journal_dir_override: Path = None):
    """スケジュールExcelを読み、OneDriveの日誌ファイルに新しい月シートを追加する"""
    old_cwd = os.getcwd()
    try:
        # export_by_grade_subject.py のあるディレクトリを一時的にパスに追加
        sys.path.insert(0, str(script_dir))
        os.chdir(str(script_dir))

        from export_by_grade_subject import (
            GRADES, SUBJS, GRADE_LABEL, _display_subj_name,
            pick_schedule_in_same_folder, choose_target_sheets, collect_events,
            create_month_sheet, set_header_cells,
            fill_sheet_main, fill_sheet_x, save_year_workbook,
            open_or_create_year_workbook, ensure_hidden_template_sheet,
            TEMPLATE_MAIN, TEMPLATE_X,
        )

        journal_dir = journal_dir_override if journal_dir_override else get_journal_dir()
        created_count = 0

        for month_str in months:
            year, month = int(month_str.split("-")[0]), int(month_str.split("-")[1])

            try:
                y, m, sch = pick_schedule_in_same_folder([month])
            except FileNotFoundError:
                print(f"  [SKIP] {month}月のスケジュールが見つかりません")
                continue

            if m != month:
                print(f"  [SKIP] {month}月のスケジュールが見つかりません（{m}月のみ）")
                continue

            import openpyxl
            wb_s = openpyxl.load_workbook(sch, data_only=True, keep_vba=True)
            targets = choose_target_sheets(wb_s)
            if not targets:
                print(f"  [SKIP] 教務部用シートが見つかりません")
                continue

            from collections import defaultdict

            for campus, sname, sh in targets:
                all_events = collect_events(sh, month, campus, year, month)
                buckets = defaultdict(list)
                for e in all_events:
                    if ("英語補講" in e.text) or ("数学補講" in e.text):
                        continue
                    buckets[(e.grade, e.subj, e.klass)].append(e)

                for grade in GRADES:
                    for subj in SUBJS:
                        s_list = buckets.get((grade, subj, "S"), [])
                        a_list = buckets.get((grade, subj, "A"), [])
                        b_list = buckets.get((grade, subj, "B"), [])
                        x_list = buckets.get((grade, subj, "X"), [])

                        subj_j = _display_subj_name(grade, subj)
                        grade_j = GRADE_LABEL[grade]

                        # S/A/B メインファイル
                        if any([s_list, a_list, b_list]):
                            fname = f"{campus}{grade_j}{subj_j}_{year}.xlsx"
                            wb_path = find_workbook_in_journal(journal_dir, fname)
                            if wb_path is None:
                                continue

                            wb = open_or_create_year_workbook(wb_path, TEMPLATE_MAIN, "__TEMPLATE_MAIN__")
                            tmpl = ensure_hidden_template_sheet(wb, TEMPLATE_MAIN, "__TEMPLATE_MAIN__")
                            ws_out = create_month_sheet(wb, tmpl, year, month)
                            if ws_out is None:
                                continue  # 既にある
                            set_header_cells(ws_out, campus, grade_j, subj_j, month)
                            fill_sheet_main(ws_out, month, {"S": s_list, "A": a_list, "B": b_list}, teacher_blank=False)
                            save_year_workbook(wb, wb_path)
                            print(f"  [NEW] {fname} + {ws_out.title}")
                            created_count += 1

                        # X専用ファイル
                        if x_list:
                            fname = f"{campus}{grade_j}X{subj_j}_{year}.xlsx"
                            wb_path = find_workbook_in_journal(journal_dir, fname)
                            if wb_path is None:
                                continue

                            wb = open_or_create_year_workbook(wb_path, TEMPLATE_X, "__TEMPLATE_X__")
                            tmpl = ensure_hidden_template_sheet(wb, TEMPLATE_X, "__TEMPLATE_X__")
                            ws_out = create_month_sheet(wb, tmpl, year, month)
                            if ws_out is None:
                                continue
                            set_header_cells(ws_out, campus, grade_j, subj_j, month)
                            fill_sheet_x(ws_out, month, x_list, teacher_blank=False)
                            save_year_workbook(wb, wb_path)
                            print(f"  [NEW] {fname} + {ws_out.title}")
                            created_count += 1

                # 補講ファイル
                for keyword, title in [("英語補講", "英語補講"), ("数学補講", "数学補講")]:
                    hits = [e for e in all_events if keyword in e.text]
                    if not hits:
                        continue
                    fname = f"{campus}{title}_{year}.xlsx"
                    wb_path = find_workbook_in_journal(journal_dir, fname)
                    if wb_path is None:
                        continue

                    wb = open_or_create_year_workbook(wb_path, TEMPLATE_MAIN, "__TEMPLATE_MAIN__")
                    tmpl = ensure_hidden_template_sheet(wb, TEMPLATE_MAIN, "__TEMPLATE_MAIN__")
                    ws_out = create_month_sheet(wb, tmpl, year, month)
                    if ws_out is None:
                        continue
                    set_header_cells(ws_out, campus, "", title, month)
                    fill_sheet_main(ws_out, month, {"S": hits, "A": [], "B": []}, teacher_blank=True)
                    save_year_workbook(wb, wb_path)
                    print(f"  [NEW] {fname} + {ws_out.title}")
                    created_count += 1

            wb_s.close()

        os.chdir(old_cwd)
        return created_count

    except Exception as e:
        os.chdir(old_cwd)
        print(f"  [ERROR] 月シート作成中にエラー: {e}")
        import traceback
        traceback.print_exc()
        return 0


def run(args, **kwargs):
    print(f"  > {' '.join(str(a) for a in args)}")
    result = subprocess.run(args, **kwargs)
    if result.returncode != 0:
        print(f"[ERROR] コマンド失敗 (exit {result.returncode})")
        sys.exit(1)
    return result


def main():
    script_dir = Path(__file__).parent.resolve()
    month_arg = sys.argv[1] if len(sys.argv) > 1 else None
    repo_dir = get_repo_dir()
    months = get_months_to_process(month_arg)

    print("=" * 48)
    print("  スケジュール＋授業日誌 → GitHub push")
    print(f"  対象月: {', '.join(months)}")
    print("=" * 48)
    print()

    # --- 0. rclone でクラウドから日誌Excelをダウンロード ---
    cloud_journal_dir = None
    try:
        from download_journal_from_cloud import download_journal, upload_journal
        print("[0/5] OneDrive クラウドから授業日誌をダウンロード中...")
        cloud_journal_dir = download_journal()
        print()
    except FileNotFoundError as e:
        print(f"[0/5] rclone 未設定のためローカル同期フォルダを使用: {e}")
        print()
    except Exception as e:
        print(f"[0/5] クラウドダウンロード失敗。ローカル同期フォルダを使用: {e}")
        print()

    # --- 1. スケジュールJSON再生成 ---
    print("[1/5] スケジュールExcelから JSON を生成中...")
    export_script = script_dir / "export_schedule_json.py"
    if export_script.exists():
        run([sys.executable, str(export_script)], cwd=str(script_dir))
        # 生成されたJSONをリポにコピー
        for pat in ["schedule_*.json"]:
            for f in script_dir.glob(pat):
                dst = repo_dir / f.name
                shutil.copy2(f, dst)
                print(f"  → {dst.name}")
    else:
        print("[SKIP] export_schedule_json.py が見つかりません。スキップします。")
    print()

    # --- 1.5. 月シート作成（必要な場合） ---
    print("[1.5/5] 授業日誌Excelに新しい月シートを追加中...")
    created = create_month_sheets(script_dir, months, journal_dir_override=cloud_journal_dir)
    if created > 0:
        print(f"  → {created} 個のシートを作成しました")
    else:
        print("  → 新しいシートはありません（既に作成済み）")
    print()

    # --- 1.6. 既存シートの日付・講師を最新スケジュールに同期 ---
    print("[1.6/5] 既存シートの日付・講師をスケジュールに同期中...")
    import re as _re
    _RE_SCH = _re.compile(r"^(\d{4})年0?(\d{1,2})月スケジュール\.xlsm$")
    update_journal_py = repo_dir / "update_journal.py"
    if update_journal_py.exists():
        for m in months:
            month_num = int(m.split("-")[1])
            # script_dir からスケジュールファイルを探す
            sch_path = None
            for p in script_dir.iterdir():
                mm = _RE_SCH.match(p.name)
                if mm and int(mm.group(2)) == month_num:
                    sch_path = p
                    break
            if sch_path is None:
                print(f"  [SKIP] {month_num}月のスケジュールが見つかりません")
                continue
            print(f"  → {m} ({sch_path.name})")
            result = subprocess.run(
                [sys.executable, str(update_journal_py),
                 "--schedule", str(sch_path.resolve())],
                cwd=str(repo_dir),
                capture_output=True, text=True, encoding="utf-8", errors="replace"
            )
            if result.stdout:
                for line in result.stdout.strip().split("\n"):
                    print(f"    {line}")
            if result.returncode != 0:
                print(f"  [WARN] update_journal 失敗 (exit {result.returncode})")
                if result.stderr:
                    print(f"    {result.stderr.strip()}")
    else:
        print("  [SKIP] update_journal.py が見つかりません")
    print()

    # --- 日誌JSON退避（マージ用） ---
    old_jsons = {}
    for m in months:
        repo_journal = repo_dir / f"journal_{m}.json"
        if repo_journal.exists():
            old_jsons[m] = json.loads(repo_journal.read_text(encoding="utf-8"))

    for m in months:
        print(f"--- {m} の処理 ---")

        # --- 2. 日誌キャンパス間コピー ---
        print(f"[2/5] オンラインペア授業の日誌をキャンパス間でコピー中... ({m})")
        sync_script = script_dir / "sync_journal_across_campus.py"
        if sync_script.exists():
            sync_cmd = [sys.executable, str(sync_script), "--month", m]
            if cloud_journal_dir:
                sync_cmd += ["--journal-dir", str(cloud_journal_dir)]
            run(sync_cmd)
        else:
            print("[SKIP] sync_journal_across_campus.py が見つかりません。スキップします。")
        print()

        # --- 3. 授業日誌JSON抽出 ---
        print(f"[3/5] OneDrive の授業日誌Excelから JSON を抽出中... ({m})")
        extract_cmd = [sys.executable, str(script_dir / "extract_journal_to_json.py"), "--month", m]
        if cloud_journal_dir:
            extract_cmd += ["--journal-dir", str(cloud_journal_dir)]
        run(extract_cmd)
        print()

    # --- 3.2. 変更されたExcelをクラウドにアップロード ---
    if cloud_journal_dir:
        print("[3.2/5] 変更されたExcelをクラウドにアップロード中...")
        try:
            upload_journal(cloud_journal_dir)
        except Exception as e:
            print(f"  [WARN] アップロード失敗: {e}")
        print()

    # --- 3.5. 日誌データ マージ ---
    FIELDS = ["content", "page", "report", "recordingUrl",
              "teacher", "absence", "curriculumProgress", "note",
              "sessionNumber", "monthNum", "weekNum"]

    def entry_has_data(entry):
        if not isinstance(entry, dict):
            return False
        return any(entry.get(f) for f in FIELDS) or bool(entry.get("homework"))

    def merge_entry(old_entry, new_entry):
        """フィールド単位でマージ: 新しい方に値があれば採用、なければ前回を維持"""
        merged = {}
        for f in FIELDS:
            new_val = new_entry.get(f, "")
            old_val = old_entry.get(f, "")
            merged[f] = new_val if new_val else old_val
        # homework は中身があるほうを採用
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
                    # 両方にデータあり → フィールド単位でマージ
                    merged_entries[key] = merge_entry(old_val, new_val)
                    updated_count += 1
                elif entry_has_data(new_val):
                    # 新しいデータだけ → 採用
                    merged_entries[key] = new_val
                    updated_count += 1
                elif entry_has_data(old_val):
                    # 前回だけデータあり → 前回を維持
                    merged_entries[key] = old_val
                    kept_count += 1
                else:
                    # 両方空
                    merged_entries[key] = new_val if key in new_entries else old_val

            if kept_count > 0:
                print(f"  {m}: {updated_count}件更新, {kept_count}件前回データ維持（同期遅延を保護）")
            else:
                print(f"  {m}: OK（{updated_count}件更新）")

            # マージ結果を書き戻し
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

    # --- 4. git add & commit ---
    print(f"[4/5] git commit... ({repo_dir})")

    files_to_add = ["schedule_latest.json", "journal_latest.json"]
    files_to_add += [f.name for f in repo_dir.glob("schedule_2*-*.json")]
    files_to_add += [f.name for f in repo_dir.glob("journal_2*-*.json")]

    run(["git", "add"] + files_to_add, cwd=repo_dir)

    # 差分チェック
    result = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        cwd=repo_dir
    )
    if result.returncode == 0:
        print("[INFO] 変更なし — commit/push をスキップします。")
        print()
        print("完了（変更なし）")
        return

    run(["git", "commit", "-m", "Update schedule & journal JSON"], cwd=repo_dir)
    print()

    # --- 5. git pull --rebase & push ---
    print("[5/5] GitHub に push 中...")
    # リモートに先行コミットがある場合に備えて rebase
    pull_result = subprocess.run(
        ["git", "pull", "--rebase", "origin", "main"],
        cwd=repo_dir,
        capture_output=True, text=True
    )
    if pull_result.returncode != 0:
        print(f"[WARN] git pull --rebase 失敗（続行します）: {pull_result.stderr.strip()}")

    push_result = subprocess.run(
        ["git", "push"], cwd=repo_dir,
        capture_output=True, text=True
    )
    if push_result.returncode != 0:
        print(f"[ERROR] git push 失敗: {push_result.stderr.strip()}")
        print("[ERROR] データは commit 済みです。次回の実行時に再push されます。")
        sys.exit(1)
    print()

    print("=" * 48)
    print("  完了！ GitHub Pages に反映されます。")
    print("=" * 48)


if __name__ == "__main__":
    main()
