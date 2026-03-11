#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse
import json
import os
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

# --- 設定（既存の動作を維持） ---
BASE_ROW = {"S": 7, "A": 27, "B": 47, "X": 7}
BLOCK_WIDTH = 10
FIRST_BLOCK_COL = 2

CAMPUS_JP = {"hon": "本校", "minami": "南教室"}
GRADE_JP = {"e4": "小４", "e5": "小５", "e6": "小６", "j1": "中１", "j2": "中２", "j3": "中３"}
SUBJECT_JP = {"eng": "英語", "math": "数学", "jp": "国語", "sci": "理科", "soc": "社会", "arith": "算数"}

TARGET_FOLDER_NAME = "09　授業日誌"

def get_journal_dir():
    user_home = Path.home()
    candidates = [
        user_home / "OneDrive" / "●勉強クラブ共有" / TARGET_FOLDER_NAME,
        Path(r"C:\Users\kudok\OneDrive\●勉強クラブ共有\09　授業日誌")
    ]
    for c in candidates:
        if c.exists(): return c
    return None

def find_workbook(journal_dir, filename):
    # まず指定の名前で探す
    p = journal_dir / filename
    if p.exists(): return p
    # 見つからない場合はサブフォルダ内を全検索
    for candidate in journal_dir.rglob(filename):
        if candidate.is_file(): return candidate
    return None

def generate(repo_dir, journal_dir):
    repo_dir = Path(repo_dir)
    sched_path = repo_dir / "schedule_latest.json"
    if not sched_path.exists():
        print(f"[エラー] {sched_path.name} が見つかりません。")
        return

    events = json.loads(sched_path.read_text(encoding="utf-8"))
    map_path = repo_dir / "journal_map_latest.json"
    if not map_path.exists(): map_path = repo_dir / "journal_map_2026-03.json"
    slots_map = json.loads(map_path.read_text(encoding="utf-8")) if map_path.exists() else {"slots": {}}
    
    entries = {}
    temp_dir = repo_dir / "__tmp__"
    if temp_dir.exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
    temp_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n--- データ読み取り詳細レポート ---")

    for ev in events:
        key = f"{ev.get('date','')}|{str(ev.get('time','')).replace('~','～').strip()}|{ev.get('campus','')}|{ev.get('groupKey','')}|{str(ev.get('room',''))}"
        gk = ev.get("groupKey", "")
        try:
            l_idx = slots_map.get("slots", {}).get(gk, []).index(key) + 1
        except: l_idx = 1
        
        # クラス名(S, A, B)をファイル名に含めるかどうかの柔軟な判定
        cls_code = ev.get('class', '')
        possible_filenames = [
            f"{CAMPUS_JP.get(ev['campus'])}{GRADE_JP.get(ev['grade'])}{cls_code}{SUBJECT_JP.get(ev['subject'])}_2026.xlsx",
            f"{CAMPUS_JP.get(ev['campus'])}{GRADE_JP.get(ev['grade'])}{SUBJECT_JP.get(ev['subject'])}_2026.xlsx"
        ]
        
        wb_path = None
        for fname in possible_filenames:
            wb_path = find_workbook(journal_dir, fname)
            if wb_path: break
        
        content, page, hw, rec = "", "", [], ""
        
        if wb_path:
            print(f"【発見】 {wb_path.name}")
            try:
                tmp_path = temp_dir / f"{int(time.time()*1000)}_{wb_path.name}"
                shutil.copy2(wb_path, tmp_path)
                wb = load_workbook(tmp_path, data_only=True, read_only=True)
                m_idx = int(ev['date'].split("-")[1])
                sheet_name = f"{m_idx}月"
                
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    br = BASE_ROW.get(cls_code, 7)
                    sc = FIRST_BLOCK_COL + (l_idx - 1) * BLOCK_WIDTH
                    
                    content = str(ws.cell(br, sc+2).value or "").strip()
                    page = str(ws.cell(br+1, sc+4).value or "").strip()
                    print(f"  -> {ev['date']} {cls_code}クラス 読込成功: 「{content[:10]}...」")
                wb.close()
            except Exception as e:
                print(f"  -> [失敗] 読み取りエラー: {e}")
        else:
            # 2S数学などの場合に、どの名前で探して見つからなかったかを表示
            print(f"【未発見】 探した名前: {possible_filenames[0]}")

        entries[key] = {"content": content, "page": page, "homework": hw, "recordingUrl": rec}

    (repo_dir / "journal_latest.json").write_text(json.dumps({"month": "latest", "generatedAt": datetime.now().isoformat(), "entries": entries}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"--- レポート終了 ---\n")

def git_push(repo_dir):
    print("[INFO] GitHubへ送信中...")
    subprocess.run(["git", "-C", str(repo_dir), "add", "."], check=False)
    subprocess.run(["git", "-C", str(repo_dir), "commit", "-m", "Debug Update"], check=False)
    subprocess.run(["git", "-C", str(repo_dir), "push"], check=False)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode")
    parser.add_argument("--do-backup", action="store_true")
    args = parser.parse_args()
    current_repo = Path(__file__).parent.resolve()
    j_dir = get_journal_dir()
    if j_dir:
        generate(current_repo, j_dir)
        if args.mode == "auto": git_push(current_repo)
        print("[INFO] done.")