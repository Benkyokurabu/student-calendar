#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""update_journal.py
授業日誌のExcelからJSONを生成し、GitHubへ反映するプログラム本体です。
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from openpyxl import load_workbook

# 抽出ルール設定
BASE_ROW = {"S": 7, "A": 27, "B": 47, "X": 7}
BLOCK_WIDTH = 10
FIRST_BLOCK_COL = 2

CAMPUS_JP = {"hon": "本校", "minami": "南教室"}
GRADE_JP = {"e4": "小４", "e5": "小５", "e6": "小６", "j1": "中１", "j2": "中２", "j3": "中３"}
SUBJECT_JP = {"eng": "英語", "math": "数学", "jp": "国語", "sci": "理科", "soc": "社会", "arith": "算数"}

def safe_mkdir(p: Path) -> None: p.mkdir(parents=True, exist_ok=True)
def read_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f: return json.load(f)
def atomic_write_text(path: Path, text: str) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(path)
def normalize_str(v: Any) -> str: return str(v).strip() if v is not None else ""

def build_event_key(ev: Dict[str, Any]) -> str:
    t = str(ev.get('time','')).replace('~','～').strip()
    return f"{ev.get('date','')}|{t}|{ev.get('campus','')}|{ev.get('groupKey','')}|{str(ev.get('room',''))}"

def find_workbook(journal_dir: Path, filename: str) -> Optional[Path]:
    # 深い階層のサブフォルダ（南教室 ＞ 中１ など）まで自動で潜って検索します
    if (journal_dir / filename).exists(): return journal_dir / filename
    for candidate in journal_dir.rglob(filename):
        if candidate.is_file(): return candidate
    return None

def load_schedule_files(repo_dir: Path) -> Tuple[Dict[str, List[Dict[str, Any]]], Optional[str]]:
    month_events = {}
    for p in sorted(repo_dir.glob("schedule_????-??.json")):
        m = re.search(r"schedule_(\d{4}-\d{2})\.json$", p.name)
        if m: month_events[m.group(1)] = read_json(p)
    latest_month = None; latest_path = repo_dir / "schedule_latest.json"
    if latest_path.exists():
        le = read_json(latest_path)
        if le:
            ms = sorted({str(ev.get("date", ""))[:7] for ev in le if ev.get("date")})
            if ms: latest_month = ms[-1]; month_events.setdefault(latest_month, le)
    return month_events, latest_month

def generate_journal_for_month(repo_dir, month, events, journal_dir, temp_dir):
    needed = {f"{CAMPUS_JP.get(str(ev.get('campus')), '')}{GRADE_JP.get(str(ev.get('grade')), '')}{'X' if str(ev.get('class'))=='X' else ''}{SUBJECT_JP.get(str(ev.get('subject')), '')}_{str(ev.get('date'))[:4]}.xlsx" for ev in events}
    wb_handles = {}
    for fname in needed:
        p = find_workbook(journal_dir, fname)
        if p:
            copy_p = temp_dir / fname
            shutil.copy2(p, copy_p)
            wb_handles[fname] = load_workbook(copy_p, read_only=True, data_only=True)
    entries = {}
    for ev in events:
        ek = build_event_key(ev); content = ""; page = ""; hw = []; rec = ""
        entries[ek] = {"content": content, "page": page, "homework": hw, "recordingUrl": rec}
    for h in wb_handles.values(): h.close()
    return {"month": month, "generatedAt": datetime.now().isoformat(), "entries": entries}

def git_push(repo_dir):
    # GitHubへの自動送信機能を実行します
    subprocess.run(["git", "-C", str(repo_dir), "add", "."], check=False)
    subprocess.run(["git", "-C", str(repo_dir), "commit", "-m", "Auto Update Journal"], check=False)
    subprocess.run(["git", "-C", str(repo_dir), "push"], check=False)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--journal-dir"); ap.add_argument("--backup-dir"); ap.add_argument("--mode"); ap.add_argument("--do-backup", action="store_true")
    args = ap.parse_args()
    repo_dir = Path(__file__).parent.resolve(); journal_dir = Path(args.journal_dir).resolve()
    m_events, latest_m = load_schedule_files(repo_dir)
    temp_dir = repo_dir / "__tmp__"; safe_mkdir(temp_dir)
    for month, evs in m_events.items():
        jobj = generate_journal_for_month(repo_dir, month, evs, journal_dir, temp_dir)
        atomic_write_text(repo_dir / f"journal_{month}.json", json.dumps(jobj, ensure_ascii=False, indent=2))
    if latest_m: shutil.copy2(repo_dir / f"journal_{latest_m}.json", repo_dir / "journal_latest.json")
    if args.mode == "auto": git_push(repo_dir)
    print("[INFO] done.")

if __name__ == "__main__": main()