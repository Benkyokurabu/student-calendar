#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
プログラム③：報告抽出

役割
- 授業日誌Excelから、Web掲載用の本文だけを抽出して JSON にする
- 読み取る項目
  - D7   授業内容       -> content
  - F8   ページ         -> page
  - E9:E10 宿題        -> homework
  - E11  記録           -> report
  - D17  録画URL        -> recordingUrl
  - B14  担当講師       -> teacher
  - C18  欠席           -> absence
  - I20  カリキュラム進捗 -> curriculumProgress
  - ヘッダー: 第○回 / ○月○週  -> sessionNumber / monthNum / weekNum

出力
- journal_latest.json
- journal_YYYY-MM.json

前提
- カレ ンダー側 repo に schedule_latest.json があること
- 可能なら journal_map_latest.json もあること
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import tempfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

TARGET_FOLDER_NAME = "09　授業日誌"
MONTH_SHEET_RE = re.compile(r"^(\d{4})-(\d{2})(?:\(再(\d+)\))?$")

CAMPUS_JP = {"hon": "本校", "minami": "南教室"}
GRADE_JP = {
    "e4": "小４",
    "e5": "小５",
    "e6": "小６",
    "j1": "中１",
    "j2": "中２",
    "j3": "中３",
}
SUBJECT_JP = {
    "eng": "英語",
    "math": "数学",
    "jp": "国語",
    "sci": "理科",
    "soc": "社会",
    "arith": "算数",
}

FIRST_BLOCK_COL = 2
BLOCK_WIDTH = 10
ROW_STEP = 20
BASE_TOP_MAIN = {"S": 6, "A": 26, "B": 46}
BASE_TOP_X = 6


def get_default_repo_dir() -> Optional[Path]:
    user_home = Path.home()
    candidates = [
        user_home / "OneDrive" / "デスクトップ" / "生徒スケジュール表",
        Path(r"C:\Users\kudok\OneDrive\デスクトップ\生徒スケジュール表"),
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def get_default_journal_dir() -> Optional[Path]:
    user_home = Path.home()
    candidates = [
        user_home / "OneDrive" / "●勉強クラブ共有" / TARGET_FOLDER_NAME,
        Path(r"C:\Users\kudok\OneDrive\●勉強クラブ共有\09　授業日誌"),
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def merged_top_left_cell(ws, row: int, col: int):
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            return ws.cell(mr.min_row, mr.min_col)
    return ws.cell(row, col)


def read_merged_text(ws, row: int, col: int) -> str:
    v = merged_top_left_cell(ws, row, col).value
    return "" if v is None else str(v).strip()


def normalize_time(s: str) -> str:
    return str(s or "").replace("~", "～").strip()


def make_entry_key(ev: dict) -> str:
    return "|".join(
        [
            str(ev.get("date", "")),
            normalize_time(ev.get("time", "")),
            str(ev.get("campus", "")),
            str(ev.get("groupKey", "")),
            str(ev.get("room", "")),
        ]
    )


SKIP_DIRS = {"_backup", "退避", "__pycache__"}


def find_workbook(journal_dir: Path, filename: str) -> Optional[Path]:
    p = journal_dir / filename
    if p.exists():
        return p
    for candidate in journal_dir.rglob(filename):
        if candidate.is_file():
            # _backup / 退避 など除外
            if SKIP_DIRS & {part for part in candidate.relative_to(journal_dir).parts}:
                continue
            return candidate
    return None


def find_month_sheet(wb: openpyxl.Workbook, year: int, month: int):
    target = f"{year:04d}-{month:02d}"
    if target in wb.sheetnames:
        return wb[target]

    cands: List[Tuple[int, str]] = []
    for name in wb.sheetnames:
        m = MONTH_SHEET_RE.match(name)
        if not m:
            continue
        y = int(m.group(1))
        mo = int(m.group(2))
        retry = int(m.group(3) or 0)
        if y == year and mo == month:
            cands.append((retry, name))

    if cands:
        cands.sort()
        return wb[cands[-1][1]]
    return None


def build_group_slot_map(events: List[dict], slots_map: Optional[dict]) -> Dict[str, int]:
    result: Dict[str, int] = {}

    # 1) map がある場合はそれを優先
    if slots_map and isinstance(slots_map, dict) and isinstance(slots_map.get("slots"), dict):
        for ev in events:
            gk = str(ev.get("groupKey", ""))
            key = make_entry_key(ev)
            lst = slots_map.get("slots", {}).get(gk, [])
            try:
                result[key] = lst.index(key) + 1
                continue
            except Exception:
                pass

    # 2) fallback: schedule_latest.json 内の出現順で groupKey ごとに連番
    counters: Dict[str, int] = defaultdict(int)
    seen: Dict[Tuple[str, str], int] = {}
    for ev in events:
        gk = str(ev.get("groupKey", ""))
        key = make_entry_key(ev)
        pair = (gk, key)
        if pair in seen:
            result[key] = seen[pair]
            continue
        counters[gk] += 1
        seen[pair] = counters[gk]
        result[key] = counters[gk]

    return result


def subject_for_filename(grade_key: str, subj_key: str) -> str:
    """小学生(e4,e5,e6)の math は '算数' に変換"""
    if subj_key == "math" and grade_key in ("e4", "e5", "e6"):
        return "算数"
    return SUBJECT_JP.get(subj_key, "")


def workbook_filename_from_event(ev: dict, year: int) -> Optional[str]:
    campus = CAMPUS_JP.get(str(ev.get("campus", "")))
    grade = GRADE_JP.get(str(ev.get("grade", "")))
    grade_key = str(ev.get("grade", ""))
    subj_key = str(ev.get("subject", ""))
    subject = subject_for_filename(grade_key, subj_key)
    klass = str(ev.get("class", "") or "")
    if not campus or not grade or not subject:
        return None

    if klass == "X":
        return f"{campus}{grade}X{subject}_{year}.xlsx"
    return f"{campus}{grade}{subject}_{year}.xlsx"


def read_block(ws, top_row: int, left_col: int) -> dict:
    content = read_merged_text(ws, top_row + 1, left_col + 2)       # D7
    page = read_merged_text(ws, top_row + 2, left_col + 4)          # F8
    hw1 = read_merged_text(ws, top_row + 3, left_col + 3)           # E9
    hw2 = read_merged_text(ws, top_row + 4, left_col + 3)           # E10
    report = read_merged_text(ws, top_row + 5, left_col + 3)        # E11
    recording_url = read_merged_text(ws, top_row + 11, left_col + 2)  # D17
    teacher = read_merged_text(ws, top_row + 8, left_col)            # B14 (担当)
    absence = read_merged_text(ws, top_row + 12, left_col + 1)      # C18 (欠席)
    curriculum_sign = read_merged_text(ws, top_row + 14, left_col + 6)  # H20 (進捗符号: +/-/±)
    curriculum_val = read_merged_text(ws, top_row + 14, left_col + 7)   # I20 (進捗数値)
    curriculum = (curriculum_sign + curriculum_val) if (curriculum_sign or curriculum_val) else ""
    note = read_merged_text(ws, top_row + 15, left_col + 1)          # C21 (備考)

    homework = [x for x in [hw1, hw2] if x]
    return {
        "content": content,
        "page": page,
        "homework": homework,
        "report": report,
        "recordingUrl": recording_url,
        "teacher": teacher,
        "absence": absence,
        "curriculumProgress": curriculum,
        "note": note,
    }


class WorkbookCache:
    """同じExcelを何度も開かないようにキャッシュする"""
    def __init__(self, journal_dir: Path, temp_dir: Path):
        self.journal_dir = journal_dir
        self.temp_dir = temp_dir
        self._wb_cache: Dict[str, Optional[openpyxl.Workbook]] = {}
        self._path_cache: Dict[str, Optional[Path]] = {}

    def get_workbook(self, filename: str) -> Optional[openpyxl.Workbook]:
        if filename in self._wb_cache:
            return self._wb_cache[filename]

        wb_path = find_workbook(self.journal_dir, filename)
        if wb_path is None:
            self._wb_cache[filename] = None
            return None

        try:
            tmp_file = self.temp_dir / filename
            if not tmp_file.exists():
                shutil.copy2(wb_path, tmp_file)
            wb = openpyxl.load_workbook(tmp_file, data_only=True, read_only=False)
            self._wb_cache[filename] = wb
            return wb
        except Exception:
            self._wb_cache[filename] = None
            return None

    def close_all(self):
        for wb in self._wb_cache.values():
            if wb:
                try:
                    wb.close()
                except Exception:
                    pass


def read_slot_header(ws, left_col: int) -> dict:
    """スロットヘッダー情報を読み取る（第○回、○月○週）"""
    session = read_merged_text(ws, 2, left_col + 4)   # F2 相当
    month_n = read_merged_text(ws, 3, left_col + 3)    # E3 相当
    week_n = read_merged_text(ws, 3, left_col + 5)     # G3 相当
    return {
        "sessionNumber": session,
        "monthNum": month_n,
        "weekNum": week_n,
    }


def extract_entry_for_event(ev: dict, slot_index: int, wb_cache: WorkbookCache) -> dict:
    empty = {"content": "", "page": "", "homework": [], "report": "", "recordingUrl": "",
             "teacher": "", "absence": "", "curriculumProgress": "", "note": "",
             "sessionNumber": "", "monthNum": "", "weekNum": ""}
    date = str(ev.get("date", ""))
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", date)
    if not m:
        return empty

    year = int(m.group(1))
    month = int(m.group(2))
    klass = str(ev.get("class", "") or "")

    filename = workbook_filename_from_event(ev, year)
    if not filename:
        return empty

    wb = wb_cache.get_workbook(filename)
    if wb is None:
        return empty

    ws = find_month_sheet(wb, year, month)
    if ws is None:
        return empty

    left_col = FIRST_BLOCK_COL + (max(1, slot_index) - 1) * BLOCK_WIDTH
    if klass == "X":
        top_row = BASE_TOP_X
    else:
        top_row = BASE_TOP_MAIN.get(klass)
        if top_row is None:
            return empty

    block = read_block(ws, top_row, left_col)
    block.update(read_slot_header(ws, left_col))
    return block


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def determine_month(events: List[dict], month_arg: Optional[str]) -> str:
    if month_arg:
        if not re.match(r"^\d{4}-\d{2}$", month_arg):
            raise ValueError("--month は YYYY-MM 形式で指定してください。")
        return month_arg

    months = sorted({str(ev.get("date", ""))[:7] for ev in events if str(ev.get("date", ""))[:7]})
    if not months:
        raise ValueError("schedule_latest.json から年月を判定できませんでした。")
    return months[-1]


def find_schedule_json(repo_dir: Path, month_arg: Optional[str]) -> Tuple[Path, str]:
    """スケジュールJSONを探す。--month 指定時は月別ファイルを優先。"""
    if month_arg:
        if not re.match(r"^\d{4}-\d{2}$", month_arg):
            raise ValueError("--month は YYYY-MM 形式で指定してください。")
        # 月別ファイルを優先
        month_path = repo_dir / f"schedule_{month_arg}.json"
        if month_path.exists():
            events = load_json(month_path)
            print(f"[INFO] スケジュール: {month_path.name}")
            return month_path, month_arg
        # なければ latest から該当月をフィルタ
        latest = repo_dir / "schedule_latest.json"
        if latest.exists():
            events = load_json(latest)
            has_month = any(str(ev.get("date", ""))[:7] == month_arg for ev in events)
            if has_month:
                print(f"[INFO] スケジュール: {latest.name} (月フィルタ: {month_arg})")
                return latest, month_arg
        raise FileNotFoundError(
            f"schedule_{month_arg}.json も schedule_latest.json({month_arg}分) も見つかりません: {repo_dir}"
        )

    # --month 未指定 → latest から自動判定
    latest = repo_dir / "schedule_latest.json"
    if not latest.exists():
        raise FileNotFoundError(f"schedule_latest.json が見つかりません: {repo_dir}")
    events = load_json(latest)
    target_month = determine_month(events, None)
    print(f"[INFO] スケジュール: {latest.name} (自動判定: {target_month})")
    return latest, target_month


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo-dir", help="student-calendar リポジトリの場所。未指定時は自動探索")
    ap.add_argument("--journal-dir", help="授業日誌フォルダ。未指定時は OneDrive 共有フォルダを自動探索")
    ap.add_argument("--month", help="出力対象月（YYYY-MM）。未指定時は schedule_latest.json から自動判定")
    args = ap.parse_args()

    repo_dir = Path(args.repo_dir).resolve() if args.repo_dir else get_default_repo_dir()
    if repo_dir is None or not repo_dir.exists():
        raise FileNotFoundError("student-calendar リポジトリが見つかりません。--repo-dir で指定してください。")
    journal_dir = Path(args.journal_dir) if args.journal_dir else get_default_journal_dir()
    if journal_dir is None or not journal_dir.exists():
        raise FileNotFoundError("授業日誌フォルダが見つかりません。--journal-dir で指定してください。")

    schedule_path, target_month = find_schedule_json(repo_dir, args.month)
    events = load_json(schedule_path)

    map_latest = repo_dir / "journal_map_latest.json"
    map_month = repo_dir / f"journal_map_{target_month}.json"
    slots_map = None
    if map_latest.exists():
        slots_map = load_json(map_latest)
    elif map_month.exists():
        slots_map = load_json(map_month)

    slot_map = build_group_slot_map(events, slots_map)

    entries = {}
    with tempfile.TemporaryDirectory() as tmp:
        wb_cache = WorkbookCache(journal_dir, Path(tmp))
        try:
            for ev in events:
                if str(ev.get("date", ""))[:7] != target_month:
                    continue
                key = make_entry_key(ev)
                slot_index = slot_map.get(key, 1)
                entries[key] = extract_entry_for_event(ev, slot_index, wb_cache)
        finally:
            wb_cache.close_all()

    output = {
        "month": target_month,
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "entries": entries,
    }

    latest_path = repo_dir / "journal_latest.json"
    month_path = repo_dir / f"journal_{target_month}.json"
    latest_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    month_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[OK] 出力: {latest_path}")
    print(f"[OK] 出力: {month_path}")
    print("[INFO] 抽出対象: content / page / homework / report / recordingUrl / teacher / absence / curriculumProgress / sessionNumber / monthNum / weekNum")


if __name__ == "__main__":
    main()
