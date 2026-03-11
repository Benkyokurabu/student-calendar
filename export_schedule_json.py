# -*- coding: utf-8 -*-
"""
スケジュール表（xlsm）から「月次JSON（schedule_YYYY-MM.json）」を生成するスクリプト

設計方針（デグレ防止）:
- 既存の export_by_grade_subject.py（授業日誌側）は一切変更しない
- 読み取りロジック（collect_events / choose_target_sheets 等）は既存を流用
- 本スクリプトは JSON出力だけを担当する（別ファイル運用）

仕様（確定事項）:
- campus: hon / minami
- grade: e4/e5/e6, j1/j2/j3
- subject: arith/math/eng/jp/sci/soc（特は区別しない）
- room: "1".."5"（①〜⑤、または 1/２/2/第２教室 なども数字文字列に正規化）
- groupKey: {campus}_{grade}_{class}_{subject}（roomは含めない）
- label: 例「中3S 理科」「小5A 算数」

追加仕様（対面判定ルール変更）:
- 授業セル内の2行目が「対面」の授業は、
  JSONに faceToFace=true を付与し、displayTitle を「(1行目)+対面」で出力する。
  例: "３Ｓ国\\n対面" → displayTitle="3S国対面"
  ※既存UI互換のため、label/groupKey など既存キーは変更しない。

使い方:
1) 明示指定（推奨）
   python export_schedule_json.py --schedule "2026年3月スケジュール.xlsm"

2) フォルダから自動検出（xlsmが1つなら案内）
   python export_schedule_json.py
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl

# 同じフォルダに export_by_grade_subject.py がある前提
import export_by_grade_subject as legacy


# ===== 正規化辞書 =====

CAMPUS_CODE = {
    "本校": "hon",
    "南教室": "minami",
}

GRADE_CODE = {
    "1": "j1",  # 中1
    "2": "j2",  # 中2
    "3": "j3",  # 中3
    "4": "e4",  # 小4
    "5": "e5",  # 小5
    "6": "e6",  # 小6
}

# ①②③… → "1".."9"
ROOM_MAP = {
    "①": "1",
    "②": "2",
    "③": "3",
    "④": "4",
    "⑤": "5",
    "⑥": "6",
    "⑦": "7",
    "⑧": "8",
    "⑨": "9",
}

# 全角数字→半角数字
ZEN2HAN_TRANS = str.maketrans("０１２３４５６７８９", "0123456789")


def normalize_room(room_raw: str) -> str:
    """
    教室表記を数字文字列に正規化する
    対応例：
      ① → "1"
      2 → "2"
      ２ → "2"
      ③教室 → "3"
      第４教室 → "4"
    """
    if not room_raw:
        return ""

    s = str(room_raw).strip()

    # ①②③… を最優先
    for k, v in ROOM_MAP.items():
        if k in s:
            return v

    # 全角数字 → 半角数字
    s2 = s.translate(ZEN2HAN_TRANS)

    # 数字を抽出
    m = re.search(r"(\d+)", s2)
    if m:
        return m.group(1)

    return ""


def normalize_subject(grade_digit: str, subj_legacy: str) -> str:
    """
    legacyのsubjは主に「英/数/国/理/社」。
    算数はlegacy側で「数」扱いになることがあるため、学年で arith/math を分ける。
    特は区別しない（specialは見ない）。
    """
    s = (subj_legacy or "").strip()
    if s == "英":
        return "eng"
    if s == "国":
        return "jp"
    if s == "理":
        return "sci"
    if s == "社":
        return "soc"
    if s == "数":
        # 小学生は算数、中学生は数学
        if grade_digit in ("4", "5", "6"):
            return "arith"
        return "math"
    return ""


def grade_label_jp(grade_digit: str) -> str:
    return legacy.GRADE_LABEL.get(str(grade_digit), "")


def subject_label_jp(subject_code: str) -> str:
    return {
        "arith": "算数",
        "math": "数学",
        "eng": "英語",
        "jp": "国語",
        "sci": "理科",
        "soc": "社会",
    }.get(subject_code, "")


def build_group_key(campus_code: str, grade_code: str, klass: str, subject_code: str) -> str:
    k = (klass or "").strip()
    return f"{campus_code}_{grade_code}_{k}_{subject_code}"


def build_label(grade_digit: str, klass: str, subject_code: str) -> str:
    g = grade_label_jp(grade_digit)  # "中３" 等
    s = subject_label_jp(subject_code)
    k = (klass or "").strip()
    return f"{g}{k} {s}".strip()


def ymd(year: int, month: int, day: str) -> str:
    d = str(day).strip()
    d = d.translate(ZEN2HAN_TRANS)
    if not d.isdigit():
        return ""
    return f"{year}-{month:02d}-{int(d):02d}"


# ===== 「対面」検出（セル2行目） =====

def _split_lines(cell_text: str) -> List[str]:
    if not cell_text:
        return []
    s = str(cell_text).replace("\r\n", "\n").replace("\r", "\n")
    return [ln.strip() for ln in s.split("\n") if ln.strip() != ""]


def is_face_to_face_cell(text: str) -> bool:
    """
    仕様：セル内の「2行目」に「対面」と書かれている授業を対面扱いにする。
    例: "３Ｓ国\\n対面" -> True
    """
    lines = _split_lines(text)
    if len(lines) < 2:
        return False
    return "対面" in lines[1]


def normalize_first_line_for_display(text: str) -> str:
    """
    1行目を表示用に整形（全角数字→半角、全角S/A/B/X→半角大文字、空白除去）
    例: "３Ｓ理" -> "3S理"
    """
    lines = _split_lines(text)
    if not lines:
        return ""
    s = lines[0]
    s = s.translate(ZEN2HAN_TRANS)
    s = s.replace(" ", "").replace("\u3000", "")
    s = s.replace("Ｓ", "S").replace("Ａ", "A").replace("Ｂ", "B").replace("Ｘ", "X")
    return s


# ===== スケジュールファイル検出 =====

def parse_year_month_from_filename(path: Path) -> Optional[Tuple[int, int]]:
    name = path.name
    name = name.translate(ZEN2HAN_TRANS)
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", name)
    if not m:
        return None
    year = int(m.group(1))
    month = int(m.group(2))
    if not (1 <= month <= 12):
        return None
    return year, month


def find_schedule_file(schedule_arg: Optional[str] = None) -> Tuple[int, int, Path]:
    here = Path(__file__).resolve().parent

    if schedule_arg:
        p = Path(schedule_arg)
        if not p.is_absolute():
            p = (here / p).resolve()
        if not p.exists():
            raise FileNotFoundError(f"--schedule で指定されたファイルが見つかりません: {p}")
        if p.suffix.lower() != ".xlsm":
            raise ValueError(f".xlsm を指定してください: {p}")
        ym = parse_year_month_from_filename(p)
        if not ym:
            raise ValueError(
                "ファイル名から年月を判定できません。\n"
                "ファイル名を『YYYY年M月...xlsm』（全角数字でも可）にするか、年/月を含めてください。\n"
                f"指定ファイル: {p.name}"
            )
        return ym[0], ym[1], p

    xlsm_files = sorted(here.glob("*.xlsm"))
    if not xlsm_files:
        raise FileNotFoundError("同フォルダに .xlsm が見つかりません。--schedule で指定してください。")

    ym_files: List[Tuple[int, int, Path]] = []
    unknown_files: List[Path] = []
    for f in xlsm_files:
        ym = parse_year_month_from_filename(f)
        if ym:
            ym_files.append((ym[0], ym[1], f))
        else:
            unknown_files.append(f)

    if len(ym_files) == 1:
        return ym_files[0]

    if len(ym_files) >= 2:
        msg = "年月が判定できる .xlsm が複数あります。--schedule で1つ指定してください。\n候補:\n"
        msg += "\n".join([f"- {p.name}" for (_, _, p) in ym_files])
        raise RuntimeError(msg)

    if len(unknown_files) == 1:
        raise ValueError(
            "xlsm は1つ見つかりましたが、ファイル名から年月を判定できません。\n"
            "ファイル名を『YYYY年M月...xlsm』（全角数字でも可）に変更するか、--schedule で指定してください。\n"
            f"見つかったファイル: {unknown_files[0].name}"
        )

    msg = "年月が判定できない .xlsm が複数あります。--schedule で指定してください。\n候補:\n"
    msg += "\n".join([f"- {p.name}" for p in unknown_files])
    raise RuntimeError(msg)


# ===== JSON生成 =====

def convert_events_to_schedule_json(
    events: List[legacy.Event],
    campus_label: str,
    year: int,
    month: int,
) -> List[Dict[str, Any]]:
    campus_code = CAMPUS_CODE.get(campus_label, "")
    if not campus_code:
        raise ValueError(f"campus_label が想定外です: {campus_label}")

    out: List[Dict[str, Any]] = []

    for e in events:
        grade_digit = str(e.grade).strip()
        grade_code = GRADE_CODE.get(grade_digit, "")
        if not grade_code:
            continue

        subject_code = normalize_subject(grade_digit, e.subj)
        if not subject_code:
            continue

        klass = (e.klass or "").strip()
        date_str = ymd(year, month, e.day)
        if not date_str:
            continue

        room = normalize_room(e.classroom)
        label = build_label(grade_digit, klass, subject_code)
        gk = build_group_key(campus_code, grade_code, klass, subject_code)

        raw_text = (e.text or "").strip()
        face = is_face_to_face_cell(raw_text)

        # 対面のときだけ、表示名を別出し（既存UIの表示崩れ回避）
        display_title = ""
        if face:
            base = normalize_first_line_for_display(raw_text)
            display_title = f"{base}対面" if base else "対面"

        out.append(
            {
                "date": date_str,
                "time": (e.time or "").strip(),
                "grade": grade_code,
                "class": klass,
                "subject": subject_code,
                "campus": campus_code,
                "room": room,
                "groupKey": gk,
                "label": label,
                # --- 追加（後方互換） ---
                "faceToFace": bool(face),
                "displayTitle": display_title,
            }
        )

    out.sort(key=lambda x: (x["date"], x.get("time", ""), x.get("label", "")))
    return out


def export_month_schedule_json(schedule_arg: Optional[str] = None) -> Path:
    year, month, sch_path = find_schedule_file(schedule_arg)

    wb = openpyxl.load_workbook(sch_path, data_only=True, keep_vba=True)
    targets = legacy.choose_target_sheets(wb)
    if not targets:
        raise RuntimeError("対象シート（本校/南教室の教務部用）が見つかりません。")

    all_rows: List[Dict[str, Any]] = []
    for campus_label, _sheetname, sh in targets:
        events = legacy.collect_events(sh, month)
        rows = convert_events_to_schedule_json(events, campus_label, year, month)
        all_rows.extend(rows)

    json_text = json.dumps(all_rows, ensure_ascii=False, indent=2)

    out_name = f"schedule_{year}-{month:02d}.json"
    out_path = Path(sch_path).parent / out_name
    out_path.write_text(json_text, encoding="utf-8")

    latest_path = Path(sch_path).parent / "schedule_latest.json"
    latest_path.write_text(json_text, encoding="utf-8")

    return out_path


def main():
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--schedule", type=str, default=None, help="対象のスケジュールxlsmを指定（推奨）")
    args = ap.parse_args()

    out_path = export_month_schedule_json(args.schedule)
    latest_path = out_path.parent / "schedule_latest.json"
    print(f"[OK] JSONを書き出しました: {out_path}")
    print(f"[OK] JSONを書き出しました: {latest_path}")


if __name__ == "__main__":
    main()