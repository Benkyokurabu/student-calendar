#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rclone を使って OneDrive クラウドから授業日誌Excelを直接ダウンロードする。
OneDrive デスクトップ同期に依存しないため、同期が壊れていても動作する。

使い方:
    from download_journal_from_cloud import download_journal, upload_journal

    # ダウンロード: クラウド → ローカルフォルダ
    local_dir = download_journal()

    # アップロード: ローカルフォルダ → クラウド（sync後に変更を戻す）
    upload_journal(local_dir)
"""

import json
import os
import subprocess
import sys
from pathlib import Path

# rclone リモート名（rclone config で設定した名前）
RCLONE_REMOTE = "onedrive"

# OneDrive上の日誌フォルダパス
CLOUD_JOURNAL_PATH = "●勉強クラブ共有/09　授業日誌"

# ローカルのダウンロード先（スクリプトと同じフォルダの _cloud_journal）
DEFAULT_LOCAL_DIR = Path(__file__).parent / "_cloud_journal"

# OneDrive は「:」をフォルダ名に使えるが、Windows は使えない。rclone は
# OneDrive 上の「:」を Windows 側では U+201B + U+FF1A として可逆変換する。
# 単独の全角コロン U+FF1A は別名なので、再アップロードすると重複が生じる。
RCLONE_ENCODED_COLON = "\u201b\uff1a"
FULLWIDTH_COLON = "\uff1a"


def validate_cloud_tree_names(local_dir: Path) -> None:
    """同じ論理名になるフォルダが複数あれば本番更新を停止する。"""
    logical_paths = {}
    collisions = []
    for path in local_dir.rglob("*"):
        if not path.is_dir():
            continue
        relative = path.relative_to(local_dir)
        if any(part.startswith("_backup") or part.startswith("退避") for part in relative.parts):
            continue
        logical = tuple(
            part.replace(RCLONE_ENCODED_COLON, ":").replace(FULLWIDTH_COLON, ":")
            for part in relative.parts
        )
        previous = logical_paths.setdefault(logical, relative)
        if previous != relative:
            collisions.append(f"{previous} <-> {relative}")

    if collisions:
        raise RuntimeError(
            "危険なOneDriveフォルダ名の重複を検出したため公開を停止: "
            + ", ".join(sorted(set(collisions)))
        )

# rclone.exe のパス（winget インストール先）
RCLONE_EXE_CANDIDATES = [
    # winget でインストールされた場所
    Path.home() / "AppData" / "Local" / "Microsoft" / "WinGet" / "Links" / "rclone.exe",
]


def find_rclone() -> str:
    """rclone.exe のパスを見つける"""
    for candidate in RCLONE_EXE_CANDIDATES:
        if candidate.exists():
            return str(candidate)
    # winget Packages 配下をバージョン非依存で探索
    pkg_dir = Path.home() / "AppData" / "Local" / "Microsoft" / "WinGet" / "Packages"
    if pkg_dir.exists():
        for p in pkg_dir.glob("Rclone.Rclone*/rclone-*/rclone.exe"):
            if p.is_file():
                return str(p)
    # PATH上にあるか試す
    try:
        subprocess.run(["rclone", "version"], capture_output=True, check=True)
        return "rclone"
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass
    raise FileNotFoundError(
        "rclone が見つかりません。winget install Rclone.Rclone でインストールしてください。"
    )


def _logical_remote_path(path_text: str) -> tuple[str, ...]:
    return tuple(
        part.replace(RCLONE_ENCODED_COLON, ":").replace(FULLWIDTH_COLON, ":")
        for part in path_text.replace("\\", "/").split("/")
        if part
    )


def validate_remote_tree_names(rclone: str, remote_path: str) -> None:
    """OneDrive側の実名を読み、表記違いの重複があれば書き込み前に停止する。"""
    command = [
        rclone,
        "lsjson",
        remote_path,
        "--dirs-only",
        "--recursive",
        "--exclude", "_backup*/**",
        "--exclude", "**/_backup*/**",
        "--exclude", "退避*/**",
        "--exclude", "**/退避/**",
    ]
    result = subprocess.run(
        command, capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    if result.returncode != 0:
        raise RuntimeError(
            "OneDriveフォルダ名の事前監査に失敗したため公開を停止: "
            + (result.stderr.strip() or f"exit {result.returncode}")
        )
    try:
        entries = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError("OneDriveフォルダ名の事前監査結果を解析できないため公開を停止") from exc

    logical_paths = {}
    collisions = []
    for entry in entries:
        original = str(entry.get("Path") or "")
        if not original:
            continue
        parts = tuple(part for part in original.replace("\\", "/").split("/") if part)
        if any(part.startswith("_backup") or part.startswith("退避") for part in parts):
            continue
        logical = _logical_remote_path(original)
        previous = logical_paths.setdefault(logical, original)
        if previous != original:
            collisions.append(f"{previous} <-> {original}")
    if collisions:
        raise RuntimeError(
            "OneDrive上に表記違いの重複フォルダを検出したため公開を停止: "
            + ", ".join(sorted(set(collisions)))
        )


def download_journal(local_dir: Path = None, exclude_backup: bool = True) -> Path:
    """OneDrive クラウドから授業日誌フォルダをダウンロードする。

    Args:
        local_dir: ダウンロード先。未指定時は _cloud_journal/
        exclude_backup: _backup フォルダを除外するか

    Returns:
        ダウンロード先の Path
    """
    if local_dir is None:
        local_dir = DEFAULT_LOCAL_DIR

    local_dir.mkdir(parents=True, exist_ok=True)

    rclone = find_rclone()
    remote_path = f"{RCLONE_REMOTE}:{CLOUD_JOURNAL_PATH}"

    # ローカルの確認だけに頼らず、書き込み直前にOneDrive上の実名も監査する。
    validate_remote_tree_names(rclone, remote_path)

    cmd = [
        rclone, "sync",
        remote_path,
        str(local_dir),
        "--transfers", "8",
        "--checkers", "16",
        "--fast-list",
    ]

    if exclude_backup:
        cmd += [
            "--exclude", "_backup*/**",
            "--exclude", "**/_backup*/**",
            "--exclude", "退避*/**",
            "--exclude", "**/退避/**",
        ]

    print(f"  rclone: {remote_path} → {local_dir}")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")

    if result.returncode != 0:
        print(f"  [ERROR] rclone download 失敗: {result.stderr.strip()}")
        raise RuntimeError(f"rclone download failed (exit {result.returncode})")

    validate_cloud_tree_names(local_dir)

    # ダウンロード結果のファイル数を表示
    xlsx_count = sum(
        1 for path in local_dir.rglob("*.xlsx")
        if not any(part.startswith("_backup") or part.startswith("退避")
                   for part in path.relative_to(local_dir).parts[:-1])
    )
    print(f"  → {xlsx_count} 個のExcelファイルをダウンロード済み")

    return local_dir


def upload_journal(local_dir: Path = None):
    """ローカルの変更をOneDriveクラウドにアップロードする。
    sync_journal_across_campus で変更されたファイルをクラウドに戻す。

    Args:
        local_dir: アップロード元。未指定時は _cloud_journal/
    """
    if local_dir is None:
        local_dir = DEFAULT_LOCAL_DIR

    if not local_dir.exists():
        print("  [SKIP] アップロード元フォルダがありません")
        return

    validate_cloud_tree_names(local_dir)

    rclone = find_rclone()
    remote_path = f"{RCLONE_REMOTE}:{CLOUD_JOURNAL_PATH}"

    validate_remote_tree_names(rclone, remote_path)

    cmd = [
        rclone, "copy",
        str(local_dir),
        remote_path,
        "--transfers", "8",
        "--checkers", "16",
        "--update",  # 新しいファイルのみアップロード
        "--exclude", "_backup*/**",
        "--exclude", "**/_backup*/**",
        "--exclude", "退避*/**",
        "--exclude", "**/退避/**",
    ]

    print(f"  rclone: {local_dir} → {remote_path}")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")

    if result.returncode != 0:
        print(f"  [ERROR] rclone upload 失敗: {result.stderr.strip()}")
        raise RuntimeError(f"rclone upload failed (exit {result.returncode})")

    # 書き込み後にも再監査し、表記違いが生じていないことを確かめる。
    validate_remote_tree_names(rclone, remote_path)
    print("  → アップロード完了（OneDriveフォルダ名の再監査済み）")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="OneDriveクラウドから授業日誌をダウンロード")
    ap.add_argument("--upload", action="store_true", help="ダウンロードではなくアップロード")
    ap.add_argument("--dir", help="ローカルフォルダ")
    args = ap.parse_args()

    d = Path(args.dir) if args.dir else None
    if args.upload:
        upload_journal(d)
    else:
        download_journal(d)
