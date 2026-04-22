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

# rclone.exe のパス（winget インストール先）
RCLONE_EXE_CANDIDATES = [
    # winget でインストールされた場所
    Path.home() / "AppData" / "Local" / "Microsoft" / "WinGet" / "Links" / "rclone.exe",
    # 直接パス
    Path(r"C:\Users\kudok\AppData\Local\Microsoft\WinGet\Packages\Rclone.Rclone_Microsoft.Winget.Source_8wekyb3d8bbwe\rclone-v1.73.5-windows-amd64\rclone.exe"),
]


def find_rclone() -> str:
    """rclone.exe のパスを見つける"""
    for candidate in RCLONE_EXE_CANDIDATES:
        if candidate.exists():
            return str(candidate)
    # PATH上にあるか試す
    try:
        subprocess.run(["rclone", "version"], capture_output=True, check=True)
        return "rclone"
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass
    raise FileNotFoundError(
        "rclone が見つかりません。winget install Rclone.Rclone でインストールしてください。"
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

    cmd = [
        rclone, "sync",
        remote_path,
        str(local_dir),
        "--transfers", "8",
        "--checkers", "16",
        "--fast-list",
    ]

    if exclude_backup:
        cmd += ["--exclude", "_backup/**", "--exclude", "退避/**"]

    print(f"  rclone: {remote_path} → {local_dir}")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")

    if result.returncode != 0:
        print(f"  [ERROR] rclone download 失敗: {result.stderr.strip()}")
        raise RuntimeError(f"rclone download failed (exit {result.returncode})")

    # ダウンロード結果のファイル数を表示
    xlsx_count = len(list(local_dir.rglob("*.xlsx")))
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

    rclone = find_rclone()
    remote_path = f"{RCLONE_REMOTE}:{CLOUD_JOURNAL_PATH}"

    cmd = [
        rclone, "copy",
        str(local_dir),
        remote_path,
        "--transfers", "8",
        "--checkers", "16",
        "--update",  # 新しいファイルのみアップロード
    ]

    print(f"  rclone: {local_dir} → {remote_path}")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")

    if result.returncode != 0:
        print(f"  [WARN] rclone upload 失敗: {result.stderr.strip()}")
    else:
        print("  → アップロード完了")


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
