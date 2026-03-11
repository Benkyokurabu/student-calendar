@echo off
chcp 65001 >nul
setlocal

REM ========= 設定（ここだけ編集） =========
REM 授業日誌Excelが入っている「公開OneDriveフォルダ（同期済み）」のパス
set "BENKYO_JOURNAL_DIR=C:\OneDrive_Public\授業日誌"
REM =======================================

REM このcmdは student-calendar リポジトリ直下に置いて実行してください
python update_journal.py --mode manual

endlocal
