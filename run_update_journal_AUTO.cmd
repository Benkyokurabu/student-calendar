@echo off
chcp 65001 >nul
cd /d %~dp0
echo ===================================================
echo 授業日誌 全自動更新システム (起動中...)
echo ===================================================

REM 指定された絶対パスを使用してフォルダを設定します
set "JOURNAL_DIR=C:\Users\kudok\OneDrive\●勉強クラブ共有\09　授業日誌"
set "BACKUP_DIR=C:\Users\kudok\OneDrive\●勉強クラブ共有\授業日誌_バックアップ"

echo フォルダ確認中: %JOURNAL_DIR%

REM Pythonを実行してJSON作成・GitHub送信まで一気に行います
python update_journal.py --journal-dir "%JOURNAL_DIR%" --backup-dir "%BACKUP_DIR%" --mode auto --do-backup

echo.
echo ---------------------------------------------------
echo 処理が完了しました。
timeout /t 30