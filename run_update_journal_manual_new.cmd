@echo off
chcp 65001 >nul
echo ===================================================
echo 授業日誌 JSON 手動更新スクリプト (新フォルダ構成対応版)
echo ===================================================
echo.

set "BASE_DIR=%USERPROFILE%\OneDrive"
set "TARGET_DIR="

echo 対象フォルダ「09　授業日誌」を検索しています...

REM 優先パスをチェック
set "KNOWN_DIR=%BASE_DIR%\●勉強クラブ共有\09　授業日誌"
if exist "%KNOWN_DIR%" (
    set "TARGET_DIR=%KNOWN_DIR%"
    goto :FOUND
)

REM 見つからない場合は2階層目まで検索
for /d %%A in ("%BASE_DIR%\*") do (
    if exist "%%A\09　授業日誌" (
        set "TARGET_DIR=%%A\09　授業日誌"
        goto :FOUND
    )
)

:FOUND
if "%TARGET_DIR%"=="" (
    echo.
    echo [エラー] 「09　授業日誌」フォルダが見つかりませんでした。
    echo パスが正しいか確認してください。
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] フォルダを確認しました: %TARGET_DIR%
echo これから全てのサブフォルダ内のエクセルを検索し、JSONを更新します。
echo ---------------------------------------------------

python update_journal.py --journal-dir "%TARGET_DIR%" --mode manual --no-git

echo ---------------------------------------------------
echo 処理が終わりました。内容を確認してください。
pause