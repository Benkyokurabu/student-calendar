@echo off
setlocal
cd /d "%~dp0"

set "LOGDIR=%~dp0logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"

set "LOGFILE=%LOGDIR%\zoom_recording_json.log"

set "PYTHON_EXE=python"
where python >nul 2>nul
if errorlevel 1 set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"

echo [%date% %time%] START >> "%LOGFILE%" 2>&1
"%PYTHON_EXE%" "%~dp0publish_zoom_recording_url_json.py" >> "%LOGFILE%" 2>&1
set "EXITCODE=%ERRORLEVEL%"
echo [%date% %time%] END exit=%EXITCODE% >> "%LOGFILE%" 2>&1

set "ALERTFILE=%USERPROFILE%\OneDrive\デスクトップ\★★Zoom録画URL公開エラー★★.txt"
if %EXITCODE% NEQ 0 (
  echo 授業名: Zoom録画URL公開処理> "%ALERTFILE%"
  echo 書き込むはずだったURL: Excelには書き込みません。公開用JSON生成・GitHub反映で失敗しました。>> "%ALERTFILE%"
  echo.>> "%ALERTFILE%"
  echo ログを確認してください:>> "%ALERTFILE%"
  echo %LOGFILE%>> "%ALERTFILE%"
) else (
  if exist "%ALERTFILE%" del "%ALERTFILE%"
)
exit /b %EXITCODE%
