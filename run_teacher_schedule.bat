@echo off
chcp 65001 > nul
cd /d "%~dp0"

python ".\export_teacher_schedule_json.py" --schedule ".\2026年３月スケジュール.xlsm"

echo.
pause