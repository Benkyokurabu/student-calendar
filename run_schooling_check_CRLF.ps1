Set-Location -LiteralPath $PSScriptRoot

$py = $null
if (Test-Path ".\.venv\Scripts\python.exe") { $py = ".\.venv\Scripts\python.exe" }
elseif (Get-Command py -ErrorAction SilentlyContinue) { $py = "py" }
elseif (Get-Command python -ErrorAction SilentlyContinue) { $py = "python" }

if (-not $py) {
  Write-Host "[ERROR] Python not found." -ForegroundColor Red
  Write-Host "Place this file in the same folder where .venv exists, or install Python."
  Read-Host "Press Enter to exit"
  exit 1
}

Write-Host "========================================"
Write-Host "Schooling Border Check"
Write-Host "========================================"
& $py "check_schooling_border.py" "2026年2月スケジュール.xlsm"
Write-Host ""
Read-Host "Press Enter to exit"
