# calendar_server.ps1
# Student Calendar local server launcher (PowerShell)
# - Open calendar.html via http://localhost (avoids file:/// fetch block)
# - Picks a free port from 8000..8099
# - Binds to localhost only

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $here

function Get-FreePort([int]$Start=8000,[int]$End=8099) {
  for ($p=$Start; $p -le $End; $p++) {
    $l = $null
    try {
      $l = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $p)
      $l.Start()
      $l.Stop()
      return $p
    } catch {
      if ($l) { try { $l.Stop() } catch {} }
    }
  }
  throw "No free port found in $Start..$End"
}

function Get-PythonCmd() {
  $c = Get-Command py -ErrorAction SilentlyContinue
  if ($c) { return "py" }
  $c = Get-Command python -ErrorAction SilentlyContinue
  if ($c) { return "python" }
  throw "Python not found in PowerShell (py/python)."
}

if (-not (Test-Path -LiteralPath (Join-Path $here "calendar.html"))) {
  Write-Host "[ERROR] calendar.html not found in this folder." -ForegroundColor Red
  Write-Host ("Folder: " + $here)
  Read-Host "Press Enter to close"
  exit 1
}

$port = Get-FreePort 8000 8099
$url  = "http://localhost:$port/calendar.html"

Write-Host ""
Write-Host "========================================"
Write-Host "Student Calendar Local Server"
Write-Host "========================================"
Write-Host ("Folder: " + $here)
Write-Host ("Port  : " + $port)
Write-Host ("URL   : " + $url)
Write-Host ("JSON  : http://localhost:$port/schedule_2026-02.json")
Write-Host ""
Write-Host "A browser will open. To stop the server: press Ctrl+C in this window."
Write-Host ""

Start-Process $url | Out-Null

$py = Get-PythonCmd
& $py -m http.server $port --bind 127.0.0.1
