param([Parameter(Mandatory=$true)][string]$PreparationDirectory)
$ErrorActionPreference = 'Stop'
$zoomRepair = (Resolve-Path -LiteralPath $PreparationDirectory).Path
$zoomStage = Join-Path $zoomRepair 'stable-runtime'
$zoomRuntime = Join-Path $env:LOCALAPPDATA 'BenkyoClub\zoom-publisher'
$zoomPython = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python313\python.exe'
$zoomLauncher = Join-Path $zoomRuntime 'zoomURL\run_zoom_publisher_hidden.vbs'
$zoomManifest = Get-Content -LiteralPath (Join-Path $zoomRepair 'stable-runtime-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($entry in $zoomManifest.PSObject.Properties) {
    $source = Join-Path $zoomStage $entry.Name
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.Value) { throw "Staged file changed: $($entry.Name)" }
    $destination = Join-Path $zoomRuntime $entry.Name
    New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination
    if ((Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.Value) { throw "Installed file mismatch: $($entry.Name)" }
}

# Exercise the actual new installation before changing any scheduled action.
& $zoomPython -X utf8 -B (Join-Path $zoomRuntime 'zoomURL\run_zoom_publisher.py') --dry-run
if ($LASTEXITCODE -ne 0) { throw 'New runtime validation failed; scheduled tasks are unchanged.' }

$zoomTasks = @(Get-ScheduledTask | Where-Object { $_.TaskName -like 'BenkyoZoomRecordingURLJson_*' })
if ($zoomTasks.Count -ne 60) { throw "Unexpected Zoom task count: $($zoomTasks.Count)" }
$healthTasks = @(Get-ScheduledTask -TaskName 'BenkyoSystemHealthCheck' -ErrorAction SilentlyContinue)
$zoomBefore = @(($zoomTasks + $healthTasks) | ForEach-Object { [pscustomobject]@{Name=$_.TaskName;Path=$_.TaskPath;Xml=(Export-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath)} })
$zoomBefore | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $zoomRepair 'tasks-before.json') -Encoding UTF8
& $zoomPython -X utf8 -B (Join-Path $PSScriptRoot 'backup_and_update_setup.py') $zoomRepair
if ($LASTEXITCODE -ne 0) { throw 'Cloud backup/setup validation failed; scheduled tasks are unchanged.' }

$zoomUpdated = @()
try {
    foreach ($old in $zoomBefore) {
        $actionLauncher = if ($old.Name -eq 'BenkyoSystemHealthCheck') { Join-Path $zoomRuntime 'run_system_health_check_hidden.vbs' } else { $zoomLauncher }
        [xml]$xml = $old.Xml
        $manager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $manager.AddNamespace('t', $xml.DocumentElement.NamespaceURI)
        $execNode = $xml.SelectSingleNode('/t:Task/t:Actions/t:Exec', $manager)
        if ($null -eq $execNode) { throw "Unexpected action: $($old.Name)" }
        $execNode.SelectSingleNode('t:Command', $manager).InnerText = Join-Path $env:WINDIR 'System32\wscript.exe'
        $execNode.SelectSingleNode('t:Arguments', $manager).InnerText = '"' + $actionLauncher + '"'
        $working = $execNode.SelectSingleNode('t:WorkingDirectory', $manager)
        if ($null -eq $working) { $working = $xml.CreateElement('WorkingDirectory', $xml.DocumentElement.NamespaceURI); $execNode.AppendChild($working) | Out-Null }
        $working.InnerText = $zoomRuntime
        Register-ScheduledTask -TaskName $old.Name -TaskPath $old.Path -Xml $xml.OuterXml -Force | Out-Null
        $zoomUpdated += $old
    }
    $verified = @(Get-ScheduledTask | Where-Object { $_.TaskName -like 'BenkyoZoomRecordingURLJson_*' })
    if ($verified.Count -ne 60 -or @($verified | Where-Object { $_.Actions.Arguments -ne ('"' + $zoomLauncher + '"') }).Count) { throw 'Scheduled action verification failed.' }
} catch {
    foreach ($old in $zoomUpdated) { Register-ScheduledTask -TaskName $old.Name -TaskPath $old.Path -Xml $old.Xml -Force | Out-Null }
    throw
}
Start-ScheduledTask -TaskName 'BenkyoZoomRecordingURLJson_1818'
if ($healthTasks.Count) { Start-ScheduledTask -TaskName 'BenkyoSystemHealthCheck' }
Write-Output 'Installed and verified 60 Zoom task actions outside OneDrive; started one verification run.'
Write-Output ('Runtime log: ' + (Join-Path $zoomRuntime 'logs\zoom_recording_json.log'))
