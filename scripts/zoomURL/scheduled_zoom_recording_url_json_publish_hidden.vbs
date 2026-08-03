Option Explicit

Dim shell, fso, scriptDir, commandPath, exitCode
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
commandPath = fso.BuildPath(scriptDir, "scheduled_zoom_recording_url_json_publish.cmd")
exitCode = shell.Run("cmd.exe /c """ & commandPath & """", 0, True)
WScript.Quit exitCode
