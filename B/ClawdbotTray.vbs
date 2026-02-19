' OpenClaw Gateway Tray Launcher v3.0
' Proper single-instance via named mutex check
' NEVER restarts or closes the gateway - launches PS1 once

Option Explicit
On Error Resume Next

Dim objShell, objFSO, scriptPath, wmi, procs

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check if ClawdbotTray.ps1 is already running via process list
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set procs = wmi.ExecQuery("SELECT ProcessId,CommandLine,CreationDate FROM Win32_Process WHERE Name='powershell.exe'")
Dim proc, runningProcs, oldestProc, oldestDate
Set runningProcs = CreateObject("Scripting.Dictionary")

' Find all ClawdbotTray instances
For Each proc In procs
    If InStr(1, proc.CommandLine, "ClawdbotTray.ps1", vbTextCompare) > 0 Then
        runningProcs.Add proc.ProcessId, proc.CreationDate
    End If
Next

' If any instance exists, exit (don't start another)
If runningProcs.Count > 0 Then
    WScript.Quit 0
End If

Err.Clear

scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ClawdbotTray.ps1"

' Verify script exists
If Not objFSO.FileExists(scriptPath) Then
    WScript.Quit 1
End If

' Set critical environment variables
Dim procEnv, userEnv, oauthToken
Set procEnv = objShell.Environment("Process")
Set userEnv = objShell.Environment("User")
oauthToken = userEnv("CLAUDE_CODE_OAUTH_TOKEN")
If oauthToken <> "" Then procEnv("CLAUDE_CODE_OAUTH_TOKEN") = oauthToken

procEnv("SHELL") = objShell.ExpandEnvironmentStrings("%COMSPEC%")
procEnv("OPENCLAW_SHELL") = "cmd"
procEnv("OPENCLAW_NO_WSL") = "1"
procEnv("OPENCLAW_NO_PTY") = "1"

' Launch PowerShell once - hidden, no window
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -File """ & scriptPath & """", 0, False
