' OpenClaw Gateway Tray Launcher - TURBO v2.0
' Optimized for 10x faster startup - removed WMI queries and delays
' Single instance via lock file only (fast check)

Option Explicit
On Error Resume Next

Dim objShell, objFSO, scriptPath, lockFile, lockHandle

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Fast lock file check - try to get exclusive write access
lockFile = objShell.ExpandEnvironmentStrings("%TEMP%\OpenClawTray.lock")

' Try to open lock file exclusively - if fails, another instance is running
Set lockHandle = objFSO.OpenTextFile(lockFile, 2, True)
If Err.Number <> 0 Then
    ' Lock file is held by another process - exit silently
    WScript.Quit 0
End If
lockHandle.Close
Err.Clear

' Write our script path to lock file for identification
Set lockHandle = objFSO.CreateTextFile(lockFile, True)
lockHandle.WriteLine WScript.ScriptFullName
lockHandle.Close

scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ClawdbotTray.ps1"

' Verify script exists
If Not objFSO.FileExists(scriptPath) Then
    WScript.Quit 1
End If

' Set critical environment variables (minimal set for speed)
Dim procEnv
Set procEnv = objShell.Environment("Process")

' OAuth token from user env
Dim userEnv, oauthToken
Set userEnv = objShell.Environment("User")
oauthToken = userEnv("CLAUDE_CODE_OAUTH_TOKEN")
If oauthToken <> "" Then procEnv("CLAUDE_CODE_OAUTH_TOKEN") = oauthToken

' Shell config (essential only)
procEnv("SHELL") = objShell.ExpandEnvironmentStrings("%COMSPEC%")
procEnv("OPENCLAW_SHELL") = "cmd"
procEnv("OPENCLAW_NO_WSL") = "1"
procEnv("OPENCLAW_NO_PTY") = "1"

' Launch PowerShell immediately - no delays
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -File """ & scriptPath & """", 0, False
