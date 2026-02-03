' OpenClaw Gateway Tray Launcher
' Runs PowerShell script completely hidden (no console window)
' Ensures only ONE instance runs - this is the ONLY gateway entry point
' Auto-fixes OAuth token by reading from user environment
' Updated for OpenClaw stable (2026.2.x)

Option Explicit

Dim objShell, objFSO, objWMI, objEnv
Dim colProcesses, objProcess, scriptPath, oauthToken
Dim lockFile, lockStream, existingPid

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set objEnv = objShell.Environment("User")

' Lock file path for single instance check
lockFile = objShell.ExpandEnvironmentStrings("%TEMP%\OpenClawTray.lock")

' Check for existing instance via lock file - silently exit if already running
If objFSO.FileExists(lockFile) Then
    On Error Resume Next
    existingPid = Trim(objFSO.OpenTextFile(lockFile, 1).ReadLine())
    On Error GoTo 0
    If existingPid <> "" And IsNumeric(existingPid) Then
        Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & CLng(existingPid))
        If colProcesses.Count > 0 Then
            ' Already running - silently exit (no popup)
            WScript.Quit 0
        End If
    End If
    ' Stale lock file - remove it
    On Error Resume Next
    objFSO.DeleteFile lockFile, True
    On Error GoTo 0
End If

' Kill any existing tray PowerShell processes
Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'powershell.exe'")
For Each objProcess In colProcesses
    If InStr(objProcess.CommandLine, "ClawdbotTray") > 0 Or InStr(objProcess.CommandLine, "OpenClawTray") > 0 Then
        On Error Resume Next
        objProcess.Terminate()
        On Error GoTo 0
    End If
Next

' Kill any orphaned gateway processes (openclaw, clawdbot, moltbot)
Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'node.exe'")
For Each objProcess In colProcesses
    If objProcess.CommandLine <> "" Then
        If (InStr(LCase(objProcess.CommandLine), "openclaw") > 0 Or _
            InStr(LCase(objProcess.CommandLine), "clawdbot") > 0 Or _
            InStr(LCase(objProcess.CommandLine), "moltbot") > 0) And _
            InStr(LCase(objProcess.CommandLine), "gateway") > 0 Then
            On Error Resume Next
            objProcess.Terminate()
            On Error GoTo 0
        End If
    End If
Next

' Wait for processes to terminate
WScript.Sleep 500

scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ClawdbotTray.ps1"

' Verify script exists
If Not objFSO.FileExists(scriptPath) Then
    MsgBox "Error: ClawdbotTray.ps1 not found at:" & vbCrLf & scriptPath, vbCritical, "OpenClaw Gateway"
    WScript.Quit 1
End If

' Ensure CLAUDE_CODE_OAUTH_TOKEN is in process environment (inherit from user env)
oauthToken = objEnv("CLAUDE_CODE_OAUTH_TOKEN")
If oauthToken <> "" Then
    objShell.Environment("Process")("CLAUDE_CODE_OAUTH_TOKEN") = oauthToken
End If

' Run PowerShell hidden with bypass execution policy
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -File """ & scriptPath & """", 0, False
