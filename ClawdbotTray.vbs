' Clawdbot Gateway Tray Launcher
' Runs PowerShell script completely hidden (no console window)

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ClawdbotTray.ps1"

' Run PowerShell hidden with bypass execution policy
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """", 0, False
