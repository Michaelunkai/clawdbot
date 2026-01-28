# Create startup shortcut for Clawdbot Gateway Tray
$WshShell = New-Object -ComObject WScript.Shell
$startupPath = [Environment]::GetFolderPath('Startup')
$shortcutPath = Join-Path $startupPath "Clawdbot Gateway.lnk"

$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = "$env:USERPROFILE\.clawdbot\ClawdbotTray.vbs"
$Shortcut.WorkingDirectory = "$env:USERPROFILE\.clawdbot"
$Shortcut.Description = "Clawdbot Gateway System Tray"
$Shortcut.Save()

Write-Host "Startup shortcut created at: $shortcutPath" -ForegroundColor Green
