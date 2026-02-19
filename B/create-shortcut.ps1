# Create startup shortcut for Clawdbot Gateway Tray
$WshShell = New-Object -ComObject WScript.Shell
$startupPath = [Environment]::GetFolderPath('Startup')
$shortcutPath = Join-Path $startupPath "Clawdbot Gateway.lnk"

$basePath = Split-Path -Parent $MyInvocation.MyCommand.Path

$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = Join-Path $basePath "ClawdbotTray.vbs"
$Shortcut.WorkingDirectory = $basePath
$Shortcut.Description = "Clawdbot Gateway System Tray"
$Shortcut.Save()

Write-Host "Startup shortcut created at: $shortcutPath" -ForegroundColor Green
Write-Host "Target: $($Shortcut.TargetPath)" -ForegroundColor Cyan
