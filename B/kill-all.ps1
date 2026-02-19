# Kill all OpenClaw/Clawdbot instances
Write-Host "Stopping OpenClaw/Clawdbot processes..." -ForegroundColor Yellow

# Kill tray app PowerShell processes
Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" | ForEach-Object {
    if ($_.CommandLine -like "*ClawdbotTray*" -or $_.CommandLine -like "*OpenClawTray*") {
        Write-Host "Killing tray app PID $($_.ProcessId)"
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

Start-Sleep -Seconds 1

# Kill gateway node processes (openclaw, clawdbot, moltbot)
Get-WmiObject Win32_Process -Filter "Name='node.exe'" | ForEach-Object {
    if ($_.CommandLine -like "*openclaw*" -or $_.CommandLine -like "*clawdbot*" -or $_.CommandLine -like "*moltbot*") {
        Write-Host "Killing gateway PID $($_.ProcessId)"
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

# Clean up lock files
Remove-Item "$env:TEMP\OpenClawTray.lock" -Force -ErrorAction SilentlyContinue
Remove-Item "$env:TEMP\ClawdbotTray.lock" -Force -ErrorAction SilentlyContinue

Write-Host "Done" -ForegroundColor Green
