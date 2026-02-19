# Kill Duplicate ClawdbotTray Instances
$ErrorActionPreference = 'SilentlyContinue'

$instances = @(Get-Process powershell | Where-Object {
    $cmd = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue).CommandLine
    $cmd -like "*ClawdbotTray.ps1*"
})

if ($instances.Count -le 1) {
    Write-Host "Only one or zero ClawdbotTray instances - OK"
    exit 0
}

$sorted = $instances | Sort-Object StartTime
$keep = $sorted[0]
$kill = $sorted | Select-Object -Skip 1

Write-Host "Keeping PID $($keep.Id), killing $($kill.Count) duplicates"

foreach ($proc in $kill) {
    Stop-Process -Id $proc.Id -Force
    Write-Host "Killed PID $($proc.Id)"
}
