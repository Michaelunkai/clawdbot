# Verify and fix Clawdbot Gateway startup setup
Write-Host "=== CLAWDBOT STARTUP VERIFICATION ===" -ForegroundColor Cyan

$newPath = "F:\study\AI_ML\AI_and_Machine_Learning\Artificial_Intelligence\cli\claudecode\wrappers\ClawdBot\B"
$startupPath = [Environment]::GetFolderPath('Startup')
$shortcutPath = Join-Path $startupPath "Clawdbot Gateway.lnk"

# 1. Check startup shortcut
Write-Host "`n[1] Checking startup shortcut..." -ForegroundColor Yellow
$ws = New-Object -ComObject WScript.Shell
if (Test-Path $shortcutPath) {
    $shortcut = $ws.CreateShortcut($shortcutPath)
    Write-Host "  Current target: $($shortcut.TargetPath)"
    Write-Host "  Expected target: $newPath\ClawdbotTray.vbs"

    if ($shortcut.TargetPath -ne "$newPath\ClawdbotTray.vbs") {
        Write-Host "  FIXING: Updating shortcut target..." -ForegroundColor Red
        $shortcut.TargetPath = "$newPath\ClawdbotTray.vbs"
        $shortcut.WorkingDirectory = $newPath
        $shortcut.Save()
        Write-Host "  FIXED!" -ForegroundColor Green
    } else {
        Write-Host "  OK!" -ForegroundColor Green
    }
} else {
    Write-Host "  MISSING: Creating shortcut..." -ForegroundColor Red
    $shortcut = $ws.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "$newPath\ClawdbotTray.vbs"
    $shortcut.WorkingDirectory = $newPath
    $shortcut.Description = "Clawdbot Gateway System Tray"
    $shortcut.Save()
    Write-Host "  CREATED!" -ForegroundColor Green
}

# 2. Check for old files in .clawdbot
Write-Host "`n[2] Checking for old tray files in .clawdbot..." -ForegroundColor Yellow
$oldFiles = @(
    "$env:USERPROFILE\.clawdbot\ClawdbotTray.ps1",
    "$env:USERPROFILE\.clawdbot\ClawdbotTray.vbs",
    "$env:USERPROFILE\.clawdbot\create-shortcut.ps1"
)
foreach ($file in $oldFiles) {
    if (Test-Path $file) {
        Write-Host "  Removing old file: $file" -ForegroundColor Red
        Remove-Item $file -Force
    }
}
Write-Host "  OK! No old files." -ForegroundColor Green

# 3. Check registry Run keys for duplicates
Write-Host "`n[3] Checking registry startup entries..." -ForegroundColor Yellow
$runKeys = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
)
foreach ($key in $runKeys) {
    if (Test-Path $key) {
        $entries = Get-ItemProperty $key -ErrorAction SilentlyContinue
        $entries.PSObject.Properties | Where-Object { $_.Value -like "*clawdbot*" -or $_.Value -like "*ClawdbotTray*" } | ForEach-Object {
            Write-Host "  Removing registry entry: $($_.Name)" -ForegroundColor Red
            Remove-ItemProperty -Path $key -Name $_.Name -ErrorAction SilentlyContinue
        }
    }
}
Write-Host "  OK! No duplicate registry entries." -ForegroundColor Green

# 4. Check scheduled tasks
Write-Host "`n[4] Checking scheduled tasks..." -ForegroundColor Yellow
$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "*clawdbot*" -or $_.TaskName -like "*Clawdbot*" }
if ($tasks) {
    foreach ($task in $tasks) {
        Write-Host "  Found task: $($task.TaskName)" -ForegroundColor Red
        # Don't auto-remove, just warn
    }
} else {
    Write-Host "  OK! No scheduled tasks." -ForegroundColor Green
}

# 5. Verify new files exist
Write-Host "`n[5] Verifying new installation files..." -ForegroundColor Yellow
$requiredFiles = @(
    "$newPath\ClawdbotTray.ps1",
    "$newPath\ClawdbotTray.vbs"
)
$allExist = $true
foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Host "  EXISTS: $file" -ForegroundColor Green
    } else {
        Write-Host "  MISSING: $file" -ForegroundColor Red
        $allExist = $false
    }
}

# 6. Summary
Write-Host "`n=== VERIFICATION COMPLETE ===" -ForegroundColor Cyan
if ($allExist) {
    Write-Host "Startup is configured correctly!" -ForegroundColor Green
    Write-Host "Location: $newPath" -ForegroundColor Cyan
} else {
    Write-Host "Some files are missing!" -ForegroundColor Red
}
