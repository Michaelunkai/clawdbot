# Nuclear option - kill everything and start bulletproof version
Write-Host "Killing all ClawdBotManager processes..." -ForegroundColor Red

# Kill with extreme prejudice
Get-Process -Name "ClawdBotManager" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# Kill again in case it restarted
Get-Process -Name "ClawdBotManager" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# Kill OpenClaw gateway too
Get-Process -Name "node" -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "*openclaw*" } | Stop-Process -Force
Start-Sleep -Seconds 2

Write-Host "Building bulletproof version..." -ForegroundColor Yellow
Set-Location "F:\study\AI_ML\AI_and_Machine_Learning\Artificial_Intelligence\cli\claudecode\wrappers\ClawdBot\b\clawdbotmanager"

# Build to isolated directory
$buildDir = "F:\study\AI_ML\AI_and_Machine_Learning\Artificial_Intelligence\cli\claudecode\wrappers\ClawdBot\b\BULLETPROOF"
Remove-Item -Path $buildDir -Recurse -Force -ErrorAction SilentlyContinue
dotnet publish -c Release -r win-x64 --self-contained -o $buildDir -p:PublishSingleFile=true

Write-Host "Starting bulletproof version..." -ForegroundColor Green
Start-Process "$buildDir\ClawdBotManager.exe" -WorkingDirectory $buildDir

Write-Host "DONE! Bulletproof version with auto-restart is now running!" -ForegroundColor Cyan
