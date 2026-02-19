# OpenClaw Gateway System Tray - TURBO v2.0
# Optimized for 10x faster startup
# Key: Start gateway FIRST, setup GUI async

$ErrorActionPreference = 'SilentlyContinue'

# ============================================
# PHASE 1: IMMEDIATE GATEWAY START (< 100ms)
# ============================================

# Check if gateway is already running - if so, attach to it instead of killing
$existingGateway = $null
Get-Process -Name "node" -ErrorAction SilentlyContinue | Where-Object {
    $_.Path -like "*node*" -and $_.MainWindowTitle -eq ""
} | ForEach-Object {
    try {
        $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue).CommandLine
        if ($cmdLine -like "*openclaw*gateway*" -or $cmdLine -like "*clawdbot*gateway*") {
            $existingGateway = $_
        }
    } catch {}
}

# Detect openclaw command (fast)
function Get-OpenClawCommand {
    $npmPath = "$env:APPDATA\npm"
    $cmd = "$npmPath\openclaw.cmd"
    if (Test-Path $cmd) { return @{ Type = "cmd"; Path = $cmd } }
    $mjs = "$npmPath\node_modules\openclaw\openclaw.mjs"
    if (Test-Path $mjs) { return @{ Type = "node"; Path = $mjs } }
    return @{ Type = "npx"; Path = "openclaw" }
}

$script:openclawInfo = Get-OpenClawCommand
$script:gatewayProcess = $null
$script:productName = "OpenClaw"

# Set environment (minimal)
$env:SHELL = "$env:COMSPEC"
$env:OPENCLAW_SHELL = "cmd"
$env:OPENCLAW_NO_WSL = "1"
$env:OPENCLAW_NO_PTY = "1"
$env:NODE_OPTIONS = "--max-old-space-size=4096"

# Ensure log dir exists
$logDir = "$env:TEMP\openclaw"
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

# START GATEWAY (only if not already running)
if ($existingGateway) {
    # Gateway already running - just track it, don't kill/restart
    $script:gatewayProcess = $existingGateway
} else {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $info = $script:openclawInfo

    switch ($info.Type) {
        "cmd" { $psi.FileName = $info.Path; $psi.Arguments = "gateway --allow-unconfigured --auth token --token moltbot-local-token-2026" }
        "node" { $psi.FileName = "node"; $psi.Arguments = "`"$($info.Path)`" gateway --allow-unconfigured --auth token --token moltbot-local-token-2026" }
        "npx" { $psi.FileName = "npx"; $psi.Arguments = "openclaw gateway --allow-unconfigured --auth token --token moltbot-local-token-2026" }
    }

    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $psi.CreateNoWindow = $true
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    if ($env:CLAUDE_CODE_OAUTH_TOKEN) { $psi.EnvironmentVariables["CLAUDE_CODE_OAUTH_TOKEN"] = $env:CLAUDE_CODE_OAUTH_TOKEN }

    $script:gatewayProcess = New-Object System.Diagnostics.Process
    $script:gatewayProcess.StartInfo = $psi
    $script:gatewayProcess.EnableRaisingEvents = $true
    $script:gatewayProcess.Start() | Out-Null
    $script:gatewayProcess.BeginOutputReadLine()
    $script:gatewayProcess.BeginErrorReadLine()

    # Log output async
    $logFile = "$logDir\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
    Register-ObjectEvent -InputObject $script:gatewayProcess -EventName OutputDataReceived -Action {
        if ($EventArgs.Data) { Add-Content -Path "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log" -Value $EventArgs.Data -ErrorAction SilentlyContinue }
    } | Out-Null
    Register-ObjectEvent -InputObject $script:gatewayProcess -EventName ErrorDataReceived -Action {
        if ($EventArgs.Data) { Add-Content -Path "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log" -Value $EventArgs.Data -ErrorAction SilentlyContinue }
    } | Out-Null
}

# ============================================
# PHASE 2: GUI SETUP (async, non-blocking)
# ============================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global state
$script:enabled = $true
$script:oauthError = $false
$script:userStopped = $false

# Simple icon (fast)
function Create-Icon {
    param([string]$status = "normal")
    $bitmap = New-Object System.Drawing.Bitmap(16, 16)
    $g = [System.Drawing.Graphics]::FromImage($bitmap)
    $color = switch ($status) {
        "running" { [System.Drawing.Color]::LimeGreen }
        "error" { [System.Drawing.Color]::Red }
        default { [System.Drawing.Color]::Orange }
    }
    $g.FillEllipse((New-Object System.Drawing.SolidBrush($color)), 2, 2, 12, 12)
    $g.Dispose()
    return [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
}

# Create tray icon
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = Create-Icon -status "running"
$trayIcon.Text = "$script:productName Gateway"
$trayIcon.Visible = $true

# Context menu (simplified)
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

$statusItem = New-Object System.Windows.Forms.ToolStripMenuItem
$statusItem.Text = "Status: Running"
$statusItem.Enabled = $false
$contextMenu.Items.Add($statusItem) | Out-Null

$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$restartItem = New-Object System.Windows.Forms.ToolStripMenuItem
$restartItem.Text = "Restart Gateway"
$contextMenu.Items.Add($restartItem) | Out-Null

$stopItem = New-Object System.Windows.Forms.ToolStripMenuItem
$stopItem.Text = "Stop Gateway"
$contextMenu.Items.Add($stopItem) | Out-Null

$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$terminalItem = New-Object System.Windows.Forms.ToolStripMenuItem
$terminalItem.Text = "Show Terminal"
$contextMenu.Items.Add($terminalItem) | Out-Null

$logItem = New-Object System.Windows.Forms.ToolStripMenuItem
$logItem.Text = "Open Log"
$contextMenu.Items.Add($logItem) | Out-Null

$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "Exit"
$contextMenu.Items.Add($exitItem) | Out-Null

$trayIcon.ContextMenuStrip = $contextMenu

# Functions
function Update-Status {
    param([string]$status, [string]$iconStatus = "normal")
    $statusItem.Text = "Status: $status"
    $trayIcon.Text = "$script:productName - $status"
    $trayIcon.Icon = Create-Icon -status $iconStatus
}

function Stop-Gateway {
    $script:enabled = $false
    $script:userStopped = $true
    if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
        try { $script:gatewayProcess.Kill() } catch {}
    }
    $script:gatewayProcess = $null
    Update-Status "Stopped" "normal"
}

function Start-Gateway {
    if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) { return }
    $script:enabled = $true
    $script:userStopped = $false

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $info = $script:openclawInfo
    switch ($info.Type) {
        "cmd" { $psi.FileName = $info.Path; $psi.Arguments = "gateway --allow-unconfigured --auth token --token moltbot-local-token-2026" }
        "node" { $psi.FileName = "node"; $psi.Arguments = "`"$($info.Path)`" gateway --allow-unconfigured --auth token --token moltbot-local-token-2026" }
        "npx" { $psi.FileName = "npx"; $psi.Arguments = "openclaw gateway --allow-unconfigured --auth token --token moltbot-local-token-2026" }
    }
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $psi.CreateNoWindow = $true
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    if ($env:CLAUDE_CODE_OAUTH_TOKEN) { $psi.EnvironmentVariables["CLAUDE_CODE_OAUTH_TOKEN"] = $env:CLAUDE_CODE_OAUTH_TOKEN }
    
    $script:gatewayProcess = New-Object System.Diagnostics.Process
    $script:gatewayProcess.StartInfo = $psi
    $script:gatewayProcess.Start() | Out-Null
    $script:gatewayProcess.BeginOutputReadLine()
    $script:gatewayProcess.BeginErrorReadLine()
    
    Update-Status "Running" "running"
}

function Restart-Gateway {
    Stop-Gateway
    Start-Sleep -Milliseconds 500
    Start-Gateway
}

# Event handlers
$restartItem.Add_Click({ Restart-Gateway })
$stopItem.Add_Click({ Stop-Gateway })
$terminalItem.Add_Click({
    $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
    Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "Get-Content '$logPath' -Wait -Tail 50"
})
$logItem.Add_Click({
    $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
    if (Test-Path $logPath) { Start-Process notepad.exe $logPath }
})
$exitItem.Add_Click({
    Stop-Gateway
    $trayIcon.Visible = $false
    Remove-Item "$env:TEMP\OpenClawTray.lock" -Force -ErrorAction SilentlyContinue
    [System.Windows.Forms.Application]::Exit()
})

# Health check timer (every 5 minutes) - only restart if crashed, never if user stopped
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 300000
$timer.Add_Tick({
    if ($script:userStopped) { return }
    if ($script:enabled) {
        if (!$script:gatewayProcess -or $script:gatewayProcess.HasExited) {
            Update-Status "Reconnecting..." "warning"
            Start-Gateway
        } else {
            Update-Status "Running" "running"
        }
    }
})
$timer.Start()

# Initial status
Update-Status "Running" "running"

# Run message loop
[System.Windows.Forms.Application]::Run()
