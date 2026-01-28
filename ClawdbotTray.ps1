# Clawdbot Gateway System Tray Application
# Run hidden at startup, provides tray icon with controls

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global state
$script:gatewayProcess = $null
$script:enabled = $true
$script:logFile = "$env:TEMP\clawdbot\clawdbot-$(Get-Date -Format 'yyyy-MM-dd').log"

# Create NotifyIcon
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = [System.Drawing.SystemIcons]::Application
$trayIcon.Text = "Clawdbot Gateway"
$trayIcon.Visible = $true

# Try to use a custom icon if available
$iconPath = "$env:USERPROFILE\.clawdbot\lobster.ico"
if (Test-Path $iconPath) {
    $trayIcon.Icon = New-Object System.Drawing.Icon($iconPath)
}

# Context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Status item (non-clickable header)
$statusItem = New-Object System.Windows.Forms.ToolStripMenuItem
$statusItem.Text = "Status: Starting..."
$statusItem.Enabled = $false
$contextMenu.Items.Add($statusItem) | Out-Null

$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

# Start/Enable
$startItem = New-Object System.Windows.Forms.ToolStripMenuItem
$startItem.Text = "Start Gateway"
$contextMenu.Items.Add($startItem) | Out-Null

# Stop/Disable
$stopItem = New-Object System.Windows.Forms.ToolStripMenuItem
$stopItem.Text = "Stop Gateway"
$contextMenu.Items.Add($stopItem) | Out-Null

# Restart
$restartItem = New-Object System.Windows.Forms.ToolStripMenuItem
$restartItem.Text = "Restart Gateway"
$contextMenu.Items.Add($restartItem) | Out-Null

$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

# Show Terminal
$terminalItem = New-Object System.Windows.Forms.ToolStripMenuItem
$terminalItem.Text = "Show Live Terminal"
$contextMenu.Items.Add($terminalItem) | Out-Null

# Open Log File
$logItem = New-Object System.Windows.Forms.ToolStripMenuItem
$logItem.Text = "Open Log File"
$contextMenu.Items.Add($logItem) | Out-Null

# Doctor
$doctorItem = New-Object System.Windows.Forms.ToolStripMenuItem
$doctorItem.Text = "Run Doctor"
$contextMenu.Items.Add($doctorItem) | Out-Null

$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

# Exit
$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "Exit"
$contextMenu.Items.Add($exitItem) | Out-Null

$trayIcon.ContextMenuStrip = $contextMenu

# Functions
function Update-Status {
    param([string]$status, [string]$tooltip)
    $statusItem.Text = "Status: $status"
    $trayIcon.Text = "Clawdbot Gateway - $status"
    if ($tooltip) {
        $trayIcon.BalloonTipTitle = "Clawdbot Gateway"
        $trayIcon.BalloonTipText = $tooltip
        $trayIcon.ShowBalloonTip(2000)
    }
}

function Start-Gateway {
    if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
        Update-Status "Running" "Gateway already running"
        return
    }

    $script:enabled = $true
    Update-Status "Starting..." $null

    # Start gateway process hidden
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "cmd.exe"
    $psi.Arguments = "/c clawdbot gateway"
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $psi.CreateNoWindow = $true
    $psi.UseShellExecute = $false

    try {
        $script:gatewayProcess = [System.Diagnostics.Process]::Start($psi)
        Start-Sleep -Milliseconds 2000

        if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
            Update-Status "Running" "Gateway started successfully"
            $startItem.Enabled = $false
            $stopItem.Enabled = $true
            $restartItem.Enabled = $true
        } else {
            Update-Status "Failed" "Gateway failed to start"
            $startItem.Enabled = $true
            $stopItem.Enabled = $false
        }
    } catch {
        Update-Status "Error" "Failed to start: $_"
        $startItem.Enabled = $true
        $stopItem.Enabled = $false
    }
}

function Stop-Gateway {
    $script:enabled = $false
    Update-Status "Stopping..." $null

    # Stop using clawdbot command first
    try {
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c clawdbot gateway stop" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
    } catch {}

    # Kill any remaining gateway processes
    Get-Process -Name "node" -ErrorAction SilentlyContinue | Where-Object {
        $_.CommandLine -like "*clawdbot*gateway*"
    } | Stop-Process -Force -ErrorAction SilentlyContinue

    if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
        $script:gatewayProcess.Kill()
    }
    $script:gatewayProcess = $null

    Start-Sleep -Milliseconds 500
    Update-Status "Stopped" "Gateway stopped"
    $startItem.Enabled = $true
    $stopItem.Enabled = $false
    $restartItem.Enabled = $false
}

function Restart-Gateway {
    Stop-Gateway
    Start-Sleep -Milliseconds 1000
    Start-Gateway
}

function Show-Terminal {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "Write-Host 'Clawdbot Gateway Live Output' -ForegroundColor Cyan; Write-Host '============================' -ForegroundColor Cyan; Get-Content '$script:logFile' -Wait -Tail 50"
}

function Open-LogFile {
    if (Test-Path $script:logFile) {
        Start-Process notepad.exe $script:logFile
    } else {
        [System.Windows.Forms.MessageBox]::Show("Log file not found: $script:logFile", "Clawdbot", "OK", "Information")
    }
}

function Run-Doctor {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "clawdbot doctor; Read-Host 'Press Enter to close'"
}

# Event handlers
$startItem.Add_Click({ Start-Gateway })
$stopItem.Add_Click({ Stop-Gateway })
$restartItem.Add_Click({ Restart-Gateway })
$terminalItem.Add_Click({ Show-Terminal })
$logItem.Add_Click({ Open-LogFile })
$doctorItem.Add_Click({ Run-Doctor })

$exitItem.Add_Click({
    Stop-Gateway
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
    [System.Windows.Forms.Application]::Exit()
})

# Double-click to show status
$trayIcon.Add_DoubleClick({
    $status = if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) { "Running" } else { "Stopped" }
    [System.Windows.Forms.MessageBox]::Show("Gateway Status: $status`nPID: $(if($script:gatewayProcess){$script:gatewayProcess.Id}else{'N/A'})`nLog: $script:logFile", "Clawdbot Gateway", "OK", "Information")
})

# Initial state
$startItem.Enabled = $true
$stopItem.Enabled = $false
$restartItem.Enabled = $false

# Auto-start gateway
Start-Gateway

# Timer to check gateway health
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10000  # Check every 10 seconds
$timer.Add_Tick({
    if ($script:enabled) {
        if (!$script:gatewayProcess -or $script:gatewayProcess.HasExited) {
            Update-Status "Crashed - Restarting..." $null
            Start-Gateway
        }
    }
})
$timer.Start()

# Run message loop
[System.Windows.Forms.Application]::Run()
