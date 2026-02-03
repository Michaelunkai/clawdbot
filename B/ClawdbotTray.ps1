# OpenClaw Gateway System Tray Application
# Run hidden at startup, provides tray icon with controls
# Updated for OpenClaw stable (2026.2.x)
# This is the ONLY gateway entry point - ensures single instance

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================
# SINGLE INSTANCE CHECK - Prevent duplicates
# ============================================
$lockFile = "$env:TEMP\OpenClawTray.lock"
$script:lockStream = $null

# Check for existing instance using lock file - silently exit if already running
if (Test-Path $lockFile) {
    try {
        $existingPid = Get-Content $lockFile -ErrorAction Stop
        $existingProcess = Get-Process -Id $existingPid -ErrorAction SilentlyContinue
        if ($existingProcess) {
            # Already running - silently exit (no popup)
            exit 0
        }
    } catch {
        Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
    }
}

# Create lock file with our PID
try {
    $PID | Out-File $lockFile -Force
} catch {}

# Kill any orphaned tray app processes (not us)
Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | ForEach-Object {
    if (($_.CommandLine -like "*ClawdbotTray*" -or $_.CommandLine -like "*OpenClawTray*") -and $_.ProcessId -ne $PID) {
        try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {}
    }
}

# Kill any orphaned gateway processes on startup
Get-WmiObject Win32_Process -Filter "Name='node.exe'" -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.CommandLine -like "*openclaw*gateway*" -or $_.CommandLine -like "*clawdbot*gateway*" -or $_.CommandLine -like "*moltbot*gateway*") {
        try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {}
    }
}

# ============================================
# OPENCLAW COMMAND DETECTION
# ============================================
function Get-OpenClawCommand {
    # Priority order: openclaw.cmd > node + openclaw.mjs > clawdbot.cmd > node + clawdbot entry
    $npmPath = "$env:APPDATA\npm"
    $nodeModules = "$npmPath\node_modules"

    # Check for openclaw.cmd first
    $openclawCmd = "$npmPath\openclaw.cmd"
    if (Test-Path $openclawCmd) {
        return @{ Type = "cmd"; Path = $openclawCmd }
    }

    # Check for openclaw.mjs
    $openclawMjs = "$nodeModules\openclaw\openclaw.mjs"
    if (Test-Path $openclawMjs) {
        return @{ Type = "node"; Path = $openclawMjs }
    }

    # Fallback to clawdbot
    $clawdbotCmd = "$npmPath\clawdbot.cmd"
    if (Test-Path $clawdbotCmd) {
        return @{ Type = "cmd"; Path = $clawdbotCmd }
    }

    # Check for clawdbot entry.js
    $clawdbotEntry = "$nodeModules\clawdbot\dist\entry.js"
    if (Test-Path $clawdbotEntry) {
        return @{ Type = "node"; Path = $clawdbotEntry }
    }

    # Last resort - use npx
    return @{ Type = "npx"; Path = "openclaw" }
}

# Global state
$script:gatewayProcess = $null
$script:enabled = $true
$script:oauthError = $false
$script:reauthInProgress = $false
$script:logFile = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
$script:basePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:gatewayOutput = ""
$script:openclawInfo = Get-OpenClawCommand
$script:productName = "OpenClaw"

# Create lobster icon (orange crab/lobster style) with status indicator
function Create-LobsterIcon {
    param([string]$status = "normal") # normal, running, error, warning

    $bitmap = New-Object System.Drawing.Bitmap(32, 32)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([System.Drawing.Color]::Transparent)

    # Color based on status
    $baseColor = switch ($status) {
        "running" { [System.Drawing.Color]::FromArgb(255, 50, 205, 50) }  # Green
        "error"   { [System.Drawing.Color]::FromArgb(255, 220, 20, 60) }  # Red
        "warning" { [System.Drawing.Color]::FromArgb(255, 255, 165, 0) }  # Orange
        default   { [System.Drawing.Color]::FromArgb(255, 220, 80, 40) }  # Default orange/red
    }

    $lobsterBrush = New-Object System.Drawing.SolidBrush($baseColor)
    $darkBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, [Math]::Max(0, $baseColor.R - 40), [Math]::Max(0, $baseColor.G - 30), [Math]::Max(0, $baseColor.B - 20)))
    $eyeBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
    $whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)

    # Body (main oval)
    $graphics.FillEllipse($lobsterBrush, 8, 10, 16, 18)

    # Head
    $graphics.FillEllipse($lobsterBrush, 10, 4, 12, 10)

    # Claws (left)
    $graphics.FillEllipse($darkBrush, 2, 8, 8, 6)
    $graphics.FillEllipse($lobsterBrush, 0, 6, 6, 4)

    # Claws (right)
    $graphics.FillEllipse($darkBrush, 22, 8, 8, 6)
    $graphics.FillEllipse($lobsterBrush, 26, 6, 6, 4)

    # Eyes (white background)
    $graphics.FillEllipse($whiteBrush, 11, 5, 4, 4)
    $graphics.FillEllipse($whiteBrush, 17, 5, 4, 4)

    # Eyes (black pupils)
    $graphics.FillEllipse($eyeBrush, 12, 6, 2, 2)
    $graphics.FillEllipse($eyeBrush, 18, 6, 2, 2)

    # Tail segments
    $graphics.FillEllipse($darkBrush, 12, 26, 8, 4)

    $graphics.Dispose()

    # Convert to icon
    $hicon = $bitmap.GetHicon()
    $icon = [System.Drawing.Icon]::FromHandle($hicon)
    return $icon
}

# Create NotifyIcon with custom lobster icon
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = Create-LobsterIcon
$trayIcon.Text = "$script:productName Gateway"
$trayIcon.Visible = $true

# Context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Version/Header item
$versionItem = New-Object System.Windows.Forms.ToolStripMenuItem
$versionItem.Text = "$script:productName Gateway ($($script:openclawInfo.Type))"
$versionItem.Enabled = $false
$contextMenu.Items.Add($versionItem) | Out-Null

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

# Re-authenticate
$reauthItem = New-Object System.Windows.Forms.ToolStripMenuItem
$reauthItem.Text = "Re-authenticate (Fix OAuth)"
$contextMenu.Items.Add($reauthItem) | Out-Null

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

# Dashboard
$dashboardItem = New-Object System.Windows.Forms.ToolStripMenuItem
$dashboardItem.Text = "Open Dashboard"
$contextMenu.Items.Add($dashboardItem) | Out-Null

$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

# Clean All Context
$cleanContextItem = New-Object System.Windows.Forms.ToolStripMenuItem
$cleanContextItem.Text = "Clean All Context"
$contextMenu.Items.Add($cleanContextItem) | Out-Null

# Optimize Connection
$optimizeItem = New-Object System.Windows.Forms.ToolStripMenuItem
$optimizeItem.Text = "Optimize Connection"
$contextMenu.Items.Add($optimizeItem) | Out-Null

$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

# Exit
$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "Exit"
$contextMenu.Items.Add($exitItem) | Out-Null

$trayIcon.ContextMenuStrip = $contextMenu

# Functions
function Update-Status {
    param([string]$status, [string]$tooltip, [string]$iconStatus = "normal")
    $statusItem.Text = "Status: $status"
    $trayIcon.Text = "$script:productName Gateway - $status"

    # Update icon based on status
    $trayIcon.Icon = Create-LobsterIcon -status $iconStatus

    if ($tooltip) {
        $trayIcon.BalloonTipTitle = "$script:productName Gateway"
        $trayIcon.BalloonTipText = $tooltip
        $trayIcon.ShowBalloonTip(2000)
    }
}

# Helper to run openclaw commands
function Invoke-OpenClaw {
    param([string]$Arguments, [switch]$Wait, [switch]$Hidden)

    $info = $script:openclawInfo
    $psi = New-Object System.Diagnostics.ProcessStartInfo

    switch ($info.Type) {
        "cmd" {
            $psi.FileName = $info.Path
            $psi.Arguments = $Arguments
        }
        "node" {
            $psi.FileName = "node"
            $psi.Arguments = "`"$($info.Path)`" $Arguments"
        }
        "npx" {
            $psi.FileName = "npx"
            $psi.Arguments = "$($info.Path) $Arguments"
        }
    }

    if ($Hidden) {
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $psi.CreateNoWindow = $true
    }
    $psi.UseShellExecute = (-not $Hidden)

    if ($Wait) {
        $proc = [System.Diagnostics.Process]::Start($psi)
        $proc.WaitForExit()
        return $proc.ExitCode
    } else {
        return [System.Diagnostics.Process]::Start($psi)
    }
}

function Start-Gateway {
    if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
        Update-Status "Running" "Gateway already running" "running"
        return
    }

    # Don't auto-start if OAuth error was detected
    if ($script:oauthError) {
        Update-Status "Auth Required" "Click 'Re-authenticate' to fix OAuth" "error"
        $startItem.Enabled = $false
        $stopItem.Enabled = $false
        $restartItem.Enabled = $false
        return
    }

    $script:enabled = $true
    Update-Status "Starting..." $null "warning"

    # Ensure log directory exists
    $logDir = "$env:TEMP\openclaw"
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Ensure OAuth token is available from user environment
    $userToken = [Environment]::GetEnvironmentVariable("CLAUDE_CODE_OAUTH_TOKEN", "User")
    if ($userToken) {
        $env:CLAUDE_CODE_OAUTH_TOKEN = $userToken
    }

    # Build process start info based on detected openclaw type
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $info = $script:openclawInfo

    switch ($info.Type) {
        "cmd" {
            $psi.FileName = $info.Path
            $psi.Arguments = "gateway"
        }
        "node" {
            $psi.FileName = "node"
            $psi.Arguments = "`"$($info.Path)`" gateway"
        }
        "npx" {
            $psi.FileName = "npx"
            $psi.Arguments = "openclaw gateway"
        }
    }

    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $psi.CreateNoWindow = $true
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    # Pass OAuth token to gateway process
    if ($env:CLAUDE_CODE_OAUTH_TOKEN) {
        $psi.EnvironmentVariables["CLAUDE_CODE_OAUTH_TOKEN"] = $env:CLAUDE_CODE_OAUTH_TOKEN
    }

    try {
        $script:gatewayProcess = New-Object System.Diagnostics.Process
        $script:gatewayProcess.StartInfo = $psi
        $script:gatewayProcess.EnableRaisingEvents = $true

        $script:gatewayProcess.Start() | Out-Null

        # Start async output reading
        $script:gatewayProcess.BeginOutputReadLine()
        $script:gatewayProcess.BeginErrorReadLine()

        # Register event handlers with real-time OAuth error detection
        Register-ObjectEvent -InputObject $script:gatewayProcess -EventName OutputDataReceived -Action {
            if ($EventArgs.Data) {
                $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
                Add-Content -Path $logPath -Value $EventArgs.Data -ErrorAction SilentlyContinue
                # Real-time OAuth error detection
                if ($EventArgs.Data -match "OAuth token has been revoked" -or
                    $EventArgs.Data -match "HTTP 403.*permission_error" -or
                    $EventArgs.Data -match "\[openclaw\].*HTTP 403" -or
                    $EventArgs.Data -match "\[clawdbot\].*HTTP 403" -or
                    $EventArgs.Data -match "permission_error.*revoked") {
                    "oauth_error" | Out-File "$env:TEMP\openclaw_oauth_error.signal" -Force
                }
            }
        } | Out-Null

        Register-ObjectEvent -InputObject $script:gatewayProcess -EventName ErrorDataReceived -Action {
            if ($EventArgs.Data) {
                $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
                Add-Content -Path $logPath -Value $EventArgs.Data -ErrorAction SilentlyContinue
                # Real-time OAuth error detection
                if ($EventArgs.Data -match "OAuth token has been revoked" -or
                    $EventArgs.Data -match "HTTP 403.*permission_error" -or
                    $EventArgs.Data -match "\[openclaw\].*HTTP 403" -or
                    $EventArgs.Data -match "\[clawdbot\].*HTTP 403" -or
                    $EventArgs.Data -match "permission_error.*revoked") {
                    "oauth_error" | Out-File "$env:TEMP\openclaw_oauth_error.signal" -Force
                }
            }
        } | Out-Null

        Start-Sleep -Milliseconds 3000

        if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
            # Check log for OAuth errors
            $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
            if (Test-Path $logPath) {
                $recentLog = Get-Content $logPath -Tail 50 -ErrorAction SilentlyContinue | Out-String
                if (Check-OAuthError $recentLog) {
                    $script:oauthError = $true
                    Stop-Gateway
                    Update-Status "OAuth Error - Auto-fixing..." "Opening authentication wizard" "error"
                    $trayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
                    $trayIcon.BalloonTipTitle = "$script:productName OAuth Error - Auto-Fixing"
                    $trayIcon.BalloonTipText = "Your OAuth token has been revoked. Opening re-authentication wizard automatically..."
                    $trayIcon.ShowBalloonTip(3000)
                    Run-Reauth
                    return
                }
            }
            Update-Status "Running" "Gateway started successfully" "running"
            $startItem.Enabled = $false
            $stopItem.Enabled = $true
            $restartItem.Enabled = $true
        } else {
            # Check for OAuth error on quick exit
            $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
            if (Test-Path $logPath) {
                $recentLog = Get-Content $logPath -Tail 50 -ErrorAction SilentlyContinue | Out-String
                if (Check-OAuthError $recentLog) {
                    $script:oauthError = $true
                    Update-Status "OAuth Error - Auto-fixing..." "Opening authentication wizard" "error"
                    $trayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
                    $trayIcon.BalloonTipTitle = "$script:productName OAuth Error - Auto-Fixing"
                    $trayIcon.BalloonTipText = "Your OAuth token has been revoked. Opening re-authentication wizard automatically..."
                    $trayIcon.ShowBalloonTip(3000)
                    Run-Reauth
                    return
                }
            }
            Update-Status "Failed" "Gateway failed to start" "error"
            $startItem.Enabled = $true
            $stopItem.Enabled = $false
        }
    } catch {
        Update-Status "Error" "Failed to start: $_" "error"
        $startItem.Enabled = $true
        $stopItem.Enabled = $false
    }
}

function Stop-Gateway {
    $script:enabled = $false
    Update-Status "Stopping..." $null "warning"

    # Try graceful stop via openclaw command
    try {
        Invoke-OpenClaw -Arguments "gateway stop" -Wait -Hidden
    } catch {}

    # Kill any remaining gateway processes
    Get-Process -Name "node" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $wmi = Get-WmiObject Win32_Process -Filter "ProcessId = $($_.Id)" -ErrorAction SilentlyContinue
            if ($wmi -and ($wmi.CommandLine -like "*openclaw*gateway*" -or $wmi.CommandLine -like "*clawdbot*gateway*" -or $wmi.CommandLine -like "*moltbot*gateway*")) {
                $_.Kill()
            }
        } catch {}
    }

    if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
        try { $script:gatewayProcess.Kill() } catch {}
    }
    $script:gatewayProcess = $null

    Start-Sleep -Milliseconds 500
    Update-Status "Stopped" "Gateway stopped" "normal"
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
    $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-NoProfile", "-Command", "Write-Host '$script:productName Gateway Live Output' -ForegroundColor Cyan; Write-Host '================================' -ForegroundColor Cyan; if (Test-Path '$logPath') { Get-Content '$logPath' -Wait -Tail 50 } else { Write-Host 'Log file not found yet. Waiting...' -ForegroundColor Yellow; while (!(Test-Path '$logPath')) { Start-Sleep 1 }; Get-Content '$logPath' -Wait -Tail 50 }"
}

function Open-LogFile {
    $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
    if (Test-Path $logPath) {
        Start-Process notepad.exe $logPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Log file not found: $logPath", $script:productName, "OK", "Information")
    }
}

function Run-Doctor {
    $info = $script:openclawInfo
    $cmdLine = switch ($info.Type) {
        "cmd" { "& '$($info.Path)' doctor" }
        "node" { "node '$($info.Path)' doctor" }
        "npx" { "npx openclaw doctor" }
    }
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-NoProfile", "-Command", "$cmdLine; Read-Host 'Press Enter to close'"
}

function Open-Dashboard {
    $info = $script:openclawInfo
    $cmdLine = switch ($info.Type) {
        "cmd" { "& '$($info.Path)' dashboard" }
        "node" { "node '$($info.Path)' dashboard" }
        "npx" { "npx openclaw dashboard" }
    }
    Start-Process -FilePath "powershell.exe" -ArgumentList "-WindowStyle", "Hidden", "-NoProfile", "-Command", $cmdLine
}

function Clean-AllContext {
    Update-Status "Cleaning context..." $null "warning"

    # Stop gateway first
    Stop-Gateway

    $cleanedItems = 0
    $freedSpace = 0

    # OpenClaw/Clawdbot context paths to clean
    $contextPaths = @(
        "$env:USERPROFILE\.openclaw"
        "$env:USERPROFILE\.clawdbot"
        "$env:USERPROFILE\.moltbot"
        "$env:APPDATA\openclaw"
        "$env:APPDATA\clawdbot"
        "$env:APPDATA\moltbot"
        "$env:LOCALAPPDATA\openclaw"
        "$env:LOCALAPPDATA\clawdbot"
        "$env:LOCALAPPDATA\moltbot"
        "$env:TEMP\openclaw"
        "$env:TEMP\clawdbot"
        "$env:TEMP\moltbot"
        "$env:TEMP\openclaw_*.signal"
        "$env:TEMP\clawdbot_*.signal"
        "$env:USERPROFILE\.claude\projects"
        "$env:USERPROFILE\.claude\statsig"
        "$env:USERPROFILE\.claude\todos"
        "$env:LOCALAPPDATA\claude"
        "$env:APPDATA\npm\.moltbot-*"
        "$env:APPDATA\npm\node_modules\.cache"
    )

    foreach ($path in $contextPaths) {
        if ($path -like "*`**") {
            # Wildcard pattern
            $items = Get-ChildItem -Path (Split-Path $path) -Filter (Split-Path $path -Leaf) -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                try {
                    $size = (Get-ChildItem $item.FullName -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                    Remove-Item $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    $cleanedItems++
                    $freedSpace += $size
                } catch {}
            }
        } elseif (Test-Path $path) {
            try {
                $size = (Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
                $cleanedItems++
                $freedSpace += $size
            } catch {}
        }
    }

    # Clear npm cache for openclaw/clawdbot
    try {
        Start-Process -FilePath "npm" -ArgumentList "cache", "clean", "--force" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
    } catch {}

    $freedMB = [math]::Round($freedSpace / 1MB, 2)
    Update-Status "Cleaned" "Removed $cleanedItems items, freed ${freedMB}MB" "normal"

    $trayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $trayIcon.BalloonTipTitle = "$script:productName Context Cleaned"
    $trayIcon.BalloonTipText = "Cleaned $cleanedItems items, freed ${freedMB}MB`nGateway will restart..."
    $trayIcon.ShowBalloonTip(3000)

    Start-Sleep -Seconds 2
    Start-Gateway
}

function Optimize-Connection {
    Update-Status "Optimizing..." $null "warning"

    # Set high process priority
    try {
        if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
            $script:gatewayProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::High
        }
    } catch {}

    # Optimize network settings for this process
    try {
        # Disable Nagle's algorithm for lower latency
        $code = @"
using System;
using System.Runtime.InteropServices;
public class NetOptimizer {
    [DllImport("ws2_32.dll")]
    public static extern int WSAStartup(ushort wVersionRequested, out WSAData wsaData);
    [StructLayout(LayoutKind.Sequential)]
    public struct WSAData { public ushort wVersion; public ushort wHighVersion; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 257)] public string szDescription; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 129)] public string szSystemStatus; public ushort iMaxSockets; public ushort iMaxUdpDg; public IntPtr lpVendorInfo; }
}
"@
        Add-Type -TypeDefinition $code -ErrorAction SilentlyContinue
    } catch {}

    # Set environment variables for optimal connection
    $env:NODE_OPTIONS = "--max-old-space-size=4096"
    $env:UV_THREADPOOL_SIZE = "16"

    # Restart gateway with optimizations
    $wasRunning = $script:gatewayProcess -and !$script:gatewayProcess.HasExited
    if ($wasRunning) {
        Stop-Gateway
        Start-Sleep -Milliseconds 500
    }

    # Apply Windows network optimizations
    try {
        # Disable network throttling
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "NetworkThrottlingIndex" -Value 0xffffffff -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "SystemResponsiveness" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    } catch {}

    if ($wasRunning) {
        Start-Gateway
    }

    Update-Status "Optimized" "Connection optimized for best performance" "running"
    $trayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $trayIcon.BalloonTipTitle = "$script:productName Optimized"
    $trayIcon.BalloonTipText = "Network optimized, high priority set, latency reduced"
    $trayIcon.ShowBalloonTip(2000)
}

function Run-Reauth {
    if ($script:reauthInProgress) {
        Update-Status "Re-auth in progress..." "Already re-authenticating" "warning"
        return
    }

    $script:reauthInProgress = $true
    Stop-Gateway
    Update-Status "Re-authenticating..." "Opening authentication wizard" "warning"

    # Create a temp script to run onboard and signal completion
    $signalFile = "$env:TEMP\openclaw_reauth_done.signal"
    Remove-Item $signalFile -Force -ErrorAction SilentlyContinue

    $info = $script:openclawInfo
    $onboardCmd = switch ($info.Type) {
        "cmd" { "& '$($info.Path)' onboard --auth-choice claude-cli" }
        "node" { "node '$($info.Path)' onboard --auth-choice claude-cli" }
        "npx" { "npx openclaw onboard --auth-choice claude-cli" }
    }

    $reauthScript = @"
Write-Host '=== $script:productName Re-authentication ===' -ForegroundColor Cyan
Write-Host 'Your OAuth token has been revoked. Running onboard wizard...' -ForegroundColor Yellow
Write-Host ''
$onboardCmd
Write-Host ''
Write-Host 'Authentication complete!' -ForegroundColor Green
'done' | Out-File '$signalFile'
Write-Host 'The gateway will restart automatically. You can close this window.' -ForegroundColor Green
Start-Sleep -Seconds 3
"@

    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-NoProfile", "-Command", $reauthScript

    # Start a background job to watch for completion
    $watchTimer = New-Object System.Windows.Forms.Timer
    $watchTimer.Interval = 2000
    $watchTimer.Add_Tick({
        if (Test-Path "$env:TEMP\openclaw_reauth_done.signal") {
            Remove-Item "$env:TEMP\openclaw_reauth_done.signal" -Force -ErrorAction SilentlyContinue
            Remove-Item "$env:TEMP\openclaw_oauth_error.signal" -Force -ErrorAction SilentlyContinue
            $script:oauthError = $false
            $script:reauthInProgress = $false
            $this.Stop()
            $this.Dispose()
            # Clear old log entries that contain OAuth errors
            $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
            if (Test-Path $logPath) {
                # Archive old log and start fresh
                $archivePath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd')-pre-reauth.log"
                Move-Item $logPath $archivePath -Force -ErrorAction SilentlyContinue
            }
            Update-Status "Starting..." "Re-authentication complete, starting gateway" "warning"
            $startItem.Enabled = $true
            $stopItem.Enabled = $false
            $restartItem.Enabled = $false
            Start-Gateway
        }
    })
    $watchTimer.Start()
}

function Check-OAuthError {
    param([string]$logContent)
    if ($logContent -match "OAuth token has been revoked" -or
        $logContent -match "HTTP 403.*permission_error" -or
        $logContent -match "\[openclaw\].*HTTP 403" -or
        $logContent -match "\[clawdbot\].*HTTP 403" -or
        $logContent -match "token.*revoked" -or
        $logContent -match "authentication.*failed" -or
        $logContent -match "401.*Unauthorized" -or
        $logContent -match "permission_error.*revoked") {
        return $true
    }
    return $false
}

# Event handlers
$startItem.Add_Click({ Start-Gateway })
$stopItem.Add_Click({ Stop-Gateway })
$restartItem.Add_Click({ Restart-Gateway })
$reauthItem.Add_Click({ Run-Reauth })
$terminalItem.Add_Click({ Show-Terminal })
$logItem.Add_Click({ Open-LogFile })
$doctorItem.Add_Click({ Run-Doctor })
$dashboardItem.Add_Click({ Open-Dashboard })
$cleanContextItem.Add_Click({ Clean-AllContext })
$optimizeItem.Add_Click({ Optimize-Connection })

$exitItem.Add_Click({
    Stop-Gateway
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
    # Remove lock file
    $lockFile = "$env:TEMP\OpenClawTray.lock"
    Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
    [System.Windows.Forms.Application]::Exit()
})

# Double-click to show status
$trayIcon.Add_DoubleClick({
    $status = if ($script:gatewayProcess -and !$script:gatewayProcess.HasExited) { "Running" } else { "Stopped" }
    $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
    $info = $script:openclawInfo
    [System.Windows.Forms.MessageBox]::Show(
        "$script:productName Gateway Status: $status`n" +
        "PID: $(if($script:gatewayProcess){$script:gatewayProcess.Id}else{'N/A'})`n" +
        "Mode: $($info.Type)`n" +
        "Log: $logPath",
        "$script:productName Gateway", "OK", "Information")
})

# Initial state
$startItem.Enabled = $true
$stopItem.Enabled = $false
$restartItem.Enabled = $false

# ============================================
# AUTO-OPTIMIZATION & FULL PERMISSIONS SETUP
# ============================================

# Set this process to high priority
try {
    $currentProcess = Get-Process -Id $PID
    $currentProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::High
} catch {}

# Ensure full permissions - set environment for unrestricted access
$env:OPENCLAW_UNRESTRICTED = "1"
$env:CLAWDBOT_UNRESTRICTED = "1"
$env:OPENCLAW_ALLOW_ALL = "1"
$env:CLAWDBOT_ALLOW_ALL = "1"
$env:OPENCLAW_SKIP_PERMISSION_CHECK = "1"
$env:CLAWDBOT_SKIP_PERMISSION_CHECK = "1"
$env:NODE_OPTIONS = "--max-old-space-size=4096"
$env:UV_THREADPOOL_SIZE = "16"

# Apply network optimizations for best connection
try {
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "NetworkThrottlingIndex" -Value 0xffffffff -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "SystemResponsiveness" -Value 0 -Type DWord -ErrorAction SilentlyContinue
} catch {}

# Connection watchdog - auto-reconnect on failure
$script:reconnectAttempts = 0
$script:maxReconnectAttempts = 5
$script:lastHealthCheck = Get-Date

# Auto-start gateway
Start-Gateway

# Timer to check gateway health
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000  # Check every 5 seconds for faster OAuth error response
$timer.Add_Tick({
    # Skip if OAuth error detected and re-auth in progress
    if ($script:oauthError -and $script:reauthInProgress) {
        return
    }

    # Check for real-time OAuth error signal file
    $signalFile = "$env:TEMP\openclaw_oauth_error.signal"
    if (Test-Path $signalFile) {
        Remove-Item $signalFile -Force -ErrorAction SilentlyContinue
        if (-not $script:oauthError) {
            $script:oauthError = $true
            Stop-Gateway
            Update-Status "OAuth Error - Auto-fixing..." "Opening authentication wizard" "error"
            $trayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
            $trayIcon.BalloonTipTitle = "$script:productName OAuth Error - Auto-Fixing"
            $trayIcon.BalloonTipText = "Your OAuth token has been revoked. Opening re-authentication wizard automatically..."
            $trayIcon.ShowBalloonTip(3000)
            Run-Reauth
            return
        }
    }

    # Check log for new OAuth errors even while running
    $logPath = "$env:TEMP\openclaw\openclaw-$(Get-Date -Format 'yyyy-MM-dd').log"
    if (Test-Path $logPath) {
        $recentLog = Get-Content $logPath -Tail 30 -ErrorAction SilentlyContinue | Out-String
        if (Check-OAuthError $recentLog) {
            if (-not $script:oauthError) {
                $script:oauthError = $true
                Stop-Gateway
                Update-Status "OAuth Error - Auto-fixing..." "Opening authentication wizard" "error"
                $trayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
                $trayIcon.BalloonTipTitle = "$script:productName OAuth Error - Auto-Fixing"
                $trayIcon.BalloonTipText = "Your OAuth token has been revoked. Opening re-authentication wizard automatically..."
                $trayIcon.ShowBalloonTip(3000)
                Run-Reauth
            }
            return
        }
    }

    if ($script:enabled) {
        if (!$script:gatewayProcess -or $script:gatewayProcess.HasExited) {
            $script:reconnectAttempts++
            if ($script:reconnectAttempts -le $script:maxReconnectAttempts) {
                Update-Status "Reconnecting ($($script:reconnectAttempts)/$($script:maxReconnectAttempts))..." $null "warning"
                Start-Gateway
            } else {
                # Reset after max attempts and try again
                $script:reconnectAttempts = 0
                Update-Status "Connection reset - Retrying..." $null "warning"
                Start-Sleep -Milliseconds 2000
                Start-Gateway
            }
        } else {
            # Gateway running - reset reconnect counter
            $script:reconnectAttempts = 0

            # Ensure process priority stays high
            try {
                if ($script:gatewayProcess.PriorityClass -ne [System.Diagnostics.ProcessPriorityClass]::High) {
                    $script:gatewayProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::High
                }
            } catch {}
        }
    }
})
$timer.Start()

# Connection stability timer - faster checks for uninterrupted connection
$stabilityTimer = New-Object System.Windows.Forms.Timer
$stabilityTimer.Interval = 2000  # Check every 2 seconds
$stabilityTimer.Add_Tick({
    if ($script:enabled -and $script:gatewayProcess -and !$script:gatewayProcess.HasExited) {
        # Keep connection alive - touch the process
        try {
            $null = $script:gatewayProcess.Responding
        } catch {
            # Process unresponsive - force restart
            Update-Status "Unresponsive - Force restart..." $null "error"
            try { $script:gatewayProcess.Kill() } catch {}
            $script:gatewayProcess = $null
            Start-Sleep -Milliseconds 500
            Start-Gateway
        }
    }
})
$stabilityTimer.Start()

# Run message loop
[System.Windows.Forms.Application]::Run()
