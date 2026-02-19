using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ClawdBotManager;

static class Program
{
    static NotifyIcon? _trayIcon;
    static Process? _gatewayProcess;
    static ToolStripMenuItem? _statusItem;
    static System.Windows.Forms.Timer? _healthCheckTimer;
    static bool _userStopped;
    static int _consecutiveFailures = 0;
    static DateTime _lastStartAttempt = DateTime.MinValue;
    static int _healthCheckFailures = 0;
    static string _logDir = Path.Combine(Path.GetTempPath(), "openclaw");
    static Mutex? _mutex;
    static readonly string AppName = "ClawdBotManager";
    static readonly string ExePath = Environment.ProcessPath ?? AppContext.BaseDirectory;
    static readonly object _startLock = new object();
    static volatile bool _isRestarting = false;
    static SynchronizationContext? _uiContext;

    [STAThread]
    static void Main()
    {
        bool createdNew;
        _mutex = new Mutex(true, "ClawdBotManager_SingleInstance_2026", out createdNew);
        if (!createdNew) return;

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Directory.CreateDirectory(_logDir);

        _uiContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
        SynchronizationContext.SetSynchronizationContext(_uiContext);

        // Clean up old competitors
        EliminateOpenClawRunner();
        KillOrphanGatewayProcesses();

        // Set env vars for node
        SetupEnvironment();

        // Registry startup only - NO scheduled tasks, NO firewall (those cause popups)
        EnsureRegistryStartup();

        // Show tray FIRST (instant, no delay for user)
        SetupTray();

        // Start gateway in background thread (no UI blocking)
        var startThread = new Thread(() =>
        {
            WaitForNetwork();
            StartGateway();
        });
        startThread.IsBackground = true;
        startThread.Name = "GatewayStartThread";
        startThread.Start();

        Application.Run();

        _mutex?.ReleaseMutex();
        _mutex?.Dispose();
    }

    static void EliminateOpenClawRunner()
    {
        try
        {
            foreach (var proc in Process.GetProcessesByName("OpenClawRunner"))
            {
                try { proc.Kill(entireProcessTree: true); } catch { }
            }

            var registryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
            using var key = Registry.CurrentUser.OpenSubKey(registryKey, true);
            if (key?.GetValue("OpenClawRunner") != null)
                key.DeleteValue("OpenClawRunner", false);
        }
        catch { }
    }

    static void WaitForNetwork()
    {
        LogLine("[NETWORK] Waiting for internet...");
        for (int i = 0; i < 30; i++)
        {
            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };
                var result = client.GetAsync("https://api.telegram.org").Result;
                LogLine($"[NETWORK] Telegram API reachable after {i * 2}s");
                return;
            }
            catch { Thread.Sleep(2000); }
        }
        LogLine("[NETWORK] Could not reach Telegram after 60s, starting anyway");
    }

    static void KillOrphanGatewayProcesses()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c \"for /f \"tokens=5\" %a in ('netstat -aon ^| findstr :18789 ^| findstr LISTEN') do taskkill /F /PID %a 2>nul\"",
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using var proc = Process.Start(psi);
            proc?.WaitForExit(5000);
        }
        catch { }
    }

    static void SetupEnvironment()
    {
        Environment.SetEnvironmentVariable("SHELL", Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe");
        Environment.SetEnvironmentVariable("OPENCLAW_SHELL", "cmd");
        Environment.SetEnvironmentVariable("OPENCLAW_NO_WSL", "1");
        Environment.SetEnvironmentVariable("OPENCLAW_NO_PTY", "1");

        // Node.js stability flags
        Environment.SetEnvironmentVariable("NODE_OPTIONS",
            "--unhandled-rejections=warn " +          // Don't crash on Baileys WhatsApp errors
            "--max-old-space-size=8192 " +            // 8GB memory limit
            "--dns-result-order=ipv4first " +         // IPv4 first (fixes Telegram DNS issues)
            "--max-http-header-size=16384");           // Larger headers

        Environment.SetEnvironmentVariable("NODE_TLS_REJECT_UNAUTHORIZED", "0");
        Environment.SetEnvironmentVariable("UV_THREADPOOL_SIZE", "64");

        // Session keepalive
        Environment.SetEnvironmentVariable("OPENCLAW_AGENT_TIMEOUT", "300000");
        Environment.SetEnvironmentVariable("OPENCLAW_RUN_TIMEOUT", "300000");
        Environment.SetEnvironmentVariable("OPENCLAW_LLM_TIMEOUT", "120000");
        Environment.SetEnvironmentVariable("OPENCLAW_STREAM_TIMEOUT", "180000");
        Environment.SetEnvironmentVariable("OPENCLAW_SESSION_KEEPALIVE", "true");
        Environment.SetEnvironmentVariable("OPENCLAW_AUTO_RECONNECT", "true");
        Environment.SetEnvironmentVariable("OPENCLAW_RECONNECT_INTERVAL", "3000");
        Environment.SetEnvironmentVariable("OPENCLAW_MAX_RECONNECT_ATTEMPTS", "999");

        var oauthToken = Environment.GetEnvironmentVariable("CLAUDE_CODE_OAUTH_TOKEN", EnvironmentVariableTarget.User);
        if (!string.IsNullOrEmpty(oauthToken))
            Environment.SetEnvironmentVariable("CLAUDE_CODE_OAUTH_TOKEN", oauthToken);
    }

    static void EnsureRegistryStartup()
    {
        try
        {
            var registryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
            using var key = Registry.CurrentUser.OpenSubKey(registryKey, true);
            if (key != null)
            {
                var desired = $"\"{ExePath}\"";
                if ((key.GetValue(AppName) as string) != desired)
                {
                    key.SetValue(AppName, desired);
                    LogLine($"[STARTUP] Registry Run entry set");
                }
            }

            // Clean up old scheduled tasks that cause popups
            CleanupOldScheduledTasks();
        }
        catch (Exception ex)
        {
            LogLine($"[STARTUP] Registry error: {ex.Message}");
        }
    }

    static void CleanupOldScheduledTasks()
    {
        // Delete ALL old ClawdBotManager scheduled tasks - they cause popup windows
        string[] oldTasks = {
            "ClawdBotManager", "ClawdBotManager_AutoStart",
            "ClawdBotManager_BootStart", "ClawdBotManager_Logon"
        };
        foreach (var taskName in oldTasks)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "schtasks.exe",
                    Arguments = $"/Delete /TN \"{taskName}\" /F",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using var proc = Process.Start(psi);
                proc?.WaitForExit(3000);
            }
            catch { }
        }
    }

    static bool IsPortListening(int port)
    {
        try
        {
            using var socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            socket.ReceiveTimeout = 1000;
            socket.SendTimeout = 1000;
            socket.Connect(IPAddress.Loopback, port);
            socket.Close();
            return true;
        }
        catch { return false; }
    }

    static (string fileName, string args) GetOpenClawCommand()
    {
        // ALWAYS use node.exe directly - NEVER .cmd files (they flash a terminal window)
        var nodePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "nodejs", "node.exe");
        var npmPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm");
        var gatewayArgs = "gateway --allow-unconfigured --auth token --token moltbot-local-token-2026";

        var mjsPath = Path.Combine(npmPath, "node_modules", "openclaw", "openclaw.mjs");
        if (File.Exists(nodePath) && File.Exists(mjsPath))
        {
            LogLine($"[GATEWAY] Using node.exe + openclaw.mjs");
            return (nodePath, $"\"{mjsPath}\" {gatewayArgs}");
        }

        var entryJs = Path.Combine(npmPath, "node_modules", "openclaw", "dist", "entry.js");
        if (File.Exists(nodePath) && File.Exists(entryJs))
        {
            LogLine($"[GATEWAY] Using node.exe + entry.js");
            return (nodePath, $"\"{entryJs}\" {gatewayArgs}");
        }

        LogLine("[GATEWAY] WARNING: Falling back to cmd.exe /c openclaw.cmd");
        var cmdPath = Path.Combine(npmPath, "openclaw.cmd");
        return ("cmd.exe", $"/c \"{cmdPath}\" {gatewayArgs}");
    }

    static void StartGateway()
    {
        lock (_startLock)
        {
            if (_gatewayProcess != null && !_gatewayProcess.HasExited) return;

            _userStopped = false;
            _isRestarting = false;

            if (IsPortListening(18789))
            {
                KillOrphanGatewayProcesses();
                Thread.Sleep(2000);
            }

            var (fileName, args) = GetOpenClawCommand();

            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = args,
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            try
            {
                _gatewayProcess = new Process { StartInfo = psi, EnableRaisingEvents = true };
                _gatewayProcess.OutputDataReceived += (s, e) => LogLine(e.Data);
                _gatewayProcess.ErrorDataReceived += (s, e) => LogLine(e.Data);
                _gatewayProcess.Exited += OnGatewayExited;
                _gatewayProcess.Start();
                _gatewayProcess.BeginOutputReadLine();
                _gatewayProcess.BeginErrorReadLine();

                _consecutiveFailures = 0;
                _lastStartAttempt = DateTime.Now;
                InvokeOnUI(() => UpdateStatus("Running", Color.FromArgb(0, 220, 255)));
                LogLine("[GATEWAY] Started successfully");
            }
            catch (Exception ex)
            {
                LogLine($"[GATEWAY] Failed to start: {ex.Message}");
                InvokeOnUI(() => UpdateStatus("Start Failed", Color.FromArgb(255, 60, 60)));
                _consecutiveFailures++;
                ScheduleRestart(GetBackoffDelay());
            }
        }
    }

    static void InvokeOnUI(Action action)
    {
        if (_uiContext != null)
            _uiContext.Post(_ => { try { action(); } catch { } }, null);
        else
            try { action(); } catch { }
    }

    static void OnGatewayExited(object? sender, EventArgs e)
    {
        if (_userStopped || _isRestarting) return;

        _consecutiveFailures++;
        var uptime = DateTime.Now - _lastStartAttempt;
        LogLine($"[GATEWAY] Exited after {uptime.TotalSeconds:F1}s (failure #{_consecutiveFailures})");

        if (uptime.TotalSeconds > 300) _consecutiveFailures = 1;

        var delay = GetBackoffDelay();
        InvokeOnUI(() => UpdateStatus($"Restarting in {delay / 1000}s...", Color.FromArgb(255, 200, 50)));
        ScheduleRestart(delay);
    }

    static int GetBackoffDelay()
    {
        return _consecutiveFailures switch
        {
            1 => 5000,
            2 => 10000,
            3 => 20000,
            4 => 30000,
            _ => 60000
        };
    }

    static void ScheduleRestart(int delayMs)
    {
        if (_isRestarting) return;
        _isRestarting = true;

        var thread = new Thread(() =>
        {
            try
            {
                Thread.Sleep(delayMs);
                if (!_userStopped) StartGateway();
                else _isRestarting = false;
            }
            catch
            {
                _isRestarting = false;
                Thread.Sleep(10000);
                if (!_userStopped) StartGateway();
            }
        });
        thread.IsBackground = true;
        thread.Name = "AutoRestartThread";
        thread.Start();
    }

    static void StopGateway()
    {
        _userStopped = true;
        if (_gatewayProcess != null && !_gatewayProcess.HasExited)
        {
            try { _gatewayProcess.Kill(entireProcessTree: true); } catch { }
        }
        _gatewayProcess = null;
        _consecutiveFailures = 0;
        UpdateStatus("Stopped", Color.FromArgb(255, 100, 100));
    }

    static void LogLine(string? line)
    {
        if (string.IsNullOrEmpty(line)) return;
        try
        {
            var logFile = Path.Combine(_logDir, $"openclaw-{DateTime.Now:yyyy-MM-dd}.log");
            File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {line}{Environment.NewLine}");
        }
        catch { }
    }

    static Icon CreateIcon(Color accentColor)
    {
        var bmp = new Bitmap(16, 16);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using var bgBrush = new SolidBrush(Color.FromArgb(255, 15, 15, 30));
            g.FillEllipse(bgBrush, 0, 0, 15, 15);
            using var ringPen = new Pen(Color.FromArgb(180, accentColor.R, accentColor.G, accentColor.B), 1.2f);
            g.DrawEllipse(ringPen, 1, 1, 13, 13);
            using var clawPen = new Pen(accentColor, 1.8f);
            clawPen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
            clawPen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
            g.DrawLine(clawPen, 4.5f, 3.5f, 5.5f, 12.5f);
            g.DrawLine(clawPen, 8f, 3f, 8f, 13f);
            g.DrawLine(clawPen, 11.5f, 3.5f, 10.5f, 12.5f);
        }
        return Icon.FromHandle(bmp.GetHicon());
    }

    static void UpdateStatus(string text, Color color)
    {
        if (_trayIcon == null) return;
        try
        {
            _trayIcon.Icon?.Dispose();
            _trayIcon.Icon = CreateIcon(color);
            _trayIcon.Text = $"ClawdBot - {text}";
            if (_statusItem != null) _statusItem.Text = $"Status: {text}";
        }
        catch { }
    }

    static void SetupTray()
    {
        _trayIcon = new NotifyIcon
        {
            Icon = CreateIcon(Color.FromArgb(0, 220, 255)),
            Text = "ClawdBot - Starting...",
            Visible = true
        };

        var menu = new ContextMenuStrip();
        _statusItem = new ToolStripMenuItem("Status: Starting...") { Enabled = false };
        menu.Items.Add(_statusItem);
        menu.Items.Add(new ToolStripSeparator());

        var startItem = new ToolStripMenuItem("Start Gateway");
        startItem.Click += (s, e) => { _userStopped = false; new Thread(() => StartGateway()) { IsBackground = true }.Start(); };
        menu.Items.Add(startItem);

        var stopItem = new ToolStripMenuItem("Stop Gateway");
        stopItem.Click += (s, e) => StopGateway();
        menu.Items.Add(stopItem);

        var restartItem = new ToolStripMenuItem("Restart Gateway");
        restartItem.Click += (s, e) =>
        {
            StopGateway();
            Thread.Sleep(1000);
            _userStopped = false;
            new Thread(() => StartGateway()) { IsBackground = true }.Start();
        };
        menu.Items.Add(restartItem);

        menu.Items.Add(new ToolStripSeparator());

        var logItem = new ToolStripMenuItem("Open Log");
        logItem.Click += (s, e) =>
        {
            var logFile = Path.Combine(_logDir, $"openclaw-{DateTime.Now:yyyy-MM-dd}.log");
            if (File.Exists(logFile)) Process.Start("notepad.exe", logFile);
        };
        menu.Items.Add(logItem);

        menu.Items.Add(new ToolStripSeparator());

        var exitItem = new ToolStripMenuItem("Exit");
        exitItem.Click += (s, e) =>
        {
            StopGateway();
            _healthCheckTimer?.Stop();
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            Application.Exit();
        };
        menu.Items.Add(exitItem);

        _trayIcon.ContextMenuStrip = menu;

        // Health check every 2 minutes - only restart if process is truly dead
        _healthCheckTimer = new System.Windows.Forms.Timer { Interval = 120000 };
        _healthCheckTimer.Tick += (s, e) => PerformHealthCheck();
        _healthCheckTimer.Start();
    }

    static void PerformHealthCheck()
    {
        try
        {
            if (_userStopped || _isRestarting) return;

            // Only check if process has crashed (not responding to port is NOT a reason to kill it)
            if (_gatewayProcess == null || _gatewayProcess.HasExited)
            {
                _healthCheckFailures++;
                LogLine($"[HEALTH] Gateway process dead (check #{_healthCheckFailures})");

                if (_healthCheckFailures >= 2 && !_isRestarting)
                {
                    LogLine("[HEALTH] Restarting dead gateway");
                    _healthCheckFailures = 0;
                    _consecutiveFailures = 0;
                    ScheduleRestart(3000);
                }
                return;
            }

            // Process is alive - just log status
            _healthCheckFailures = 0;
            var uptime = DateTime.Now - _lastStartAttempt;
            bool portUp = IsPortListening(18789);

            if ((int)uptime.TotalMinutes % 10 == 0)
            {
                try
                {
                    _gatewayProcess.Refresh();
                    var memMB = _gatewayProcess.WorkingSet64 / 1024 / 1024;
                    LogLine($"[HEALTH] Uptime {uptime.TotalMinutes:F0}m, {memMB}MB RAM, port={portUp}");
                }
                catch { }
            }

            if (portUp)
                InvokeOnUI(() => UpdateStatus("Running", Color.FromArgb(0, 220, 255)));
        }
        catch { }
    }
}
