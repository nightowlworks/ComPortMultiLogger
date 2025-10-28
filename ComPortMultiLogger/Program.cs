using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

// Aliases to avoid Timer ambiguity
using WinFormsTimer = System.Windows.Forms.Timer;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        AppLogger.Init(); // prepare app logger

        // Global fail-safe handlers
        Application.ThreadException += (s, e) =>
        {
            AppLogger.LogException("ThreadException", e.Exception);
            MessageBox.Show("Ein unerwarteter Fehler ist aufgetreten.\n" +
                            e.Exception.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            CrashDumper.TryWriteMiniDump("thread_exception");
        };
        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
        {
            var ex = e.ExceptionObject as Exception;
            AppLogger.LogException("UnhandledException", ex ?? new Exception("Unknown unhandled exception"));
            CrashDumper.TryWriteMiniDump("unhandled_exception");
        };

        using var mtx = new Mutex(true, "M81_DataTransfer_SingleInstance", out bool first);
        if (!first)
        {
            MessageBox.Show("M81 DataTransfer läuft bereits.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        AppLogger.Log("=== M81 DataTransfer starting ===");

        Application.Run(new MainForm());

        AppLogger.Log("=== M81 DataTransfer exiting ===");
    }
}

/// <summary>Global defaults and metadata.</summary>
public static class Defaults
{
    public const string BaseFolder = @"C:\BaSyTec\Drivers\OSI\";
    public const string FixedFileName = "do_not_delete.txt";
    public const int FixedBaud = 9600;
    public const string AppVersion = "v1.0.8";

    // Disk guard (bytes)
    public const long MinFreeBytes = 200L * 1024L * 1024L; // 200 MB

    // Reconnect idle threshold (seconds). App waits at least this long without frames before reconnect.
    public const int ReconnectIdleSeconds = 20;

    // Cooldown between reconnect requests per logger (seconds) to avoid reconnect storms.
    public const int ReconnectRequestCooldownSeconds = 5;

    // Fallback white PNG (64x64) as Base64 if Assets\logo.png is missing.
    public const string LogoBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAQAAABt9U0VAAAACXBIWXMAAAsSAAALEgHS3X78AAABc0lEQVR4nO2Z0U7DUBCFv2g+1Wk" +
        "lNwH2q3kQyG6p2h9F1p7QmXwS8rJwQyqg0Qor8H2gq1eR7Ywz4qkL4T8m2C4j1kzv9mEw+g2I0H0e4Q3rLQz4l1H8aB2t3p0i7k7xqj3kH" +
        "WqgQqk3y0mY9hUTqGx8lqvQh7GvG3H7HfJqkYQwQ1qg9eY6mZ6mC6m8ZbXy2cXr7u1w6Yz2k4n+7v0jY6C3nS6S5QvD6oG6b0KkKpKj8p" +
        "wq2lU0V2hXwq5YQ0r5j7r7Q0bQw3r8x5s7z0r/8rR8/4m6p9gM8m2b7wAjyM0sJ8o5Jq4c9wYQm8m2b4w0Wl0mQ5OIG5m3m8gG0k0k0l" +
        "0k0k0m8m8n8q8n8o8o8p8p8q8q8r8r8s8s8t8t8u8u8v8v8w8w8x8x8y8y8z8z8z8z8z8z8z8z8z8z8z8/9yoYwF6h7GkSx9QAAAAASUVORK5CYII=";
}

// =============================== APP LOGGER & DUMPS ===============================
public static class AppLogger
{
    private static readonly object _lock = new();
    private static string _logDir = Path.Combine(AppContext.BaseDirectory, "logs");
    private static string _logPath = Path.Combine(_logDir, "app.log");
    private const long MaxBytes = 1024 * 1024; // 1 MB
    private const int Backups = 5;

    public static void Init()
    {
        try { Directory.CreateDirectory(_logDir); } catch { }
    }

    public static void Log(string msg)
    {
        try
        {
            lock (_lock)
            {
                RotateIfNeeded();
                File.AppendAllText(_logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {msg}{Environment.NewLine}", new UTF8Encoding(false));
            }
        }
        catch { }
    }

    public static void Debug(string msg) => Log("[DEBUG] " + msg);

    public static void LogException(string where, Exception ex)
    {
        Log($"[{where}] {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
    }

    private static void RotateIfNeeded()
    {
        try
        {
            if (File.Exists(_logPath) && new FileInfo(_logPath).Length >= MaxBytes)
            {
                for (int i = Backups - 1; i >= 1; i--)
                {
                    string src = _logPath + "." + i;
                    string dst = _logPath + "." + (i + 1);
                    if (File.Exists(src)) File.Copy(src, dst, true);
                }
                File.Copy(_logPath, _logPath + ".1", true);
                File.WriteAllText(_logPath, string.Empty, new UTF8Encoding(false));
                Log("Log rotated");
            }
        }
        catch { }
    }
}

public static class CrashDumper
{
    // P/Invoke MiniDumpWriteDump
    [Flags]
    private enum MINIDUMP_TYPE : uint
    {
        MiniDumpNormal = 0x00000000,
        MiniDumpWithDataSegs = 0x00000001,
        MiniDumpWithFullMemory = 0x00000002,
        MiniDumpWithHandleData = 0x00000004,
        MiniDumpWithUnloadedModules = 0x00000020,
        MiniDumpWithThreadInfo = 0x00001000
    }

    [DllImport("Dbghelp.dll", SetLastError = true)]
    private static extern bool MiniDumpWriteDump(
        IntPtr hProcess, uint processId, SafeHandle hFile, MINIDUMP_TYPE dumpType,
        IntPtr exceptionParam, IntPtr userStreamParam, IntPtr callbackParam);

    public static void TryWriteMiniDump(string reason)
    {
        try
        {
            string dumpDir = Path.Combine(AppContext.BaseDirectory, "dumps");
            Directory.CreateDirectory(dumpDir);
            string path = Path.Combine(dumpDir, $"crash_{DateTime.Now:yyyyMMdd_HHmmss}_{reason}.dmp");
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            var p = Process.GetCurrentProcess();
            bool ok = MiniDumpWriteDump(p.Handle, (uint)p.Id, fs.SafeFileHandle,
                MINIDUMP_TYPE.MiniDumpWithThreadInfo | MINIDUMP_TYPE.MiniDumpWithUnloadedModules, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            AppLogger.Log(ok ? $"MiniDump written: {path}" : $"MiniDump failed: {Marshal.GetLastWin32Error()}");
        }
        catch (Exception ex)
        {
            AppLogger.Log($"MiniDump exception: {ex.Message}");
        }
    }
}

// =============================== MAIN FORM ===============================
public sealed class MainForm : Form
{
    // Header
    private readonly Label lblTitle = new() { AutoSize = true };
    private readonly Label lblSubtitle = new() { AutoSize = true };
    private readonly PictureBox picLogo = new() { SizeMode = PictureBoxSizeMode.Zoom, Width = 199, Height = 49 };

    // Selector controls
    private readonly ComboBox cmbPorts = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };
    private readonly TextBox txtFolder = new() { Width = 460, PlaceholderText = @"Ordner für do_not_delete.txt" };
    private readonly Button btnChooseFolder = new() { Text = "Ordner", Width = 90, Height = 32 };
    private readonly Button btnAdd = new() { Text = "Port hinzufügen", Width = 150, Height = 32 };
    private readonly Button btnAddSim = new() { Text = "Sim-Port", Width = 100, Height = 32 };
    private readonly Button btnRefreshPorts = new() { Text = "Ports aktualisieren", Width = 140, Height = 32 };

    // Loggers list
    private readonly ListView lv = new()
    {
        Dock = DockStyle.Fill,
        FullRowSelect = true,
        GridLines = true,
        View = View.Details,
        HideSelection = true
    };

    // Control buttons
    private readonly Button btnStart = new() { Text = "Start", Enabled = false, Width = 92, Height = 32 };
    private readonly Button btnStop = new() { Text = "Stop", Enabled = false, Width = 92, Height = 32, ForeColor = Color.Red };
    private readonly Button btnRemove = new() { Text = "Entfernen", Enabled = false, Width = 110, Height = 32 };
    private readonly Button btnOpenFolder = new() { Text = "Ordner öffnen", Enabled = false, Width = 130, Height = 32 };
    private readonly Button btnClearLive = new() { Text = "Live leeren", Enabled = false, Width = 110, Height = 32 };

    // Split layout
    private readonly SplitContainer split = new()
    {
        Dock = DockStyle.Fill,
        Orientation = Orientation.Horizontal,
        SplitterWidth = 6
    };

    // Live table – one row per COM (latest value only)
    private readonly DataGridView dgvLive = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = false,
        AllowUserToResizeRows = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        MultiSelect = false,
        RowHeadersVisible = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        EnableHeadersVisualStyles = false
    };

    // Footer
    private readonly StatusStrip status = new();
    private readonly ToolStripStatusLabel slVersion = new();
    private readonly ToolStripStatusLabel slNotes = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly ToolStripStatusLabel slState = new();

    // Panels as fields for color switching
    private readonly FlowLayoutPanel pnlTopPanel;
    private readonly Panel headerPanel;

    // Error tracking & debounce
    private readonly HashSet<string> _loggersWithError = new(StringComparer.OrdinalIgnoreCase);
    private readonly WinFormsTimer _errorUiTimer = new() { Interval = 500 }; // debounce UI switching
    private bool _wantErrorUi = false;
    private bool _isErrorUi = false;

    // Watchdog timer
    private readonly WinFormsTimer _watchdogTimer = new() { Interval = 2000 }; // 2s scan

    // Colors
    private readonly Color _errorBack = Color.FromArgb(255, 247, 205); // yellowish
    private readonly Color _origFormBack;
    private readonly Color _origTopBack;
    private readonly Color _origHeaderBack;
    private readonly Color _origSplitP1Back;
    private readonly Color _origSplitP2Back;

    // Per-port row color cache and live row map
    private readonly Dictionary<string, Color> _portColorMap = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DataGridViewRow> _liveRowsByPort = new(StringComparer.OrdinalIgnoreCase);

    // Active loggers
    private readonly Dictionary<string, ComLogger> loggers = new();

    // Device change / port snapshot
    private const int WM_DEVICECHANGE = 0x0219;
    private const int DBT_DEVNODES_CHANGED = 0x0007;
    private string[] _lastPortSnapshot = Array.Empty<string>();

    public MainForm()
    {
        Text = "M81 DataTransfer";
        Width = 1280;
        Height = 880;
        StartPosition = FormStartPosition.CenterScreen;
        Shown += (_, __) => { TrySetDefaultSplit(); };
        Resize += (_, __) => { TrySetDefaultSplit(); };

        // Reduce flicker (form)
        this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
        this.UpdateStyles();

        // ===== Selector (TOP)
        pnlTopPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 90,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Padding = new Padding(10)
        };
        var lblPort = new Label { Text = "Port:", AutoSize = true, Padding = new Padding(0, 8, 4, 0) };
        var lblFolder = new Label { Text = "Ordner:", AutoSize = true, Padding = new Padding(12, 8, 4, 0) };
        pnlTopPanel.Controls.Add(lblPort);
        pnlTopPanel.Controls.Add(cmbPorts);
        pnlTopPanel.Controls.Add(btnRefreshPorts);
        pnlTopPanel.Controls.Add(lblFolder);
        pnlTopPanel.Controls.Add(txtFolder);
        pnlTopPanel.Controls.Add(btnChooseFolder);
        pnlTopPanel.Controls.Add(btnAdd);
        pnlTopPanel.Controls.Add(btnAddSim);

        // ===== Header (under selector; dark)
        headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 110,
            Padding = new Padding(20),
            BackColor = Color.FromArgb(40, 40, 40)
        };
        lblTitle.Text = "M81 DataTransfer";
        lblTitle.Font = new Font("Segoe UI", 24, FontStyle.Bold);
        lblTitle.ForeColor = Color.White;

        lblSubtitle.Text = "Frames: 089…980 (ohne Punkt).  •  T1–T5: [±dd.dd], T6: [±dd.ddddddd]  •  Dist: [±ddd.dddd]  •  Output: T1–T6";
        lblSubtitle.Font = new Font("Segoe UI", 10, FontStyle.Regular);
        lblSubtitle.ForeColor = Color.Gainsboro;

        LoadLogo();
        picLogo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        picLogo.BackColor = Color.Transparent;
        picLogo.Location = new Point(10, 6);

        var headerLeft = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
        headerLeft.Controls.Add(lblTitle);
        headerLeft.Controls.Add(lblSubtitle);
        lblTitle.Location = new Point(0, 0);
        lblSubtitle.Left = lblTitle.Left;
        lblSubtitle.Top = lblTitle.Bottom + 6; // spacing

        var headerRight = new Panel { Dock = DockStyle.Right, Width = picLogo.Width + 24, BackColor = Color.Transparent };
        headerRight.Controls.Add(picLogo);

        headerPanel.Controls.Add(headerLeft);
        headerPanel.Controls.Add(headerRight);

        // ===== Loggers ListView — minimal columns (no ID, no Rate)
        lv.Columns.Add("Port", 120);             // 0
        lv.Columns.Add("Ordner", 700);           // 1
        lv.Columns.Add("Status", 260);           // 2
        lv.Columns.Add("Letzter Fehler", 360);   // 3
        lv.Columns.Add("Age", 80);               // 4
        TryEnableDoubleBuffer(lv);

        // Split container
        split.Panel1.Controls.Add(lv);

        // ===== Live Grid columns (one row per COM, updated)
        TryEnableDoubleBuffer(dgvLive);
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeit", Name = "Time", FillWeight = 120 });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "COM", Name = "COM", FillWeight = 80 });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T1", Name = "T1" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T2", Name = "T2" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T3", Name = "T3" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T4", Name = "T4" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T5", Name = "T5" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T6", Name = "T6" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "RAW", Name = "RAW", FillWeight = 240 });

        // Do not highlight clicked/updated rows in LIVE grid
        dgvLive.DefaultCellStyle.SelectionBackColor = dgvLive.DefaultCellStyle.BackColor;
        dgvLive.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;
        dgvLive.SelectionChanged += (_, __) => { try { dgvLive.ClearSelection(); } catch { } };
        dgvLive.RowsAdded += (_, e) =>
        {
            for (int i = 0; i < e.RowCount; i++)
            {
                var row = dgvLive.Rows[e.RowIndex + i];
                row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.SelectionForeColor = row.DefaultCellStyle.ForeColor;
            }
        };

        split.Panel2.Controls.Add(dgvLive);

        // ===== Buttons
        var pnlButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(12)
        };
        pnlButtons.Controls.AddRange(new Control[] { btnStart, btnStop, btnRemove, btnOpenFolder, btnClearLive });

        // ===== Footer (status) – include reconnect duration
        slVersion.Text = $"Version {Defaults.AppVersion}";
        slNotes.Text = $"Baudrate: {Defaults.FixedBaud} • Dateiname: {Defaults.FixedFileName} • Reconnect nach: {Defaults.ReconnectIdleSeconds}s Idle";
        slState.Text = "Bereit";
        status.Items.AddRange(new ToolStripItem[] { slVersion, slNotes, slState });

        // ===== Add controls (Selector top, header below)
        Controls.Add(split);
        Controls.Add(pnlTopPanel);
        Controls.Add(headerPanel);
        Controls.Add(pnlButtons);
        Controls.Add(status);

        // Events
        btnChooseFolder.Click += (_, __) => ChooseFolder();
        btnAdd.Click += (_, __) => AddLoggerFromUi();
        btnAddSim.Click += (_, __) => AddSimLogger();
        btnRefreshPorts.Click += (_, __) => RefreshPorts();

        lv.SelectedIndexChanged += (_, __) => UpdateButtons();
        lv.DoubleClick += (_, __) => StartOrStopSelected();

        btnStart.Click += (_, __) => StartSelected();
        btnStop.Click += (_, __) => StopSelected();
        btnRemove.Click += (_, __) => RemoveSelected();
        btnOpenFolder.Click += (_, __) => OpenFolderSelected();
        btnClearLive.Click += (_, __) => { try { dgvLive.Rows.Clear(); _liveRowsByPort.Clear(); } catch { } btnClearLive.Enabled = false; };

        FormClosing += (_, __) =>
        {
            try
            {
                foreach (var lg in loggers.Values) lg.Dispose();
                // On next start selected COMs shall be empty:
                AppSettings.Save(new AppSettings { DefaultFolder = NormalizeFolder(txtFolder.Text) });
            }
            catch { }
        };

        // Init
        var settings = AppSettings.Load();
        RefreshPorts();
        txtFolder.Text = NormalizeFolder(settings.DefaultFolder ?? Defaults.BaseFolder);
        UpdateButtons();

        // Port snapshot
        _lastPortSnapshot = SerialPort.GetPortNames().OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();

        // Originalfarben merken (für Warnmodus)
        _origFormBack = this.BackColor;
        _origTopBack = pnlTopPanel.BackColor;
        _origHeaderBack = headerPanel.BackColor;
        _origSplitP1Back = split.Panel1.BackColor;
        _origSplitP2Back = split.Panel2.BackColor;

        // Debounce-Timer für Error-UI
        _errorUiTimer.Tick += (_, __) =>
        {
            if (_wantErrorUi != _isErrorUi)
            {
                ApplyGlobalErrorVisual(_wantErrorUi);
                _isErrorUi = _wantErrorUi;
            }
        };
        _errorUiTimer.Start();

        // Watchdog: update Age column and auto-reconnect on inactivity/unplug
        _watchdogTimer.Tick += (_, __) => WatchdogScan();
        _watchdogTimer.Start();
    }

    // ----- WM_DEVICECHANGE to detect USB serial arrival/removal -----
    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);

        if (m.Msg == WM_DEVICECHANGE && (int)m.WParam == DBT_DEVNODES_CHANGED)
        {
            try
            {
                var now = SerialPort.GetPortNames().OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();
                var before = _lastPortSnapshot;
                _lastPortSnapshot = now;

                AppLogger.Debug("WM_DEVICECHANGE: ports now = " + string.Join(",", now));

                // refresh combo
                cmbPorts.Items.Clear();
                cmbPorts.Items.AddRange(now);
                if (cmbPorts.Items.Count > 0 && cmbPorts.SelectedIndex < 0) cmbPorts.SelectedIndex = 0;

                // nudge running loggers only when state changed for their port
                foreach (ListViewItem it in lv.Items)
                {
                    if (it.Tag is ComLogger lg && lg.IsRunning)
                    {
                        bool existsNow = now.Contains(lg.Config.PortName, StringComparer.OrdinalIgnoreCase);
                        bool existedBefore = before.Contains(lg.Config.PortName, StringComparer.OrdinalIgnoreCase);
                        if (existsNow != existedBefore)
                        {
                            lg.RequestReconnect(existsNow ? "device_arrival" : "device_removed");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.LogException("WM_DEVICECHANGE", ex);
            }
        }
    }

    private static void TryEnableDoubleBuffer(Control c)
    {
        try
        {
            typeof(Control).GetProperty("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(c, true);
        }
        catch { }
    }

    private void TrySetDefaultSplit()
    {
        try
        {
            // Panel1 ~30% der Fläche
            if (split.Height > 0)
                split.SplitterDistance = Math.Max(120, (int)(split.Height * 0.30));
        }
        catch { }
    }

    private void LoadLogo()
    {
        try
        {
            string logoPath = Path.Combine(AppContext.BaseDirectory, "Assets", "logo.png");
            if (File.Exists(logoPath))
            {
                using var fs = new FileStream(logoPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var bmp = new Bitmap(fs);
                picLogo.Image = new Bitmap(bmp);
            }
            else
            {
                using var ms = new MemoryStream(Convert.FromBase64String(Defaults.LogoBase64));
                picLogo.Image = Image.FromStream(ms);
            }
        }
        catch (Exception ex)
        {
            AppLogger.LogException("LoadLogo", ex);
        }
    }

    private static string NormalizeFolder(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return Defaults.BaseFolder;
        path = path.Replace('/', '\\');
        if (!path.EndsWith("\\")) path += "\\";
        return path;
    }

    private void ChooseFolder()
    {
        using var dlg = new FolderBrowserDialog
        {
            Description = $"Ordner wählen (Dateiname ist fest: {Defaults.FixedFileName})",
            ShowNewFolderButton = true
        };
        var start = txtFolder.Text.Trim();
        dlg.SelectedPath = Directory.Exists(start) ? start : NormalizeFolder(Defaults.BaseFolder);
        if (dlg.ShowDialog(this) == DialogResult.OK)
            txtFolder.Text = NormalizeFolder(dlg.SelectedPath);
    }

    private void RefreshPorts()
    {
        try
        {
            var ports = SerialPort.GetPortNames().OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToArray();
            cmbPorts.Items.Clear();
            cmbPorts.Items.AddRange(ports);
            if (cmbPorts.Items.Count > 0 && cmbPorts.SelectedIndex < 0)
                cmbPorts.SelectedIndex = 0;
            slState.Text = "Ports aktualisiert";
            AppLogger.Debug("Ports: " + string.Join(",", ports));
        }
        catch (Exception ex)
        {
            slState.Text = "Fehler beim Abfragen der Ports";
            AppLogger.LogException("RefreshPorts", ex);
        }
    }

    private void AddLoggerFromUi()
    {
        var port = cmbPorts.SelectedItem as string;
        if (string.IsNullOrWhiteSpace(port)) { slState.Text = "Bitte Port wählen"; return; }

        var folder = NormalizeFolder(txtFolder.Text.Trim());
        try { Directory.CreateDirectory(folder); }
        catch (Exception ex) { slState.Text = $"Ordnerfehler: {ex.Message}"; return; }

        AddLogger(port, folder, simulated: false);
    }

    private void AddSimLogger()
    {
        var folder = NormalizeFolder(txtFolder.Text.Trim());
        try { Directory.CreateDirectory(folder); }
        catch (Exception ex) { slState.Text = $"Ordnerfehler: {ex.Message}"; return; }

        string simPort = $"SIM-{DateTime.Now:HHmmss}";
        AddLogger(simPort, folder, simulated: true);
    }

    private void AddLogger(string port, string folder, bool simulated)
    {
        if (loggers.Values.Any(l => string.Equals(l.Config.PortName, port, StringComparison.OrdinalIgnoreCase)))
        {
            slState.Text = $"Port {port} ist bereits hinzugefügt.";
            return;
        }

        var cfg = new LoggerConfig
        {
            PortName = port,
            FolderPath = folder,
            Simulated = simulated,
            AutoRebind = true
        };

        var logger = new ComLogger(port, cfg); // use port as Id now
        logger.StatusChanged += OnLoggerStatus;
        logger.MetricsUpdated += OnLoggerMetrics;
        logger.LiveRow += OnLoggerLiveRow;

        loggers[port] = logger;

        var item = new ListViewItem(new[]
        {
            cfg.PortName, cfg.FolderPath,
            simulated ? "Simuliert (gestoppt)" : "Gestoppt", "-", "-"
        })
        { Name = port, Tag = logger, UseItemStyleForSubItems = false };

        var color = GetSoftColorForPort(cfg.PortName);
        item.BackColor = color;

        lv.Items.Add(item);
        lv.SelectedItems.Clear();
        item.Selected = true;

        // Create or get live row for this port (single row per port)
        EnsureLiveRowForPort(cfg.PortName);

        slState.Text = simulated ? $"Sim-Port {port} hinzugefügt" : $"Port {port} hinzugefügt";
        AppLogger.Log($"Logger added: port={cfg.PortName}, folder={cfg.FolderPath}, sim={simulated}");
        UpdateButtons();
    }

    private DataGridViewRow EnsureLiveRowForPort(string port)
    {
        if (_liveRowsByPort.TryGetValue(port, out var row)) return row;

        int idx = dgvLive.Rows.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), port, "", "", "", "", "", "", "");
        row = dgvLive.Rows[idx];
        var color = GetSoftColorForPort(port);
        row.DefaultCellStyle.BackColor = color;
        row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor;
        row.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;
        _liveRowsByPort[port] = row;
        btnClearLive.Enabled = dgvLive.Rows.Count > 0;
        return row;
    }

    private Color GetSoftColorForPort(string port)
    {
        if (_portColorMap.TryGetValue(port, out var c)) return c;
        // Hash to HSL → soft pastel
        int hash = port.Aggregate(17, (a, ch) => unchecked(a * 31 + ch));
        double hue = (hash & 0xFFFF) / (double)0xFFFF * 360.0;  // 0..360
        double sat = 0.35; // soft
        double light = 0.92; // near white background
        var color = HslToColor(hue, sat, light);
        _portColorMap[port] = color;
        return color;
    }

    private static Color HslToColor(double h, double s, double l)
    {
        h = h % 360.0; if (h < 0) h += 360.0;
        double c = (1 - Math.Abs(2 * l - 1)) * s;
        double x = c * (1 - Math.Abs((h / 60.0) % 2 - 1));
        double m = l - c / 2;
        double r = 0, g = 0, b = 0;
        if (h < 60) { r = c; g = x; b = 0; }
        else if (h < 120) { r = x; g = c; b = 0; }
        else if (h < 180) { r = 0; g = c; b = x; }
        else if (h < 240) { r = 0; g = x; b = c; }
        else if (h < 300) { r = x; g = 0; b = c; }
        else { r = c; g = 0; b = x; }
        return Color.FromArgb(
            255,
            (int)Math.Round((r + m) * 255),
            (int)Math.Round((g + m) * 255),
            (int)Math.Round((b + m) * 255));
    }

    // Debounced global error visual switching
    private void ApplyGlobalErrorVisual(bool on)
    {
        if (on)
        {
            this.BackColor = _errorBack;
            pnlTopPanel.BackColor = _errorBack;
            headerPanel.BackColor = _errorBack;
            split.Panel1.BackColor = _errorBack;
            split.Panel2.BackColor = _errorBack;
            slState.BackColor = Color.FromArgb(255, 235, 150);
        }
        else
        {
            this.BackColor = _origFormBack;
            pnlTopPanel.BackColor = _origTopBack;
            headerPanel.BackColor = _origHeaderBack;
            split.Panel1.BackColor = _origSplitP1Back;
            split.Panel2.BackColor = _origSplitP2Back;
            slState.BackColor = SystemColors.Control;
        }
    }

    private void WatchdogScan()
    {
        try
        {
            foreach (ListViewItem it in lv.Items)
            {
                if (it.Tag is not ComLogger lg) continue;

                // update Age column (index 4)
                double ageSec = lg.LastFrameUtc == DateTime.MinValue ? double.NaN
                    : (DateTime.UtcNow - lg.LastFrameUtc).TotalSeconds;

                it.SubItems[4].Text = double.IsNaN(ageSec) ? "-" : $"{ageSec:0.0}s";

                // If running and idle ≥ threshold → request reconnect (throttled inside logger)
                if (lg.IsRunning && (!double.IsNaN(ageSec) && ageSec >= Defaults.ReconnectIdleSeconds))
                {
                    var sub = it.SubItems[2]; // Status column
                    sub.ForeColor = Color.DarkOrange;
                    sub.Font = new Font(lv.Font, FontStyle.Bold);
                    sub.Text = "Reconnect …";
                    lg.RequestReconnect("watchdog_idle");
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.LogException("WatchdogScan", ex);
        }
    }

    private void OnLoggerStatus(object? sender, LoggerStatus e)
    {
        if (IsDisposed || Disposing) return;
        BeginInvoke(new Action(() =>
        {
            if (!lv.Items.ContainsKey(e.Id)) return; // Id == Port
            var it = lv.Items[e.Id];
            it.SubItems[2].Text = e.StatusText; // Status
            it.SubItems[3].Text = string.IsNullOrEmpty(e.LastError) ? "-" : e.LastError; // LastError

            if (!string.IsNullOrWhiteSpace(e.LastError))
                _loggersWithError.Add(e.Id);
            else
                _loggersWithError.Remove(e.Id);
            _wantErrorUi = _loggersWithError.Count > 0;

            if (it.Tag is ComLogger logger)
            {
                var statusSub = it.SubItems[2];
                if (logger.IsRunning && !statusSub.Text.StartsWith("Reconnect", StringComparison.OrdinalIgnoreCase))
                {
                    statusSub.ForeColor = Color.Green;
                    statusSub.Font = new Font(lv.Font, FontStyle.Bold);
                }
                else if (!logger.IsRunning)
                {
                    statusSub.ForeColor = lv.ForeColor;
                    statusSub.Font = lv.Font;
                }
            }

            slState.Text = e.StatusText;
            UpdateButtons();
        }));
    }

    private void OnLoggerMetrics(object? sender, LoggerMetrics e)
    {
        // rate column removed – no-op
    }

    private void OnLoggerLiveRow(object? sender, LoggerLive e)
    {
        if (IsDisposed || Disposing) return;
        BeginInvoke(new Action(() =>
        {
            try
            {
                var row = EnsureLiveRowForPort(e.Port);
                row.Cells["Time"].Value = e.Ts.ToString("yyyy-MM-dd HH:mm:ss.fff");
                row.Cells["COM"].Value = e.Port;
                row.Cells["T1"].Value = e.Temps.Length > 0 ? e.Temps[0] : "";
                row.Cells["T2"].Value = e.Temps.Length > 1 ? e.Temps[1] : "";
                row.Cells["T3"].Value = e.Temps.Length > 2 ? e.Temps[2] : "";
                row.Cells["T4"].Value = e.Temps.Length > 3 ? e.Temps[3] : "";
                row.Cells["T5"].Value = e.Temps.Length > 4 ? e.Temps[4] : "";
                row.Cells["T6"].Value = e.Temps.Length > 5 ? e.Temps[5] : "";
                row.Cells["RAW"].Value = e.Raw;

                // Do NOT select or highlight anything
                dgvLive.ClearSelection();

                btnClearLive.Enabled = dgvLive.Rows.Count > 0;
            }
            catch (Exception ex)
            {
                slState.Text = $"Live-Ansicht-Fehler: {ex.Message}";
                AppLogger.LogException("OnLoggerLiveRow", ex);
            }
        }));
    }

    private void UpdateButtons()
    {
        bool has = lv.SelectedItems.Count == 1;
        btnRemove.Enabled = has;
        btnOpenFolder.Enabled = has;
        btnClearLive.Enabled = dgvLive.Rows.Count > 0;

        if (!has)
        {
            btnStart.Enabled = false;
            btnStop.Enabled = false;
            return;
        }

        if (lv.SelectedItems[0].Tag is ComLogger logger)
        {
            btnStart.Enabled = !logger.IsRunning;
            btnStop.Enabled = logger.IsRunning;
        }
        else
        {
            btnStart.Enabled = false;
            btnStop.Enabled = false;
        }
    }

    private void StartOrStopSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        if (lv.SelectedItems[0].Tag is not ComLogger logger) return;
        if (!logger.IsRunning) StartSelected(); else StopSelected();
    }

    private void StartSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        if (lv.SelectedItems[0].Tag is not ComLogger logger) return;

        if (!logger.Config.Simulated)
        {
            // allow start even if port missing; reconnect loop will wait
            if (!SerialPort.GetPortNames().Contains(logger.Config.PortName, StringComparer.OrdinalIgnoreCase))
            {
                slState.Text = $"Port {logger.Config.PortName} nicht vorhanden – warte auf Gerät …";
                logger.RequestReconnect("start_missing");
                return;
            }
            if (loggers.Values.Any(l => !ReferenceEquals(l, logger) && l.IsRunning &&
                string.Equals(l.Config.PortName, logger.Config.PortName, StringComparison.OrdinalIgnoreCase)))
            {
                slState.Text = $"Port {logger.Config.PortName} wird bereits verwendet";
                return;
            }
        }

        _ = logger.StartAsync();
    }

    private void StopSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        if (lv.SelectedItems[0].Tag is not ComLogger logger) return;
        logger.Stop();
    }

    private void RemoveSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        var sel = lv.SelectedItems[0];
        if (sel is null) return;
        var id = sel.Name; // port
        if (!string.IsNullOrEmpty(id) && loggers.TryGetValue(id, out var logger))
        {
            logger.Dispose();
            loggers.Remove(id);
            lv.Items.RemoveByKey(id);
            _loggersWithError.Remove(id);
            _wantErrorUi = _loggersWithError.Count > 0;

            // remove live row
            if (_liveRowsByPort.TryGetValue(id, out var row))
            {
                dgvLive.Rows.Remove(row);
                _liveRowsByPort.Remove(id);
            }

            slState.Text = $"Port {id} entfernt";
            UpdateButtons();
        }
    }

    private void OpenFolderSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        if (lv.SelectedItems[0].Tag is not ComLogger logger) return;
        var folder = logger.Config.FolderPath;
        try
        {
            Directory.CreateDirectory(folder);
            Process.Start(new ProcessStartInfo { FileName = folder, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            slState.Text = $"Ordner öffnen fehlgeschlagen: {ex.Message}";
            AppLogger.LogException("OpenFolderSelected", ex);
        }
    }
}

// ========================== COM LOGGER ENGINE ===========================
public sealed class ComLogger : IDisposable
{
    public LoggerConfig Config { get; }
    public bool IsRunning => _simTaskRunning || _serial?.IsOpen == true;

    public event EventHandler<LoggerStatus>? StatusChanged;
    public event EventHandler<LoggerLines>? LinesUpdated;
    public event EventHandler<LoggerMetrics>? MetricsUpdated;
    public event EventHandler<LoggerLive>? LiveRow;

    private SerialPort? _serial;
    private readonly object _serialLock = new();
    private readonly object _fileLock = new();
    private readonly ConcurrentQueue<string> _last100 = new();
    private readonly StringBuilder _buf = new();
    private string? _lastError;
    private CancellationTokenSource? _cts;
    private int _reconnectAttempt;
    private bool _reconnectLoopRunning;
    private DateTime _nextReconnectAllowedUtc = DateTime.MinValue; // throttle
    private readonly WinFormsTimer _rateTimer = new() { Interval = 1000 };
    private readonly WinFormsTimer _idleTimer = new() { Interval = Defaults.ReconnectIdleSeconds * 1000 }; // 20s idle
    private int _linesThisSecond = 0;

    // For simulation mode
    private Task? _simTask;
    private bool _simTaskRunning;

    // Strict frame (terminator "980" without dot)
    // Pattern: 089 + t1..t5 [±dd.dd] + t6 [±dd.ddddddd] + x1,x2 [±ddd.dddd] + 980
    private static readonly Regex FrameRegex = new(
        @"^089" +
        @"(?<t1>[+-]\d{2}\.\d{2})" +
        @"(?<t2>[+-]\d{2}\.\d{2})" +
        @"(?<t3>[+-]\d{2}\.\d{2})" +
        @"(?<t4>[+-]\d{2}\.\d{2})" +
        @"(?<t5>[+-]\d{2}\.\d{2})" +
        @"(?<t6>[+-]\d{2}\.\d{7})" +
        @"(?<x1>[+-]\d{3}\.\d{4})" +
        @"(?<x2>[+-]\d{3}\.\d{4})" +
        @"980$",
        RegexOptions.Compiled);

    private const int MaxBuffer = 64 * 1024;

    // Watchdog info: store ticks atomically
    private long _lastFrameTicks; // 0 == unset
    public DateTime LastFrameUtc
    {
        get
        {
            long ticks = Interlocked.Read(ref _lastFrameTicks);
            return ticks == 0 ? DateTime.MinValue : new DateTime(ticks, DateTimeKind.Utc);
        }
        set
        {
            long ticks = (value.Kind == DateTimeKind.Utc ? value : value.ToUniversalTime()).Ticks;
            Interlocked.Exchange(ref _lastFrameTicks, ticks);
        }
    }

    // encoder for pure UTF-8 without BOM
    private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

    // Jitter
    private static readonly ThreadLocal<Random> _rnd = new(() => new Random(unchecked(Environment.TickCount * 31 + Thread.CurrentThread.ManagedThreadId)));

    public string Id { get; }

    public ComLogger(string id, LoggerConfig config)
    {
        Id = id;                 // Id == Port name
        Config = config;

        _rateTimer.Tick += (_, __) =>
        {
            int rate = Interlocked.Exchange(ref _linesThisSecond, 0);
            MetricsUpdated?.Invoke(this, new LoggerMetrics(Id, rate));
        };
        _rateTimer.Start();

        // Idle timer: if no frames for a while, try reconnect (20s)
        _idleTimer.Tick += (_, __) =>
        {
            if (!IsRunning || Config.Simulated) return;
            var last = LastFrameUtc;
            var age = last == DateTime.MinValue ? TimeSpan.MaxValue : DateTime.UtcNow - last;
            if (age.TotalSeconds >= Defaults.ReconnectIdleSeconds)
            {
                AppLogger.Debug($"IdleTimer: no data for {age.TotalSeconds:0.0}s on {Config.PortName}, requesting reconnect");
                RequestReconnect("idle_timer");
            }
        };
        _idleTimer.Start();
    }

    public async Task StartAsync()
    {
        if (IsRunning) return;

        try
        {
            if (Config.Simulated)
            {
                StartSimulated();
                RaiseStatus($"Läuft – {Config.PortName} (Sim).");
                return;
            }

            if (!SerialPort.GetPortNames().Contains(Config.PortName, StringComparer.OrdinalIgnoreCase))
            {
                RaiseStatus($"Port {Config.PortName} nicht vorhanden", "Port fehlt");
                RequestReconnect("start_missing");
                return;
            }

            OpenSerial();
            _cts ??= new CancellationTokenSource();
            _reconnectAttempt = 0;
            RaiseStatus($"Läuft – {Config.PortName} @ {Defaults.FixedBaud}.");
        }
        catch (UnauthorizedAccessException ex)
        {
            LogError("Start (Port belegt/kein Zugriff)", ex);
            RaiseStatus("Start fehlgeschlagen (Zugriff verweigert). Reconnect …", ex.Message);
            RequestReconnect("start_unauthorized");
        }
        catch (IOException ex)
        {
            LogError("Start (I/O)", ex);
            RaiseStatus("Start fehlgeschlagen (I/O). Reconnect …", ex.Message);
            RequestReconnect("start_io");
        }
        catch (Exception ex)
        {
            LogError("Start", ex);
            RaiseStatus("Start fehlgeschlagen. Reconnect …", ex.Message);
            RequestReconnect("start_other");
        }
    }

    public void Stop()
    {
        try { _cts?.Cancel(); _rateTimer.Stop(); } catch { }

        if (Config.Simulated)
        {
            _simTaskRunning = false;
            try { _simTask?.Wait(200); } catch { }
        }

        try
        {
            lock (_serialLock)
            {
                if (_serial != null)
                {
                    _serial.ErrorReceived -= SerialOnError;
                    _serial.PinChanged -= SerialOnPinChanged;
                    _serial.DataReceived -= SerialOnDataReceived;
                    if (_serial.IsOpen) _serial.Close();
                    _serial.Dispose();
                    _serial = null;
                }
            }
        }
        catch { }

        _reconnectLoopRunning = false;

        RaiseStatus(Config.Simulated ? "Simuliert (gestoppt)." : "Gestoppt.");
    }

    public void Dispose() => Stop();

    private void OpenSerial()
    {
        lock (_serialLock)
        {
            if (_serial != null) return;

            var sp = new SerialPort(Config.PortName, Defaults.FixedBaud)
            {
                Parity = Parity.None,
                DataBits = 8,
                StopBits = StopBits.One,
                Handshake = Handshake.None,
                ReadTimeout = 100,
                WriteTimeout = 500,
                NewLine = "\n",
                Encoding = Encoding.UTF8,
                DtrEnable = true,
                RtsEnable = true
            };
            sp.DataReceived += SerialOnDataReceived;
            sp.ErrorReceived += SerialOnError;
            sp.PinChanged += SerialOnPinChanged;
            sp.Open();
            _serial = sp;
            AppLogger.Log($"Serial opened: {Config.PortName} @ {Defaults.FixedBaud}");
        }
    }

    private void SerialOnError(object? s, SerialErrorReceivedEventArgs e)
    {
        AppLogger.Log($"Serial error on {Config.PortName}: {e.EventType}");
        RequestReconnect("serial_error_" + e.EventType);
    }

    private void SerialOnPinChanged(object? s, SerialPinChangedEventArgs e)
    {
        AppLogger.Log($"Pin changed on {Config.PortName}: {e.EventType}");
        if (e.EventType == SerialPinChange.Break || e.EventType == SerialPinChange.CDChanged ||
            e.EventType == SerialPinChange.DsrChanged || e.EventType == SerialPinChange.CtsChanged)
        {
            RequestReconnect("pin_change_" + e.EventType);
        }
    }

    public void RequestReconnect(string reason)
    {
        // throttle reconnect requests so they don't fire too often/early
        var nowUtc = DateTime.UtcNow;
        if (nowUtc < _nextReconnectAllowedUtc)
        {
            AppLogger.Debug($"Reconnect suppressed (cooldown) on {Config.PortName}, reason={reason}");
            return;
        }
        _nextReconnectAllowedUtc = nowUtc.AddSeconds(Defaults.ReconnectRequestCooldownSeconds);

        if (_reconnectLoopRunning)
        {
            AppLogger.Debug($"Reconnect already running on {Config.PortName}, reason={reason}");
            return;
        }

        _reconnectLoopRunning = true;
        _cts ??= new CancellationTokenSource();

        Task.Run(async () =>
        {
            try
            {
                AppLogger.Debug($"Reconnect loop begin: port={Config.PortName}, reason={reason}");
                string? lastTried = null;

                while (!_cts.IsCancellationRequested)
                {
                    try
                    {
                        // fully close
                        lock (_serialLock)
                        {
                            if (_serial != null)
                            {
                                _serial.ErrorReceived -= SerialOnError;
                                _serial.PinChanged -= SerialOnPinChanged;
                                _serial.DataReceived -= SerialOnDataReceived;
                                if (_serial.IsOpen) _serial.Close();
                                _serial.Dispose();
                                _serial = null;
                            }
                        }

                        var ports = SerialPort.GetPortNames();
                        string? target = null;

                        if (ports.Contains(Config.PortName, StringComparer.OrdinalIgnoreCase))
                        {
                            target = Config.PortName;
                        }
                        else if (Config.AutoRebind && ports.Length > 0)
                        {
                            // Heuristic: choose first available by name order
                            target = ports.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).First();
                            if (!string.Equals(target, Config.PortName, StringComparison.OrdinalIgnoreCase))
                                AppLogger.Log($"AutoRebind: {Config.PortName} not found; trying {target}");
                        }

                        if (target != null)
                        {
                            // Try open
                            lock (_serialLock)
                            {
                                var sp = new SerialPort(target, Defaults.FixedBaud)
                                {
                                    Parity = Parity.None,
                                    DataBits = 8,
                                    StopBits = StopBits.One,
                                    Handshake = Handshake.None,
                                    ReadTimeout = 100,
                                    WriteTimeout = 500,
                                    NewLine = "\n",
                                    Encoding = Encoding.UTF8,
                                    DtrEnable = true,
                                    RtsEnable = true
                                };
                                sp.DataReceived += SerialOnDataReceived;
                                sp.ErrorReceived += SerialOnError;
                                sp.PinChanged += SerialOnPinChanged;
                                sp.Open();
                                _serial = sp;
                            }

                            if (!string.Equals(target, Config.PortName, StringComparison.OrdinalIgnoreCase))
                                Config.PortName = target; // rebind to new COM number

                            _reconnectAttempt = 0;
                            RaiseStatus($"Wieder verbunden – {Config.PortName} @ {Defaults.FixedBaud}.");
                            AppLogger.Log($"Serial reopened: {Config.PortName}");
                            _reconnectLoopRunning = false;
                            return;
                        }

                        // Still no suitable port
                        string joined = string.Join(",", ports);
                        if (lastTried != joined)
                        {
                            AppLogger.Debug("Reconnect waiting: no suitable COM port yet. Available: " + joined);
                            lastTried = joined;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError("ReconnectLoopOpen", ex);
                        RaiseStatus("Reconnect fehlgeschlagen.", ex.Message);
                    }

                    _reconnectAttempt++;
                    int baseDelay = (int)Math.Min(30000, Math.Pow(2, Math.Min(6, _reconnectAttempt)) * 250);
                    int jitter = _rnd.Value!.Next(-200, 200);
                    int delay = Math.Max(500, baseDelay + jitter);
                    await Task.Delay(delay, _cts.Token);
                }
            }
            catch { }
            finally { _reconnectLoopRunning = false; }
        });
    }

    private void SerialOnDataReceived(object? sender, SerialDataReceivedEventArgs e)
    {
        try
        {
            SerialPort? s;
            lock (_serialLock) s = _serial;
            if (s is null || !s.IsOpen) return;

            var chunk = s.ReadExisting();
            if (string.IsNullOrEmpty(chunk)) return;

            AppLogger.Debug($"DataReceived: chunk len={chunk.Length}, preview='{SafePreview(chunk)}'");

            lock (_buf)
            {
                _buf.Append(chunk);

                if (_buf.Length > MaxBuffer)
                {
                    var all = _buf.ToString();
                    int startIdx = all.LastIndexOf("089", StringComparison.Ordinal);
                    _buf.Clear();
                    if (startIdx >= 0)
                    {
                        var tail = all.AsSpan(startIdx);
                        _buf.Append(tail);
                        AppLogger.Debug($"Buffer trimmed, kept tail from last '089', tailLen={tail.Length}");
                    }
                    else
                    {
                        AppLogger.Debug("Buffer trimmed, no '089' found, buffer cleared");
                    }
                }

                ExtractFramesFromBuffer();
            }
        }
        catch (Exception ex)
        {
            LogError("DataReceived", ex);
            RequestReconnect("data_received_exception");
        }
    }

    private static string SafePreview(string s)
    {
        s = s.Replace("\r", "\\r").Replace("\n", "\\n");
        return s.Length <= 120 ? s : s.Substring(0, 120) + "...";
    }

    // Frame extraction with terminator "980" and optional CR/LF
    private void ExtractFramesFromBuffer()
    {
        const string Terminator = "980";
        while (true)
        {
            string all = _buf.ToString();
            int start = all.IndexOf("089", StringComparison.Ordinal);
            if (start < 0)
            {
                _buf.Clear();
                return;
            }

            int end = all.IndexOf(Terminator, start, StringComparison.Ordinal);
            if (end < 0)
            {
                if (start > 0) _buf.Remove(0, start);
                AppLogger.Debug($"Partial frame waiting, bufferLen={_buf.Length}");
                return;
            }

            int frameLen = end - start + Terminator.Length;
            string candidate = all.Substring(start, frameLen);

            // remove consumed including trailing CR/LF
            int removeLen = start + frameLen;
            if (all.Length > removeLen && (all[removeLen] == '\r' || all[removeLen] == '\n'))
            {
                removeLen++;
                if (all.Length > removeLen && all[removeLen - 1] == '\r' && all[removeLen] == '\n') removeLen++;
            }
            _buf.Remove(0, removeLen);

            var m = FrameRegex.Match(candidate);
            if (!m.Success)
            {
                AppLogger.Log($"Invalid frame format (len={candidate.Length}): '{SafePreview(candidate)}'");
                continue;
            }

            // Parse temps
            var tempsDouble = new[] { "t1", "t2", "t3", "t4", "t5", "t6" }
                .Select(k => ParseDoubleInvariant(m.Groups[k].Value))
                .ToArray();

            // Output format: +20.8800,+20.8400,... (4 decimals, with sign)
            string[] tempsFormatted = tempsDouble
                .Select(v => v.HasValue
                    ? v.Value.ToString("+0.0000;-0.0000", CultureInfo.InvariantCulture)
                    : "+0.0000")
                .ToArray();

            string fileLine = string.Join(",", tempsFormatted);

            DateTime ts = DateTime.Now;

            // Live + debug
            AppendToLive($"{ts:yyyy-MM-dd HH:mm:ss.fff} | RAW={candidate} | T= {fileLine}");
            AppLogger.Debug($"Frame OK: port={Config.PortName}, temps={fileLine}, rawLen={candidate.Length}");

            LiveRow?.Invoke(this, new LoggerLive(Id, Config.PortName, ts, tempsFormatted, candidate));

            // Only keep last value in do_not_delete.txt (atomic)
            WriteLastValueSafe(fileLine, out bool _);

            // update watchdog timestamp even if file write failed
            LastFrameUtc = DateTime.UtcNow;

            Interlocked.Increment(ref _linesThisSecond);
        }
    }

    private static double? ParseDoubleInvariant(string s)
    {
        return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var v)
            ? v : (double?)null;
    }

    private void AppendToLive(string line)
    {
        _last100.Enqueue(line);
        while (_last100.Count > 100 && _last100.TryDequeue(out _)) { }
        LinesUpdated?.Invoke(this, new LoggerLines(Id, _last100.ToArray()));
    }

    // Disk guard + atomic replace with retries (UTF-8 without BOM)
    private void WriteLastValueSafe(string fileLine, out bool diskOk)
    {
        diskOk = true;
        try
        {
            string target = Config.OutputPath;
            string? root = Path.GetPathRoot(target);
            if (!string.IsNullOrEmpty(root))
            {
                try
                {
                    var di = new DriveInfo(root);
                    if (di.AvailableFreeSpace < Defaults.MinFreeBytes)
                    {
                        RaiseStatus("Wenig Speicher – Logging pausiert", "Freier Speicher < 200 MB");
                        AppLogger.Log("Low disk space: logging paused");
                        diskOk = false;
                        return;
                    }
                }
                catch (Exception ex) { AppLogger.LogException("DriveInfo", ex); }
            }

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    lock (_fileLock)
                    {
                        Directory.CreateDirectory(Config.FolderPath);
                        string tmp = target + ".tmp";
                        File.WriteAllText(tmp, fileLine + Environment.NewLine, Utf8NoBom);
                        try
                        {
                            File.Replace(tmp, target, null);
                        }
                        catch (FileNotFoundException)
                        {
                            File.Move(tmp, target);
                        }
                    }
                    AppLogger.Debug($"Last value written: '{fileLine}' -> {target}");
                    return;
                }
                catch (IOException ex)
                {
                    AppLogger.LogException("Write(IO)", ex);
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {
                    LogError("Write", ex);
                    RaiseStatus("Schreibfehler.", ex.Message);
                    return;
                }
            }
            LogError("Write", new IOException("Mehrfache Schreibversuche fehlgeschlagen."));
            RaiseStatus("Schreibfehler (wiederholt).", "Mehrfache Schreibversuche fehlgeschlagen.");
        }
        catch (Exception ex)
        {
            diskOk = false;
            LogError("WriteOuter", ex);
            RaiseStatus("Schreibfehler (äußerer).", ex.Message);
        }
    }

    private void LogError(string where, Exception ex)
    {
        _lastError = ex.Message;
        AppLogger.LogException(where + $" [{Config.PortName}]", ex);
    }

    private void RaiseStatus(string text, string? lastError = null)
    {
        StatusChanged?.Invoke(this, new LoggerStatus(Id, text, lastError));
    }

    // Simulation mode (valid frames @ 1 Hz)
    private void StartSimulated()
    {
        _simTaskRunning = true;
        _cts ??= new CancellationTokenSource();
        var token = _cts.Token;

        _simTask = Task.Run(async () =>
        {
            var rnd = _rnd.Value!;
            while (_simTaskRunning && !token.IsCancellationRequested)
            {
                try
                {
                    string Temp2(double v) => v.ToString("+00.00;-00.00", CultureInfo.InvariantCulture);
                    string Temp7(double v)
                    {
                        string sign = v >= 0 ? "+" : "-";
                        v = Math.Abs(v);
                        return sign + v.ToString("00.0000000", CultureInfo.InvariantCulture);
                    }
                    string Dist(double v)
                    {
                        string sign = v >= 0 ? "+" : "-";
                        v = Math.Abs(v);
                        return sign + v.ToString("000.0004".Replace('4', '0'), CultureInfo.InvariantCulture); // keep 4 decimals
                    }

                    double t1 = 20.80 + (rnd.NextDouble() - 0.5) * 0.2;
                    double t2 = 20.84 + (rnd.NextDouble() - 0.5) * 0.2;
                    double t3 = 21.09 + (rnd.NextDouble() - 0.5) * 0.2;
                    double t4 = 0.05 + (rnd.NextDouble() - 0.5) * 0.02;
                    double t5 = 0.05 + (rnd.NextDouble() - 0.5) * 0.02;
                    double t6 = 0.00 + (rnd.NextDouble() - 0.5) * 0.02;

                    double x1 = (rnd.NextDouble() - 0.5) * 200;
                    double x2 = (rnd.NextDouble() - 0.5) * 200;

                    // Dist formatting helper corrected:
                    string Dist4(double v)
                    {
                        string sign = v >= 0 ? "+" : "-";
                        v = Math.Abs(v);
                        return sign + v.ToString("000.0000", CultureInfo.InvariantCulture);
                    }

                    string raw = "089"
                                 + Temp2(t1) + Temp2(t2) + Temp2(t3) + Temp2(t4) + Temp2(t5) + Temp7(t6)
                                 + Dist4(x1) + Dist4(x2)
                                 + "980";

                    lock (_buf)
                    {
                        _buf.Append(raw + "\r\n");
                        ExtractFramesFromBuffer();
                    }
                    await Task.Delay(1000, token);
                }
                catch (Exception ex)
                {
                    AppLogger.LogException("SimulatedLoop", ex);
                }
            }
        }, token);
    }
}

// ======================== MODELS & PERSISTENCE ===========================
public sealed record LoggerConfig
{
    public string PortName { get; set; } = "COM1";
    public string FolderPath { get; set; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ComPortLogger");
    public bool Simulated { get; set; } = false;
    public bool AutoRebind { get; set; } = true;

    public string OutputPath => Path.Combine(FolderPath, Defaults.FixedFileName);
}

public sealed record LoggerStatus(string Id, string StatusText, string? LastError);
public sealed record LoggerLines(string Id, string[] Last100);
public sealed record LoggerMetrics(string Id, int LinesPerSecond);
public sealed record LoggerLive(string Id, string Port, DateTime Ts, string[] Temps, string Raw);

public sealed class AppSettings
{
    public string? DefaultFolder { get; set; }

    private static string AppDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ComPortLogger");
    private static string ConfigPath => Path.Combine(AppDir, "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath, Encoding.UTF8);
                var s = JsonSerializer.Deserialize<AppSettings>(json);
                if (s is not null) return s;
            }
        }
        catch (Exception ex)
        {
            AppLogger.LogException("SettingsLoad", ex);
        }

        Directory.CreateDirectory(AppDir);
        var def = new AppSettings { DefaultFolder = Defaults.BaseFolder };
        Save(def);
        return def;
    }

    public static void Save(AppSettings s)
    {
        try
        {
            Directory.CreateDirectory(AppDir);
            var json = JsonSerializer.Serialize(s, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigPath, json, new UTF8Encoding(false));
        }
        catch (Exception ex)
        {
            AppLogger.LogException("SettingsSave", ex);
        }
    }
}
