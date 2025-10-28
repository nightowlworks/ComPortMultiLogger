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
    public const string AppVersion = "v1.0.1";

    // Disk guard (bytes)
    public const long MinFreeBytes = 200L * 1024L * 1024L; // 200 MB

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
    private const long MaxBytes = 512 * 1024; // 512 KB
    private const int Backups = 3;

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
                File.AppendAllText(_logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {msg}{Environment.NewLine}", Encoding.UTF8);
            }
        }
        catch { }
    }

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
    private readonly Button btnAdd = new() { Text = "Logger hinzufügen", Width = 150, Height = 32 };
    private readonly Button btnAddSim = new() { Text = "Sim-Port", Width = 100, Height = 32 };
    private readonly Button btnRefreshPorts = new() { Text = "Ports aktualisieren", Width = 140, Height = 32 };

    // Loggers list
    private readonly ListView lv = new()
    {
        Dock = DockStyle.Fill,
        FullRowSelect = true,
        GridLines = true,
        View = View.Details
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

    // Live table
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
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
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

    // Per-port row color cache
    private readonly Dictionary<string, Color> _portColorMap = new(StringComparer.OrdinalIgnoreCase);

    // Caps
    private const int MaxLiveRows = 200;

    private readonly Dictionary<string, ComLogger> loggers = new();

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
        lblTitle.Text = "M81 Data Transfer";
        lblTitle.Font = new Font("Segoe UI", 24, FontStyle.Bold);
        lblTitle.ForeColor = Color.White;

        lblSubtitle.Text = " Frames: 089…980. • 6× Temp. [±dd.dd], 2× Distance [±dd.dddd]";
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

        // ===== Loggers ListView — removed "Datei" column
        lv.Columns.Add("ID", 70);                // 0
        lv.Columns.Add("Port", 90);              // 1
        lv.Columns.Add("Ordner", 640);           // 2
        lv.Columns.Add("Status", 200);           // 3
        lv.Columns.Add("Rate (Zeilen/s)", 120);  // 4
        lv.Columns.Add("Letzter Fehler", 300);   // 5
        lv.Columns.Add("Age", 80);               // 6

        TryEnableDoubleBuffer(lv);
        TryEnableDoubleBuffer(dgvLive);

        // Split container
        split.Panel1.Controls.Add(lv);
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

        // ===== Live Grid columns
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeit", Name = "Time", FillWeight = 120 });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "COM", Name = "COM", FillWeight = 80 });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T1", Name = "T1" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T2", Name = "T2" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T3", Name = "T3" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T4", Name = "T4" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T5", Name = "T5" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T6", Name = "T6" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "RAW", Name = "RAW", FillWeight = 240 });

        // ===== Footer (status)
        slVersion.Text = $"Version {Defaults.AppVersion}";
        slNotes.Text = $"Baudrate: {Defaults.FixedBaud} • Dateiname: {Defaults.FixedFileName}";
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
        btnClearLive.Click += (_, __) => { try { dgvLive.Rows.Clear(); } catch { } btnClearLive.Enabled = false; };

        FormClosing += (_, __) =>
        {
            try
            {
                foreach (var lg in loggers.Values) lg.Dispose();
                AppSettings.Save(CaptureSettings());
            }
            catch { }
        };

        // Init
        var settings = AppSettings.Load();
        RefreshPorts();
        if (string.IsNullOrWhiteSpace(settings.DefaultFolder))
            settings.DefaultFolder = Defaults.BaseFolder;
        txtFolder.Text = NormalizeFolder(settings.DefaultFolder);
        RestoreSettings(settings);
        UpdateButtons();

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

        // Watchdog: update Age column and stalled state
        _watchdogTimer.Tick += (_, __) => WatchdogScan();
        _watchdogTimer.Start();
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
        catch { /* optional */ }
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
        if (!SerialPort.GetPortNames().Contains(port, StringComparer.OrdinalIgnoreCase))
        {
            slState.Text = $"Port {port} nicht mehr vorhanden";
            return;
        }

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
        var cfg = new LoggerConfig { PortName = port, FolderPath = folder, Simulated = simulated };

        var id = Guid.NewGuid().ToString("N")[..8];
        var logger = new ComLogger(id, cfg);
        logger.StatusChanged += OnLoggerStatus;
        logger.MetricsUpdated += OnLoggerMetrics;
        logger.LiveRow += OnLoggerLiveRow;

        loggers[id] = logger;

        var item = new ListViewItem(new[]
        {
            id, cfg.PortName, cfg.FolderPath,
            simulated ? "Simuliert (gestoppt)" : "Gestoppt", "0", "-", "-"
        })
        { Name = id, Tag = logger, UseItemStyleForSubItems = false };

        // Apply soft color to row
        var color = GetSoftColorForPort(cfg.PortName);
        item.BackColor = color;

        lv.Items.Add(item);
        lv.SelectedItems.Clear();
        item.Selected = true;

        slState.Text = simulated ? $"Sim-Logger {id} hinzugefügt" : $"Logger {id} hinzugefügt";
        UpdateButtons();
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

                // update Age column (index 6)
                double ageSec = lg.LastFrameUtc == DateTime.MinValue ? double.NaN
                    : (DateTime.UtcNow - lg.LastFrameUtc).TotalSeconds;

                it.SubItems[6].Text = double.IsNaN(ageSec) ? "-" : $"{ageSec:0.0}s";

                // Stalled threshold 10s when running
                if (lg.IsRunning && !double.IsNaN(ageSec) && ageSec >= 10.0)
                {
                    var sub = it.SubItems[3]; // Status column
                    sub.ForeColor = Color.DarkOrange;
                    sub.Font = new Font(lv.Font, FontStyle.Bold);
                    sub.Text = "Stalled – keine Frames";
                    lg.TryWatchdogReopen();
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
            if (!lv.Items.ContainsKey(e.Id)) return;
            var it = lv.Items[e.Id];
            it.SubItems[3].Text = e.StatusText; // Status
            it.SubItems[5].Text = string.IsNullOrEmpty(e.LastError) ? "-" : e.LastError; // LastError

            // Fehler-Set pflegen, Zielzustand für Error-UI setzen (debounced)
            if (!string.IsNullOrWhiteSpace(e.LastError))
                _loggersWithError.Add(e.Id);
            else
                _loggersWithError.Remove(e.Id);
            _wantErrorUi = _loggersWithError.Count > 0;

            // Status SubItem grün & fett wenn logger läuft
            if (it.Tag is ComLogger logger)
            {
                var statusSub = it.SubItems[3];
                if (logger.IsRunning && !statusSub.Text.StartsWith("Stalled", StringComparison.OrdinalIgnoreCase))
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
        if (IsDisposed || Disposing) return;
        BeginInvoke(new Action(() =>
        {
            if (!lv.Items.ContainsKey(e.Id)) return;
            lv.Items[e.Id].SubItems[4].Text = e.LinesPerSecond.ToString(); // Rate column
        }));
    }

    private void OnLoggerLiveRow(object? sender, LoggerLive e)
    {
        if (IsDisposed || Disposing) return;
        BeginInvoke(new Action(() =>
        {
            try
            {
                int rowIndex = dgvLive.Rows.Add(
                    e.Ts.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                    e.Port,
                    e.Temps.Length > 0 ? e.Temps[0] : "",
                    e.Temps.Length > 1 ? e.Temps[1] : "",
                    e.Temps.Length > 2 ? e.Temps[2] : "",
                    e.Temps.Length > 3 ? e.Temps[3] : "",
                    e.Temps.Length > 4 ? e.Temps[4] : "",
                    e.Temps.Length > 5 ? e.Temps[5] : "",
                    e.Raw
                );

                // Color the entire row based on port
                var color = GetSoftColorForPort(e.Port);
                dgvLive.Rows[rowIndex].DefaultCellStyle.BackColor = color;

                while (dgvLive.Rows.Count > MaxLiveRows)
                    dgvLive.Rows.RemoveAt(0);

                if (dgvLive.Rows.Count > 0)
                {
                    int lastIdx = dgvLive.Rows.Count - 1;
                    dgvLive.FirstDisplayedScrollingRowIndex = Math.Max(0, lastIdx);
                    dgvLive.ClearSelection();
                    dgvLive.Rows[lastIdx].Selected = true;
                }

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
            if (!SerialPort.GetPortNames().Contains(logger.Config.PortName, StringComparer.OrdinalIgnoreCase))
            {
                slState.Text = $"Port {logger.Config.PortName} nicht vorhanden";
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
        var id = sel.Name;
        if (!string.IsNullOrEmpty(id) && loggers.TryGetValue(id, out var logger))
        {
            logger.Dispose();
            loggers.Remove(id);
            lv.Items.RemoveByKey(id);
            _loggersWithError.Remove(id);
            _wantErrorUi = _loggersWithError.Count > 0;
            slState.Text = $"Logger {id} entfernt";
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

    private void RestoreSettings(AppSettings s)
    {
        foreach (var cfg in s.Loggers)
        {
            cfg.FolderPath = NormalizeFolder(cfg.FolderPath);
            var id = Guid.NewGuid().ToString("N")[..8];
            var logger = new ComLogger(id, cfg);
            logger.StatusChanged += OnLoggerStatus;
            logger.MetricsUpdated += OnLoggerMetrics;
            logger.LiveRow += OnLoggerLiveRow;
            loggers[id] = logger;

            var item = new ListViewItem(new[]
            {
                id, cfg.PortName, cfg.FolderPath,
                cfg.Simulated ? "Simuliert (gestoppt)" : "Gestoppt", "0", "-", "-"
            })
            { Name = id, Tag = logger, UseItemStyleForSubItems = false };

            var color = GetSoftColorForPort(cfg.PortName);
            item.BackColor = color;

            lv.Items.Add(item);
        }

        txtFolder.Text = NormalizeFolder(s.DefaultFolder ?? Defaults.BaseFolder);
    }

    private AppSettings CaptureSettings()
    {
        return new AppSettings
        {
            DefaultFolder = NormalizeFolder(txtFolder.Text),
            Loggers = loggers.Values.Select(l => l.Config).ToList()
        };
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
    private readonly WinFormsTimer _rateTimer = new() { Interval = 1000 };
    private int _linesThisSecond = 0;

    // For simulation mode
    private Task? _simTask;
    private bool _simTaskRunning;

    // Strict frame: 089 + 6 temps (+/-dd.dd) + 2 distances (+/-dd.dddd) + 980.
    private static readonly Regex FrameRegex = new(
        @"^089" +
        @"(?<t1>[+-]\d{2}\.\d{2})" +
        @"(?<t2>[+-]\d{2}\.\d{2})" +
        @"(?<t3>[+-]\d{2}\.\d{2})" +
        @"(?<t4>[+-]\d{2}\.\d{2})" +
        @"(?<t5>[+-]\d{2}\.\d{2})" +
        @"(?<t6>[+-]\d{2}\.\d{2})" +
        @"(?<x1>[+-]\d{2}\.\d{4})" +
        @"(?<x2>[+-]\d{2}\.\d{4})" +
        @"980\.$",
        RegexOptions.Compiled);

    private const int MaxBuffer = 16 * 1024;

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

    // Jitter
    private static readonly ThreadLocal<Random> _rnd = new(() => new Random(unchecked(Environment.TickCount * 31 + Thread.CurrentThread.ManagedThreadId)));

    public string Id { get; }

    public ComLogger(string id, LoggerConfig config)
    {
        Id = id;
        Config = config;

        _rateTimer.Tick += (_, __) =>
        {
            int rate = Interlocked.Exchange(ref _linesThisSecond, 0);
            MetricsUpdated?.Invoke(this, new LoggerMetrics(Id, rate));
        };
        _rateTimer.Start();
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
                // Auto-rebind (light): if exactly one other port exists, try it
                var ports = SerialPort.GetPortNames();
                if (ports.Length == 1)
                {
                    Config.PortName = ports[0];
                    AppLogger.Log($"Auto-rebind: using single available port {Config.PortName}");
                }
                else
                {
                    RaiseStatus($"Port {Config.PortName} nicht vorhanden", "Port fehlt");
                    return;
                }
            }

            OpenSerial();
            _cts = new CancellationTokenSource();
            _reconnectAttempt = 0;
            RaiseStatus($"Läuft – {Config.PortName} @ {Defaults.FixedBaud}.");
        }
        catch (UnauthorizedAccessException ex)
        {
            LogError("Start (Port belegt/kein Zugriff)", ex);
            RaiseStatus("Start fehlgeschlagen (Zugriff verweigert). Reconnect …", ex.Message);
            await ReconnectLoopAsync();
        }
        catch (IOException ex)
        {
            LogError("Start (I/O)", ex);
            RaiseStatus("Start fehlgeschlagen (I/O). Reconnect …", ex.Message);
            await ReconnectLoopAsync();
        }
        catch (Exception ex)
        {
            LogError("Start", ex);
            RaiseStatus("Start fehlgeschlagen. Reconnect …", ex.Message);
            await ReconnectLoopAsync();
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
                    _serial.DataReceived -= SerialOnDataReceived;
                    if (_serial.IsOpen) _serial.Close();
                    _serial.Dispose();
                    _serial = null;
                }
            }
        }
        catch { }

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
                Encoding = Encoding.UTF8
            };
            sp.DataReceived += SerialOnDataReceived;
            sp.Open();
            _serial = sp;
        }
    }

    private async Task ReconnectLoopAsync()
    {
        while (_cts is { IsCancellationRequested: false })
        {
            _reconnectAttempt++;
            int baseDelay = (int)Math.Min(30000, Math.Pow(2, Math.Min(6, _reconnectAttempt)) * 250);
            int jitter = _rnd.Value!.Next(-200, 200);
            int delay = Math.Max(100, baseDelay + jitter);
            RaiseStatus($"Reconnect-Versuch {_reconnectAttempt} in {delay} ms …");
            try { await Task.Delay(delay, _cts!.Token); } catch { return; }

            try
            {
                if (Config.Simulated)
                {
                    StartSimulated();
                    _reconnectAttempt = 0;
                    RaiseStatus($"Wieder verbunden – {Config.PortName} (Sim).");
                    return;
                }

                var ports = SerialPort.GetPortNames();
                if (!ports.Contains(Config.PortName, StringComparer.OrdinalIgnoreCase))
                {
                    if (ports.Length == 1)
                    {
                        Config.PortName = ports[0];
                        AppLogger.Log($"Auto-rebind during reconnect: {Config.PortName}");
                    }
                    else
                    {
                        continue;
                    }
                }

                OpenSerial();
                _reconnectAttempt = 0;
                RaiseStatus($"Wieder verbunden – {Config.PortName} @ {Defaults.FixedBaud}.");
                return;
            }
            catch (Exception ex)
            {
                LogError("Reconnect", ex);
                RaiseStatus("Reconnect fehlgeschlagen.", ex.Message);
            }
        }
    }

    // Watchdog tries to reopen on stall
    public void TryWatchdogReopen()
    {
        if (Config.Simulated) return;
        try
        {
            lock (_serialLock)
            {
                if (_serial != null)
                {
                    _serial.DataReceived -= SerialOnDataReceived;
                    if (_serial.IsOpen) _serial.Close();
                    _serial.Dispose();
                    _serial = null;
                }
            }
            OpenSerial();
            _reconnectAttempt = 0;
            RaiseStatus($"Watchdog Reopen – {Config.PortName} @ {Defaults.FixedBaud}.");
        }
        catch (Exception ex)
        {
            LogError("WatchdogReopen", ex);
            RaiseStatus("Watchdog Reopen fehlgeschlagen.", ex.Message);
        }
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

            lock (_buf)
            {
                _buf.Append(chunk);

                if (_buf.Length > MaxBuffer)
                {
                    var all = _buf.ToString();
                    int startIdx = all.LastIndexOf("089", StringComparison.Ordinal);
                    _buf.Clear();
                    if (startIdx >= 0) _buf.Append(all.AsSpan(startIdx));
                }

                ExtractFramesFromBuffer();
            }
        }
        catch (Exception ex)
        {
            LogError("DataReceived", ex);
            Task.Run(async () =>
            {
                try
                {
                    lock (_serialLock)
                    {
                        if (_serial != null)
                        {
                            _serial.DataReceived -= SerialOnDataReceived;
                            if (_serial.IsOpen) _serial.Close();
                            _serial.Dispose();
                            _serial = null;
                        }
                    }
                }
                catch { }
                await ReconnectLoopAsync();
            });
        }
    }

    private void ExtractFramesFromBuffer()
    {
        while (true)
        {
            string all = _buf.ToString();
            int start = all.IndexOf("089", StringComparison.Ordinal);
            if (start < 0)
            {
                _buf.Clear();
                return;
            }
            int end = all.IndexOf("980.", start, StringComparison.Ordinal);
            if (end < 0 || end + "980.".Length > all.Length)
            {
                if (start > 0) _buf.Remove(0, start);
                return;
            }
            int frameLen = end - start + "980.".Length;
            string candidate = all.Substring(start, frameLen);
            _buf.Remove(0, start + frameLen);

            var m = FrameRegex.Match(candidate);
            if (!m.Success) continue;

            // Parse temps
            var tempsDouble = new[] { "t1", "t2", "t3", "t4", "t5", "t6" }
                .Select(k => ParseDoubleInvariant(m.Groups[k].Value))
                .ToArray();

            string[] tempsFormatted = tempsDouble
                .Select(v => v.HasValue
                    ? v.Value.ToString("+0.0000;-0.0000", CultureInfo.InvariantCulture)
                    : "+0.0000")
                .ToArray();

            // Compose file content line
            string fileLine = string.Join(",", tempsFormatted);

            // Live info & raise
            DateTime ts = DateTime.Now;
            AppendToLive($"{ts:yyyy-MM-dd HH:mm:ss.fff} | RAW={candidate} | T= {fileLine}");
            LiveRow?.Invoke(this, new LoggerLive(Id, Config.PortName, ts, tempsFormatted, candidate));

            // Only keep last value in do_not_delete.txt (atomic)
            WriteLastValueSafe(fileLine, out bool diskOk);

            // update metrics + watchdog
            if (diskOk) LastFrameUtc = DateTime.UtcNow;
            Interlocked.Increment(ref _linesThisSecond);
        }
    }

    private static double? ParseDoubleInvariant(string s)
    {
        return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : (double?)null;
    }

    private void AppendToLive(string line)
    {
        _last100.Enqueue(line);
        while (_last100.Count > 100 && _last100.TryDequeue(out _)) { }
        LinesUpdated?.Invoke(this, new LoggerLines(Id, _last100.ToArray()));
    }

    // Disk guard + atomic replace with retries
    private void WriteLastValueSafe(string fileLine, out bool diskOk)
    {
        diskOk = true;
        try
        {
            string target = Config.OutputPath;
            string root = Path.GetPathRoot(target)!;
            try
            {
                var di = new DriveInfo(root);
                if (di.AvailableFreeSpace < Defaults.MinFreeBytes)
                {
                    RaiseStatus("Wenig Speicher – Logging pausiert", "Freier Speicher < 200 MB");
                    diskOk = false;
                    return;
                }
            }
            catch { /* ignore disk info errors */ }

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    lock (_fileLock)
                    {
                        Directory.CreateDirectory(Config.FolderPath);
                        string tmp = target + ".tmp";
                        File.WriteAllText(tmp, fileLine + Environment.NewLine, Encoding.UTF8);
                        try
                        {
                            File.Replace(tmp, target, null);
                        }
                        catch (FileNotFoundException)
                        {
                            File.Move(tmp, target);
                        }
                    }
                    return;
                }
                catch (IOException)
                {
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
        AppLogger.LogException(where, ex);
    }

    private void RaiseStatus(string text, string? lastError = null)
    {
        StatusChanged?.Invoke(this, new LoggerStatus(Id, text, lastError));
    }

    // Simulation mode: generate valid frames 1 Hz
    private void StartSimulated()
    {
        _simTaskRunning = true;
        _cts = new CancellationTokenSource();
        var token = _cts.Token;

        _simTask = Task.Run(async () =>
        {
            var rnd = _rnd.Value!;
            while (_simTaskRunning && !token.IsCancellationRequested)
            {
                try
                {
                    // temperatures ±dd.dd
                    string Temp(double v) => v.ToString("+00.00;-00.00", CultureInfo.InvariantCulture);
                    string Dist(double v) => v.ToString("+00.0000;-00.0000", CultureInfo.InvariantCulture);

                    double t1 = rnd.NextDouble() * 10 - 5 + 22.4;
                    double t2 = rnd.NextDouble() * 10 - 5 + 22.3;
                    double t3 = rnd.NextDouble() * 10 - 5 + 22.5;
                    double t4 = rnd.NextDouble() * 0.2 - 0.1;
                    double t5 = rnd.NextDouble() * 0.2 - 0.1;
                    double t6 = rnd.NextDouble() * 0.2 - 0.1;
                    double x1 = rnd.NextDouble() * 0.5 - 0.25;
                    double x2 = rnd.NextDouble() * 0.5 - 0.25;

                    string raw = "089" + Temp(t1) + Temp(t2) + Temp(t3) + Temp(t4) + Temp(t5) + Temp(t6)
                                 + Dist(x1) + Dist(x2) + "980.";
                    // Pump it through the same parser path
                    lock (_buf)
                    {
                        _buf.Append(raw);
                        ExtractFramesFromBuffer();
                    }
                    await Task.Delay(1000, token); // 1 Hz
                }
                catch { }
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

    public string OutputPath => Path.Combine(FolderPath, Defaults.FixedFileName);
}

public sealed record LoggerStatus(string Id, string StatusText, string? LastError);
public sealed record LoggerLines(string Id, string[] Last100);
public sealed record LoggerMetrics(string Id, int LinesPerSecond);
public sealed record LoggerLive(string Id, string Port, DateTime Ts, string[] Temps, string Raw);

public sealed class AppSettings
{
    public List<LoggerConfig> Loggers { get; set; } = new();
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
