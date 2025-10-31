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
    public const string BaseFolder = @"C:\BaSyTec\Drivers\OSI";
    public const string FixedFileName = "do_not_delete.txt";
    public const int FixedBaud = 9600;
    public const string AppVersion = "v1.0.91";

    // Disk guard (bytes)
    public const long MinFreeBytes = 200L * 1024L * 1024L; // 200 MB

    // Reconnect idle threshold (seconds). App waits at least this long without frames before reconnect.
    public const int ReconnectIdleSeconds = 20;

    // Polling interval for ensure-open loop (milliseconds)
    public const int EnsurePollMs = 2000;

    // Cooldown between explicit reconnect requests per logger (seconds)
    public const int ReconnectRequestCooldownSeconds = 5;

    // After opening the port, give the device a grace period to boot before we consider it idle
    public const int PostOpenGraceSeconds = 15;

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
    private readonly TextBox txtFolder = new() { Width = 400, PlaceholderText = @"Ordner für do_not_delete.txt" };
    private readonly Button btnChooseFolder = new() { Text = "Ordner", Width = 90, Height = 32 };
    private readonly Button btnAdd = new() { Text = "Port hinzufügen", Width = 150, Height = 32 };
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
    private readonly Button btnStart = new() { Text = "START", Enabled = false, Width = 92, Height = 32 };
    private readonly Button btnStop = new() { Text = "STOP", Enabled = false, Width = 92, Height = 32, ForeColor = Color.Red };
    private readonly Button btnRemove = new() { Text = "Entfernen", Enabled = false, Width = 110, Height = 32 };
    private readonly Button btnOpenFolder = new() { Text = "Ordner öffnen", Enabled = false, Width = 130, Height = 32 };

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

    // Colors for error highlighting in live grid
    private readonly Color _liveErrorBack = Color.Red;
    private readonly Color _liveErrorFore = Color.White;

    // Original background colors (for restoring)
    private readonly Color _origFormBack;
    private readonly Color _origTopBack;
    private readonly Color _origHeaderBack;
    private readonly Color _origSplitP1Back;
    private readonly Color _origSplitP2Back;

    // Track live rows
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

        // Reduce overall line height / font size for tighter rows
        lv.Font = new Font("Segoe UI", 9.0f);
        dgvLive.RowTemplate.Height = 20;
        dgvLive.DefaultCellStyle.Padding = new Padding(0);
        dgvLive.DefaultCellStyle.Font = new Font("Segoe UI", 9.0f);

        // ===== Header (under selector; dark)
        headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 100,
            Padding = new Padding(20),
            BackColor = Color.FromArgb(40, 40, 40)
        };
        lblTitle.Text = "M81 DataTransfer";
        lblTitle.Font = new Font("Segoe UI", 24, FontStyle.Bold);
        lblTitle.ForeColor = Color.White;

        LoadLogo();
        picLogo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        picLogo.BackColor = Color.Transparent;
        picLogo.Location = new Point(10, 6);

        var headerLeft = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
        headerLeft.Controls.Add(lblTitle);
        headerLeft.Controls.Add(lblSubtitle);
        lblTitle.Location = new Point(0, 0);

        var headerRight = new Panel { Dock = DockStyle.Right, Width = picLogo.Width + 24, BackColor = Color.Transparent };
        headerRight.Controls.Add(picLogo);

        headerPanel.Controls.Add(headerLeft);
        headerPanel.Controls.Add(headerRight);

        // ===== Loggers ListView — keep: Port, Ordner, Status, Age (no row coloring)
        lv.Columns.Add("Port", 120);             // 0
        lv.Columns.Add("Ordner", 460);           // 1
        lv.Columns.Add("Status", 360);           // 2
        lv.Columns.Add("Age", 80);               // 3
        TryEnableDoubleBuffer(lv);
        lv.ListViewItemSorter = new ComListViewComparer(0);

        // Split container
        split.Panel1.Controls.Add(lv);

        // ===== Live Grid columns (COM first; no default row coloring)
        TryEnableDoubleBuffer(dgvLive);
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "COM", Name = "COM", FillWeight = 80 });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeit", Name = "Time", FillWeight = 120 });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T1", Name = "T1" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T2", Name = "T2" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T3", Name = "T3" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T4", Name = "T4" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T5", Name = "T5" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T6", Name = "T6" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "RAW", Name = "RAW", FillWeight = 240 });

        // Custom sort for COM column (numeric-aware)
        dgvLive.SortCompare += (s, e) =>
        {
            if (e.Column.Name == "COM")
            {
                e.SortResult = ComOrder.CompareComNames(e.CellValue1?.ToString(), e.CellValue2?.ToString());
                e.Handled = true;
            }
        };

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
        pnlButtons.Controls.AddRange(new Control[] { btnStart, btnStop, btnRemove, btnOpenFolder });

        // ===== Footer (status) – include reconnect duration
        slVersion.Text = $"Version {Defaults.AppVersion}";
        slNotes.Text = $"Baudrate: {Defaults.FixedBaud}";
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
        btnRefreshPorts.Click += (_, __) => RefreshPorts();

        lv.SelectedIndexChanged += (_, __) => UpdateButtons();
        lv.DoubleClick += (_, __) => StartOrStopSelected();

        btnStart.Click += (_, __) => StartSelected();
        btnStop.Click += (_, __) => StopSelected();
        btnRemove.Click += (_, __) => RemoveSelected();
        btnOpenFolder.Click += (_, __) => OpenFolderSelected();

        // Closing confirmation and persistence
        FormClosing += (s, e) =>
        {
            var dr = MessageBox.Show(
                this,
                "Sind Sie sicher, dass Sie das Programm schließen möchten?",
                "Programm beenden",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (dr != DialogResult.Yes)
            {
                e.Cancel = true;
                return;
            }

            try
            {
                foreach (var lg in loggers.Values) lg.Dispose();
                // Save selected COMs + folders for restore
                var settings = new AppSettings
                {
                    DefaultFolder = NormalizeFolder(txtFolder.Text),
                    Saved = lv.Items.Cast<ListViewItem>()
                        .Select(it => new LoggerConfigSnapshot
                        {
                            PortName = it.SubItems[0].Text,
                            FolderPath = (it.Tag as ComLogger)?.Config.FolderPath ?? it.SubItems[1].Text
                        })
                        .ToList()
                };
                AppSettings.Save(settings);
            }
            catch { }
        };

        // Init
        var settings = AppSettings.Load();
        RefreshPorts();
        txtFolder.Text = NormalizeFolder(settings.DefaultFolder ?? Defaults.BaseFolder);

        // Restore saved loggers (do not autostart)
        if (settings.Saved != null)
        {
            foreach (var snap in settings.Saved)
            {
                try { AddLogger(snap.PortName, NormalizeFolder(snap.FolderPath)); }
                catch (Exception ex) { AppLogger.LogException("RestoreLogger", ex); }
            }
            SortListViewByPort();
            SortLiveByPort();
        }

        UpdateButtons();

        // Port snapshot
        _lastPortSnapshot = SerialPort.GetPortNames()
            .OrderBy(p => p, Comparer<string>.Create(ComOrder.CompareComNames))
            .ToArray();

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

        // Watchdog: update Age column and auto-reconnect on inactivity/unplug + live row red highlight
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
                var now = SerialPort.GetPortNames()
                    .OrderBy(p => p, Comparer<string>.Create(ComOrder.CompareComNames))
                    .ToArray();
                var before = _lastPortSnapshot;
                _lastPortSnapshot = now;

                AppLogger.Debug("WM_DEVICECHANGE: ports now = " + string.Join(",", now));

                // refresh combo
                cmbPorts.Items.Clear();
                cmbPorts.Items.AddRange(now);
                if (cmbPorts.Items.Count > 0 && cmbPorts.SelectedIndex < 0) cmbPorts.SelectedIndex = 0;

                // tell running loggers to ensure/open soon
                foreach (ListViewItem it in lv.Items)
                {
                    if (it.Tag is ComLogger lg && lg.WantsRunning)
                    {
                        lg.NudgeEnsure("device_change");
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
            var ports = SerialPort.GetPortNames()
                .OrderBy(s => s, Comparer<string>.Create(ComOrder.CompareComNames))
                .ToArray();
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

        AddLogger(port, folder);
        SortListViewByPort();
        SortLiveByPort();
    }

    private void AddLogger(string port, string folder)
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
            AutoRebind = true
        };

        var logger = new ComLogger(port, cfg); // use port as Id
        logger.StatusChanged += OnLoggerStatus;
        logger.LiveRow += OnLoggerLiveRow;

        loggers[port] = logger;

        var item = new ListViewItem(new[]
        {
            cfg.PortName, cfg.FolderPath,
            "Gestoppt", "-"
        })
        { Name = port, Tag = logger, UseItemStyleForSubItems = false };

        // No row coloring anymore
        lv.Items.Add(item);
        lv.SelectedItems.Clear();
        item.Selected = true;

        // Create or get live row for this port (single row per port)
        EnsureLiveRowForPort(cfg.PortName);

        slState.Text = $"Port {port} hinzugefügt";
        AppLogger.Log($"Logger added: port={cfg.PortName}, folder={cfg.FolderPath}");
        UpdateButtons();
    }

    private DataGridViewRow EnsureLiveRowForPort(string port)
    {
        if (_liveRowsByPort.TryGetValue(port, out var row)) return row;

        int idx = dgvLive.Rows.Add(port, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), "", "", "", "", "", "", "");
        row = dgvLive.Rows[idx];

        // No default row coloring
        row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor;
        row.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;

        _liveRowsByPort[port] = row;
        return row;
    }

    // Debounced global error visual switching
    private readonly Color _errorBack = Color.FromArgb(255, 247, 205); // gelbliches Warn-Highlight
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

                // update Age column (index 3)
                double ageSec = lg.LastFrameUtc == DateTime.MinValue ? double.NaN
                    : (DateTime.UtcNow - lg.LastFrameUtc).TotalSeconds;

                it.SubItems[3].Text = double.IsNaN(ageSec) ? "-" : $"{ageSec:0.0}s";

                // Live grid red highlight if running & idle ≥ threshold
                if (_liveRowsByPort.TryGetValue(lg.Config.PortName, out var liveRow))
                {
                    bool alert = lg.WantsRunning && !double.IsNaN(ageSec) && ageSec >= Defaults.ReconnectIdleSeconds;
                    if (alert)
                    {
                        liveRow.DefaultCellStyle.BackColor = _liveErrorBack;
                        liveRow.DefaultCellStyle.ForeColor = _liveErrorFore;
                        liveRow.DefaultCellStyle.SelectionBackColor = _liveErrorBack;
                        liveRow.DefaultCellStyle.SelectionForeColor = _liveErrorFore;
                    }
                    else
                    {
                        // restore defaults
                        liveRow.DefaultCellStyle.BackColor = dgvLive.DefaultCellStyle.BackColor;
                        liveRow.DefaultCellStyle.ForeColor = dgvLive.DefaultCellStyle.ForeColor;
                        liveRow.DefaultCellStyle.SelectionBackColor = dgvLive.DefaultCellStyle.BackColor;
                        liveRow.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;
                    }
                }

                // If running and idle ≥ threshold → nudge ensure loop (it will reopen with grace)
                if (lg.WantsRunning && (!double.IsNaN(ageSec) && ageSec >= Defaults.ReconnectIdleSeconds))
                {
                    var sub = it.SubItems[2]; // Status column
                    sub.ForeColor = Color.DarkOrange;
                    sub.Font = new Font(lv.Font, FontStyle.Bold);
                    sub.Text = "Reconnect …";
                    lg.NudgeEnsure("watchdog_idle");
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

            _wantErrorUi = _loggersWithError.Count > 0; // keep global logic minimal

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

                    // also clear any red highlight when stopped
                    if (_liveRowsByPort.TryGetValue(logger.Config.PortName, out var liveRow))
                    {
                        liveRow.DefaultCellStyle.BackColor = dgvLive.DefaultCellStyle.BackColor;
                        liveRow.DefaultCellStyle.ForeColor = dgvLive.DefaultCellStyle.ForeColor;
                        liveRow.DefaultCellStyle.SelectionBackColor = dgvLive.DefaultCellStyle.BackColor;
                        liveRow.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;
                    }
                }
            }

            slState.Text = e.StatusText;
            UpdateButtons();
        }));
    }

    private void OnLoggerLiveRow(object? sender, LoggerLive e)
    {
        if (IsDisposed || Disposing) return;
        BeginInvoke(new Action(() =>
        {
            try
            {
                var row = EnsureLiveRowForPort(e.Port);
                row.Cells["COM"].Value = e.Port;
                row.Cells["Time"].Value = e.Ts.ToString("yyyy-MM-dd HH:mm:ss.fff");
                row.Cells["T1"].Value = e.Temps.Length > 0 ? e.Temps[0] : "";
                row.Cells["T2"].Value = e.Temps.Length > 1 ? e.Temps[1] : "";
                row.Cells["T3"].Value = e.Temps.Length > 2 ? e.Temps[2] : "";
                row.Cells["T4"].Value = e.Temps.Length > 3 ? e.Temps[3] : "";
                row.Cells["T5"].Value = e.Temps.Length > 4 ? e.Temps[4] : "";
                row.Cells["T6"].Value = e.Temps.Length > 5 ? e.Temps[5] : "";
                row.Cells["RAW"].Value = e.Raw;

                // New data: clear any alert highlight immediately
                row.DefaultCellStyle.BackColor = dgvLive.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.ForeColor = dgvLive.DefaultCellStyle.ForeColor;
                row.DefaultCellStyle.SelectionBackColor = dgvLive.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;

                // Do NOT select or highlight anything
                dgvLive.ClearSelection();

                // Keep live grid sorted by COM ascending
                SortLiveByPort();
            }
            catch (Exception ex)
            {
                slState.Text = $"Live-Ansicht-Fehler: {ex.Message}";
                AppLogger.LogException("OnLoggerLiveRow", ex);
            }
        }));
    }

    private void SortListViewByPort()
    {
        try { lv.Sort(); } catch { }
    }

    private void SortLiveByPort()
    {
        try
        {
            if (dgvLive.Columns.Count == 0) return;
            dgvLive.Sort(dgvLive.Columns["COM"], System.ComponentModel.ListSortDirection.Ascending);
        }
        catch { }
    }

    private void UpdateButtons()
    {
        bool has = lv.SelectedItems.Count == 1;
        btnRemove.Enabled = has;
        btnOpenFolder.Enabled = has;

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

        _ = logger.StartAsync();
    }

    private void StopSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        if (lv.SelectedItems[0].Tag is not ComLogger logger) return;

        var dr = MessageBox.Show(
            this,
            $"Sind Sie sicher, dass Sie den Port {logger.Config.PortName} stoppen möchten?",
            "Port stoppen",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);

        if (dr != DialogResult.Yes) return;

        logger.Stop();
    }

    private void RemoveSelected()
    {
        if (lv.SelectedItems.Count != 1) return;

        // Ask user "are you sure?"
        var dr = MessageBox.Show(
            this,
            "Sind Sie sicher, dass Sie den ausgewählten Port entfernen möchten?",
            "Port entfernen",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);
        if (dr != DialogResult.Yes) return;

        var sel = lv.SelectedItems[0];
        if (sel is null) return;
        var id = sel.Name; // port
        if (!string.IsNullOrEmpty(id) && loggers.TryGetValue(id, out var logger))
        {
            logger.Dispose();
            loggers.Remove(id);

            // remove live row & ensure no color left
            if (_liveRowsByPort.TryGetValue(id, out var row))
            {
                // restore default colors (not strictly necessary since we remove the row, but keeps UI clean)
                row.DefaultCellStyle.BackColor = dgvLive.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.ForeColor = dgvLive.DefaultCellStyle.ForeColor;
                row.DefaultCellStyle.SelectionBackColor = dgvLive.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;

                dgvLive.Rows.Remove(row);
                _liveRowsByPort.Remove(id);
            }

            lv.Items.RemoveByKey(id);

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
    public bool IsRunning => _serial?.IsOpen == true;
    public bool WantsRunning => _desiredRunning;

    public event EventHandler<LoggerStatus>? StatusChanged;
    public event EventHandler<LoggerLines>? LinesUpdated;
    public event EventHandler<LoggerLive>? LiveRow;

    private SerialPort? _serial;
    private readonly object _serialLock = new();
    private readonly object _fileLock = new();
    private readonly ConcurrentQueue<string> _last100 = new();
    private readonly StringBuilder _buf = new();
    private string? _lastError;
    private CancellationTokenSource? _cts;
    private bool _desiredRunning;
    private readonly WinFormsTimer _idleTimer = new() { Interval = Defaults.ReconnectIdleSeconds * 1000 }; // 20s idle

    // Ensure-open background loop
    private CancellationTokenSource? _ensureCts;
    private Task? _ensureTask;
    private DateTime _graceUntilUtc = DateTime.MinValue;
    private DateTime _nextReconnectAllowedUtc = DateTime.MinValue;

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

    public string Id { get; }

    public ComLogger(string id, LoggerConfig config)
    {
        Id = id;                 // Id == Port name
        Config = config;

        // Idle timer: if no frames for a while, ask ensure-loop to reopen
        _idleTimer.Tick += (_, __) =>
        {
            if (!WantsRunning) return;

            // respect grace window after (re)open
            if (DateTime.UtcNow < _graceUntilUtc) return;

            var last = LastFrameUtc;
            var age = last == DateTime.MinValue ? TimeSpan.MaxValue : DateTime.UtcNow - last;
            if (age.TotalSeconds >= Defaults.ReconnectIdleSeconds)
            {
                AppLogger.Debug($"IdleTimer: no data {age.TotalSeconds:0.0}s on {Config.PortName}, nudge ensure");
                NudgeEnsure("idle_timer");
            }
        };
        _idleTimer.Start();
    }

    public async Task StartAsync()
    {
        if (_desiredRunning) return;

        _desiredRunning = true;
        _cts ??= new CancellationTokenSource();

        // Start ensure loop
        StartEnsureLoop();

        // Try initial open quickly; don't await the loop
        _ = Task.Run(() => EnsureOpenOnceAsync("start"));
    }

    public void Stop()
    {
        _desiredRunning = false;

        try
        {
            _ensureCts?.Cancel();
            _ensureTask = null;
        }
        catch { }

        SafeClose("stop");

        RaiseStatus("Gestoppt.");
    }

    public void Dispose() => Stop();

    private void StartEnsureLoop()
    {
        if (_ensureTask != null && !_ensureTask.IsCompleted) return;

        _ensureCts = new CancellationTokenSource();
        var token = _ensureCts.Token;

        _ensureTask = Task.Run(async () =>
        {
            AppLogger.Debug($"Ensure loop started for {Config.PortName}");
            while (!token.IsCancellationRequested)
            {
                try
                {
                    if (!_desiredRunning)
                    {
                        await Task.Delay(Defaults.EnsurePollMs, token);
                        continue;
                    }

                    // If port is not open → try open when present (or rebind)
                    if (!IsOpen())
                    {
                        await EnsureOpenOnceAsync("ensure_loop");
                        await Task.Delay(Defaults.EnsurePollMs, token);
                        continue;
                    }

                    // If open but no data for too long → close (ensure will reopen)
                    if (DateTime.UtcNow >= _graceUntilUtc)
                    {
                        var last = LastFrameUtc;
                        var age = last == DateTime.MinValue ? TimeSpan.MaxValue : DateTime.UtcNow - last;
                        if (age.TotalSeconds >= Defaults.ReconnectIdleSeconds)
                        {
                            AppLogger.Debug($"Ensure: idle {age.TotalSeconds:0.0}s on {Config.PortName}, closing for reopen");
                            SafeClose("ensure_idle");
                        }
                    }
                }
                catch (TaskCanceledException) { }
                catch (Exception ex)
                {
                    AppLogger.LogException("EnsureLoop", ex);
                }

                try { await Task.Delay(Defaults.EnsurePollMs, token); } catch { }
            }
            AppLogger.Debug($"Ensure loop stopped for {Config.PortName}");
        }, token);
    }

    private bool IsOpen()
    {
        lock (_serialLock) return _serial?.IsOpen == true;
    }

    public void NudgeEnsure(string reason)
    {
        AppLogger.Debug($"NudgeEnsure({reason}) for {Config.PortName}");
        // Force immediate attempt by allowing next reconnect now
        _nextReconnectAllowedUtc = DateTime.MinValue;
        // If currently open, we let idle timer/ensure logic handle; if closed, ensure loop will try soon.
    }

    private async Task EnsureOpenOnceAsync(string reason)
    {
        try
        {
            if (!_desiredRunning) return;

            var now = DateTime.UtcNow;
            if (now < _nextReconnectAllowedUtc) return; // throttle explicit attempts
            _nextReconnectAllowedUtc = now.AddSeconds(Defaults.ReconnectRequestCooldownSeconds);

            var ports = SerialPort.GetPortNames();
            string? target = null;

            if (ports.Contains(Config.PortName, StringComparer.OrdinalIgnoreCase))
            {
                target = Config.PortName;
            }
            else if (Config.AutoRebind && ports.Length > 0)
            {
                // Heuristic: choose first available by name order (numeric-aware)
                target = ports.OrderBy(p => p, Comparer<string>.Create(ComOrder.CompareComNames)).First();
                if (!string.Equals(target, Config.PortName, StringComparison.OrdinalIgnoreCase))
                    AppLogger.Log($"AutoRebind: {Config.PortName} not found; trying {target}");
            }

            if (target == null)
            {
                AppLogger.Debug($"EnsureOpen: no suitable COM port yet for {Config.PortName}");
                RaiseStatus($"Warte auf Gerät – {Config.PortName} …");
                return;
            }

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
                sp.DiscardInBuffer();
                sp.DiscardOutBuffer();
                _serial = sp;
            }

            if (!string.Equals(target, Config.PortName, StringComparison.OrdinalIgnoreCase))
                Config.PortName = target;

            // After open, give Arduino time to boot (grace)
            _graceUntilUtc = DateTime.UtcNow.AddSeconds(Defaults.PostOpenGraceSeconds);
            LastFrameUtc = DateTime.UtcNow; // reset idle age
            RaiseStatus($"Läuft – {Config.PortName} @ {Defaults.FixedBaud}.");
            AppLogger.Log($"Serial opened: {Config.PortName} @ {Defaults.FixedBaud} (reason={reason})");
        }
        catch (UnauthorizedAccessException ex)
        {
            AppLogger.LogException("EnsureOpen(Unauthorized)", ex);
            RaiseStatus("Port belegt/kein Zugriff – erneuter Versuch …");
        }
        catch (IOException ex)
        {
            AppLogger.LogException("EnsureOpen(IO)", ex);
            RaiseStatus("I/O-Fehler – erneuter Versuch …");
        }
        catch (Exception ex)
        {
            AppLogger.LogException("EnsureOpen(Other)", ex);
            RaiseStatus("Fehler beim Öffnen – erneuter Versuch …");
        }
    }

    private void SafeClose(string reason)
    {
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
            RaiseStatus("Getrennt – Reconnect läuft …");
            AppLogger.Log($"Serial closed ({reason}) on {Config.PortName}");
        }
        catch (Exception ex)
        {
            AppLogger.LogException("SafeClose", ex);
        }
    }

    private void SerialOnError(object? s, SerialErrorReceivedEventArgs e)
    {
        AppLogger.Log($"Serial error on {Config.PortName}: {e.EventType}");
        SafeClose("serial_error_" + e.EventType);
        NudgeEnsure("serial_error");
    }

    private void SerialOnPinChanged(object? s, SerialPinChangedEventArgs e)
    {
        AppLogger.Log($"Pin changed on {Config.PortName}: {e.EventType}");
        if (e.EventType == SerialPinChange.Break || e.EventType == SerialPinChange.CDChanged ||
            e.EventType == SerialPinChange.DsrChanged || e.EventType == SerialPinChange.CtsChanged)
        {
            SafeClose("pin_change_" + e.EventType);
            NudgeEnsure("pin_change");
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

            AppLogger.Debug($"DataReceived {Config.PortName}: len={chunk.Length}, preview='{SafePreview(chunk)}'");

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
            AppLogger.LogException("DataReceived", ex);
            SafeClose("data_received_exception");
            NudgeEnsure("data_received_exception");
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
                AppLogger.Log($"Invalid frame format (len={candidate.Length}, port={Config.PortName}): '{SafePreview(candidate)}'");
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

            // mark activity
            LastFrameUtc = DateTime.UtcNow;

            // Live + debug
            AppendToLive($"{ts:yyyy-MM-dd HH:mm:ss.fff} | RAW={candidate} | T= {fileLine}");
            AppLogger.Debug($"Frame OK {Config.PortName}: temps={fileLine}, rawLen={candidate.Length}");

            LiveRow?.Invoke(this, new LoggerLive(Id, Config.PortName, ts, tempsFormatted, candidate));

            // Only keep last value in do_not_delete.txt (atomic)
            WriteLastValueSafe(fileLine);
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
    private void WriteLastValueSafe(string fileLine)
    {
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
                        RaiseStatus("Wenig Speicher – Logging pausiert");
                        AppLogger.Log("Low disk space: logging paused");
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
                    AppLogger.LogException("Write(Other)", ex);
                    RaiseStatus("Schreibfehler.");
                    return;
                }
            }
            AppLogger.Log("Write failed repeatedly.");
            RaiseStatus("Schreibfehler (wiederholt).");
        }
        catch (Exception ex)
        {
            AppLogger.LogException("WriteOuter", ex);
            RaiseStatus("Schreibfehler (äußerer).");
        }
    }

    private void RaiseStatus(string text)
    {
        StatusChanged?.Invoke(this, new LoggerStatus(Id, text, null));
    }
}

// ======================== MODELS & PERSISTENCE ===========================
public sealed record LoggerConfig
{
    public string PortName { get; set; } = "COM1";
    public string FolderPath { get; set; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ComPortLogger");
    public bool AutoRebind { get; set; } = true;

    public string OutputPath => Path.Combine(FolderPath, Defaults.FixedFileName);
}

public sealed record LoggerStatus(string Id, string StatusText, string? LastError);
public sealed record LoggerLines(string Id, string[] Last100);
public sealed record LoggerLive(string Id, string Port, DateTime Ts, string[] Temps, string Raw);

public sealed class AppSettings
{
    public string? DefaultFolder { get; set; }
    public List<LoggerConfigSnapshot> Saved { get; set; } = new();

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
        var def = new AppSettings { DefaultFolder = Defaults.BaseFolder, Saved = new List<LoggerConfigSnapshot>() };
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

public sealed class LoggerConfigSnapshot
{
    public string PortName { get; set; } = "COM1";
    public string FolderPath { get; set; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ComPortLogger");
}

// ======================== COM NAME ORDER HELPERS =========================
public static class ComOrder
{
    public static int CompareComNames(string? a, string? b)
    {
        if (ReferenceEquals(a, b)) return 0;
        if (a is null) return -1;
        if (b is null) return 1;

        // Try parse like "COM123"
        if (TryParseCom(a, out var prefixA, out var numA) && TryParseCom(b, out var prefixB, out var numB))
        {
            int p = string.Compare(prefixA, prefixB, StringComparison.OrdinalIgnoreCase);
            if (p != 0) return p;
            return numA.CompareTo(numB);
        }

        // Fallback to case-insensitive ordinal
        return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParseCom(string s, out string prefix, out int num)
    {
        prefix = s;
        num = 0;

        // Find trailing integer
        int i = s.Length - 1;
        while (i >= 0 && char.IsDigit(s[i])) i--;
        // i now at last non-digit; digits are i+1..end
        if (i < s.Length - 1)
        {
            var digits = s.Substring(i + 1);
            if (int.TryParse(digits, out num))
            {
                prefix = s.Substring(0, i + 1);
                return true;
            }
        }
        return false;
    }
}

public sealed class ComListViewComparer : System.Collections.IComparer
{
    private readonly int _column;
    public ComListViewComparer(int column) { _column = column; }
    public int Compare(object? x, object? y)
    {
        if (x is not ListViewItem a || y is not ListViewItem b) return 0;
        var sa = a.SubItems[_column].Text;
        var sb = b.SubItems[_column].Text;
        return ComOrder.CompareComNames(sa, sb);
    }
}
