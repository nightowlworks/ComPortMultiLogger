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

// Alias, um Timer-Mehrdeutigkeit zu vermeiden
using WinFormsTimer = System.Windows.Forms.Timer;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        AppLogger.Init(); // Logger vorbereiten

        // Globale Fail-Safes
        Application.ThreadException += (s, e) =>
        {
            AppLogger.LogException("ThreadException", e.Exception);
            MessageBox.Show("Ein unerwarteter Fehler ist aufgetreten.\n" +
                            e.Exception.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            CrashDumper.TryWriteMiniDump("thread_exception");
        };
        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
        {
            var ex = e.ExceptionObject as Exception ?? new Exception("Unknown unhandled exception");
            AppLogger.LogException("UnhandledException", ex);
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

// =============================== Defaults ===============================
public static class Defaults
{
    public const string BaseFolder = @"C:\BaSyTec\Drivers\OSI\";
    public const string FixedFileName = "do_not_delete.txt";
    public const int FixedBaud = 9600;
    public const string AppVersion = "v1.1.0";

    // Disk guard (bytes)
    public const long MinFreeBytes = 200L * 1024L * 1024L; // 200 MB

    // Reconnect idle threshold (seconds)
    public const int ReconnectIdleSeconds = 20;

    // Polling interval for ensure-open loop (milliseconds)
    public const int EnsurePollMs = 2000;

    // Cooldown zwischen Reconnect-Versuchen (seconds)
    public const int ReconnectRequestCooldownSeconds = 5;

    // Grace nach Öffnen (z.B. Arduino-Boot) (seconds)
    public const int PostOpenGraceSeconds = 15;

    // Fallback-Logo PNG (weiß, 64x64) als Base64 falls Assets\logo.png fehlt
    public const string LogoBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAQAAABt9U0VAAAACXBIWXMAAAsSAAALEgHS3X78AAABc0lEQVR4nO2Z0U7DUBCFv2g+1Wk" +
        "lNwH2q3kQyG6p2h9F1p7QmXwS8rJwQyqg0Qor8H2gq1eR7Ywz4qkL4T8m2C4j1kzv9mEw+g2I0H0e4Q3rLQz4l1H8aB2t3p0i7k7xqj3kH" +
        "WqgQqk3y0mY9hUTqGx8lqvQh7GvG3H7HfJqkYQwQ1qg9eY6mZ6mC6m8ZbXy2cXr7u1w6Yz2k4n+7v0jY6C3nS6S5QvD6oG6b0KkKpKj8p" +
        "wq2lU0V2hXwq5YQ0r5j7r7Q0bQw3r8x5s7z0r/8rR8/4m6p9gM8m2b7wAjyM0sJ8o5Jq4c9wYQm8m2b4w0Wl0mQ5OIG5m3m8gG0k0k0l" +
        "0k0k0m8m8n8q8n8o8o8p8p8q8q8r8r8s8s8t8t8u8u8v8v8w8w8x8x8y8y8z8z8z8z8z8z8z8z8z8z8z8/9yoYwF6h7GkSx9QAAAAASUVORK5CYII=";
}

// =============================== Logging & Dumps ===============================
public enum LogLevel { Error = 0, Info = 1, Debug = 2 }

public static class AppLogger
{
    private static readonly object _lock = new();
    private static readonly string _logDir = Path.Combine(AppContext.BaseDirectory, "logs");
    private static readonly string _logPath = Path.Combine(_logDir, "app.log");
    private const long MaxBytes = 1024 * 1024; // 1 MB
    private const int Backups = 5;

    public static LogLevel Level { get; set; } = LogLevel.Info;

    public static string LogPath => _logPath;

    public static void Init()
    {
        try { Directory.CreateDirectory(_logDir); } catch { }
    }

    public static void SetLevel(LogLevel lvl)
    {
        Level = lvl;
        Log($"[LOG] Level set to {lvl}");
    }

    public static void ClearLog()
    {
        try
        {
            lock (_lock)
            {
                Directory.CreateDirectory(_logDir);
                File.WriteAllText(_logPath, string.Empty, new UTF8Encoding(false));
            }
            Log("[LOG] Cleared");
        }
        catch { }
    }

    public static void Log(string msg) => Write(msg);

    public static void Info(string msg)
    {
        if (Level >= LogLevel.Info) Write("[INFO] " + msg);
    }

    public static void Debug(string msg)
    {
        if (Level >= LogLevel.Debug) Write("[DEBUG] " + msg);
    }

    public static void LogException(string where, Exception ex)
    {
        Write($"[{where}] {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
    }

    private static void Write(string msg)
    {
        try
        {
            lock (_lock)
            {
                RotateIfNeeded();
                File.AppendAllText(_logPath,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {msg}{Environment.NewLine}",
                    new UTF8Encoding(false));
            }
        }
        catch { }
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
                Write("[LOG] Rotated");
            }
        }
        catch { }
    }
}

public static class CrashDumper
{
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

// =============================== Main Form ===============================
public sealed class MainForm : Form
{
    // Header
    private readonly Label lblTitle = new() { AutoSize = true };
    private readonly PictureBox picLogo = new() { SizeMode = PictureBoxSizeMode.Zoom, Width = 199, Height = 49 };

    // Selector controls
    private readonly ComboBox cmbPorts = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
    private readonly TextBox txtFolder = new() { Width = 480, PlaceholderText = @"Ordner für do_not_delete.txt" };
    private readonly Button btnChooseFolder = new() { Text = "Ordner", Width = 90, Height = 32 };
    private readonly Button btnAdd = new() { Text = "Port hinzufügen", Width = 150, Height = 32 };
    private readonly Button btnRefreshPorts = new() { Text = "Ports aktualisieren", Width = 140, Height = 32 };

    // Log level + log buttons
    private readonly ComboBox cmbLogLevel = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };
    private readonly Button btnOpenLog = new() { Text = "App-Log öffnen", Width = 130, Height = 32 };
    private readonly Button btnClearLog = new() { Text = "App-Log leeren", Width = 130, Height = 32 };

    // Upper table (ports)
    private readonly ListView lv = new()
    {
        Dock = DockStyle.Fill,
        FullRowSelect = true,
        GridLines = true,
        View = View.Details,
        HideSelection = false,
        OwnerDraw = true // für Status-Kreis & Zeilenhöhe & unified selection
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

    // Live table – one row per COM (latest only), COM first col
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

    // Panels
    private readonly FlowLayoutPanel pnlTopPanel;
    private readonly Panel headerPanel;

    // Watchdog timer
    private readonly WinFormsTimer _watchdogTimer = new() { Interval = 2000 }; // 2s

    // Per-port row color & live row map
    private readonly Dictionary<string, Color> _portColorMap = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DataGridViewRow> _liveRowsByPort = new(StringComparer.OrdinalIgnoreCase);

    // Active loggers
    private readonly Dictionary<string, ComLogger> loggers = new();

    // Device change / port snapshot
    private const int WM_DEVICECHANGE = 0x0219;
    private const int DBT_DEVNODES_CHANGED = 0x0007;
    private string[] _lastPortSnapshot = Array.Empty<string>();

    // Row height via SmallImageList
    private readonly ImageList _rowHeightImages = new() { ImageSize = new Size(1, 34) }; // ~34px Höhe

    // Owner-draw: bigger status circle
    private const int StatusCircleDiameter = 22; // noch größer
    private const int StatusCircleMargin = 8;

    // Selection colors
    private readonly Color _selBack = SystemColors.Highlight;
    private readonly Color _selText = SystemColors.HighlightText;

    public MainForm()
    {
        Text = "M81 DataTransfer";
        Width = 1280;
        Height = 900;
        StartPosition = FormStartPosition.CenterScreen;
        Shown += (_, __) => { TrySetDefaultSplit(); };
        Resize += (_, __) => { TrySetDefaultSplit(); };

        // Flicker reduzieren
        this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
        this.UpdateStyles();

        // ===== Selector (TOP)
        pnlTopPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 96,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Padding = new Padding(10)
        };
        var lblPort = new Label { Text = "Port:", AutoSize = true, Padding = new Padding(0, 8, 4, 0) };
        var lblFolder = new Label { Text = "Ordner:", AutoSize = true, Padding = new Padding(12, 8, 4, 0) };
        var lblLogLvl = new Label { Text = "Log-Level:", AutoSize = true, Padding = new Padding(12, 8, 4, 0) };
        pnlTopPanel.Controls.Add(lblPort);
        pnlTopPanel.Controls.Add(cmbPorts);
        pnlTopPanel.Controls.Add(btnRefreshPorts);
        pnlTopPanel.Controls.Add(lblFolder);
        pnlTopPanel.Controls.Add(txtFolder);
        pnlTopPanel.Controls.Add(btnChooseFolder);
        pnlTopPanel.Controls.Add(btnAdd);
        pnlTopPanel.Controls.Add(lblLogLvl);
        pnlTopPanel.Controls.Add(cmbLogLevel);
        pnlTopPanel.Controls.Add(btnOpenLog);
        pnlTopPanel.Controls.Add(btnClearLog);

        // ===== Header (unter selector; dunkler Balken, höher gemacht)
        headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 110, // höher
            Padding = new Padding(20),
            BackColor = Color.FromArgb(40, 40, 40)
        };
        lblTitle.Text = "M81 DataTransfer";
        lblTitle.Font = new Font("Segoe UI", 26, FontStyle.Bold); // etwas größer
        lblTitle.ForeColor = Color.White;

        LoadLogo();
        picLogo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        picLogo.BackColor = Color.Transparent;
        picLogo.Location = new Point(10, 6);

        var headerLeft = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
        headerLeft.Controls.Add(lblTitle);
        lblTitle.Location = new Point(0, 0);

        var headerRight = new Panel { Dock = DockStyle.Right, Width = picLogo.Width + 24, BackColor = Color.Transparent };
        headerRight.Controls.Add(picLogo);

        headerPanel.Controls.Add(headerLeft);
        headerPanel.Controls.Add(headerRight);

        // ===== Upper ListView (Ports)
        lv.SmallImageList = _rowHeightImages; // Row-Höhe erhöhen
        lv.Columns.Add("", 50);               // 0: Status-Kreis
        lv.Columns.Add("Port", 160);          // 1
        lv.Columns.Add("Ordner", 520);        // 2
        lv.Columns.Add("Status", 450);        // 3
        lv.Columns.Add("Age", 80);            // 4
        TryEnableDoubleBuffer(lv);

        // Owner draw events
        lv.DrawColumnHeader += (s, e) => e.DrawDefault = true;
        lv.DrawItem += (s, e) => { /* nichts */ };
        lv.DrawSubItem += Lv_DrawSubItem;

        // ===== Split + Live Grid
        split.Panel1.Controls.Add(lv);

        TryEnableDoubleBuffer(dgvLive);
        // COM zuerst
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "COM", Name = "COM", FillWeight = 80 });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeit", Name = "Time", FillWeight = 120 });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T1", Name = "T1" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T2", Name = "T2" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T3", Name = "T3" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T4", Name = "T4" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T5", Name = "T5" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "T6", Name = "T6" });
        dgvLive.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "RAW", Name = "RAW", FillWeight = 240 });

        // Kein Highlighting im Live-Grid
        dgvLive.DefaultCellStyle.SelectionBackColor = dgvLive.DefaultCellStyle.BackColor;
        dgvLive.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;
        dgvLive.SelectionChanged += (_, __) => { try { dgvLive.ClearSelection(); } catch { } };
        dgvLive.RowsAdded += (_, e) =>
        {
            for (int i = 0; i < e.RowCount; i++)
            {
                var row = dgvLive.Rows[e.RowIndex + i];
                row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;
            }
        };

        split.Panel2.Controls.Add(dgvLive);

        // ===== Buttons unten
        var pnlButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(12)
        };
        pnlButtons.Controls.AddRange(new Control[] { btnStart, btnStop, btnRemove, btnOpenFolder, btnClearLive });

        // ===== Footer
        slVersion.Text = $"Version {Defaults.AppVersion}";
        slNotes.Text = $"Baudrate: {Defaults.FixedBaud} • Dateiname: {Defaults.FixedFileName} • Reconnect nach: {Defaults.ReconnectIdleSeconds}s Idle";
        slState.Text = "Bereit";
        status.Items.AddRange(new ToolStripItem[] { slVersion, slNotes, slState });

        // ===== Controls hinzufügen
        Controls.Add(split);
        Controls.Add(pnlTopPanel);
        Controls.Add(headerPanel);
        Controls.Add(pnlButtons);
        Controls.Add(status);

        // Events
        btnChooseFolder.Click += (_, __) => ChooseFolder();
        btnAdd.Click += (_, __) => AddLoggerFromUi();
        btnRefreshPorts.Click += (_, __) => RefreshPorts();

        btnOpenLog.Click += (_, __) => { try { Process.Start(new ProcessStartInfo { FileName = AppLogger.LogPath, UseShellExecute = true }); } catch (Exception ex) { AppLogger.LogException("OpenLog", ex); } };
        btnClearLog.Click += (_, __) => { AppLogger.ClearLog(); slState.Text = "Log geleert."; };

        cmbLogLevel.Items.AddRange(new object[] { "Error", "Info", "Debug" });
        cmbLogLevel.SelectedIndexChanged += (_, __) =>
        {
            var lvl = cmbLogLevel.SelectedIndex switch
            {
                0 => LogLevel.Error,
                2 => LogLevel.Debug,
                _ => LogLevel.Info
            };
            AppLogger.SetLevel(lvl);
            // Persist
            var s = AppSettings.Load();
            s.LogLevel = lvl;
            AppSettings.Save(s);
        };

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
                // Beim Neustart soll die Tabelle leer sein:
                var s = AppSettings.Load();
                s.DefaultFolder = NormalizeFolder(txtFolder.Text);
                AppSettings.Save(s);
            }
            catch { }
        };

        // Init
        var settings = AppSettings.Load();
        AppLogger.SetLevel(settings.LogLevel);
        cmbLogLevel.SelectedIndex = settings.LogLevel switch
        {
            LogLevel.Error => 0,
            LogLevel.Debug => 2,
            _ => 1
        };

        RefreshPorts();
        txtFolder.Text = NormalizeFolder(settings.DefaultFolder ?? Defaults.BaseFolder);
        UpdateButtons();

        _lastPortSnapshot = SerialPort.GetPortNames().OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();

        // Watchdog
        _watchdogTimer.Tick += (_, __) => WatchdogScan();
        _watchdogTimer.Start();
    }

    // Owner-draw für SubItems (einheitliche Zeilen-Selektion + Status-Kreis)
    private void Lv_DrawSubItem(object? sender, DrawListViewSubItemEventArgs e)
    {
        var item = e.Item;
        var lg = item.Tag as ComLogger;

        // Einheitliche Selektion: nur einmal (in Spalte 0) den gesamten Zeilenbereich füllen
        if (item.Selected && e.ColumnIndex == 0)
        {
            var fullRow = item.Bounds;
            using var sb = new SolidBrush(_selBack);
            e.Graphics.FillRectangle(sb, fullRow);
        }
        else if (!item.Selected)
        {
            // Normaler Hintergrund wird vom System/Owner-Draw pro Subitem schon gefüllt via DrawBackground()
            e.DrawBackground();
        }

        // Spalte 0: Kreis
        if (e.ColumnIndex == 0)
        {
            var bounds = e.Bounds;
            var cx = bounds.Left + StatusCircleMargin + StatusCircleDiameter / 2;
            var cy = bounds.Top + (bounds.Height / 2);
            var r = StatusCircleDiameter / 2;

            var color = Color.Gray; // default
            if (lg != null)
            {
                var open = lg.IsOpen();
                var last = lg.LastFrameUtc;
                var age = last == DateTime.MinValue ? TimeSpan.MaxValue : (DateTime.UtcNow - last);

                if (!open)
                {
                    color = Color.Gray;
                }
                else if (age.TotalSeconds <= Defaults.ReconnectIdleSeconds)
                {
                    color = Color.ForestGreen; // aktiv
                }
                else
                {
                    color = Color.DarkOrange; // offen, aber zu lange keine Daten
                }

                if (lg.HadRecentErrorUtc.AddSeconds(5) > DateTime.UtcNow)
                    color = Color.Red; // kürzlich Fehler
            }

            using var b = new SolidBrush(color);
            using var p = new Pen(Color.Black, 1f);
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            var rect = new Rectangle(cx - r, cy - r, StatusCircleDiameter, StatusCircleDiameter);
            e.Graphics.FillEllipse(b, rect);
            e.Graphics.DrawEllipse(p, rect);
        }
        else
        {
            // Textfarbe je nach Selektion
            var fore = item.Selected ? _selText : e.SubItem.ForeColor;
            TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis;
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.SubItem.Font ?? e.Item.Font, e.Bounds, fore, flags);
        }

        // Fokusrahmen
        if ((e.ItemState & ListViewItemStates.Focused) != 0 && e.ColumnIndex == 0)
            e.DrawFocusRectangle(e.Item.Bounds);
    }

    // ----- WM_DEVICECHANGE: USB Serial Arrival/Removal -----
    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);

        if (m.Msg == WM_DEVICECHANGE && (int)m.WParam == DBT_DEVNODES_CHANGED)
        {
            try
            {
                var now = SerialPort.GetPortNames().OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();
                _lastPortSnapshot = now;

                AppLogger.Debug("WM_DEVICECHANGE: ports now = " + string.Join(",", now));

                // Refresh Combo only
                cmbPorts.Items.Clear();
                cmbPorts.Items.AddRange(now);
                if (cmbPorts.Items.Count > 0 && cmbPorts.SelectedIndex < 0) cmbPorts.SelectedIndex = 0;

                // Laufende Logger nudgen
                foreach (ListViewItem it in lv.Items)
                    if (it.Tag is ComLogger lg && lg.WantsRunning)
                        lg.NudgeEnsure("device_change");
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
            if (split.Height > 0)
                split.SplitterDistance = Math.Max(140, (int)(split.Height * 0.30));
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

        AddLogger(port, folder);
    }

    private void AddLogger(string port, string folder)
    {
        if (loggers.ContainsKey(port))
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

        var logger = new ComLogger(port, cfg); // Id == Port
        logger.StatusChanged += OnLoggerStatus;
        logger.LiveRow += OnLoggerLiveRow;

        loggers[port] = logger;

        // Zeile oben – noch nicht gestartet → Status "Gestoppt"
        var item = new ListViewItem(new[]
        {
            "", // Kreis
            cfg.PortName,
            cfg.FolderPath,
            "Gestoppt",
            "-"
        })
        { Name = port, Tag = logger, UseItemStyleForSubItems = false };

        // Pastell-Hintergrund je Port
        var color = GetSoftColorForPort(cfg.PortName);
        item.BackColor = color;

        lv.Items.Add(item);
        lv.SelectedItems.Clear();
        item.Selected = true;

        slState.Text = $"Port {port} hinzugefügt";
        AppLogger.Info($"Port added: {cfg.PortName}, folder={cfg.FolderPath}");
        UpdateButtons();
    }

    private Color GetSoftColorForPort(string port)
    {
        if (_portColorMap.TryGetValue(port, out var c)) return c;
        int hash = port.Aggregate(17, (a, ch) => unchecked(a * 31 + ch));
        double hue = (hash & 0xFFFF) / (double)0xFFFF * 360.0;
        double sat = 0.35;
        double light = 0.92;
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

    private void WatchdogScan()
    {
        try
        {
            foreach (ListViewItem it in lv.Items)
            {
                if (it.Tag is not ComLogger lg) continue;

                // Age == Zeit seit *letzter gültiger Nachricht* (bleibt über Reconnects bestehen)
                double ageSec = lg.LastFrameUtc == DateTime.MinValue
                    ? double.NaN
                    : (DateTime.UtcNow - lg.LastFrameUtc).TotalSeconds;
                it.SubItems[4].Text = double.IsNaN(ageSec) ? "-" : $"{ageSec:0.0}s";

                // Status-Text
                var statusSub = it.SubItems[3];
                if (lg.IsOpen())
                {
                    if (!double.IsNaN(ageSec) && ageSec > Defaults.ReconnectIdleSeconds)
                    {
                        statusSub.Text = "Offen – keine Daten (Reconnect folgt)";
                        statusSub.ForeColor = Color.DarkOrange;
                        statusSub.Font = new Font(lv.Font, FontStyle.Bold);
                        lg.NudgeEnsure("watchdog_idle");
                    }
                    else
                    {
                        statusSub.Text = "Läuft";
                        statusSub.ForeColor = Color.Green;
                        statusSub.Font = new Font(lv.Font, FontStyle.Bold);
                    }
                }
                else
                {
                    statusSub.Text = lg.WantsRunning ? "Warte auf Gerät …" : "Gestoppt";
                    statusSub.ForeColor = lv.ForeColor;
                    statusSub.Font = lv.Font;
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
            it.SubItems[3].Text = e.StatusText; // Status
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

    private DataGridViewRow EnsureLiveRowForPort(string port)
    {
        if (_liveRowsByPort.TryGetValue(port, out var row)) return row;

        int idx = dgvLive.Rows.Add(port, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), "", "", "", "", "", "", "");
        row = dgvLive.Rows[idx];
        var color = GetSoftColorForPort(port);
        row.DefaultCellStyle.BackColor = color;
        row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor;
        row.DefaultCellStyle.SelectionForeColor = dgvLive.DefaultCellStyle.ForeColor;
        _liveRowsByPort[port] = row;
        btnClearLive.Enabled = dgvLive.Rows.Count > 0;
        return row;
    }

    private void UpdateButtons()
    {
        bool has = lv.SelectedItems.Count == 1;
        btnRemove.Enabled = has;
        btnOpenFolder.Enabled = has;
        btnClearLive.Enabled = dgvLive.Rows.Count > 0;

        if (!has) { btnStart.Enabled = false; btnStop.Enabled = false; return; }

        if (lv.SelectedItems[0].Tag is ComLogger logger)
        {
            btnStart.Enabled = !logger.WantsRunning || !logger.IsOpen();
            btnStop.Enabled = logger.WantsRunning;
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
        if (!logger.WantsRunning) StartSelected(); else StopSelected();
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
        logger.Stop();
    }

    private void RemoveSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        var sel = lv.SelectedItems[0];
        var id = sel.Name; // port
        if (!string.IsNullOrEmpty(id) && loggers.TryGetValue(id, out var logger))
        {
            logger.Dispose();
            loggers.Remove(id);
            lv.Items.RemoveByKey(id);

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

// ========================== COM Logger Engine ===========================
public sealed class ComLogger : IDisposable
{
    public LoggerConfig Config { get; }
    public bool WantsRunning => _desiredRunning;

    public event EventHandler<LoggerStatus>? StatusChanged;
    public event EventHandler<LoggerLive>? LiveRow;

    private SerialPort? _serial;
    private readonly object _serialLock = new();
    private readonly object _fileLock = new();
    private readonly StringBuilder _buf = new();

    private CancellationTokenSource? _ensureCts;
    private Task? _ensureTask;
    private bool _desiredRunning;
    private DateTime _graceUntilUtc = DateTime.MinValue;
    private DateTime _nextReconnectAllowedUtc = DateTime.MinValue;

    // Watchdog info: last frame (UTC) atomic ticks
    private long _lastFrameTicks; // 0 == unset
    public DateTime LastFrameUtc
    {
        get
        {
            long ticks = Interlocked.Read(ref _lastFrameTicks);
            return ticks == 0 ? DateTime.MinValue : new DateTime(ticks, DateTimeKind.Utc);
        }
        private set
        {
            long ticks = (value.Kind == DateTimeKind.Utc ? value : value.ToUniversalTime()).Ticks;
            Interlocked.Exchange(ref _lastFrameTicks, ticks);
        }
    }

    // Last recent error time (for red circle)
    public DateTime HadRecentErrorUtc { get; private set; } = DateTime.MinValue;

    // encoder UTF-8 ohne BOM
    private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

    // Frame: 089 + t1..t5 [±dd.dd] + t6 [±dd.ddddddd] + x1,x2 [±ddd.dddd] + 980 (ohne Punkt)
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

    public string Id { get; } // == Portname

    public ComLogger(string id, LoggerConfig config)
    {
        Id = id;
        Config = config;
    }

    public bool IsOpen()
    {
        lock (_serialLock) return _serial?.IsOpen == true;
    }

    public async Task StartAsync()
    {
        if (_desiredRunning) return;
        _desiredRunning = true;

        StartEnsureLoop();
        _ = Task.Run(() => EnsureOpenOnceAsync("start"));
        RaiseStatus($"Gestartet – {Config.PortName} (warte auf Gerät) …");
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

                    if (!IsOpen())
                    {
                        await EnsureOpenOnceAsync("ensure_loop");
                        await Task.Delay(Defaults.EnsurePollMs, token);
                        continue;
                    }

                    // Nur schließen, wenn wirklich idle ≥ Threshold (Keep-Open bei Aktivität)
                    if (DateTime.UtcNow >= _graceUntilUtc)
                    {
                        var last = LastFrameUtc;
                        var age = last == DateTime.MinValue ? TimeSpan.MaxValue : DateTime.UtcNow - last;
                        if (age.TotalSeconds >= Defaults.ReconnectIdleSeconds)
                        {
                            AppLogger.Info($"Ensure: idle {age.TotalSeconds:0.0}s on {Config.PortName}, closing for reopen");
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

    public void NudgeEnsure(string reason)
    {
        AppLogger.Debug($"NudgeEnsure({reason}) for {Config.PortName}");
        _nextReconnectAllowedUtc = DateTime.MinValue;
        // Ensure-Loop pollt ohnehin; dies beschleunigt den nächsten Versuch.
    }

    private async Task EnsureOpenOnceAsync(string reason)
    {
        try
        {
            if (!_desiredRunning) return;

            var now = DateTime.UtcNow;
            if (now < _nextReconnectAllowedUtc) return; // throttle
            _nextReconnectAllowedUtc = now.AddSeconds(Defaults.ReconnectRequestCooldownSeconds);

            var ports = SerialPort.GetPortNames();
            string? target = null;

            if (ports.Contains(Config.PortName, StringComparer.OrdinalIgnoreCase))
            {
                target = Config.PortName;
            }
            else if (Config.AutoRebind && ports.Length > 0)
            {
                target = ports.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).First();
                if (!string.Equals(target, Config.PortName, StringComparison.OrdinalIgnoreCase))
                    AppLogger.Info($"AutoRebind: {Config.PortName} not found; trying {target}");
            }

            if (target == null)
            {
                RaiseStatus($"Warte auf Gerät – {Config.PortName} …");
                AppLogger.Debug($"EnsureOpen: no suitable COM for {Config.PortName}");
                return;
            }

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

            // Nach Open nur Grace; KEIN Reset von LastFrameUtc (damit Age über Reconnects korrekt bleibt!)
            _graceUntilUtc = DateTime.UtcNow.AddSeconds(Defaults.PostOpenGraceSeconds);

            RaiseStatus($"Läuft – {Config.PortName} @ {Defaults.FixedBaud}.");
            AppLogger.Log($"Serial opened: {Config.PortName} @ {Defaults.FixedBaud} (reason={reason})");
        }
        catch (UnauthorizedAccessException ex)
        {
            HadRecentErrorUtc = DateTime.UtcNow;
            AppLogger.LogException("EnsureOpen(Unauthorized)", ex);
            RaiseStatus("Port belegt/kein Zugriff – erneuter Versuch …");
        }
        catch (IOException ex)
        {
            HadRecentErrorUtc = DateTime.UtcNow;
            AppLogger.LogException("EnsureOpen(IO)", ex);
            RaiseStatus("I/O-Fehler – erneuter Versuch …");
        }
        catch (Exception ex)
        {
            HadRecentErrorUtc = DateTime.UtcNow;
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
            HadRecentErrorUtc = DateTime.UtcNow;
            AppLogger.LogException("SafeClose", ex);
        }
    }

    private void SerialOnError(object? s, SerialErrorReceivedEventArgs e)
    {
        HadRecentErrorUtc = DateTime.UtcNow;
        AppLogger.Log($"Serial error on {Config.PortName}: {e.EventType}");
        SafeClose("serial_error_" + e.EventType);
        NudgeEnsure("serial_error");
    }

    private void SerialOnPinChanged(object? s, SerialPinChangedEventArgs e)
    {
        AppLogger.Info($"Pin changed on {Config.PortName}: {e.EventType}");
        if (e.EventType == SerialPinChange.Break || e.EventType == SerialPinChange.CDChanged ||
            e.EventType == SerialPinChange.DsrChanged || e.EventType == SerialPinChange.CtsChanged)
        {
            HadRecentErrorUtc = DateTime.UtcNow;
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
            HadRecentErrorUtc = DateTime.UtcNow;
            AppLogger.LogException("DataReceived", ex);
            SafeClose("data_received_exception");
            NudgeEnsure("data_received_exception");
        }
    }

    private static string SafePreview(string s)
    {
        if (s == null) return string.Empty;
        s = s.Replace("\r", "\\r").Replace("\n", "\\n");
        if (s.Length <= 120) return s;
        return s.Substring(0, 120) + "...";
    }

    // Frame extraction with terminator "980" and optional CR/LF afterwards
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

            // remove consumed incl. evtl. CR/LF
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

            string[] tempsFormatted = tempsDouble
                .Select(v => v.HasValue
                    ? v.Value.ToString("+0.0000;-0.0000", CultureInfo.InvariantCulture)
                    : "+0.0000")
                .ToArray();

            string fileLine = string.Join(",", tempsFormatted);

            DateTime ts = DateTime.Now;

            // mark activity ONLY when valid frame; do not change on open/reconnect
            LastFrameUtc = DateTime.UtcNow;

            // Live + debug
            AppLogger.Debug($"Frame OK {Config.PortName}: temps={fileLine}, rawLen={candidate.Length}");
            LiveRow?.Invoke(this, new LoggerLive(Id, Config.PortName, ts, tempsFormatted, candidate));

            // write last value atomar
            WriteLastValueSafe(fileLine);
        }
    }

    private static double? ParseDoubleInvariant(string s)
    {
        return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var v)
            ? v : (double?)null;
    }

    // Disk guard + atomic replace (UTF-8 ohne BOM)
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

// ======================== Modelle & Persistenz ===========================
public sealed record LoggerConfig
{
    public string PortName { get; set; } = "COM1";
    public string FolderPath { get; set; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ComPortLogger");
    public bool AutoRebind { get; set; } = true;

    public string OutputPath => Path.Combine(FolderPath, Defaults.FixedFileName);
}

public sealed record LoggerStatus(string Id, string StatusText, string? LastError);
public sealed record LoggerLive(string Id, string Port, DateTime Ts, string[] Temps, string Raw);

public sealed class AppSettings
{
    public string? DefaultFolder { get; set; }
    public LogLevel LogLevel { get; set; } = LogLevel.Info;

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
        var def = new AppSettings { DefaultFolder = Defaults.BaseFolder, LogLevel = LogLevel.Info };
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
