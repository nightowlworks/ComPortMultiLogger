using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}

public sealed class MainForm : Form
{
    private readonly ComboBox cmbPorts = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
    private readonly NumericUpDown numBaud = new() { Minimum = 1200, Maximum = 921600, Increment = 1200, Value = 115200, Width = 100 };
    private readonly TextBox txtFile = new() { Width = 320, PlaceholderText = "Pfad zur Ausgabedatei (z.B. C:\\Logs\\Port1.txt)" };
    private readonly Button btnBrowse = new() { Text = "Datei wählen", Width = 100 };
    private readonly NumericUpDown numRotateMb = new() { Minimum = 1, Maximum = 2048, Value = 100, Width = 80 };
    private readonly Button btnAdd = new() { Text = "Logger hinzufügen", Width = 150 };
    private readonly Button btnRefreshPorts = new() { Text = "Ports aktualisieren", Width = 140 };

    private readonly ListView lv = new()
    {
        Dock = DockStyle.Fill,
        FullRowSelect = true,
        GridLines = true,
        View = View.Details
    };

    private readonly Button btnStart = new() { Text = "Start", Enabled = false, Width = 90 };
    private readonly Button btnStop = new() { Text = "Stop", Enabled = false, Width = 90 };
    private readonly Button btnRemove = new() { Text = "Entfernen", Enabled = false, Width = 100 };
    private readonly Button btnOpenFolder = new() { Text = "Log-Ordner öffnen", Enabled = false, Width = 140 };
    private readonly Button btnClear = new() { Text = "Anzeige leeren", Enabled = false, Width = 120 };

    private readonly TextBox txtTail = new()
    {
        Dock = DockStyle.Bottom,
        Multiline = true,
        ReadOnly = true,
        ScrollBars = ScrollBars.Both,
        WordWrap = false,
        Height = 220,
        Font = new System.Drawing.Font("Consolas", 10)
    };

    private readonly Label lblStatus = new() { Text = "Bereit.", AutoSize = true, Padding = new Padding(8, 6, 0, 0) };
    private readonly Dictionary<string, ComLogger> loggers = new(); // key = id

    private AppSettings settings;

    public MainForm()
    {
        Text = "Multi-COM Datenlogger (ereignisgesteuert, robust)";
        Width = 1100;
        Height = 760;
        StartPosition = FormStartPosition.CenterScreen;

        var pnlTop = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 96,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Padding = new Padding(8)
        };

        pnlTop.Controls.Add(new Label { Text = "Port:", AutoSize = true, Padding = new Padding(0, 8, 4, 0) });
        pnlTop.Controls.Add(cmbPorts);
        pnlTop.Controls.Add(btnRefreshPorts);
        pnlTop.Controls.Add(new Label { Text = "Baud:", AutoSize = true, Padding = new Padding(8, 8, 4, 0) });
        pnlTop.Controls.Add(numBaud);
        pnlTop.Controls.Add(new Label { Text = "Datei:", AutoSize = true, Padding = new Padding(8, 8, 4, 0) });
        pnlTop.Controls.Add(txtFile);
        pnlTop.Controls.Add(btnBrowse);
        pnlTop.Controls.Add(new Label { Text = "Rotation (MB):", AutoSize = true, Padding = new Padding(8, 8, 4, 0) });
        pnlTop.Controls.Add(numRotateMb);
        pnlTop.Controls.Add(btnAdd);
        pnlTop.Controls.Add(lblStatus);

        lv.Columns.Add("ID", 70);
        lv.Columns.Add("Port", 90);
        lv.Columns.Add("Baud", 70);
        lv.Columns.Add("Datei", 360);
        lv.Columns.Add("Rotation MB", 90);
        lv.Columns.Add("Status", 180);
        lv.Columns.Add("Rate (Zeilen/s)", 120);
        lv.Columns.Add("Letzter Fehler", 200);

        var pnlButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 48,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(8)
        };
        pnlButtons.Controls.Add(btnStart);
        pnlButtons.Controls.Add(btnStop);
        pnlButtons.Controls.Add(btnRemove);
        pnlButtons.Controls.Add(btnOpenFolder);
        pnlButtons.Controls.Add(btnClear);

        Controls.Add(lv);
        Controls.Add(pnlTop);
        Controls.Add(pnlButtons);
        Controls.Add(txtTail);

        btnBrowse.Click += (_, __) => ChooseFile();
        btnAdd.Click += (_, __) => AddLoggerFromUi();
        btnRefreshPorts.Click += (_, __) => RefreshPorts();

        lv.SelectedIndexChanged += (_, __) => UpdateButtonsForSelection();
        lv.DoubleClick += (_, __) => StartOrStopSelected();

        btnStart.Click += (_, __) => StartSelected();
        btnStop.Click += (_, __) => StopSelected();
        btnRemove.Click += (_, __) => RemoveSelected();
        btnOpenFolder.Click += (_, __) => OpenFolderSelected();
        btnClear.Click += (_, __) => { txtTail.Clear(); };

        FormClosing += (_, __) =>
        {
            foreach (var lg in loggers.Values) lg.Dispose();
            AppSettings.Save(CaptureSettings());
        };

        // load settings + init
        settings = AppSettings.Load();
        RefreshPorts();
        RestoreSettings(settings);
        Status("Bereit.");
    }

    private void ChooseFile()
    {
        using var dlg = new SaveFileDialog
        {
            Title = "Ausgabedatei wählen",
            Filter = "Textdatei (*.txt)|*.txt|Alle Dateien (*.*)|*.*",
            FileName = $"COM_{(cmbPorts.SelectedItem as string) ?? "Port"}.txt",
            CheckPathExists = true,
            OverwritePrompt = false
        };
        if (dlg.ShowDialog(this) == DialogResult.OK)
            txtFile.Text = dlg.FileName;
    }

    private void RefreshPorts()
    {
        try
        {
            var ports = SerialPort.GetPortNames().OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToArray();
            cmbPorts.Items.Clear();
            cmbPorts.Items.AddRange(ports);
            if (cmbPorts.Items.Count > 0 && cmbPorts.SelectedIndex < 0) cmbPorts.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            Status("Fehler beim Abfragen der Ports.");
            Debug.WriteLine(ex);
        }
    }

    private void AddLoggerFromUi()
    {
        var port = cmbPorts.SelectedItem as string;
        if (string.IsNullOrWhiteSpace(port)) { Status("Bitte Port wählen."); return; }

        var baud = (int)numBaud.Value;

        string file = txtFile.Text.Trim();
        if (string.IsNullOrWhiteSpace(file))
        {
            // Standard: Dokumente\ComPortMultiLogger\<Port>.txt
            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var folder = Path.Combine(docs, "ComPortMultiLogger");
            Directory.CreateDirectory(folder);
            file = Path.Combine(folder, $"{port}.txt");
            txtFile.Text = file;
        }

        long rotateBytes = (long)numRotateMb.Value * 1024L * 1024L;

        try
        {
            var cfg = new LoggerConfig
            {
                PortName = port,
                BaudRate = baud,
                OutputPath = file,
                RotateBytes = rotateBytes
            };

            var id = Guid.NewGuid().ToString("N").Substring(0, 8);
            var logger = new ComLogger(id, cfg);
            logger.StatusChanged += OnLoggerStatus;
            logger.LinesUpdated += OnLoggerLines;
            logger.MetricsUpdated += OnLoggerMetrics;

            loggers.Add(id, logger);

            var item = new ListViewItem(new[]
            {
                id, cfg.PortName, cfg.BaudRate.ToString(), cfg.OutputPath,
                (cfg.RotateBytes / (1024*1024)).ToString(),
                "Gestoppt", "0", "-"
            })
            { Name = id, Tag = logger };

            lv.Items.Add(item);
            lv.SelectedItems.Clear();
            item.Selected = true;

            Status($"Logger {id} hinzugefügt.");
        }
        catch (Exception ex)
        {
            Status($"Fehler beim Hinzufügen: {ex.Message}");
        }
    }

    private void OnLoggerStatus(object? sender, LoggerStatus e)
    {
        if (IsDisposed || Disposing) return;
        BeginInvoke(new Action(() =>
        {
            if (!lv.Items.ContainsKey(e.Id)) return;
            var it = lv.Items[e.Id];
            it.SubItems[5].Text = e.StatusText;
            it.SubItems[7].Text = e.LastError ?? "-";
            if (lv.SelectedItems.Count == 1 && lv.SelectedItems[0].Name == e.Id)
                Status(e.StatusText);
        }));
    }

    private void OnLoggerMetrics(object? sender, LoggerMetrics e)
    {
        if (IsDisposed || Disposing) return;
        BeginInvoke(new Action(() =>
        {
            if (!lv.Items.ContainsKey(e.Id)) return;
            var it = lv.Items[e.Id];
            it.SubItems[6].Text = e.LinesPerSecond.ToString();
        }));
    }

    private void OnLoggerLines(object? sender, LoggerLines e)
    {
        if (IsDisposed || Disposing) return;
        BeginInvoke(new Action(() =>
        {
            if (lv.SelectedItems.Count != 1) return;
            var sel = lv.SelectedItems[0];
            if (sel.Name != e.Id) return;

            // zeige die letzten 100 Zeilen
            var sb = new StringBuilder();
            foreach (var line in e.Last100)
                sb.AppendLine(line);
            txtTail.Text = sb.ToString();
            txtTail.SelectionStart = txtTail.TextLength;
            txtTail.ScrollToCaret();
        }));
    }

    private void UpdateButtonsForSelection()
    {
        bool has = lv.SelectedItems.Count == 1;
        btnStart.Enabled = has;
        btnStop.Enabled = has;
        btnRemove.Enabled = has;
        btnOpenFolder.Enabled = has;
        btnClear.Enabled = has;

        if (!has) return;
        var logger = (ComLogger)lv.SelectedItems[0].Tag!;
        btnStart.Enabled = !logger.IsRunning;
        btnStop.Enabled = logger.IsRunning;
    }

    private void StartOrStopSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        var logger = (ComLogger)lv.SelectedItems[0].Tag!;
        if (!logger.IsRunning) StartSelected(); else StopSelected();
    }

    private void StartSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        var logger = (ComLogger)lv.SelectedItems[0].Tag!;
        _ = logger.StartAsync();
        UpdateButtonsForSelection();
    }

    private void StopSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        var logger = (ComLogger)lv.SelectedItems[0].Tag!;
        logger.Stop();
        UpdateButtonsForSelection();
    }

    private void RemoveSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        var id = lv.SelectedItems[0].Name;
        if (!loggers.TryGetValue(id, out var logger)) return;

        logger.Dispose();
        loggers.Remove(id);
        lv.Items.RemoveByKey(id);
        txtTail.Clear();
        Status($"Logger {id} entfernt.");
    }

    private void OpenFolderSelected()
    {
        if (lv.SelectedItems.Count != 1) return;
        var logger = (ComLogger)lv.SelectedItems[0].Tag!;
        try
        {
            var folder = Path.GetDirectoryName(logger.Config.OutputPath)!;
            Directory.CreateDirectory(folder);
            Process.Start(new ProcessStartInfo { FileName = folder, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            Status($"Ordner öffnen fehlgeschlagen: {ex.Message}");
        }
    }

    private void Status(string msg) => lblStatus.Text = msg;

    // --- Persistenz ---

    private void RestoreSettings(AppSettings s)
    {
        // Logger wiederherstellen
        foreach (var cfg in s.Loggers)
        {
            var id = Guid.NewGuid().ToString("N").Substring(0, 8);
            var logger = new ComLogger(id, cfg);
            logger.StatusChanged += OnLoggerStatus;
            logger.LinesUpdated += OnLoggerLines;
            logger.MetricsUpdated += OnLoggerMetrics;

            loggers[id] = logger;

            var item = new ListViewItem(new[]
            {
                id, cfg.PortName, cfg.BaudRate.ToString(), cfg.OutputPath,
                (cfg.RotateBytes/(1024*1024)).ToString(),
                "Gestoppt", "0", "-"
            })
            { Name = id, Tag = logger };

            lv.Items.Add(item);
        }

        // UI Defaults
        if (!string.IsNullOrEmpty(s.LastOutputFolder) && Directory.Exists(s.LastOutputFolder))
        {
            txtFile.Text = Path.Combine(s.LastOutputFolder, "Port.txt");
        }
        if (s.DefaultBaudRate >= numBaud.Minimum && s.DefaultBaudRate <= numBaud.Maximum)
            numBaud.Value = s.DefaultBaudRate;
    }

    private AppSettings CaptureSettings()
    {
        var s = new AppSettings
        {
            DefaultBaudRate = (int)numBaud.Value,
            LastOutputFolder = SafeFolderFromPath(txtFile.Text),
            Loggers = loggers.Values.Select(l => l.Config).ToList()
        };
        return s;
    }

    private static string? SafeFolderFromPath(string path)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            var dir = Path.GetDirectoryName(path);
            return string.IsNullOrWhiteSpace(dir) ? null : dir;
        }
        catch { return null; }
    }
}

// ===== Logger-Engine =====

public sealed class ComLogger : IDisposable
{
    public LoggerConfig Config { get; }
    public bool IsRunning => _serial != null && _serial.IsOpen;

    // events to UI
    public event EventHandler<LoggerStatus>? StatusChanged;
    public event EventHandler<LoggerLines>? LinesUpdated;
    public event EventHandler<LoggerMetrics>? MetricsUpdated;

    private SerialPort? _serial;
    private readonly object _serialLock = new();
    private readonly object _fileLock = new();

    private StreamWriter? _writer;
    private StreamWriter? _errWriter;
    private long _currentSizeBytes = 0;

    private readonly ConcurrentQueue<string> _last100 = new();
    private readonly StringBuilder _lineBuffer = new(); // holds partial line
    private string? _lastError;
    private CancellationTokenSource? _cts;
    private int _reconnectAttempt;
    private readonly System.Windows.Forms.Timer _rateTimer = new() { Interval = 1000 };
    private int _linesThisSecond = 0;

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
            EnsureWriters();
            OpenSerial();
            _cts = new CancellationTokenSource();
            _reconnectAttempt = 0;

            RaiseStatus($"Läuft – {Config.PortName} @ {Config.BaudRate}.");
        }
        catch (Exception ex)
        {
            LogError("Start", ex);
            RaiseStatus("Start fehlgeschlagen. Reconnect wird versucht …");
            await ReconnectLoopAsync();
        }
    }

    public void Stop()
    {
        try
        {
            _cts?.Cancel();
            _rateTimer.Stop();
        }
        catch { /* ignore */ }

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
        catch { /* ignore */ }

        try
        {
            lock (_fileLock)
            {
                _writer?.Dispose();
                _errWriter?.Dispose();
                _writer = null;
                _errWriter = null;
            }
        }
        catch { /* ignore */ }

        RaiseStatus("Gestoppt.");
    }

    public void Dispose() => Stop();

    private void OpenSerial()
    {
        lock (_serialLock)
        {
            if (_serial != null) return;

            var sp = new SerialPort(Config.PortName, Config.BaudRate)
            {
                Parity = Parity.None,
                DataBits = 8,
                StopBits = StopBits.One,
                Handshake = Handshake.None,
                ReadTimeout = 500,
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
            int delay = Math.Min(30000, (int)Math.Pow(2, Math.Min(6, _reconnectAttempt)) * 250);
            RaiseStatus($"Reconnect-Versuch {_reconnectAttempt} in {delay} ms …");
            try { await Task.Delay(delay, _cts!.Token); } catch { return; }

            try
            {
                EnsureWriters();
                OpenSerial();
                _reconnectAttempt = 0;
                RaiseStatus($"Wieder verbunden – {Config.PortName} @ {Config.BaudRate}.");
                return;
            }
            catch (Exception ex)
            {
                LogError("Reconnect", ex);
            }
        }
    }

    private void SerialOnDataReceived(object? sender, SerialDataReceivedEventArgs e)
    {
        try
        {
            SerialPort? s;
            lock (_serialLock) s = _serial;
            if (s == null || !s.IsOpen) return;

            var chunk = s.ReadExisting();
            if (string.IsNullOrEmpty(chunk)) return;

            lock (_lineBuffer)
            {
                _lineBuffer.Append(chunk);
                // normalize newlines and extract full lines
                string text = _lineBuffer.ToString().Replace("\r\n", "\n").Replace("\r", "\n");
                int lastNl = text.LastIndexOf('\n');
                if (lastNl >= 0)
                {
                    string full = text.Substring(0, lastNl);
                    string remainder = text.Substring(lastNl + 1);
                    _lineBuffer.Clear();
                    _lineBuffer.Append(remainder);

                    var lines = full.Split('\n');
                    foreach (var ln in lines)
                    {
                        var line = ln.TrimEnd();
                        if (line.Length == 0) continue;

                        var ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        var composed = $"{ts} | {line}";
                        AppendToQueue(composed);
                        WriteLineSafe(composed);

                        Interlocked.Increment(ref _linesThisSecond);
                    }

                    PushTail();
                }
            }
        }
        catch (Exception ex)
        {
            LogError("DataReceived", ex);
            // harte Fehler → Reconnect
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
                catch { /* ignore */ }
                await ReconnectLoopAsync();
            });
        }
    }

    private void AppendToQueue(string line)
    {
        _last100.Enqueue(line);
        while (_last100.Count > 100 && _last100.TryDequeue(out _)) { }
    }

    private void PushTail()
    {
        LinesUpdated?.Invoke(this, new LoggerLines(Id, _last100.ToArray()));
    }

    private void EnsureWriters()
    {
        lock (_fileLock)
        {
            string path = Config.OutputPath;
            string? dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir!);

            // open writer in append mode, create if missing
            _writer?.Dispose();
            _writer = new StreamWriter(new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false))
            {
                AutoFlush = true
            };
            _currentSizeBytes = new FileInfo(path).Length;

            var errPath = Path.Combine(dir ?? ".", Path.GetFileNameWithoutExtension(path) + ".errors.txt");
            _errWriter?.Dispose();
            _errWriter = new StreamWriter(new FileStream(errPath, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false))
            {
                AutoFlush = true
            };
        }
    }

    private void RotateIfNeeded()
    {
        if (_currentSizeBytes < Config.RotateBytes) return;

        lock (_fileLock)
        {
            try
            {
                string path = Config.OutputPath;
                string? dir = Path.GetDirectoryName(path);
                string name = Path.GetFileNameWithoutExtension(path);
                string ext = Path.GetExtension(path);

                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string rolled = Path.Combine(dir ?? ".", $"{name}.{stamp}{ext}");

                _writer?.Flush();
                _writer?.Dispose();
                _writer = null;

                // Move current file to rolled name
                if (File.Exists(path))
                {
                    File.Move(path, rolled, overwrite: false);
                }

                // open fresh file
                _writer = new StreamWriter(new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false))
                { AutoFlush = true };

                _currentSizeBytes = 0;

                // auch Fehlerdatei rotieren, damit sie nicht riesig wird (gleiche Grenze)
                var errPath = Path.Combine(dir ?? ".", $"{name}.errors.txt");
                if (File.Exists(errPath) && new FileInfo(errPath).Length >= Config.RotateBytes)
                {
                    var eRolled = Path.Combine(dir ?? ".", $"{name}.errors.{stamp}.txt");
                    _errWriter?.Flush();
                    _errWriter?.Dispose();
                    _errWriter = null;
                    File.Move(errPath, eRolled, overwrite: false);
                    _errWriter = new StreamWriter(new FileStream(errPath, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false))
                    { AutoFlush = true };
                }
            }
            catch (Exception ex)
            {
                LogError("Rotate", ex);
            }
        }
    }

    private void WriteLineSafe(string line)
    {
        try
        {
            lock (_fileLock)
            {
                _writer ??= new StreamWriter(new FileStream(Config.OutputPath, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false))
                { AutoFlush = true };
                _writer.WriteLine(line);
                _currentSizeBytes += Encoding.UTF8.GetByteCount(line) + Environment.NewLine.Length;
            }
            RotateIfNeeded();
        }
        catch (Exception ex)
        {
            LogError("Write", ex);
        }
    }

    private void LogError(string where, Exception ex)
    {
        try
        {
            var msg = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{where}] {ex.GetType().Name}: {ex.Message}";
            lock (_fileLock)
            {
                _errWriter ??= new StreamWriter(new FileStream(Path.ChangeExtension(Config.OutputPath, ".errors.txt"), FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false))
                { AutoFlush = true };
                _errWriter.WriteLine(msg);
            }
            _lastError = ex.Message;
            RaiseStatus($"Fehler in {where}: {ex.Message}");
        }
        catch { /* ignore */ }
    }

    private void RaiseStatus(string text)
    {
        StatusChanged?.Invoke(this, new LoggerStatus(Id, text, _lastError));
    }
}

// ===== Modelle & Persistenz =====

public sealed record LoggerConfig
{
    public string PortName { get; init; } = "COM1";
    public int BaudRate { get; init; } = 115200;
    public string OutputPath { get; init; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ComPortMultiLogger", "COM1.txt");
    public long RotateBytes { get; init; } = 100L * 1024L * 1024L; // 100 MB
}

public sealed record LoggerStatus(string Id, string StatusText, string? LastError);
public sealed record LoggerLines(string Id, string[] Last100);
public sealed record LoggerMetrics(string Id, int LinesPerSecond);

public sealed class AppSettings
{
    public List<LoggerConfig> Loggers { get; set; } = new();
    public int DefaultBaudRate { get; set; } = 115200;
    public string? LastOutputFolder { get; set; }

    private static string AppDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ComPortMultiLogger");
    private static string ConfigPath => Path.Combine(AppDir, "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath, Encoding.UTF8);
                var s = JsonSerializer.Deserialize<AppSettings>(json);
                if (s != null) return s;
            }
        }
        catch { /* ignore */ }
        Directory.CreateDirectory(AppDir);
        var def = new AppSettings();
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
        catch { /* ignore */ }
    }
}
