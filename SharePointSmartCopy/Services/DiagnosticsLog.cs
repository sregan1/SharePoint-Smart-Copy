using System.Collections.Concurrent;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Channels;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Log files, crash capture and the "Export diagnostics" bundle.
//
// Logs live in %LocalAppData%\SharePointSmartCopy\Logs:
//   activity-*.log  one per copy/verify run (written by ActivityLogWriter)
//   crash-*.log     an exception that would otherwise vanish (see InstallGlobalHandlers)
public static class DiagnosticsLog
{
    public static string LogDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SharePointSmartCopy", "Logs");

    private static readonly ConcurrentDictionary<ActivityLogWriter, byte> OpenWriters = new();
    private static int _handlersInstalled;

    // Call once at startup.
    public static void Initialize()
    {
        PruneOldLogs();
        InstallGlobalHandlers();
    }

    // Keep two months of history, and never more than 200 files, so the folder can't grow forever
    // (it was never pruned — every run added a file and nothing removed them).
    private static void PruneOldLogs()
    {
        try
        {
            if (!Directory.Exists(LogDirectory)) return;
            var files = new DirectoryInfo(LogDirectory).GetFiles("*.log")
                .OrderByDescending(f => f.LastWriteTimeUtc).ToList();
            var cutoff = DateTime.UtcNow.AddDays(-60);
            for (int i = 0; i < files.Count; i++)
            {
                if (i >= 200 || files[i].LastWriteTimeUtc < cutoff)
                    try { files[i].Delete(); } catch { /* in use or locked — leave it */ }
            }
        }
        catch { /* pruning is housekeeping; never block startup */ }
    }

    // The dispatcher handler in App already covers UI-thread exceptions. These two cover the rest:
    // an exception on a background thread (process is about to die) and a faulted Task nobody awaited
    // (silently swallowed before — the exact class of failure that was hardest to diagnose).
    private static void InstallGlobalHandlers()
    {
        if (Interlocked.Exchange(ref _handlersInstalled, 1) == 1) return;

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            WriteCrash(e.IsTerminating ? "Unhandled exception (process terminating)" : "Unhandled exception",
                e.ExceptionObject as Exception);
            FlushAll();
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            WriteCrash("Unobserved task exception", e.Exception);
            e.SetObserved();
        };
    }

    // Writes one crash-*.log. Best-effort — logging must never throw.
    public static void WriteCrash(string what, Exception? ex)
    {
        try
        {
            Directory.CreateDirectory(LogDirectory);
            var path = Path.Combine(LogDirectory, $"crash-{DateTime.Now:yyyyMMdd-HHmmss-fff}.log");
            var sb = new StringBuilder();
            AppendHeader(sb);
            sb.AppendLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  {what}");
            sb.AppendLine(ex?.ToString() ?? "(no exception object)");
            File.WriteAllText(path, sb.ToString());
        }
        catch { /* nothing more we can do */ }
    }

    // Opens a log file for one run. Header lines identify the build so a log is useful on its own.
    public static ActivityLogWriter OpenActivityLog(string path)
    {
        var header = new StringBuilder();
        AppendHeader(header);
        var writer = new ActivityLogWriter(path, header.ToString());
        OpenWriters[writer] = 0;
        return writer;
    }

    internal static void Closed(ActivityLogWriter writer) => OpenWriters.TryRemove(writer, out _);

    // Called from App.OnExit before Environment.Exit(0), which would otherwise cut the writers off
    // with lines still queued.
    public static void FlushAll()
    {
        foreach (var writer in OpenWriters.Keys.ToList())
            try { writer.Dispose(); } catch { }
    }

    private static void AppendHeader(StringBuilder sb)
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        sb.AppendLine($"# SharePoint Smart Copy {version?.Major}.{version?.Minor}.{version?.Build}");
        sb.AppendLine($"# {Environment.OSVersion} | .NET {Environment.Version} | culture {Localization.Loc.Culture.Name}");
        // Messages are written in the user's display language (they are formatted at the call site, so
        // they can't be re-rendered in English here) — the culture above says which language to expect.
    }

    // ── Export ────────────────────────────────────────────────────────────────

    // Builds a zip a user can attach to a bug report: recent logs, crash logs, system info and the
    // settings with anything identifying removed. It deliberately does NOT include copy reports or
    // history — those contain file names and site URLs from the user's tenant.
    public static void ExportDiagnostics(string zipPath, AppSettings settings)
    {
        if (File.Exists(zipPath)) File.Delete(zipPath);
        using var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create);

        if (Directory.Exists(LogDirectory))
        {
            foreach (var f in new DirectoryInfo(LogDirectory).GetFiles("*.log")
                         .OrderByDescending(f => f.LastWriteTimeUtc).Take(10))
            {
                try
                {
                    // Open with FileShare.ReadWrite: the current run's log is still being written.
                    using var src = new FileStream(f.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    var entry = zip.CreateEntry($"logs/{f.Name}", CompressionLevel.Optimal);
                    using var dst = entry.Open();
                    src.CopyTo(dst);
                }
                catch { /* skip a file that can't be read */ }
            }
        }

        var info = new StringBuilder();
        AppendHeader(info);
        info.AppendLine($"Processors: {Environment.ProcessorCount}");
        info.AppendLine($"Working set: {Environment.WorkingSet / (1024 * 1024)} MB");
        info.AppendLine($"Exported: {DateTime.Now:yyyy-MM-dd HH:mm:ss zzz}");
        WriteText(zip, "system.txt", info.ToString());
        WriteText(zip, "settings.redacted.json", RedactedSettingsJson(settings));
    }

    private static void WriteText(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var w = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        w.Write(content);
    }

    // Settings with tenant/app IDs shortened, and site URLs, domains and identity overrides removed
    // (they identify the user's organization and its people). Option flags are kept — they are what
    // explains behavior.
    internal static string RedactedSettingsJson(AppSettings settings)
    {
        var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions
        {
            Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
        });
        var root = JsonNode.Parse(json)!.AsObject();

        static string Mask(string? v) => string.IsNullOrEmpty(v) ? "" : (v.Length <= 4 ? "****" : v[..4] + "…");

        if (root["Registrations"] is JsonArray regs)
            foreach (var r in regs.OfType<JsonObject>())
            {
                r["Name"]     = "(redacted)";
                r["ClientId"] = Mask(r["ClientId"]?.GetValue<string>());
                r["TenantId"] = Mask(r["TenantId"]?.GetValue<string>());
            }
        root["SourceUrl"] = string.IsNullOrEmpty(settings.SourceUrl) ? "" : "(redacted)";
        root["TargetUrl"] = string.IsNullOrEmpty(settings.TargetUrl) ? "" : "(redacted)";
        if (root["IdentityMapping"] is JsonObject map)
        {
            map["SourceDomain"] = string.IsNullOrEmpty(settings.IdentityMapping.SourceDomain) ? "" : "(redacted)";
            map["TargetDomain"] = string.IsNullOrEmpty(settings.IdentityMapping.TargetDomain) ? "" : "(redacted)";
            map["Overrides"]    = string.IsNullOrEmpty(settings.IdentityMapping.Overrides)
                ? "" : $"({settings.IdentityMapping.Overrides.Split('\n').Length} line(s) redacted)";
        }
        return root.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }
}

// Writes log lines from a background task so a copy run never waits on disk, and keeps the file
// open (it used to be re-opened and closed for every single line).
public sealed class ActivityLogWriter : IDisposable
{
    private readonly Channel<string> _lines = Channel.CreateUnbounded<string>(
        new UnboundedChannelOptions { SingleReader = true });
    private readonly Task _pump;
    private int _disposed;

    internal ActivityLogWriter(string path, string header)
    {
        _pump = Task.Run(async () =>
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                // FileShare.Read: the user can open the log while the run is still going.
                using var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read);
                using var writer = new StreamWriter(stream, new UTF8Encoding(false));
                writer.Write(header);
                writer.Flush();
                await foreach (var line in _lines.Reader.ReadAllAsync())
                {
                    writer.WriteLine(line);
                    // Flush whenever the queue drains, so a crash loses at most the last burst.
                    if (_lines.Reader.Count == 0) writer.Flush();
                }
                writer.Flush();
            }
            catch { /* best effort: a log failure must never interrupt a copy */ }
        });
    }

    public void Write(string line) => _lines.Writer.TryWrite(line);

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) == 1) return;
        _lines.Writer.TryComplete();
        try { _pump.Wait(TimeSpan.FromSeconds(3)); } catch { }
        DiagnosticsLog.Closed(this);
    }
}
