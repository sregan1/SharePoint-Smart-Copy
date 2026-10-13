using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Crash journal for a copy run.
//
// A run's results used to be saved only when it FINISHED, so a crash, a kill or a power loss
// mid-run left no History entry, no CSV and nothing to Verify — hours of copying with no record.
// While a run is active this appends every finished row to a journal file (one JSON object per
// line) so the results survive. On the next launch, RecoverInterrupted() turns any journal that was
// never completed into a History entry marked "Interrupted", which can then be exported, verified or
// resumed like any other run. A run that ends normally deletes its journal (the normal report
// supersedes it).
//
// File format (JSONL):  line 1 = the run's header (a SavedReport with no Items), then one
// SavedReportItem per line.
public sealed class RunJournal : IDisposable
{
    private static readonly string JournalDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SharePointSmartCopy", "Journal");

    private static readonly JsonSerializerOptions Options = new()
    {
        Converters = { new JsonStringEnumConverter() },
    };

    private readonly string _path;
    private StreamWriter? _writer;

    private RunJournal(string path, StreamWriter writer)
    {
        _path   = path;
        _writer = writer;
    }

    // Starts a journal for a run. Returns null if the journal can't be created — journaling is a
    // safety net and must never stop a copy from running.
    public static RunJournal? Start(SavedReport header)
    {
        try
        {
            Directory.CreateDirectory(JournalDir);
            var path = Path.Combine(JournalDir, $"run_{header.Id}.jsonl");
            // FileShare.Read so RecoverInterrupted (or a user) can inspect it while the run is live.
            var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read);
            var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.WriteLine(JsonSerializer.Serialize(header, Options));
            writer.Flush();
            return new RunJournal(path, writer);
        }
        catch { return null; }
    }

    // Appends finished rows. Flushed every call so a crash loses at most the rows since the last tick.
    public void Append(IEnumerable<SavedReportItem> items)
    {
        var writer = _writer;
        if (writer == null) return;
        try
        {
            foreach (var item in items)
                writer.WriteLine(JsonSerializer.Serialize(item, Options));
            writer.Flush();
        }
        catch { /* best effort: journaling must never interrupt the copy */ }
    }

    // The run finished normally and its full report was saved: the journal has done its job.
    public void Complete()
    {
        Dispose();
        try { File.Delete(_path); } catch { /* leave it; harmless — recovery skips ids that already have a report */ }
    }

    public void Dispose()
    {
        var writer = Interlocked.Exchange(ref _writer, null);
        try { writer?.Dispose(); } catch { }
    }

    // Called once at startup, before any run exists: every journal still on disk belongs to a run that
    // never finished. Each becomes an "Interrupted" History entry. Returns how many were recovered.
    public static int RecoverInterrupted()
    {
        int recovered = 0;
        try
        {
            if (!Directory.Exists(JournalDir)) return 0;
            foreach (var file in Directory.GetFiles(JournalDir, "run_*.jsonl"))
            {
                try
                {
                    if (TryRecover(file)) recovered++;
                }
                catch { /* an unreadable journal is left in place; never block startup */ }
            }
        }
        catch { }
        return recovered;
    }

    private static bool TryRecover(string file)
    {
        SavedReport? report = null;
        var items = new List<SavedReportItem>();
        // FileShare.ReadWrite: if another app instance is still journaling this run, read what's there.
        using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = new StreamReader(stream, Encoding.UTF8))
        {
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                try
                {
                    if (report == null) report = JsonSerializer.Deserialize<SavedReport>(line, Options);
                    else
                    {
                        var item = JsonSerializer.Deserialize<SavedReportItem>(line, Options);
                        if (item != null) items.Add(item);
                    }
                }
                catch (JsonException)
                {
                    // The last line of a crashed run is often cut off mid-write — keep everything before it.
                    break;
                }
            }
        }

        if (report == null || string.IsNullOrEmpty(report.Id))
        {
            try { File.Delete(file); } catch { }
            return false;
        }

        // A normal report with this id already exists (the run finished but the journal wasn't
        // deleted): the report wins, the journal is just stale.
        if (ReportHistoryService.LoadFull(report.Id) is { Interrupted: false })
        {
            try { File.Delete(file); } catch { }
            return false;
        }

        report.Items          = items;
        report.Interrupted    = true;
        report.SuccessCount   = items.Count(i => i.Status == CopyStatus.Success);
        report.FailedCount    = items.Count(i => i.Status == CopyStatus.Failed);
        report.SkippedCount   = items.Count(i => i.Status == CopyStatus.Skipped);
        report.CancelledCount = items.Count(i => i.Status == CopyStatus.Cancelled);
        report.TotalCount     = items.Count;
        // Duration: from the run's start to the last time the journal was written.
        var lastWrite = new FileInfo(file).LastWriteTimeUtc;
        var elapsed   = new DateTimeOffset(lastWrite, TimeSpan.Zero) - report.Timestamp;
        report.Duration = elapsed > TimeSpan.Zero ? elapsed : TimeSpan.Zero;

        // Keep the journal if the History entry couldn't be written, so the next launch can try again.
        if (!ReportHistoryService.Save(report)) return false;
        try { File.Delete(file); } catch { }
        return true;
    }
}
