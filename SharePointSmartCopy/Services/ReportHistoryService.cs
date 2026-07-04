using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public static class ReportHistoryService
{
    private static readonly string ReportsDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SharePointSmartCopy", "Reports");

    private static readonly JsonSerializerOptions _options = new()
    {
        WriteIndented = true,
        Converters    = { new JsonStringEnumConverter() }
    };

    public static void Save(SavedReport report)
    {
        try
        {
            Directory.CreateDirectory(ReportsDir);
            File.WriteAllText(
                Path.Combine(ReportsDir, $"report_{report.Id}.json"),
                JsonSerializer.Serialize(report, _options));

            // Keep at most 50 reports; delete oldest by filename (which sorts chronologically)
            foreach (var old in Directory.GetFiles(ReportsDir, "report_*.json")
                                         .OrderByDescending(f => f)
                                         .Skip(50))
                File.Delete(old);
        }
        catch { /* non-critical */ }
    }

    // Reads via a FileStream rather than File.ReadAllText: the latter decodes the whole file to a
    // UTF-16 string, which JsonSerializer then re-encodes back to UTF-8 bytes internally for its
    // reader — a redundant round-trip conversion, plus a large string allocation, for every file.
    // Deserializing straight from the stream skips both, which matters once report files reach
    // tens of MB (a 100,000+-item, pretty-printed Items array gets large fast).
    private static T? DeserializeFile<T>(string path)
    {
        using var stream = File.OpenRead(path);
        return JsonSerializer.Deserialize<T>(stream, _options);
    }

    public static List<SavedReport> LoadAll()
    {
        if (!Directory.Exists(ReportsDir)) return [];

        var result = new List<SavedReport>();
        foreach (var file in Directory.GetFiles(ReportsDir, "report_*.json")
                                      .OrderByDescending(f => f))
        {
            try
            {
                var r = DeserializeFile<SavedReport>(file);
                if (r != null) result.Add(r);
            }
            catch { /* skip corrupt files */ }
        }
        return result;
    }

    // See SavedReportSummary for why this exists instead of always using LoadAll: deserializing
    // into a type without an Items property makes System.Text.Json skip that (often huge) JSON
    // array instead of materializing it, which is what makes listing history fast.
    public static List<SavedReportSummary> LoadSummaries()
    {
        if (!Directory.Exists(ReportsDir)) return [];

        var result = new List<SavedReportSummary>();
        foreach (var file in Directory.GetFiles(ReportsDir, "report_*.json")
                                      .OrderByDescending(f => f))
        {
            try
            {
                var r = DeserializeFile<SavedReportSummary>(file);
                if (r != null) result.Add(r);
            }
            catch { /* skip corrupt files */ }
        }
        return result;
    }

    // Loads one report's full detail (including Items) by id — used lazily once a specific run is
    // selected, exported, or verified, rather than up front for every saved report.
    public static SavedReport? LoadFull(string id)
    {
        var path = Path.Combine(ReportsDir, $"report_{id}.json");
        if (!File.Exists(path)) return null;
        try { return DeserializeFile<SavedReport>(path); }
        catch { return null; }
    }

    public static void Delete(string id)
    {
        try
        {
            var path = Path.Combine(ReportsDir, $"report_{id}.json");
            if (File.Exists(path)) File.Delete(path);
        }
        catch { /* non-critical */ }
    }
}
