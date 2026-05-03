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

    public static List<SavedReport> LoadAll()
    {
        if (!Directory.Exists(ReportsDir)) return [];

        var result = new List<SavedReport>();
        foreach (var file in Directory.GetFiles(ReportsDir, "report_*.json")
                                      .OrderByDescending(f => f))
        {
            try
            {
                var r = JsonSerializer.Deserialize<SavedReport>(File.ReadAllText(file), _options);
                if (r != null) result.Add(r);
            }
            catch { /* skip corrupt files */ }
        }
        return result;
    }

    public static void Delete(SavedReport report)
    {
        try
        {
            var path = Path.Combine(ReportsDir, $"report_{report.Id}.json");
            if (File.Exists(path)) File.Delete(path);
        }
        catch { /* non-critical */ }
    }
}
