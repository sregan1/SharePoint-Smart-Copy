using System.IO;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Forms;

// Full-length payloads for the question types the Forms API handles least predictably (Ranking, Likert / matrix): what was
// read from the source, what was sent, and what the target answered. The Forms tab log truncates every line and is not kept,
// so a copy that comes out wrong left nothing to look at. Written to %LOCALAPPDATA%\SharePointSmartCopy\Logs\formscopy-<date>.log.
// It holds question text and choices from the user's forms, so the diagnostics export leaves it out.
public static class FormsCopyTrace
{
    public static string PathFor(DateTime day) => Path.Combine(DiagnosticsLog.LogDirectory, $"formscopy-{day:yyyyMMdd}.log");

    public static void Write(string line)
    {
        try
        {
            Directory.CreateDirectory(DiagnosticsLog.LogDirectory);
            File.AppendAllText(PathFor(DateTime.Now), $"{DateTime.Now:HH:mm:ss.fff}  {line}{Environment.NewLine}");
        }
        catch { /* a trace must never break a copy */ }
    }

    public static bool IsTrickyType(string? type) =>
        type is not null && (type.Contains("rank", StringComparison.OrdinalIgnoreCase)
                             || type.Contains("likert", StringComparison.OrdinalIgnoreCase)
                             || type.Contains("matrix", StringComparison.OrdinalIgnoreCase));
}
