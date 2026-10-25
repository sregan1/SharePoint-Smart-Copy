using System.IO;
using System.Security.Cryptography;
using System.Text.Json;

namespace SharePointSmartCopy.Workflows;

// A small local record of the packages the Flows tool has written, so that adding one of them again as if it were an
// original export can be noticed. A written package keeps the original's file name, so the two are easy to confuse, and
// scanning it finds the TARGET addresses where the source ones were expected. Recognition is by content (a hash), so it
// works wherever the file is copied or renamed. Lives next to settings.json in %APPDATA%\SharePointSmartCopy.
public static class FlowOutputLog
{
    public sealed record Entry(string Hash, string Path, DateTimeOffset WrittenAt);

    private const int MaxEntries = 500;

    public static string DefaultPath { get; } = System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SharePointSmartCopy", "flows-output-log.json");

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    private static string HashOf(string file)
    {
        using var stream = File.OpenRead(file);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static List<Entry> Load(string path)
    {
        try
        {
            return File.Exists(path) ? JsonSerializer.Deserialize<List<Entry>>(File.ReadAllText(path)) ?? [] : [];
        }
        catch (Exception ex) when (ex is JsonException or IOException)
        {
            return [];   // a damaged log only means a package is not recognized; it never blocks anything
        }
    }

    // Remembers a package that was just written.
    public static void Record(string zipPath, string? logPath = null)
    {
        logPath ??= DefaultPath;
        var hash = HashOf(zipPath);
        var entries = Load(logPath);
        entries.RemoveAll(e => e.Hash == hash || string.Equals(e.Path, zipPath, StringComparison.OrdinalIgnoreCase));
        entries.Add(new Entry(hash, zipPath, DateTimeOffset.Now));
        if (entries.Count > MaxEntries) entries = entries.OrderByDescending(e => e.WrittenAt).Take(MaxEntries).ToList();

        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(logPath)!);
        var tmp = logPath + ".tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(entries, Json));
        File.Move(tmp, logPath, overwrite: true);
    }

    // The record of this exact file's content, if the tool wrote it; otherwise null.
    public static Entry? Find(string zipPath, string? logPath = null)
    {
        try
        {
            var entries = Load(logPath ?? DefaultPath);
            if (entries.Count == 0) return null;
            var hash = HashOf(zipPath);
            return entries.FirstOrDefault(e => e.Hash == hash);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return null;
        }
    }
}
