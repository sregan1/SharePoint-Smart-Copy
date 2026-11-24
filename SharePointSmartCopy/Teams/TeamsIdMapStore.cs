using System.Collections.Concurrent;
using System.IO;
using System.Text.Json;

namespace SharePointSmartCopy.Teams;

// Remembers which target team and channels a source team was copied to, so a second run of the same copy resumes
// into them instead of creating a duplicate team. Lives next to settings.json in %APPDATA%\SharePointSmartCopy.
// Keyed by "source team id" + "target account" (the same team can be copied to more than one place).
public static class TeamsIdMapStore
{
    public sealed class TeamMap
    {
        public string TargetTeamId { get; set; } = "";
        public string TargetName { get; set; } = "";
        public DateTimeOffset CopiedAt { get; set; }
        // source channel id -> target channel id
        public ConcurrentDictionary<string, string> Channels { get; set; } = new(StringComparer.Ordinal);
        // source channel id -> true once its files were copied (so a resume does not copy them again)
        public ConcurrentDictionary<string, bool> FilesDone { get; set; } = new(StringComparer.Ordinal);
        // source channel (or chat) id -> true once its messages were imported and its migration completed
        public ConcurrentDictionary<string, bool> MessagesDone { get; set; } = new(StringComparer.Ordinal);
        // source channel (or chat) id -> true once its migration mode was started (a later start takes no date)
        public ConcurrentDictionary<string, bool> MessagesStarted { get; set; } = new(StringComparer.Ordinal);
        // source message id -> imported message id, so the replies of an already imported message can still be added
        public ConcurrentDictionary<string, string> Messages { get; set; } = new(StringComparer.Ordinal);
    }

    public static string DefaultPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SharePointSmartCopy", "teams-id-map.json");

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };
    private static readonly object Gate = new();

    private static string Key(string sourceTeamId, string? targetAccount) =>
        sourceTeamId + "|" + (targetAccount ?? "").Trim().ToLowerInvariant();

    public static TeamMap? Find(string sourceTeamId, string? targetAccount, string? path = null)
    {
        lock (Gate) return Read(path ?? DefaultPath).GetValueOrDefault(Key(sourceTeamId, targetAccount));
    }

    // Writes the whole map for one team (the caller mutates a TeamMap and saves it after each step, so a crash loses at most one step).
    public static void Save(string sourceTeamId, string? targetAccount, TeamMap map, string? path = null)
    {
        lock (Gate)
        {
            path ??= DefaultPath;
            var all = Read(path);
            all[Key(sourceTeamId, targetAccount)] = map;
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            var tmp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(tmp, JsonSerializer.Serialize(all, Json));
                File.Move(tmp, path, overwrite: true);
            }
            catch
            {
                try { File.Delete(tmp); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* best effort */ }
                throw;
            }
        }
    }

    public static void Forget(string sourceTeamId, string? targetAccount, string? path = null)
    {
        lock (Gate)
        {
            path ??= DefaultPath;
            var all = Read(path);
            if (all.Remove(Key(sourceTeamId, targetAccount)))
                File.WriteAllText(path, JsonSerializer.Serialize(all, Json));
        }
    }

    private static Dictionary<string, TeamMap> Read(string path)
    {
        try
        {
            if (!File.Exists(path)) return new(StringComparer.Ordinal);
            return JsonSerializer.Deserialize<Dictionary<string, TeamMap>>(File.ReadAllText(path)) ?? new(StringComparer.Ordinal);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            // An unreadable map only means a resume starts over (a new team); it must never block a copy.
            return new(StringComparer.Ordinal);
        }
    }
}
