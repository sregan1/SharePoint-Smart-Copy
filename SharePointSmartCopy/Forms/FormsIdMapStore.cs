using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace SharePointSmartCopy.Forms;

// The latest copy of a source form: its new ID, which account it was copied to and when, and the source form's title
// (so the Flows tool can show which form an ID refers to). The account, date and title are null for entries written by
// earlier versions, which stored less.
public sealed record FormsMapEntry(string NewId, string? TargetAccount, DateTimeOffset? CopiedAt, string? SourceTitle = null);

// Persists "source form ID -> latest copy" for every form the Forms dialog copies, so the Flows dialog can
// point flows at the copies and the Forms dialog can show what was already copied. Lives next to settings.json in
// %APPDATA%\SharePointSmartCopy.
public static class FormsIdMapStore
{
    public static string DefaultPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SharePointSmartCopy", "forms-id-map.json");

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    public static Dictionary<string, FormsMapEntry> Load(string? path = null)
    {
        path ??= DefaultPath;
        var map = new Dictionary<string, FormsMapEntry>(StringComparer.Ordinal);
        try
        {
            if (!File.Exists(path)) return map;
            if (JsonNode.Parse(File.ReadAllText(path)) is not JsonObject root) return map;
            foreach (var (source, value) in root)
            {
                // Old format: "sourceId": "newId". Current format: "sourceId": { newId, targetAccount, copiedAt }.
                if (value is JsonValue v && v.TryGetValue<string>(out var oldNewId))
                    map[source] = new FormsMapEntry(oldNewId, null, null);
                else if (value is JsonObject o && o["newId"]?.GetValue<string>() is { Length: > 0 } newId)
                    map[source] = new FormsMapEntry(
                        newId,
                        o["targetAccount"]?.GetValue<string>(),
                        DateTimeOffset.TryParse(o["copiedAt"]?.GetValue<string>(), out var at) ? at : null,
                        o["sourceTitle"]?.GetValue<string>());
            }
        }
        catch (Exception ex) when (ex is JsonException or IOException or InvalidOperationException or FormatException)
        {
            // a damaged map must never block copying; it is rebuilt as forms are copied
        }
        return map;
    }

    public static void Record(string sourceFormId, string newFormId, string? targetAccount = null, string? sourceTitle = null, string? path = null)
    {
        path ??= DefaultPath;
        var map = Load(path);
        map[sourceFormId] = new FormsMapEntry(newFormId, targetAccount, DateTimeOffset.Now, sourceTitle);

        var root = new JsonObject();
        foreach (var (source, e) in map)
            root[source] = new JsonObject
            {
                ["newId"] = e.NewId,
                ["targetAccount"] = e.TargetAccount,
                ["copiedAt"] = e.CopiedAt?.ToString("o"),
                ["sourceTitle"] = e.SourceTitle,
            };
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var tmp = path + ".tmp";
        File.WriteAllText(tmp, root.ToJsonString(Json));
        File.Move(tmp, path, overwrite: true);
    }
}
