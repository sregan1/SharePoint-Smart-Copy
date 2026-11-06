using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace SharePointSmartCopy.Forms;

// One copy of a source form: its new ID, which account it was copied to and when, and the source form's title
// (so the Flows tool can show which form an ID refers to). The account, date and title are null for entries written by
// earlier versions, which stored less.
public sealed record FormsMapEntry(string NewId, string? TargetAccount, DateTimeOffset? CopiedAt, string? SourceTitle = null);

// Persists "source form ID -> copies" for every form the Forms dialog copies, so the Flows dialog can point flows at the
// copies and the Forms dialog can show what was already copied. Lives next to settings.json in
// %APPDATA%\SharePointSmartCopy.
//
// A source form keeps ONE entry PER TARGET ACCOUNT (the newest copy to that account), so copying the same form to a second
// account no longer overwrites the first. File shape, per source form:
//   { "newId", "targetAccount", "copiedAt", "sourceTitle",      <- the most recent copy overall (what older builds read)
//     "copies": [ { newId, targetAccount, copiedAt, sourceTitle }, ... ] }   <- every account's newest copy
// Files written before this (a bare ID string, or an object without "copies") load as a single entry.
public static class FormsIdMapStore
{
    public static string DefaultPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SharePointSmartCopy", "forms-id-map.json");

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    public static Dictionary<string, List<FormsMapEntry>> LoadAll(string? path = null) => Read(path ?? DefaultPath, out _);

    // The newest copy in a list (entries from earlier versions have no date and count as oldest).
    public static FormsMapEntry Latest(IEnumerable<FormsMapEntry> entries) =>
        entries.OrderByDescending(e => e.CopiedAt ?? DateTimeOffset.MinValue).First();

    // The copy of a source form that belongs to the given target account. With no account to go on, the newest copy.
    // With an account but no copy to it, an entry from an earlier version (it recorded no account, so it may be this
    // one) — and otherwise null, because a copy made for a DIFFERENT account would point flows at a form that does not
    // exist in the target.
    public static FormsMapEntry? Find(Dictionary<string, List<FormsMapEntry>> map, string sourceFormId, string? targetAccount)
    {
        if (!map.TryGetValue(sourceFormId, out var list) || list.Count == 0) return null;
        if (string.IsNullOrWhiteSpace(targetAccount)) return Latest(list);
        var mine = list.Where(e => SameAccount(e.TargetAccount, targetAccount)).ToList();
        if (mine.Count > 0) return Latest(mine);
        var unknown = list.Where(e => string.IsNullOrWhiteSpace(e.TargetAccount)).ToList();
        return unknown.Count > 0 ? Latest(unknown) : null;
    }

    private static bool SameAccount(string? a, string? b) =>
        string.Equals(a?.Trim() ?? string.Empty, b?.Trim() ?? string.Empty, StringComparison.OrdinalIgnoreCase);

    // damaged: the file exists but could not be read or parsed (as opposed to simply not being there yet).
    private static Dictionary<string, List<FormsMapEntry>> Read(string path, out bool damaged)
    {
        damaged = false;
        var map = new Dictionary<string, List<FormsMapEntry>>(StringComparer.Ordinal);
        try
        {
            if (!File.Exists(path)) return map;
            if (JsonNode.Parse(File.ReadAllText(path)) is not JsonObject root) { damaged = true; return map; }
            foreach (var (source, value) in root)
            {
                var list = new List<FormsMapEntry>();
                try
                {
                    // Oldest format: "sourceId": "newId".
                    if (value is JsonValue v && v.TryGetValue<string>(out var oldNewId))
                        list.Add(new FormsMapEntry(oldNewId, null, null));
                    else if (value is JsonObject o)
                    {
                        if (o["copies"] is JsonArray copies)
                            foreach (var item in copies)
                                if (item is JsonObject co && TryReadEntry(co) is { } ce) list.Add(ce);
                        // No "copies" (written by an earlier version): the object itself is the single entry.
                        if (list.Count == 0 && TryReadEntry(o) is { } single) list.Add(single);
                    }
                }
                catch (Exception ex) when (ex is InvalidOperationException or FormatException)
                {
                    // One entry with a field of the wrong type (a number where text belongs) loses only itself, not the
                    // whole file - which would hide every "Copied" label and make the next Record move the file aside.
                    list.Clear();
                }
                if (list.Count > 0) map[source] = list;
            }
        }
        catch (Exception ex) when (ex is JsonException or IOException or InvalidOperationException or FormatException)
        {
            // a damaged map must never block copying; it is rebuilt as forms are copied (Record keeps the old file aside)
            damaged = true;
        }
        return map;
    }

    // Reads one entry; an entry inside "copies" that is malformed is skipped on its own, so the others are kept.
    private static FormsMapEntry? TryReadEntry(JsonObject o)
    {
        try
        {
            return o["newId"]?.GetValue<string>() is { Length: > 0 } newId
                ? new FormsMapEntry(
                    newId,
                    o["targetAccount"]?.GetValue<string>(),
                    DateTimeOffset.TryParse(o["copiedAt"]?.GetValue<string>(), out var at) ? at : null,
                    o["sourceTitle"]?.GetValue<string>())
                : null;
        }
        catch (Exception ex) when (ex is InvalidOperationException or FormatException) { return null; }
    }

    private static JsonObject WriteEntry(FormsMapEntry e) => new()
    {
        ["newId"] = e.NewId,
        ["targetAccount"] = e.TargetAccount,
        ["copiedAt"] = e.CopiedAt?.ToString("o"),
        ["sourceTitle"] = e.SourceTitle,
    };

    // Record is read-modify-write, so two copies of the app (or two quick Records) could overwrite each other's entry.
    // A named mutex makes them take turns; it covers the other process as well as this one.
    public static void Record(string sourceFormId, string newFormId, string? targetAccount = null, string? sourceTitle = null, string? path = null)
    {
        using var gate = new Mutex(false, @"Local\SharePointSmartCopy.FormsIdMap");
        var held = false;
        try
        {
            try { held = gate.WaitOne(TimeSpan.FromSeconds(10)); }
            catch (AbandonedMutexException) { held = true; }   // the other process died holding it: the file is still ours to write
            if (!held) throw new IOException("The saved form IDs file is in use by another copy of the app. Try again.");
            RecordLocked(sourceFormId, newFormId, targetAccount, sourceTitle, path ?? DefaultPath);
        }
        finally { if (held) gate.ReleaseMutex(); }
    }

    private static void RecordLocked(string sourceFormId, string newFormId, string? targetAccount, string? sourceTitle, string path)
    {
        var map = Read(path, out var damaged);
        if (damaged)
        {
            // Rewriting from an empty map would wipe every earlier source -> copy entry (the Flows tool relies on them).
            // Keep the unreadable file so it can be repaired by hand.
            try { File.Move(path, path + ".corrupt-" + DateTime.Now.ToString("yyyyMMdd-HHmmss"), overwrite: true); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* best effort */ }
        }

        // Replace this account's earlier copy of the form; every other account's copy stays.
        if (!map.TryGetValue(sourceFormId, out var list)) map[sourceFormId] = list = [];
        list.RemoveAll(e => SameAccount(e.TargetAccount, targetAccount));
        list.Add(new FormsMapEntry(newFormId, targetAccount, DateTimeOffset.Now, sourceTitle));

        var root = new JsonObject();
        foreach (var (source, entries) in map)
        {
            var entry = WriteEntry(Latest(entries));   // top level = newest copy, exactly what earlier builds read
            var all = new JsonArray();
            foreach (var e in entries) all.Add(WriteEntry(e));
            entry["copies"] = all;
            root[source] = entry;
        }
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        // A temp name of its own, so nothing else writing next to the file can collide with it.
        var tmp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(tmp, root.ToJsonString(Json));
            File.Move(tmp, path, overwrite: true);
        }
        catch
        {
            try { File.Delete(tmp); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* best effort */ }
            throw;
        }
    }
}
