using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Workflows;

// What to change when moving flows from Tenant A to Tenant B. Every map is "source value" -> "target value".
// An empty target means "not mapped yet": the value is left alone and reported as unmapped.
public sealed class FlowMapping
{
    // Site collection / web URLs, e.g. https://a.sharepoint.com/sites/HR -> https://b.sharepoint.com/sites/HR.
    // Also rewrites the same URL when it appears inside email bodies and expressions (longest match first).
    public Dictionary<string, string> Sites { get; set; } = new();

    // SharePoint list GUIDs (source -> target).
    public Dictionary<string, string> Lists { get; set; } = new();

    // Forms form IDs (source -> the new form's ID from the Forms copy).
    public Dictionary<string, string> Forms { get; set; } = new();

    // SharePoint column INTERNAL names (e.g. Tax_x0020_ID). Applied only to written columns (item/<name>),
    // 'body/<name>' reads, and OData $filter/$orderby/$select.
    public Dictionary<string, string> Columns { get; set; } = new();

    // Column renames that apply only to flow steps targeting one list (key = SOURCE list GUID). The same internal
    // name can map differently on different lists; this wins over `Columns` for that list's steps.
    public Dictionary<string, Dictionary<string, string>> ListColumns { get; set; } = new();

    // Individual email addresses (recipients, shared mailboxes, approvers...).
    public Dictionary<string, string> Emails { get; set; } = new();

    // Whole-domain move, e.g. { "a.com": "b.com" }; applied after Emails to addresses not already mapped.
    public Dictionary<string, string> EmailDomains { get; set; } = new();

    // Library/folder paths and any other literal text.
    public Dictionary<string, string> Text { get; set; } = new();

    // Tenant GUID (appears in the creator block and sometimes in URLs/parameters).
    public Dictionary<string, string> TenantIds { get; set; } = new();

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
    };

    public static FlowMapping Load(string path) =>
        JsonSerializer.Deserialize<FlowMapping>(File.ReadAllText(path), Json) ?? new FlowMapping();

    public void Save(string path) => File.WriteAllText(path, JsonSerializer.Serialize(this, Json));

    // Pairs with a non-empty target, longest source first so a URL is replaced before its parent URL.
    public static IEnumerable<KeyValuePair<string, string>> Active(Dictionary<string, string> map) =>
        map.Where(kv => !string.IsNullOrWhiteSpace(kv.Value) && kv.Key != kv.Value).OrderByDescending(kv => kv.Key.Length);

    // Records column decisions per list (ListColumns). A decision whose target is empty or equal to the source is
    // "leave unchanged" and adds nothing. The global Columns map (used for expressions that can't be tied to one
    // list) gets a rename only when EVERY list that uses that column renames it the same way, because a
    // contradictory global rename would corrupt the lists that keep the name. Returns a note for each conflict.
    // Existing entries are left alone, so manual edits win.
    public List<string> ApplyColumnChoices(IEnumerable<ColumnChoice> choices)
    {
        var notes = new List<string>();
        var all = choices.ToList();

        foreach (var c in all.Where(c => !string.IsNullOrWhiteSpace(c.TargetColumn) && c.TargetColumn != c.SourceColumn))
        {
            if (!ListColumns.TryGetValue(c.SourceListId, out var perList))
                ListColumns[c.SourceListId] = perList = new Dictionary<string, string>(StringComparer.Ordinal);
            if (string.IsNullOrWhiteSpace(perList.GetValueOrDefault(c.SourceColumn)))
                perList[c.SourceColumn] = c.TargetColumn!;
        }

        foreach (var g in all.GroupBy(c => c.SourceColumn, StringComparer.Ordinal))
        {
            var targets = g.Select(c => string.IsNullOrWhiteSpace(c.TargetColumn) ? c.SourceColumn : c.TargetColumn!).Distinct().ToList();
            if (targets.Count == 1 && targets[0] != g.Key)
            {
                if (string.IsNullOrWhiteSpace(Columns.GetValueOrDefault(g.Key))) Columns[g.Key] = targets[0];
            }
            else if (targets.Count > 1)
                notes.Add(Loc.T("Flows_Note_ColumnConflict", g.Key, string.Join(", ", targets)));
        }
        return notes;
    }

    // Builds a template with every discovered value as a key and an empty target.
    public static FlowMapping TemplateFrom(IEnumerable<FlowScan> scans)
    {
        var all = scans.SelectMany(s => s.Findings).ToList();
        IEnumerable<string> Vals(params FindingKind[] kinds) =>
            all.Where(f => kinds.Contains(f.Kind)).Select(f => f.Value).Distinct(StringComparer.Ordinal).OrderBy(v => v, StringComparer.OrdinalIgnoreCase);
        var m = new FlowMapping();
        foreach (var v in Vals(FindingKind.SiteUrl)) m.Sites[v] = "";
        foreach (var v in Vals(FindingKind.ListId)) m.Lists[v] = "";
        foreach (var v in Vals(FindingKind.FormId)) m.Forms[v] = "";
        foreach (var v in Vals(FindingKind.ColumnName)) m.Columns[v] = "";
        foreach (var v in Vals(FindingKind.Email, FindingKind.Mailbox)) m.Emails[v] = "";
        foreach (var v in Vals(FindingKind.FilePath)) m.Text[v] = "";
        foreach (var v in Vals(FindingKind.TenantId)) m.TenantIds[v] = "";
        return m;
    }
}
