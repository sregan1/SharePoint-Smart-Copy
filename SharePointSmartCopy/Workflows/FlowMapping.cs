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

    // How a step writes each TARGET column (key = SOURCE list GUID, then the target column's internal name): "Value" for a
    // choice column (item/<Column>/Value), "" for a plain column such as text, number, date or yes/no (item/<Column>). A step
    // that wrote a choice column on the source and now writes a text column (or the reverse) must change its key, or the
    // import fails with "The API operation does not contain a definition for parameter 'item/<Column>/Value'".
    public Dictionary<string, Dictionary<string, string>> TargetColumnSuffixes { get; set; } = new();

    // The key suffix a SharePoint column type is written with, or null when the tool does not adjust that type (person, lookup,
    // link, multi-choice ... keep whatever the flow already has).
    public static string? ValueSuffixFor(string typeAsString) => typeAsString switch
    {
        "Choice" => "Value",
        "Text" or "Note" or "Number" or "Integer" or "Currency" or "DateTime" or "Boolean" or "Counter" => "",
        _ => null,
    };

    // Values for columns the TARGET list requires but a "Create item" step does not set (key = SOURCE list GUID, then the
    // target column's internal name). The rewriter adds `item/<column>` with this value to every Create item step on that list
    // that lacks it, because a required column left out makes the import fail ("missing required property").
    public Dictionary<string, Dictionary<string, string>> RequiredValues { get; set; } = new();

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

    public static FlowMapping Load(string path)
    {
        var m = JsonSerializer.Deserialize<FlowMapping>(File.ReadAllText(path), Json) ?? new FlowMapping();
        m.ReplaceNulls();
        return m;
    }

    // A hand-edited file can hold "Sites": null or a null value inside a section; the code that applies a mapping assumes
    // every section and value exists, so make that true here instead of failing later with a NullReferenceException.
    private void ReplaceNulls()
    {
        static Dictionary<string, string> Clean(Dictionary<string, string>? d)
        {
            var clean = new Dictionary<string, string>(StringComparer.Ordinal);
            if (d is null) return clean;
            foreach (var (k, v) in d) if (k is not null) clean[k] = v?.Trim() ?? string.Empty;
            return clean;
        }
        static Dictionary<string, Dictionary<string, string>> CleanNested(Dictionary<string, Dictionary<string, string>>? d)
        {
            var clean = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
            if (d is null) return clean;
            foreach (var (k, v) in d) if (k is not null) clean[k] = Clean(v);
            return clean;
        }
        Sites = Clean(Sites); Lists = Clean(Lists); Forms = Clean(Forms); Columns = Clean(Columns);
        Emails = Clean(Emails); EmailDomains = Clean(EmailDomains); Text = Clean(Text); TenantIds = Clean(TenantIds);
        ListColumns = CleanNested(ListColumns); RequiredValues = CleanNested(RequiredValues);
    }

    public void Save(string path) => File.WriteAllText(path, JsonSerializer.Serialize(this, Json));

    // Pairs with a non-empty target, longest source first so a URL is replaced before its parent URL.
    // Targets are trimmed: a stray space typed or saved in a target would otherwise be written into the package.
    public static IEnumerable<KeyValuePair<string, string>> Active(Dictionary<string, string> map) =>
        map.Where(kv => !string.IsNullOrWhiteSpace(kv.Value) && kv.Key != kv.Value.Trim())
           .Select(kv => new KeyValuePair<string, string>(kv.Key, kv.Value.Trim()))
           .OrderByDescending(kv => kv.Key.Length);

    // Records column decisions per list (ListColumns). A decision whose target is empty or equal to the source is
    // "leave unchanged" and adds nothing. The global Columns map (used for expressions that can't be tied to one
    // list) gets a rename only when EVERY list that uses that column renames it the same way, because a
    // contradictory global rename would corrupt the lists that keep the name. Returns a note for each conflict.
    // Existing entries are left alone, so manual edits win.
    // A column the target list does not have, "mapped" to this: the Create item steps stop sending it (the report lists it). It
    // can never be a real column name (internal names have no brackets), so it cannot clash with a target column.
    public const string OmitColumn = "[omit]";

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
            // A column left out of one list's steps is never renamed everywhere else.
            if (g.Any(c => c.TargetColumn == OmitColumn)) continue;
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
