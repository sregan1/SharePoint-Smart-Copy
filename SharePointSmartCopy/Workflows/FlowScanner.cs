using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace SharePointSmartCopy.Workflows;

public enum FindingKind
{
    SiteUrl, ListId, FormId, FormQuestionId, ColumnName, ColumnExpression, QueryFilter,
    EmbeddedUrl, Email, Mailbox, FilePath, TenantId, UserId, ConnectionRef,
    OtherResource,   // an identifier for a service the tool doesn't rewrite (Teams, Planner, Excel, ...): review manually
}

// One tenant-bound value found in a flow. Value is the raw string; Detail is a decoded/explanatory note.
// Site / ListId say which SharePoint site and list the operation that contained the value points at
// (only when the operation uses literal values, not expressions); the mapping resolver uses them to look
// the list up by name in the source and find its counterpart in the target.
public sealed record Finding(FindingKind Kind, string Value, string Location, string? Detail = null,
                             string? Site = null, string? ListId = null);

public sealed class FlowScan
{
    public required string PackageName { get; init; }
    public required string FlowName { get; init; }
    public required string Trigger { get; init; }
    // True for a flow read from a solution export (Workflows/*.json) rather than a legacy package.
    public bool IsSolution { get; init; }
    public List<Finding> Findings { get; } = [];
    // connector (apiName, e.g. shared_sharepointonline) -> number of trigger/action steps that use it
    public Dictionary<string, int> Connectors { get; } = new(StringComparer.OrdinalIgnoreCase);
}

// Reads a legacy flow package (.zip) in memory (never extracts to disk) and classifies tenant-bound values.
// It only reports; it never modifies anything.
public static partial class FlowScanner
{
    // A legacy package keeps each flow at Microsoft.Flow/flows/<id>/definition.json; a solution export keeps it at
    // Workflows/<name>-<GUID>.json. Both hold { "properties": { "definition": ..., "connectionReferences": ... } }.
    public static bool IsLegacyFlowEntry(string fullName) =>
        fullName.StartsWith("Microsoft.Flow/flows/", StringComparison.OrdinalIgnoreCase)
        && fullName.EndsWith("/definition.json", StringComparison.OrdinalIgnoreCase);

    public static bool IsSolutionFlowEntry(string fullName) =>
        fullName.StartsWith("Workflows/", StringComparison.OrdinalIgnoreCase)
        && fullName.EndsWith(".json", StringComparison.OrdinalIgnoreCase)
        && fullName.IndexOf('/', "Workflows/".Length) < 0;

    public static bool IsFlowEntry(string fullName) => IsLegacyFlowEntry(fullName) || IsSolutionFlowEntry(fullName);

    // A solution workflow's file name is "<Name>-<GUID>.json"; the name is the best label when the JSON has none.
    public static string NameFromEntry(string fullName)
    {
        var file = Path.GetFileNameWithoutExtension(fullName);
        return Regex.Replace(file, @"-[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$", "");
    }

    // The flow JSON of an entry, or null when it is not a flow (a Workflows/ file without a definition).
    public static JsonObject? TryParseFlow(string json)
    {
        var root = JsonNode.Parse(json) as JsonObject;
        return root?["properties"] is JsonObject p && p["definition"] is JsonObject ? root : null;
    }

    public static IEnumerable<FlowScan> ScanPackage(string zipPath)
    {
        using var zip = ZipFile.OpenRead(zipPath);
        var pkg = Path.GetFileNameWithoutExtension(zipPath);
        foreach (var entry in zip.Entries.Where(e => IsFlowEntry(e.FullName)))
        {
            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            var flow = TryParseFlow(reader.ReadToEnd());
            if (flow is null) continue;
            yield return ScanFlow(pkg, flow, IsSolutionFlowEntry(entry.FullName), NameFromEntry(entry.FullName));
        }
    }

    public static FlowScan ScanFlow(string packageName, JsonObject flow, bool isSolution = false, string? fallbackName = null)
    {
        var props = flow["properties"]!.AsObject();
        var def = props["definition"]!.AsObject();
        var trigger = def["triggers"]?.AsObject().FirstOrDefault();
        var scan = new FlowScan
        {
            PackageName = packageName,
            FlowName = props["displayName"]?.GetValue<string>() ?? (string.IsNullOrEmpty(fallbackName) ? "(unnamed)" : fallbackName),
            IsSolution = isSolution,
            Trigger = trigger is { Value: not null } ? $"{trigger.Value.Key} [{OperationOf(trigger.Value.Value!)}]" : "(none)",
        };

        // The creator's tenant and user are recorded in the definition.
        if (FindCreator(flow) is { } creator)
        {
            if (creator["tenantId"]?.GetValue<string>() is { } t)
                scan.Findings.Add(new(FindingKind.TenantId, t, "creator.tenantId"));
            if (creator["id"]?.GetValue<string>() is { } u)
                scan.Findings.Add(new(FindingKind.UserId, u, "creator.id", "flow creator object ID"));
        }

        if (props["connectionReferences"] is JsonObject refs)
            foreach (var (name, v) in refs)
                scan.Findings.Add(new(FindingKind.ConnectionRef, name, "connectionReferences",
                    $"{v?["apiName"]} (connection must be recreated in the target)"));

        Walk(scan, def, "definition");
        return scan;
    }

    private static JsonObject? FindCreator(JsonNode n)
    {
        if (n is JsonObject o)
        {
            if (o["creator"] is JsonObject c && c["tenantId"] is not null) return c;
            foreach (var kv in o)
                if (kv.Value is { } v && FindCreator(v) is { } r) return r;
        }
        else if (n is JsonArray a)
        {
            foreach (var v in a)
                if (v is not null && FindCreator(v) is { } r) return r;
        }
        return null;
    }

    private static string OperationOf(JsonNode node) =>
        (node["inputs"] as JsonObject)?["host"]?["operationId"]?.GetValue<string>() ?? node["type"]?.GetValue<string>() ?? "?";

    private static void Walk(FlowScan scan, JsonNode node, string path)
    {
        switch (node)
        {
            case JsonObject o:
                if (o["inputs"] is JsonObject inp && inp["host"] is JsonObject host)
                {
                    var api = host["apiId"]?.GetValue<string>()?.Split('/').Last();
                    if (!string.IsNullOrEmpty(api)) scan.Connectors[api] = scan.Connectors.GetValueOrDefault(api) + 1;
                    if (inp["parameters"] is JsonObject prm) ScanOperation(scan, host, prm, path);
                }
                foreach (var (k, v) in o)
                    if (v is not null) Walk(scan, v, $"{path}/{k}");
                break;
            case JsonArray a:
                for (var i = 0; i < a.Count; i++)
                    if (a[i] is not null) Walk(scan, a[i]!, $"{path}[{i}]");
                break;
            case JsonValue v when v.TryGetValue<string>(out var s):
                ScanString(scan, s, path);
                break;
        }
    }

    private static void ScanOperation(FlowScan scan, JsonObject host, JsonObject prm, string path)
    {
        var connector = host["apiId"]?.GetValue<string>()?.Split('/').Last() ?? "";
        var op = host["operationId"]?.GetValue<string>() ?? "";
        var where = $"{path} [{op}]";
        // Literal site/list this operation targets (null when it's an expression like @parameters(...)).
        string? Literal(string key) => prm[key] is JsonValue lv && lv.TryGetValue<string>(out var ls) && !ls.StartsWith('@') ? ls : null;
        var site = connector == "shared_sharepointonline" ? Literal("dataset") : null;
        var list = connector == "shared_sharepointonline" ? Literal("table") : null;
        foreach (var (key, val) in prm)
        {
            var s = val is JsonValue jv && jv.TryGetValue<string>(out var str) ? str : val?.ToJsonString() ?? "";
            if (connector == "shared_sharepointonline")
            {
                if (key == "dataset") scan.Findings.Add(new(FindingKind.SiteUrl, s, where));
                else if (key == "table") scan.Findings.Add(new(FindingKind.ListId, s, where, "SharePoint list GUID", site, list));
                else if (key.StartsWith("item/")) scan.Findings.Add(new(FindingKind.ColumnName, ColumnOf(key), where, "written column (internal name)", site, list));
                else if (key is "$filter" or "$orderby" or "$select") scan.Findings.Add(new(FindingKind.QueryFilter, s, where, "OData query uses column internal names", site, list));
                else if (key is "folderPath" or "path" && s.StartsWith('/')) scan.Findings.Add(new(FindingKind.FilePath, s, where));
            }
            else if (connector == "shared_microsoftforms" && key == "form_id")
                scan.Findings.Add(new(FindingKind.FormId, s, where, DecodeFormId(s)));
            else if (key == "mailboxAddress")
                scan.Findings.Add(new(FindingKind.Mailbox, s, where));
            else if (connector == "shared_onedriveforbusiness" && key == "id" && !s.StartsWith('@'))
                scan.Findings.Add(new(FindingKind.FilePath, s, where, "OneDrive file ID"));
            else if (FlowResources.IsReviewConnector(connector) && val is JsonValue rv && rv.TryGetValue<string>(out var lit)
                     && FlowResources.LooksLikeResourceId(lit))
                scan.Findings.Add(new(FindingKind.OtherResource, lit, where,
                    $"{FlowResources.FriendlyName(connector)}: {key} (not rewritten; review after import)"));
        }
    }

    private static string ColumnOf(string key)
    {
        var name = key["item/".Length..];
        var slash = name.IndexOf('/');
        return slash > 0 ? name[..slash] : name;   // item/Progress/Value -> Progress
    }

    // Form IDs are base64url(tenantGuid[16, LE] + ownerObjectId[16, LE] + 28-byte ASCII short id).
    public static string? DecodeFormId(string id) =>
        TryGetFormTenant(id) is { } tenant && TryGetFormOwner(id) is { } owner ? $"tenant {tenant}, owner {owner}" : null;

    public static Guid? TryGetFormTenant(string id) => FormIdGuid(id, 0);
    public static Guid? TryGetFormOwner(string id) => FormIdGuid(id, 16);

    private static Guid? FormIdGuid(string id, int offset)
    {
        try
        {
            var seg = id.Trim().Replace('-', '+').Replace('_', '/');
            var raw = Convert.FromBase64String(seg.PadRight((seg.Length + 3) / 4 * 4, '='));
            return raw.Length < 32 ? null : new Guid(raw[offset..(offset + 16)]);
        }
        catch (FormatException) { return null; }
    }

    private static void ScanString(FlowScan scan, string s, string path)
    {
        if (path.EndsWith("/operationMetadataId") || path.Contains("/metadata/")) return;

        foreach (Match m in SpUrl().Matches(s))
            scan.Findings.Add(new(FindingKind.EmbeddedUrl, m.Value.TrimEnd('\\', '.', ',', ')'), path, "SharePoint URL inside text/expression"));

        foreach (Match m in Question().Matches(s))
            scan.Findings.Add(new(FindingKind.FormQuestionId, m.Groups[1].Value, path, "Forms question ID (kept as-is when the form is copied)"));

        foreach (Match m in ColumnExpr().Matches(s))
        {
            var name = m.Groups[1].Value;
            if (QuestionName().IsMatch(name) || ExprNoise.Contains(name)) continue;
            scan.Findings.Add(new(FindingKind.ColumnExpression, name, path, "property read in an expression (column internal name?)"));
        }

        foreach (Match m in EmailAddr().Matches(ExprBlock().Replace(s, " ")))
        {
            // Teams chat/channel ids look like addresses ("19:abc@thread.tacv2"); they are not email.
            if (m.Value.Contains("@thread.", StringComparison.OrdinalIgnoreCase)) continue;
            scan.Findings.Add(new(FindingKind.Email, m.Value, path));
        }
    }

    private static readonly HashSet<string> ExprNoise = new(StringComparer.OrdinalIgnoreCase)
    {
        "value", "responder", "responses", "attachments", "outcome", "submitDate", "ID", "mail", "{Link}", "{Name}", "{FullPath}",
    };

    [GeneratedRegex(@"https?://[\w.-]+\.sharepoint\.com/[^\s""'<>\\)]*", RegexOptions.IgnoreCase)] private static partial Regex SpUrl();
    [GeneratedRegex(@"body/(r[0-9a-f]{32})")] private static partial Regex Question();
    [GeneratedRegex(@"^r[0-9a-f]{32}$")] private static partial Regex QuestionName();
    [GeneratedRegex(@"\['body/([^'/]+)'\]")] private static partial Regex ColumnExpr();
    [GeneratedRegex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")] private static partial Regex EmailAddr();
    [GeneratedRegex(@"@\{[^}]*\}|@[A-Za-z]+\([^)]*\)(?:\?\[[^\]]*\])*")] private static partial Regex ExprBlock();

    public static string ToMarkdown(IReadOnlyList<FlowScan> scans)
    {
        var sb = new StringBuilder("# Flow scan report\n\nTenant-bound values found per flow. Nothing was modified.\n\n");
        foreach (var s in scans)
        {
            sb.Append($"## {s.FlowName}\n\nPackage `{s.PackageName}` · trigger `{s.Trigger}`\n\n");
            foreach (var g in s.Findings.GroupBy(f => f.Kind).OrderBy(g => g.Key))
            {
                sb.Append($"### {g.Key} ({g.Select(f => f.Value).Distinct().Count()} distinct, {g.Count()} total)\n\n");
                foreach (var v in g.GroupBy(f => f.Value).OrderByDescending(x => x.Count()))
                {
                    var first = v.First();
                    var val = v.Key.Length > 120 ? v.Key[..120] + "…" : v.Key;
                    sb.Append($"- `{val.Replace("`", "'").Replace("\n", " ")}` ×{v.Count()}{(first.Detail is null ? "" : " — " + first.Detail)}\n");
                }
                sb.Append('\n');
            }
        }
        return sb.ToString();
    }
}
