using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Workflows;

public sealed record Change(string Category, string From, string To, int Count);

// Applies a FlowMapping to a flow definition. Pure and offline: the original JSON is never modified in place
// by callers; RewritePackage writes a new .zip and leaves the source package untouched.
public static partial class FlowRewriter
{
    public static List<Change> Apply(JsonObject flow, FlowMapping map)
    {
        var counts = new Dictionary<(string Cat, string From, string To), int>();
        void Hit(string cat, string from, string to, int n = 1)
        {
            if (n <= 0) return;
            counts[(cat, from, to)] = counts.GetValueOrDefault((cat, from, to)) + n;
        }

        // Every value mapping in one simultaneous pass, so a replaced value is never matched again (HR->HR2 together
        // with HR2->HR3 stays HR->HR2), and each value only matches whole (sites/HR does not touch sites/HRTeam).
        var (valueRx, rules) = BuildValueRules(map);
        string Transform(string s) => valueRx is null ? s : valueRx.Replace(s, m =>
        {
            for (var i = 0; i < rules.Count; i++)
                if (m.Groups["r" + i].Success)
                {
                    Hit(rules[i].Category, rules[i].From, rules[i].To);
                    return rules[i].To;
                }
            return m.Value;
        });

        // Global column renames (apply to expressions anywhere) and per-list renames (apply to the SharePoint
        // operation that targets that list: written fields and OData query text). Per-list wins over global.
        var activeColumns = FlowMapping.Active(map.Columns).ToList();
        var globalColumns = activeColumns.Count > 0 ? ColumnReplacer(activeColumns, queryWords: false, Hit) : null;
        var perList = map.ListColumns.ToDictionary(
            kv => kv.Key, kv => FlowMapping.Active(kv.Value).ToList(), StringComparer.OrdinalIgnoreCase);

        void Walk(JsonNode node)
        {
            switch (node)
            {
                case JsonObject o:
                    if (o["inputs"] is JsonObject inp && inp["parameters"] is JsonObject prm && IsSharePoint(inp["host"]))
                    {
                        var cols = EffectiveColumns(prm, activeColumns, perList);
                        if (cols.Count > 0)
                        {
                            RenameColumnKeys(prm, cols, Hit);          // written fields: item/<Column>[/Value]
                            RewriteQueryParameters(prm, cols, Hit);    // $filter / $orderby / $select
                        }
                    }

                    foreach (var key in o.Select(kv => kv.Key).ToList())
                    {
                        var child = o[key];
                        if (child is JsonValue v && v.TryGetValue<string>(out var s))
                        {
                            var t = Transform(s);
                            if (globalColumns is not null) t = globalColumns(t);
                            if (t != s) o[key] = t;
                        }
                        else if (child is not null) Walk(child);
                    }
                    break;
                case JsonArray a:
                    for (var i = 0; i < a.Count; i++)
                    {
                        if (a[i] is JsonValue v && v.TryGetValue<string>(out var s))
                        {
                            var t = Transform(s);
                            if (globalColumns is not null) t = globalColumns(t);
                            if (t != s) a[i] = t;
                        }
                        else if (a[i] is not null) Walk(a[i]!);
                    }
                    break;
            }
        }

        Walk(flow);
        return counts.Select(kv => new Change(kv.Key.Cat, kv.Key.From, kv.Key.To, kv.Value))
            .OrderBy(c => c.Category).ThenBy(c => c.From).ToList();
    }

    private static bool IsSharePoint(JsonNode? host) =>
        (host as JsonObject)?["apiId"]?.GetValue<string>()?.EndsWith("shared_sharepointonline", StringComparison.Ordinal) == true;

    private static void RenameColumnKeys(JsonObject prm, List<KeyValuePair<string, string>> cols, Action<string, string, string, int> hit)
    {
        foreach (var key in prm.Select(kv => kv.Key).Where(k => k.StartsWith("item/", StringComparison.Ordinal)).ToList())
        {
            var rest = key["item/".Length..];
            var slash = rest.IndexOf('/');
            var col = slash > 0 ? rest[..slash] : rest;
            var map = cols.FirstOrDefault(c => c.Key == col);
            if (map.Key is null) continue;
            var newKey = "item/" + map.Value + (slash > 0 ? rest[slash..] : "");
            var value = prm[key];
            prm.Remove(key);
            prm[newKey] = value;   // JsonObject keeps insertion order; moved to the end, which is harmless for parameters
            hit("Column", col, map.Value, 1);
        }
    }

    // The column renames that apply to one SharePoint operation: its own list's renames, plus any global rename
    // not overridden for that list. The list is identified by the literal source GUID in the operation's `table`.
    private static List<KeyValuePair<string, string>> EffectiveColumns(
        JsonObject prm, List<KeyValuePair<string, string>> global,
        Dictionary<string, List<KeyValuePair<string, string>>> perList)
    {
        var table = prm["table"] is JsonValue tv && tv.TryGetValue<string>(out var t) ? t : null;
        if (table is null || !perList.TryGetValue(table, out var own) || own.Count == 0) return global;
        var overridden = own.Select(kv => kv.Key).ToHashSet(StringComparer.Ordinal);
        return own.Concat(global.Where(kv => !overridden.Contains(kv.Key))).ToList();
    }

    private static void RewriteQueryParameters(JsonObject prm, List<KeyValuePair<string, string>> cols, Action<string, string, string, int> hit)
    {
        var replace = ColumnReplacer(cols, queryWords: true, hit);
        foreach (var key in new[] { "$filter", "$orderby", "$select" })
            if (prm[key] is JsonValue v && v.TryGetValue<string>(out var s))
            {
                var t = replace(s);
                if (t != s) prm[key] = t;
            }
    }

    // One simultaneous pass, so a rename can never be renamed again (A->B together with B->C stays A->B).
    // Matches 'body/<Column>' reads (optionally followed by /Value, /Id ...) and, for OData query text,
    // bare column names. triggerBody()?['<Column>'] style reads are not touched. The regex is built once per
    // caller, not once per string.
    private static Func<string, string> ColumnReplacer(List<KeyValuePair<string, string>> cols, bool queryWords, Action<string, string, string, int> hit)
    {
        if (cols.Count == 0) return s => s;
        var lookup = cols.GroupBy(c => c.Key, StringComparer.Ordinal).ToDictionary(g => g.Key, g => g.First().Value, StringComparer.Ordinal);
        var names = string.Join("|", lookup.Keys.OrderByDescending(k => k.Length).Select(Regex.Escape));
        var rx = new Regex(queryWords
            ? $@"(?<=body/)(?<n>{names})(?=['/])|(?<![\w])(?<n>{names})(?![\w])"
            : $@"(?<=body/)(?<n>{names})(?=['/])");
        return s => rx.Replace(s, m =>
        {
            var from = m.Groups["n"].Value;
            var to = lookup[from];
            hit("Column", from, to, 1);
            return to;
        });
    }

    private sealed record ValueRule(string Category, string From, string To);

    // Builds one alternation over every active value mapping (longest first, so a URL wins over its parent URL).
    // Each kind gets its own boundaries and case rules:
    //   sites, emails, list/tenant GUIDs: case-insensitive (SharePoint URLs and GUIDs are), whole value only;
    //   form IDs and free text: exact case (form IDs are base64url, where case matters).
    // Email domains apply only to addresses not mapped individually: the address rule matches first, at the
    // start of the address, and consumes it.
    private static (Regex? Rx, List<ValueRule> Rules) BuildValueRules(FlowMapping map)
    {
        var entries = new List<(ValueRule Rule, string Pattern)>();
        void Add(string cat, Dictionary<string, string> m, Func<string, string> pattern)
        {
            foreach (var (from, to) in FlowMapping.Active(m))
                entries.Add((new ValueRule(cat, from, to), pattern(Regex.Escape(from))));
        }

        Add("Site", map.Sites, e => $@"(?i:{e})(?![A-Za-z0-9_~-])");
        Add("List", map.Lists, e => $@"(?<![0-9A-Fa-f-])(?i:{e})(?![0-9A-Fa-f-])");
        Add("Form", map.Forms, e => $@"(?<![A-Za-z0-9_-]){e}(?![A-Za-z0-9_-])");
        Add("Email", map.Emails, e => $@"(?<![\w.%+-])(?i:{e})(?![\w-])");
        Add("Text", map.Text, e => e);
        Add("Tenant", map.TenantIds, e => $@"(?<![0-9A-Fa-f-])(?i:{e})(?![0-9A-Fa-f-])");
        // The tenant GUID also appears without dashes in resource paths (e.g. .../resourceGroups/<tenant>-<user>).
        Add("Tenant (no dashes)", FlowMapping.Active(map.TenantIds)
                .ToDictionary(kv => kv.Key.Replace("-", ""), kv => kv.Value.Replace("-", "")),
            e => $@"(?<![0-9A-Fa-f])(?i:{e})(?![0-9A-Fa-f])");
        Add("EmailDomain", map.EmailDomains, e => $@"(?<=@)(?i:{e})(?![\w-])");

        if (entries.Count == 0) return (null, []);
        var ordered = entries.OrderByDescending(x => x.Rule.From.Length).ToList();
        var pattern = string.Join("|", ordered.Select((x, i) => $"(?<r{i}>{x.Pattern})"));
        return (new Regex(pattern, RegexOptions.CultureInvariant), ordered.Select(x => x.Rule).ToList());
    }

    // Writes a copy of the package with each flow's definition.json rewritten. All other entries are copied as-is.
    // Connections stay unbound by design: they must be recreated in the target tenant on import.
    public static List<(string Flow, List<Change> Changes)> RewritePackage(string srcZip, string dstZip, FlowMapping map)
    {
        var results = new List<(string, List<Change>)>();
        using var src = ZipFile.OpenRead(srcZip);
        using var dstStream = File.Create(dstZip);
        using var dst = new ZipArchive(dstStream, ZipArchiveMode.Create);
        foreach (var entry in src.Entries)
        {
            if (entry.FullName.EndsWith('/')) continue;   // directory marker
            byte[] bytes;
            using (var ms = new MemoryStream()) { using var s = entry.Open(); s.CopyTo(ms); bytes = ms.ToArray(); }

            // Only entries that really hold a flow are rewritten (a Workflows/ file without a definition is copied as-is).
            var json = FlowScanner.IsFlowEntry(entry.FullName)
                ? FlowScanner.TryParseFlow(Encoding.UTF8.GetString(bytes).TrimStart('﻿'))
                : null;
            if (json is not null)
            {
                var hasBom = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF;
                var name = json["properties"]?["displayName"]?.GetValue<string>() ?? FlowScanner.NameFromEntry(entry.FullName);
                var changes = Apply(json, map);
                results.Add((name, changes));
                var text = json.ToJsonString(new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
                bytes = (hasBom ? Encoding.UTF8.GetPreamble() : []).Concat(Encoding.UTF8.GetBytes(text)).ToArray();
            }

            var outEntry = dst.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            using var os = outEntry.Open();
            os.Write(bytes);
        }
        return results;
    }

    // Rescans a rewritten package and returns every value that still equals a source value in the mapping
    // (mapped or not), so the report can flag anything left pointing at the source tenant. Values deliberately
    // mapped to themselves (a site or list that is not moving) are not leftovers, and neither is a URL under a
    // target site that happens to start with a source URL (sites/HR -> sites/HR-New).
    public static List<Finding> FindLeftovers(string rewrittenZip, FlowMapping map)
    {
        static IEnumerable<string> Moving(Dictionary<string, string> m, StringComparison cmp) =>
            m.Where(kv => !string.Equals(kv.Key.TrimEnd('/'), kv.Value.Trim().TrimEnd('/'), cmp)).Select(kv => kv.Key);

        var ignoreCase = new HashSet<string>(
            Moving(map.Lists, StringComparison.OrdinalIgnoreCase)
                .Concat(Moving(map.Emails, StringComparison.OrdinalIgnoreCase))
                .Concat(Moving(map.TenantIds, StringComparison.OrdinalIgnoreCase)), StringComparer.OrdinalIgnoreCase);
        var exact = new HashSet<string>(
            Moving(map.Forms, StringComparison.Ordinal).Concat(Moving(map.Text, StringComparison.Ordinal)), StringComparer.Ordinal);
        var sourceSites = Moving(map.Sites, StringComparison.OrdinalIgnoreCase).Select(k => k.TrimEnd('/')).ToList();
        var targetSites = map.Sites.Values.Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim().TrimEnd('/')).ToList();

        bool IsSourceSite(string value) =>
            sourceSites.Any(k => IsUrlUnder(value, k)) && !targetSites.Any(t => IsUrlUnder(value, t) && t.Length >= LongestSource(value));
        int LongestSource(string value) => sourceSites.Where(k => IsUrlUnder(value, k)).Max(k => k.Length);

        return FlowScanner.ScanPackage(rewrittenZip)
            .SelectMany(s => s.Findings)
            .Where(f => f.Kind != FindingKind.ConnectionRef
                        && (ignoreCase.Contains(f.Value) || exact.Contains(f.Value) || IsSourceSite(f.Value)))
            .ToList();
    }

    // True when url is base itself or something below it (base + '/', '?', '#', quote...), not base + more name.
    private static bool IsUrlUnder(string url, string baseUrl) =>
        url.StartsWith(baseUrl, StringComparison.OrdinalIgnoreCase)
        && (url.Length == baseUrl.Length || !(char.IsLetterOrDigit(url[baseUrl.Length]) || url[baseUrl.Length] is '_' or '-' or '~'));

    public static string ToMarkdown(string packageName, List<(string Flow, List<Change> Changes)> results, List<Finding> leftovers)
    {
        var sb = new StringBuilder($"## {packageName}\n\n");
        foreach (var (flow, changes) in results)
        {
            sb.Append($"### {flow}\n\n");
            if (changes.Count == 0) sb.Append("_").Append(Loc.T("Flows_Rep_NoChanges")).Append("_\n\n");
            foreach (var g in changes.GroupBy(c => c.Category))
            {
                sb.Append("**").Append(g.Key).Append("** (").Append(Loc.T("Flows_Rep_Replacements", g.Sum(c => c.Count))).Append(")\n\n");
                foreach (var c in g) sb.Append($"- `{Short(c.From)}` → `{Short(c.To)}` ×{c.Count}\n");
                sb.Append('\n');
            }
        }
        if (leftovers.Count > 0)
        {
            sb.Append("**").Append(Loc.T("Flows_Rep_StillSource")).Append("**\n\n");
            foreach (var g in leftovers.GroupBy(f => (f.Kind, f.Value)))
                sb.Append($"- {g.Key.Kind}: `{Short(g.Key.Value)}` ×{g.Count()}\n");
            sb.Append('\n');
        }
        return sb.ToString();
    }

    private static string Short(string s) => (s.Length > 110 ? s[..110] + "…" : s).Replace("`", "'").Replace("\n", " ");
}
