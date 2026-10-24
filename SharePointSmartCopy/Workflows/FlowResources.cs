using System.Text.RegularExpressions;

namespace SharePointSmartCopy.Workflows;

// How the tool deals with a service a flow uses.
public enum ResourceHandling
{
    Rewritten,        // the tool rewrites what it finds (SharePoint sites/lists/columns, Forms, email addresses)
    ReviewManually,   // identifiers were detected but are not rewritten; check them in Power Automate after import
    ConnectionOnly,   // nothing tenant-specific beyond the connection, which must be recreated anyway
}

// One row of the "what these flows use" overview.
public sealed record ResourceSummary(string Connector, string Service, int Flows, int Operations, string Detail, ResourceHandling Handling);

public static class FlowResources
{
    // apiName ("shared_xxx") -> display name, and how the tool handles it.
    private static readonly Dictionary<string, (string Name, ResourceHandling Handling)> Known = new(StringComparer.OrdinalIgnoreCase)
    {
        ["shared_sharepointonline"]       = ("SharePoint", ResourceHandling.Rewritten),
        ["shared_microsoftforms"]         = ("Microsoft Forms", ResourceHandling.Rewritten),
        ["shared_office365"]              = ("Outlook (Office 365)", ResourceHandling.Rewritten),
        ["shared_office365users"]         = ("Office 365 Users", ResourceHandling.ConnectionOnly),
        ["shared_approvals"]              = ("Approvals", ResourceHandling.ConnectionOnly),
        ["shared_onedriveforbusiness"]    = ("OneDrive for Business", ResourceHandling.ReviewManually),
        ["shared_teams"]                  = ("Microsoft Teams", ResourceHandling.ReviewManually),
        ["shared_planner"]                = ("Planner", ResourceHandling.ReviewManually),
        ["shared_todo"]                   = ("Microsoft To Do", ResourceHandling.ReviewManually),
        ["shared_excelonlinebusiness"]    = ("Excel Online (Business)", ResourceHandling.ReviewManually),
        ["shared_wordonlinebusiness"]     = ("Word Online (Business)", ResourceHandling.ReviewManually),
        ["shared_onenote"]                = ("OneNote (Business)", ResourceHandling.ReviewManually),
        ["shared_powerbi"]                = ("Power BI", ResourceHandling.ReviewManually),
        ["shared_sharepointonlinesitecontent"] = ("SharePoint", ResourceHandling.Rewritten),
    };

    // connectors whose literal IDs the scanner reports (as OtherResource) so the user knows to review them
    public static bool IsReviewConnector(string connector) =>
        Known.TryGetValue(connector, out var k) ? k.Handling == ResourceHandling.ReviewManually : !IsBuiltIn(connector);

    private static bool IsBuiltIn(string connector) => Known.ContainsKey(connector);

    public static string FriendlyName(string connector)
    {
        if (Known.TryGetValue(connector, out var k)) return k.Name;
        var n = connector.StartsWith("shared_", StringComparison.OrdinalIgnoreCase) ? connector["shared_".Length..] : connector;
        return n.Length == 0 ? connector : char.ToUpperInvariant(n[0]) + n[1..];
    }

    public static ResourceHandling HandlingOf(string connector) =>
        Known.TryGetValue(connector, out var k) ? k.Handling : ResourceHandling.ReviewManually;

    // "shared_sharepointonline_2" -> "shared_sharepointonline" (a flow can hold several connections to one connector)
    public static string BaseConnector(string apiNameOrConnection)
    {
        var s = apiNameOrConnection;
        var m = Regex.Match(s, @"^(shared_[a-z0-9]+?)(?:[_-]\d+)?$", RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[1].Value : s;
    }

    // A literal that looks like an identifier of a Teams team/channel, Planner plan, Excel file, OneDrive item...
    // (a GUID, a Teams "19:...@thread" id, a Graph drive id "b!...", or a long unbroken token).
    public static bool LooksLikeResourceId(string v)
    {
        if (string.IsNullOrWhiteSpace(v) || v.StartsWith('@') || v.Contains(' ')) return false;
        if (Guid.TryParse(v, out _)) return true;
        if (v.StartsWith("19:", StringComparison.Ordinal) || v.StartsWith("b!", StringComparison.Ordinal)) return true;
        return v.Length >= 22 && !v.Contains('@') && !v.Contains("://") && Regex.IsMatch(v, @"^[A-Za-z0-9_\-=!:.]+$");
    }

    public static List<ResourceSummary> Summarize(IEnumerable<FlowScan> scans)
    {
        var list = scans.ToList();
        var rows = new List<ResourceSummary>();
        var connectors = list.SelectMany(s => s.Connectors.Keys).Select(BaseConnector).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        foreach (var connector in connectors.OrderBy(c => HandlingOf(c)).ThenBy(c => FriendlyName(c), StringComparer.OrdinalIgnoreCase))
        {
            var usingFlows = list.Where(s => s.Connectors.Keys.Any(k => BaseConnector(k).Equals(connector, StringComparison.OrdinalIgnoreCase))).ToList();
            var ops = usingFlows.Sum(s => s.Connectors.Where(kv => BaseConnector(kv.Key).Equals(connector, StringComparison.OrdinalIgnoreCase)).Sum(kv => kv.Value));
            var findings = usingFlows.SelectMany(s => s.Findings).ToList();

            int Distinct(params FindingKind[] kinds) =>
                findings.Where(f => kinds.Contains(f.Kind)).Select(f => f.Value).Distinct(StringComparer.OrdinalIgnoreCase).Count();

            string detail = connector.ToLowerInvariant() switch
            {
                "shared_sharepointonline" => $"{Distinct(FindingKind.SiteUrl)} sites, {Distinct(FindingKind.ListId)} lists, {Distinct(FindingKind.ColumnName)} columns",
                "shared_microsoftforms"   => $"{Distinct(FindingKind.FormId)} forms",
                "shared_office365"        => $"{Distinct(FindingKind.Email)} email addresses, {Distinct(FindingKind.Mailbox)} shared mailboxes",
                "shared_office365users"   => "user lookups (no tenant-specific IDs)",
                "shared_approvals"        => "approvers are email addresses (see Outlook)",
                _ => OtherDetail(connector, usingFlows),
            };

            rows.Add(new ResourceSummary(connector, FriendlyName(connector), usingFlows.Count, ops, detail, HandlingOf(connector)));
        }
        return rows;
    }

    private static string OtherDetail(string connector, List<FlowScan> flows)
    {
        var ids = flows.SelectMany(s => s.Findings)
            .Where(f => f.Kind == FindingKind.OtherResource && f.Detail is not null
                        && f.Detail.StartsWith(FriendlyName(connector) + ":", StringComparison.OrdinalIgnoreCase))
            .Select(f => f.Value).Distinct().Count();
        return ids > 0 ? $"{ids} resource IDs found (not rewritten; review after import)" : "no resource IDs found";
    }
}
