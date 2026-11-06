using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace SharePointSmartCopy.Services;

// Copies a managed-metadata term set (group, set, terms, labels, descriptions, properties) between term stores
// through Microsoft Graph (/sites/{site}/termStore). Graph assigns every group, set and term GUID itself and
// does not accept chosen ones, so a copy always has new GUIDs; the returned map (source term GUID -> target
// term GUID) lets column bindings and item values be remapped. Copying into a new group inside ONE term store
// exercises exactly the same logic as a cross-tenant copy.
//
// Needs the delegated Graph permission TermStore.ReadWrite.All (target) / TermStore.Read.All (source), with
// admin consent, and a term store role for the signed-in user at the target.
public sealed class TermStoreCopyService
{
    public const string ReadScope  = "TermStore.Read.All";
    public const string WriteScope = "TermStore.ReadWrite.All";

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };
    private const string Base = "https://graph.microsoft.com/v1.0";

    private readonly AuthService _sourceAuth;
    private readonly AuthService _targetAuth;

    public TermStoreCopyService(AuthService sourceAuth, AuthService targetAuth)
    {
        _sourceAuth = sourceAuth;
        _targetAuth = targetAuth;
    }

    public sealed record TermNode(
        string Id,
        List<(string Name, string Lang, bool IsDefault)> Labels,
        List<(string Text, string Lang)> Descriptions,
        List<(string Key, string Value)> Props,
        List<TermNode> Children);

    public sealed record TermSetTree(
        string SetId, string SetName, string GroupName, string GroupScope,
        List<(string Name, string Lang)> LocalizedNames, List<TermNode> Roots);

    public sealed class CopyOutcome
    {
        public string TargetStoreId { get; set; } = "";
        public string TargetSetId   { get; set; } = "";
        public string SetName       { get; set; } = "";
        public int Created  { get; set; }
        public int Existing { get; set; }
        public List<string> Notes { get; } = [];
    }

    // The permission or role is missing: the message tells the user what to grant.
    public sealed class TermStoreAccessException(string message) : Exception(message);

    // ── Reading ──────────────────────────────────────────────────────────────

    public async Task<TermSetTree> ReadSetAsync(string sourceSiteUrl, string setId, CancellationToken ct = default)
    {
        // The same registration signs in for both sides in single-tenant mode, so ask for the one scope it already
        // consented to instead of prompting for a second one.
        var token = await _sourceAuth.GetGraphTokenAsync([ReferenceEquals(_sourceAuth, _targetAuth) ? WriteScope : ReadScope], ct);
        var site = await ResolveSiteIdAsync(sourceSiteUrl, token, ct);
        var set = await GetJsonAsync($"{Base}/sites/{site}/termStore/sets/{setId}?$expand=parentGroup", token, ct);

        var names = new List<(string, string)>();
        if (set["localizedNames"] is JsonArray ln)
            foreach (var n in ln) names.Add((n?["name"]?.GetValue<string>() ?? "", n?["languageTag"]?.GetValue<string>() ?? ""));
        var group = set["parentGroup"];
        var tree = new TermSetTree(setId,
            names.FirstOrDefault().Item1 ?? setId,
            group?["displayName"]?.GetValue<string>() ?? "",
            group?["scope"]?.GetValue<string>() ?? "global",
            names, []);

        tree.Roots.AddRange(await ReadChildrenAsync($"{Base}/sites/{site}/termStore/sets/{setId}/children", site, setId, token, ct));
        return tree;
    }

    private async Task<List<TermNode>> ReadChildrenAsync(string url, string site, string setId, string token, CancellationToken ct)
    {
        var result = new List<TermNode>();
        foreach (var t in await GetAllPagesAsync(url + "?$select=id,labels,descriptions,properties", token, ct))
        {
            var id = t["id"]?.GetValue<string>() ?? "";
            var node = new TermNode(id, [], [], [], []);
            if (t["labels"] is JsonArray labels)
                foreach (var l in labels)
                    node.Labels.Add((l?["name"]?.GetValue<string>() ?? "", l?["languageTag"]?.GetValue<string>() ?? "",
                                     l?["isDefault"]?.GetValue<bool>() ?? false));
            if (t["descriptions"] is JsonArray ds)
                foreach (var d in ds)
                    node.Descriptions.Add((d?["description"]?.GetValue<string>() ?? "", d?["languageTag"]?.GetValue<string>() ?? ""));
            if (t["properties"] is JsonArray ps)
                foreach (var p in ps)
                    node.Props.Add((p?["key"]?.GetValue<string>() ?? "", p?["value"]?.GetValue<string>() ?? ""));
            node.Children.AddRange(await ReadChildrenAsync(
                $"{Base}/sites/{site}/termStore/sets/{setId}/terms/{id}/children", site, setId, token, ct));
            result.Add(node);
        }
        return result;
    }

    // ── Writing ──────────────────────────────────────────────────────────────

    // Makes sure the set (and its terms) exist at the target and fills `map` with source -> target term GUIDs.
    // Idempotent: groups, sets and terms that are already there (matched by name / default label) are reused.
    // `targetGroupName` overrides the source group's name (used to copy within one term store).
    public async Task<CopyOutcome> EnsureSetAsync(
        string targetSiteUrl, TermSetTree tree, IDictionary<string, string> map,
        string? targetGroupName = null, CancellationToken ct = default)
    {
        var token = await _targetAuth.GetGraphTokenAsync([WriteScope], ct);
        var site  = await ResolveSiteIdAsync(targetSiteUrl, token, ct);
        var outcome = new CopyOutcome { SetName = tree.SetName };

        var store = await GetJsonAsync($"{Base}/sites/{site}/termStore", token, ct);
        outcome.TargetStoreId = store["id"]?.GetValue<string>() ?? "";

        var siteScoped = tree.GroupScope.Equals("siteCollection", StringComparison.OrdinalIgnoreCase) && targetGroupName == null;
        var groupId = await EnsureGroupAsync(site, siteScoped ? null : (targetGroupName ?? tree.GroupName), siteScoped, token, outcome, ct);

        // The set: matched by name inside the group.
        string? setId = null;
        foreach (var s in await GetAllPagesAsync($"{Base}/sites/{site}/termStore/groups/{groupId}/sets", token, ct))
        {
            var existingName = (s["localizedNames"] as JsonArray)?.FirstOrDefault()?["name"]?.GetValue<string>();
            if (string.Equals(existingName, tree.SetName, StringComparison.OrdinalIgnoreCase))
            { setId = s["id"]?.GetValue<string>(); break; }
        }
        if (setId == null)
        {
            var created = await PostJsonAsync($"{Base}/sites/{site}/termStore/sets", token, new JsonObject
            {
                ["localizedNames"] = new JsonArray(tree.LocalizedNames.Count > 0
                    ? tree.LocalizedNames.Select(n => (JsonNode)new JsonObject { ["name"] = n.Name, ["languageTag"] = n.Lang }).ToArray()
                    : [new JsonObject { ["name"] = tree.SetName, ["languageTag"] = "en-US" }]),
                ["parentGroup"] = new JsonObject { ["id"] = groupId },
            }, ct);
            setId = created["id"]?.GetValue<string>() ?? throw new Exception("The term store did not return the new term set.");
            outcome.Created++;
        }
        outcome.TargetSetId = setId;

        await EnsureTermsAsync(site, setId, null, tree.Roots, token, map, outcome, ct);
        return outcome;
    }

    private async Task<string> EnsureGroupAsync(string site, string? name, bool siteScoped, string token, CopyOutcome outcome, CancellationToken ct)
    {
        var groups = await GetAllPagesAsync($"{Base}/sites/{site}/termStore/groups", token, ct);
        foreach (var g in groups)
        {
            var scope = g["scope"]?.GetValue<string>() ?? "global";
            if (siteScoped ? scope.Equals("siteCollection", StringComparison.OrdinalIgnoreCase)
                           : string.Equals(g["displayName"]?.GetValue<string>(), name, StringComparison.OrdinalIgnoreCase))
                return g["id"]?.GetValue<string>() ?? "";
        }
        var body = new JsonObject { ["displayName"] = name ?? "Site Collection", ["scope"] = siteScoped ? "siteCollection" : "global" };
        var created = await PostJsonAsync($"{Base}/sites/{site}/termStore/groups", token, body, ct);
        outcome.Created++;
        return created["id"]?.GetValue<string>() ?? throw new Exception("The term store did not return the new term group.");
    }

    private async Task EnsureTermsAsync(string site, string setId, string? parentTermId, List<TermNode> nodes,
        string token, IDictionary<string, string> map, CopyOutcome outcome, CancellationToken ct)
    {
        var url = parentTermId == null
            ? $"{Base}/sites/{site}/termStore/sets/{setId}/children"
            : $"{Base}/sites/{site}/termStore/sets/{setId}/terms/{parentTermId}/children";

        // Existing children, by default label.
        var existing = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in await GetAllPagesAsync(url + "?$select=id,labels", token, ct))
        {
            var id = t["id"]?.GetValue<string>() ?? "";
            if (t["labels"] is JsonArray labels)
                foreach (var l in labels)
                    if (l?["isDefault"]?.GetValue<bool>() == true && l["name"]?.GetValue<string>() is { } n)
                        existing.TryAdd(n, id);
        }

        foreach (var node in nodes)
        {
            ct.ThrowIfCancellationRequested();
            var defaultName = node.Labels.FirstOrDefault(l => l.IsDefault).Name ?? node.Labels.FirstOrDefault().Name ?? "";
            string targetId;
            if (existing.TryGetValue(defaultName, out var found))
            {
                targetId = found;
                outcome.Existing++;
            }
            else
            {
                var body = new JsonObject
                {
                    ["labels"] = new JsonArray(node.Labels.Select(l => (JsonNode)new JsonObject
                        { ["name"] = l.Name, ["languageTag"] = l.Lang, ["isDefault"] = l.IsDefault }).ToArray()),
                };
                if (node.Descriptions.Count > 0)
                    body["descriptions"] = new JsonArray(node.Descriptions.Select(d => (JsonNode)new JsonObject
                        { ["description"] = d.Text, ["languageTag"] = d.Lang }).ToArray());
                if (node.Props.Count > 0)
                    body["properties"] = new JsonArray(node.Props.Select(p => (JsonNode)new JsonObject
                        { ["key"] = p.Key, ["value"] = p.Value }).ToArray());
                try
                {
                    var created = await PostJsonAsync(url, token, body, ct);
                    targetId = created["id"]?.GetValue<string>() ?? "";
                    outcome.Created++;
                }
                catch (Exception ex) when (ex is not OperationCanceledException and not TermStoreAccessException)
                {
                    // Typically a label in a language the target term store does not offer.
                    outcome.Notes.Add($"Term '{defaultName}' was not created: {ex.Message}");
                    continue;
                }
            }
            if (targetId.Length > 0) map[node.Id.Trim('{', '}')] = targetId;
            if (node.Children.Count > 0 && targetId.Length > 0)
                await EnsureTermsAsync(site, setId, targetId, node.Children, token, map, outcome, ct);
        }
    }

    // ── Schema rewrite ───────────────────────────────────────────────────────

    // Points a managed-metadata column's schema at the copied term set: SspId and TermSetId are replaced,
    // AnchorId is mapped when the column is bound to a branch, and the source's hidden-field GUIDs are dropped.
    public static string RewriteSchema(string schemaXml, string targetStoreId, string targetSetId, string? sourceSetId,
        IReadOnlyDictionary<string, string> map)
    {
        var el = System.Xml.Linq.XElement.Parse(schemaXml);
        foreach (var prop in el.Descendants("Property"))
        {
            var name = (string?)prop.Element("Name");
            var value = prop.Element("Value");
            if (value == null) continue;
            switch (name)
            {
                case "SspId":     value.Value = targetStoreId; break;
                case "TermSetId": value.Value = targetSetId;   break;
                case "AnchorId":
                    var anchor = value.Value.Trim('{', '}');
                    value.Value = map.TryGetValue(anchor, out var mapped) ? mapped : Guid.Empty.ToString();
                    break;
            }
        }
        return el.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    // The set GUID a managed-metadata column is bound to (null when the schema has none).
    public static string? ReadTermSetId(string? schemaXml)
    {
        if (string.IsNullOrEmpty(schemaXml)) return null;
        try
        {
            var el = System.Xml.Linq.XElement.Parse(schemaXml);
            return el.Descendants("Property")
                .Where(p => (string?)p.Element("Name") == "TermSetId")
                .Select(p => p.Element("Value")?.Value.Trim('{', '}')).FirstOrDefault();
        }
        catch { return null; }
    }

    // ── HTTP ─────────────────────────────────────────────────────────────────

    // The site id Graph wants in /sites/{id}/termStore (the "host,siteGuid,webGuid" form), read from the site URL.
    public static async Task<string> ResolveSiteIdAsync(string siteUrl, string token, CancellationToken ct)
    {
        var site = await GetJsonAsync($"{Base}/sites/{SiteKey(siteUrl)}?$select=id", token, ct);
        return site["id"]?.GetValue<string>() ?? throw new Exception("Microsoft Graph did not return the site id for " + siteUrl);
    }

    // "host:/sites/x:" style key Graph accepts in place of a site id.
    private static string SiteKey(string siteUrl)
    {
        var u = new Uri(siteUrl.TrimEnd('/'));
        var path = u.AbsolutePath.TrimEnd('/');
        return string.IsNullOrEmpty(path) ? u.Host : $"{u.Host}:{path}:";
    }

    private static async Task<JsonNode> GetJsonAsync(string url, string token, CancellationToken ct)
        => await SendAsync(HttpMethod.Get, url, token, null, ct);

    private static async Task<JsonNode> PostJsonAsync(string url, string token, JsonNode body, CancellationToken ct)
        => await SendAsync(HttpMethod.Post, url, token, body, ct);

    private static async Task<List<JsonNode>> GetAllPagesAsync(string url, string token, CancellationToken ct)
    {
        var all = new List<JsonNode>();
        string? next = url;
        while (next != null)
        {
            var page = await GetJsonAsync(next, token, ct);
            if (page["value"] is JsonArray arr)
                foreach (var n in arr) if (n != null) all.Add(n);
            next = page["@odata.nextLink"]?.GetValue<string>();
        }
        return all;
    }

    private static async Task<JsonNode> SendAsync(HttpMethod method, string url, string token, JsonNode? body, CancellationToken ct)
    {
        for (var attempt = 0; ; attempt++)
        {
            using var req = new HttpRequestMessage(method, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            if (body != null) req.Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json");
            using var resp = await Http.SendAsync(req, ct);
            var text = await resp.Content.ReadAsStringAsync(ct);

            if ((int)resp.StatusCode == 429 && attempt < 6)
            {
                var wait = resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(2 * (attempt + 1));
                await Task.Delay(wait, ct);
                continue;
            }
            if (resp.StatusCode is System.Net.HttpStatusCode.Forbidden or System.Net.HttpStatusCode.Unauthorized)
                throw new TermStoreAccessException(
                    "The term store refused the request. The app registration needs the delegated Microsoft Graph permission " +
                    "TermStore.ReadWrite.All (with admin consent), and the signed-in account needs a term store role at the target. " +
                    "See the README for details.");
            if (!resp.IsSuccessStatusCode)
                throw new Exception($"Term store HTTP {(int)resp.StatusCode}: {(text.Length > 200 ? text[..200] : text)}");
            return string.IsNullOrWhiteSpace(text) ? new JsonObject() : JsonNode.Parse(text) ?? new JsonObject();
        }
    }
}
