using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;

namespace SharePointSmartCopy.Services;

// An unmanaged solution in the target environment that a copied flow can be created in. An empty UniqueName is the
// "My flows (not in a solution)" choice. Prefix is the solution publisher's customization prefix (new connection references
// are named with it).
public sealed record FlowSolution(string UniqueName, string DisplayName, string? Prefix);

// Creates a copied flow INSIDE a solution of a Dataverse-enabled environment (the Flow API can only create flows outside any
// solution). A solution flow names its connections through Dataverse connection references, so each of the flow's
// references is pointed at one in the target: the same logical name if the target has it, otherwise an existing reference for
// that connector that is bound to a connection, otherwise a new reference created in the chosen solution (bound to one of the
// user's connections when there is one). Proven on a live environment with Tools/FlowCopyPoC (dvsol / dvread).
public sealed class DataverseFlowService
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(2) };
    private readonly string _org;
    private readonly Func<CancellationToken, Task<string>> _token;

    public DataverseFlowService(string orgUrl, Func<CancellationToken, Task<string>> token)
    {
        _org = orgUrl.TrimEnd('/');
        _token = token;
    }

    // The visible, unmanaged solutions a flow can be added to (the system "Default" solution is left out).
    public async Task<List<FlowSolution>> ListSolutionsAsync(CancellationToken ct = default)
    {
        var body = await SendAsync(HttpMethod.Get,
            "solutions?$filter=isvisible eq true and ismanaged eq false&$select=uniquename,friendlyname&$expand=publisherid($select=customizationprefix)&$orderby=friendlyname",
            null, null, ct);
        var list = new List<FlowSolution>();
        foreach (var s in JsonNode.Parse(body)?["value"]?.AsArray() ?? [])
        {
            var unique = s?["uniquename"]?.GetValue<string>();
            if (string.IsNullOrEmpty(unique) || unique.Equals("Default", StringComparison.OrdinalIgnoreCase)) continue;
            list.Add(new FlowSolution(unique, s?["friendlyname"]?.GetValue<string>() ?? unique,
                s?["publisherid"]?["customizationprefix"]?.GetValue<string>()));
        }
        return list;
    }

    private sealed record ConnectionReference(string LogicalName, string Connector, string? ConnectionId);

    private async Task<List<ConnectionReference>> ListConnectionReferencesAsync(CancellationToken ct)
    {
        var body = await SendAsync(HttpMethod.Get, "connectionreferences?$select=connectionreferencelogicalname,connectorid,connectionid&$filter=statecode eq 0", null, null, ct);
        var list = new List<ConnectionReference>();
        foreach (var r in JsonNode.Parse(body)?["value"]?.AsArray() ?? [])
        {
            var logical = r?["connectionreferencelogicalname"]?.GetValue<string>();
            var connectorId = r?["connectorid"]?.GetValue<string>() ?? string.Empty;
            if (string.IsNullOrEmpty(logical)) continue;
            list.Add(new ConnectionReference(logical, connectorId[(connectorId.LastIndexOf('/') + 1)..], r?["connectionid"]?.GetValue<string>()));
        }
        return list;
    }

    private async Task<string> CreateConnectionReferenceAsync(FlowSolution solution, string connector, string? connectionId, CancellationToken ct)
    {
        var prefix = string.IsNullOrWhiteSpace(solution.Prefix) ? "new" : solution.Prefix;
        var logical = $"{prefix}_{connector.Replace("_", string.Empty)}_{Guid.NewGuid().ToString("N")[..5]}";
        var row = new JsonObject
        {
            ["connectionreferencelogicalname"] = logical,
            ["connectionreferencedisplayname"] = $"{connector} ({solution.DisplayName})",
            ["connectorid"] = "/providers/Microsoft.PowerApps/apis/" + connector,
        };
        if (!string.IsNullOrEmpty(connectionId)) row["connectionid"] = connectionId;
        await SendAsync(HttpMethod.Post, "connectionreferences", row.ToJsonString(), solution.UniqueName, ct);
        return logical;
    }

    // Creates the (already rewritten) flow in the solution, off (draft), and returns its id (the same id the Flow API uses) and
    // how many of its connection references are not bound to a connection yet. connectionFor gives one of the user's
    // connected connections for a connector, or null; it is only asked when a new reference is needed and bind is true.
    public async Task<(string Id, int Unbound)> CreateSolutionFlowAsync(FlowSolution solution, JsonObject flow,
        Func<string, Task<string?>> connectionFor, bool bind, CancellationToken ct = default)
    {
        var p = flow["properties"] ?? throw new FlowApiException(FlowApiError.InvalidFlow, null, "Flow has no properties.");
        var definition = p["definition"]?.DeepClone() ?? throw new FlowApiException(FlowApiError.InvalidFlow, null, "Flow has no definition.");
        var existing = await ListConnectionReferencesAsync(ct);
        var refs = new JsonObject();
        var unbound = 0;
        foreach (var (key, value) in (p["connectionReferences"] as JsonObject ?? new JsonObject()).ToList())
        {
            var apiId = value?["id"]?.GetValue<string>() ?? value?["api"]?["name"]?.GetValue<string>() ?? key;
            var connector = apiId[(apiId.LastIndexOf('/') + 1)..];
            var sourceLogical = value?["connectionReferenceLogicalName"]?.GetValue<string>()
                                ?? value?["connection"]?["connectionReferenceLogicalName"]?.GetValue<string>();
            // Reuse order: the source's own logical name; a reference for the connector that is bound to a connection; any
            // reference for the connector (an unbound one is still the right thing to point at: it is set once, in Power
            // Automate, instead of a new reference being made for every copied flow).
            var match = existing.FirstOrDefault(r => sourceLogical is not null && r.LogicalName.Equals(sourceLogical, StringComparison.OrdinalIgnoreCase))
                        ?? existing.FirstOrDefault(r => r.Connector.Equals(connector, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(r.ConnectionId))
                        ?? existing.FirstOrDefault(r => r.Connector.Equals(connector, StringComparison.OrdinalIgnoreCase));
            string logical;
            if (match is not null)
            {
                logical = match.LogicalName;
                if (string.IsNullOrEmpty(match.ConnectionId)) unbound++;
            }
            else
            {
                var connectionId = bind ? await connectionFor(connector) : null;
                logical = await CreateConnectionReferenceAsync(solution, connector, connectionId, ct);
                existing.Add(new ConnectionReference(logical, connector, connectionId));
                if (connectionId is null) unbound++;
            }
            refs[key] = new JsonObject
            {
                ["runtimeSource"] = "embedded",
                ["connection"] = new JsonObject { ["connectionReferenceLogicalName"] = logical },
                ["api"] = new JsonObject { ["name"] = connector },
            };
        }

        var clientdata = new JsonObject
        {
            ["properties"] = new JsonObject { ["connectionReferences"] = refs, ["definition"] = definition },
            ["schemaVersion"] = "1.0.0.0",
        };
        var row = new JsonObject
        {
            ["category"] = 5,                 // modern (cloud) flow
            ["type"] = 1,                     // definition
            ["primaryentity"] = "none",
            ["statecode"] = 0,                // draft = off, like the Flow API copies (created stopped)
            ["name"] = p["displayName"]?.GetValue<string>() ?? "Copied flow",
            ["clientdata"] = clientdata.ToJsonString(),
        };
        var (_, entityId) = await SendWithIdAsync(HttpMethod.Post, "workflows", row.ToJsonString(), solution.UniqueName, ct);
        var id = entityId is null ? null : entityId[(entityId.LastIndexOf('(') + 1)..].TrimEnd(')');
        if (string.IsNullOrEmpty(id)) throw new FlowApiException(FlowApiError.Other, null, "Create returned no flow id.");
        return (id, unbound);
    }

    private async Task<string> SendAsync(HttpMethod method, string path, string? json, string? solution, CancellationToken ct) =>
        (await SendWithIdAsync(method, path, json, solution, ct)).Body;

    private async Task<(string Body, string? EntityId)> SendWithIdAsync(HttpMethod method, string path, string? json, string? solution, CancellationToken ct)
    {
        string token;
        try { token = await _token(ct); }
        catch (Microsoft.Identity.Client.MsalServiceException ex) when (ex.Message.Contains("AADSTS650057") || ex.Message.Contains("AADSTS65001"))
        {
            throw new FlowApiException(FlowApiError.ConsentMissing, null, ex.Message, inner: ex);
        }
        for (var attempt = 1; ; attempt++)
        {
            using var req = new HttpRequestMessage(method, $"{_org}/api/data/v9.2/{path}");
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            req.Headers.Add("OData-Version", "4.0");
            if (solution is not null) req.Headers.Add("MSCRM.SolutionUniqueName", solution);
            if (json is not null) req.Content = new StringContent(json, Encoding.UTF8, "application/json");
            HttpResponseMessage resp;
            try { resp = await Http.SendAsync(req, ct); }
            catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
            {
                throw new FlowApiException(FlowApiError.Other, null, Localization.Loc.T("Flows_Live_ErrTimeout"), inner: ex);
            }
            using (resp)
            {
                var text = await resp.Content.ReadAsStringAsync(ct);
                if (resp.IsSuccessStatusCode)
                    return (text, resp.Headers.TryGetValues("OData-EntityId", out var v) ? v.FirstOrDefault() : null);
                var code = (int)resp.StatusCode;
                // 429 = not processed, safe to send again. A gateway error says nothing about whether a POST went through, so
                // only reads are retried on those (a second POST would create a second flow or reference).
                var retryable = code == 429 || (code is 502 or 503 or 504 && method == HttpMethod.Get);
                if (retryable && attempt < 4)
                {
                    var wait = resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(5 * attempt);
                    await Task.Delay(wait > TimeSpan.FromSeconds(60) ? TimeSpan.FromSeconds(60) : wait, ct);
                    continue;
                }
                throw PowerAutomateService.Translate(code, text);
            }
        }
    }
}
