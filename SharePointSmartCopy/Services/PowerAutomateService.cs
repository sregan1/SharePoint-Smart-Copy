using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Services;

public sealed record FlowEnvironment(string Name, string DisplayName, bool IsDefault, string? DataverseUrl);

// Json is the full flow as the API returns it ({ name, properties: { displayName, definition, connectionReferences, ... } }),
// the same shape FlowScanner / FlowRewriter already handle for package flows. It is null in list results (fetch with GetFlowAsync).
public sealed record LiveFlow(string Name, string DisplayName, string State, JsonObject? Json);

public sealed record FlowConnection(string Name, string ConnectorName, string DisplayName, bool IsConnected);

public enum FlowApiError
{
    Other,
    // The app registration is missing the delegated permission for this API (AADSTS650057 and similar).
    ConsentMissing,
    // The user lacks a Dataverse security role privilege in the environment (e.g. missing prvReadconnectionreference).
    DataverseRole,
    // The user cannot see or use the environment or flow.
    Forbidden,
    // The target rejected the flow (a list or column that does not exist in the target, bad definition).
    InvalidFlow,
    NotFound,
}

public sealed class FlowApiException : Exception
{
    public FlowApiError Kind { get; }
    public int? Status { get; }
    // Raw service text (English); the dialog maps Kind to a localized message and shows this as detail.
    public string Detail { get; }
    // For DataverseRole: the privilege the service named, e.g. "prvReadconnectionreference".
    public string? MissingPrivilege { get; }

    public FlowApiException(FlowApiError kind, int? status, string detail, string? missingPrivilege = null, Exception? inner = null)
        : base(detail, inner)
    {
        Kind = kind; Status = status; Detail = detail; MissingPrivilege = missingPrivilege;
    }
}

// Reads flows from one Power Automate environment and creates them in another, through the Flow management API
// (api.flow.microsoft.com) and the Power Apps API (connection listings). Tokens come from AuthService
// (GetFlowTokenAsync / GetPowerAppsTokenAsync); the source and target sides each use their own AuthService.
// Proven by Tools/FlowCopyPoC: list, read, rewrite, create (Stopped), list connections, delete.
public sealed partial class PowerAutomateService
{
    private const string FlowApi = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple";
    private const string PowerAppsApi = "https://api.powerapps.com/providers/Microsoft.PowerApps";
    private const string ApiVersion = "api-version=2016-11-01";

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(2) };

    private readonly Func<CancellationToken, Task<string>> _flowToken;
    private readonly Func<CancellationToken, Task<string>> _powerAppsToken;

    public PowerAutomateService(Func<CancellationToken, Task<string>> flowToken, Func<CancellationToken, Task<string>> powerAppsToken)
    {
        _flowToken = flowToken;
        _powerAppsToken = powerAppsToken;
    }

    public static PowerAutomateService For(AuthService auth) =>
        new(ct => auth.GetFlowTokenAsync(ct), ct => auth.GetPowerAppsTokenAsync(ct));

    // The environments the signed-in user can see. Dataverse details need the $expand or they come back empty.
    public async Task<List<FlowEnvironment>> ListEnvironmentsAsync(CancellationToken ct = default)
    {
        var list = new List<FlowEnvironment>();
        foreach (var e in await GetAllAsync($"{FlowApi}/environments?$expand=properties.linkedEnvironmentMetadata", _flowToken, ct))
        {
            var p = e?["properties"];
            var name = e?["name"]?.GetValue<string>();
            if (name is null) continue;
            list.Add(new FlowEnvironment(
                name,
                p?["displayName"]?.GetValue<string>() ?? name,
                p?["isDefault"]?.GetValue<bool>() ?? false,
                p?["linkedEnvironmentMetadata"]?["instanceApiUrl"]?.GetValue<string>()));
        }
        return list;
    }

    // Cloud flows the user owns or can access in the environment (the definition is not included; use GetFlowAsync).
    public async Task<List<LiveFlow>> ListFlowsAsync(string environment, CancellationToken ct = default)
    {
        var list = new List<LiveFlow>();
        foreach (var f in await GetAllAsync($"{FlowApi}/environments/{Uri.EscapeDataString(environment)}/flows", _flowToken, ct))
        {
            var name = f?["name"]?.GetValue<string>();
            if (name is null) continue;
            var p = f?["properties"];
            list.Add(new LiveFlow(name, p?["displayName"]?.GetValue<string>() ?? name, p?["state"]?.GetValue<string>() ?? "", null));
        }
        return list;
    }

    public async Task<LiveFlow> GetFlowAsync(string environment, string flowName, CancellationToken ct = default)
    {
        var body = await SendAsync(HttpMethod.Get,
            $"{FlowApi}/environments/{Uri.EscapeDataString(environment)}/flows/{Uri.EscapeDataString(flowName)}", null, _flowToken, ct);
        var json = JsonNode.Parse(body)?.AsObject() ?? throw new FlowApiException(FlowApiError.Other, null, "Empty flow response.");
        var p = json["properties"];
        return new LiveFlow(flowName, p?["displayName"]?.GetValue<string>() ?? flowName, p?["state"]?.GetValue<string>() ?? "", json);
    }

    // The caller's connections for one connector (e.g. "shared_sharepointonline") in an environment.
    public async Task<List<FlowConnection>> ListConnectionsAsync(string environment, string connectorName, CancellationToken ct = default)
    {
        var url = $"{PowerAppsApi}/apis/{Uri.EscapeDataString(connectorName)}/connections?$filter=environment%20eq%20%27{Uri.EscapeDataString(environment)}%27";
        var list = new List<FlowConnection>();
        foreach (var c in await GetAllAsync(url, _powerAppsToken, ct))
        {
            var name = c?["name"]?.GetValue<string>();
            if (name is null) continue;
            var p = c?["properties"];
            // Connected only when every status says so: a connection can report "Connected" and an error at the same time,
            // and binding a flow to it makes the save fail ("connection has error status").
            var statuses = p?["statuses"]?.AsArray().Select(x => x?["status"]?.GetValue<string>() ?? string.Empty).ToList() ?? [];
            list.Add(new FlowConnection(name, connectorName, p?["displayName"]?.GetValue<string>() ?? name,
                statuses.Count > 0 && statuses.All(st => string.Equals(st, "Connected", StringComparison.OrdinalIgnoreCase))));
        }
        return list;
    }

    // Binds a connection reference to an existing connection in the target environment. Without this the flow is created
    // with the reference as it came from the source and the user fixes connections in Power Automate afterwards.
    public static void BindConnection(JsonObject connectionReferences, string referenceKey, FlowConnection connection)
    {
        if (connectionReferences[referenceKey] is not JsonObject r) return;
        r["connectionName"] = connection.Name;
        r["source"] = "Invoker";
        r.Remove("tier");
        // A flow from a solution names a Dataverse connection reference instead of a connection. Bound to a connection here,
        // that name must go: the target environment may not have that connection reference, and the copy is not in a solution.
        r.Remove("connectionReferenceLogicalName");
        if (r["connection"] is JsonObject inner) inner.Remove("connectionReferenceLogicalName");
        if (r["id"] is null) r["id"] = "/providers/Microsoft.PowerApps/apis/" + connection.ConnectorName;
    }

    // Creates the flow in the target environment, stopped so nothing runs before the user has reviewed it. `rewritten` is
    // the (already rewritten) flow JSON from GetFlowAsync / FlowRewriter.Apply. Returns the new flow's name (id).
    public async Task<string> CreateFlowAsync(string environment, JsonObject rewritten, string? displayNameOverride = null, CancellationToken ct = default)
    {
        var p = rewritten["properties"] ?? throw new FlowApiException(FlowApiError.InvalidFlow, null, "Flow has no properties.");
        var payload = new JsonObject
        {
            ["properties"] = new JsonObject
            {
                ["displayName"] = displayNameOverride ?? p["displayName"]?.GetValue<string>() ?? "Copied flow",
                ["definition"] = p["definition"]?.DeepClone() ?? throw new FlowApiException(FlowApiError.InvalidFlow, null, "Flow has no definition."),
                ["connectionReferences"] = p["connectionReferences"]?.DeepClone() ?? new JsonObject(),
                ["state"] = "Stopped",
            },
        };
        var body = await SendAsync(HttpMethod.Post, $"{FlowApi}/environments/{Uri.EscapeDataString(environment)}/flows",
            payload.ToJsonString(), _flowToken, ct);
        return JsonNode.Parse(body)?["name"]?.GetValue<string>()
               ?? throw new FlowApiException(FlowApiError.Other, null, "Create returned no flow name.");
    }

    public Task DeleteFlowAsync(string environment, string flowName, CancellationToken ct = default) =>
        SendAsync(HttpMethod.Delete, $"{FlowApi}/environments/{Uri.EscapeDataString(environment)}/flows/{Uri.EscapeDataString(flowName)}", null, _flowToken, ct);

    private async Task<string> SendAsync(HttpMethod method, string url, string? json, Func<CancellationToken, Task<string>> token, CancellationToken ct)
    {
        string accessToken;
        try { accessToken = await token(ct); }
        catch (Microsoft.Identity.Client.MsalServiceException ex) when (ex.Message.Contains("AADSTS650057") || ex.Message.Contains("AADSTS65001"))
        {
            throw new FlowApiException(FlowApiError.ConsentMissing, null, ex.Message, inner: ex);
        }

        // Throttling (429) and brief service trouble (502/503/504) are retried a few times, waiting as long as the service
        // asks (Retry-After) or a little longer each time. A request that times out is reported as an error, not as the
        // user's cancel.
        var full = url.Contains("api-version=", StringComparison.Ordinal) ? url : url + (url.Contains('?') ? "&" : "?") + ApiVersion;
        for (var attempt = 1; ; attempt++)
        {
            using var req = new HttpRequestMessage(method, full);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            if (json != null) req.Content = new StringContent(json, Encoding.UTF8, "application/json");
            HttpResponseMessage resp;
            try { resp = await Http.SendAsync(req, ct); }
            catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
            {
                throw new FlowApiException(FlowApiError.Other, null, Loc.T("Flows_Live_ErrTimeout"), inner: ex);
            }
            catch (HttpRequestException ex)
            {
                // A dropped connection is retried only for a read: a create (POST) may already have gone through on the server,
                // and sending it again would make a second flow.
                if (method == HttpMethod.Get && attempt < 4) { await Task.Delay(TimeSpan.FromSeconds(2 * attempt), ct); continue; }
                throw new FlowApiException(FlowApiError.Other, null, ex.Message, inner: ex);
            }
            using (resp)
            {
                var text = await resp.Content.ReadAsStringAsync(ct);
                if (resp.IsSuccessStatusCode) return text;
                var code = (int)resp.StatusCode;
                // 429 means the request was not processed, so any method can be sent again. A gateway error (502/503/504) says
                // nothing about whether a create (POST) went through, so only reads are retried on those.
                var retryable = code == 429 || (code is 502 or 503 or 504 && method == HttpMethod.Get);
                if (retryable && attempt < 4)
                {
                    var wait = resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(5 * attempt);
                    if (wait > TimeSpan.FromSeconds(60)) wait = TimeSpan.FromSeconds(60);
                    await Task.Delay(wait, ct);
                    continue;
                }
                throw Translate(code, text);
            }
        }
    }

    // Every item of a listing: follows "nextLink" until the service stops returning one (a large tenant has more than one page).
    private async Task<List<JsonNode?>> GetAllAsync(string url, Func<CancellationToken, Task<string>> token, CancellationToken ct)
    {
        var all = new List<JsonNode?>();
        var next = url;
        for (var page = 0; next is not null && page < 200; page++)
        {
            var root = JsonNode.Parse(await SendAsync(HttpMethod.Get, next, null, token, ct));
            all.AddRange(root?["value"]?.AsArray() ?? []);
            next = root?["nextLink"]?.GetValue<string>() ?? root?["@odata.nextLink"]?.GetValue<string>();
        }
        return all;
    }

    [GeneratedRegex(@"missing (prv\w+) privilege")]
    private static partial Regex MissingPrivilegeRegex();

    // Maps a service error to a kind the dialog can explain. The service text stays in Detail for the report.
    internal static FlowApiException Translate(int status, string body)
    {
        var detail = body;
        try { detail = JsonNode.Parse(body)?["error"]?["message"]?.GetValue<string>() ?? body; } catch (JsonException) { }

        var priv = MissingPrivilegeRegex().Match(body);
        if (priv.Success)
            return new FlowApiException(FlowApiError.DataverseRole, status, detail, priv.Groups[1].Value);
        if (body.Contains("InvalidOpenApiFlow", StringComparison.Ordinal))
            return new FlowApiException(FlowApiError.InvalidFlow, status, detail);
        return status switch
        {
            404 => new FlowApiException(FlowApiError.NotFound, status, detail),
            401 or 403 => new FlowApiException(FlowApiError.Forbidden, status, detail),
            _ => new FlowApiException(FlowApiError.Other, status, detail),
        };
    }
}
