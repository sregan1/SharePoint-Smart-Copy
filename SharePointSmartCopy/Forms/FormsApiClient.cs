using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;

namespace SharePointSmartCopy.Forms;

// Thin wrapper over the unofficial forms.office.com/formapi endpoints. The API is undocumented and can change,
// so callers work with raw JSON and treat unexpected shapes as warnings, not crashes.
public sealed record FormSummary(string Id, string Title, string Type, DateTimeOffset? Modified);

public sealed class FormsApiClient
{
    // One client for the whole app: HttpClient is meant to be shared, and every dialog action creates a FormsApiClient.
    private static readonly HttpClient SharedHttp = new() { BaseAddress = new Uri("https://forms.office.com/") };
    private const int MaxAttempts = 4;

    private readonly HttpClient _http;
    private readonly Func<CancellationToken, bool, Task<string>> _tokenProvider;

    // The provider's bool is "force a fresh token" (set after a 401 so an expired token is replaced once).
    public FormsApiClient(Func<CancellationToken, bool, Task<string>> tokenProvider, HttpClient? http = null)
    {
        _tokenProvider = tokenProvider;
        _http = http ?? SharedHttp;
    }

    public async Task<JsonDocument> GetJsonAsync(string relativeUrl, CancellationToken ct = default)
    {
        var (status, body) = await SendRawAsync(HttpMethod.Get, relativeUrl, null, ct);
        return ParseOrThrow(status, body, HttpMethod.Get, relativeUrl);
    }

    public async Task<JsonDocument> PostJsonAsync(string relativeUrl, string jsonBody, CancellationToken ct = default)
    {
        var (status, body) = await SendRawAsync(HttpMethod.Post, relativeUrl, jsonBody, ct);
        return ParseOrThrow(status, body, HttpMethod.Post, relativeUrl);
    }

    // Reads every page of an OData collection ("value" arrays chained by @odata.nextLink).
    public async Task<List<System.Text.Json.Nodes.JsonObject>> GetAllPagesAsync(string relativeUrl, string what, CancellationToken ct = default)
    {
        var items = new List<System.Text.Json.Nodes.JsonObject>();
        string? url = relativeUrl;
        while (url is not null)
        {
            var (status, body) = await SendRawAsync(HttpMethod.Get, url, null, ct);
            if (status != 200)
                throw new HttpRequestException($"{what} failed: {status} {(body.Length > 300 ? body[..300] : body)}", null, (System.Net.HttpStatusCode)status);
            var page = System.Text.Json.Nodes.JsonNode.Parse(body)!.AsObject();
            foreach (var f in page["value"]?.AsArray() ?? [])
                if (f is System.Text.Json.Nodes.JsonObject o) items.Add(o);
            url = page["@odata.nextLink"]?.GetValue<string>();
            if (url is not null && _http.BaseAddress is { } b && url.StartsWith(b.ToString(), StringComparison.OrdinalIgnoreCase))
                url = url[b.ToString().Length..];
        }
        return items;
    }

    // The title of one form by its ID, or null when it cannot be read (no access, deleted, or an unexpected answer). The
    // list below holds only the account's own forms, so this is how the title of a form shared with the account is found.
    public async Task<string?> GetFormTitleAsync(string formId, CancellationToken ct = default)
    {
        var (status, body) = await SendRawAsync(HttpMethod.Get, $"formapi/api/forms('{formId}')", null, ct);
        if (status != 200) return null;
        try { return System.Text.Json.Nodes.JsonNode.Parse(body)?["title"]?.GetValue<string>(); }
        catch (JsonException) { return null; }
        catch (InvalidOperationException) { return null; }
    }

    // The signed-in user's own forms (not group-owned ones, not forms shared with them), newest first as the
    // service returns them. Soft-deleted entries (recycle bin / API-created drafts) are skipped.
    public async Task<List<FormSummary>> ListFormsAsync(CancellationToken ct = default)
    {
        var forms = new List<FormSummary>();
        foreach (var o in await GetAllPagesAsync("formapi/api/forms", "Listing forms", ct))
        {
            if (o["softDeleted"]?.GetValue<int>() is > 0) continue;
            DateTimeOffset? modified = DateTimeOffset.TryParse(o["modifiedDate"]?.GetValue<string>(), out var m) ? m : null;
            forms.Add(new FormSummary(
                o["id"]!.GetValue<string>(),
                o["title"]?.GetValue<string>() ?? "",
                o["type"]?.GetValue<string>() ?? "form",
                modified));
        }
        return forms;
    }

    // Any method; never throws on HTTP errors, returns status and raw body. Throttling (429), transient server
    // errors (502/503/504) and an expired token (401, retried once with a fresh token) are retried here, so a
    // bulk copy doesn't turn a brief throttle into a lost question or setting.
    public async Task<(int Status, string Body)> SendRawAsync(HttpMethod method, string relativeUrl, string? jsonBody = null, CancellationToken ct = default)
    {
        var refreshed = false;
        var forceToken = false;
        for (var attempt = 1; ; attempt++)
        {
            using var req = new HttpRequestMessage(method, relativeUrl);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await _tokenProvider(ct, forceToken));
            req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            if (jsonBody != null) req.Content = new StringContent(jsonBody, System.Text.Encoding.UTF8, "application/json");
            forceToken = false;

            using var resp = await _http.SendAsync(req, ct);
            var status = (int)resp.StatusCode;
            var body = await resp.Content.ReadAsStringAsync(ct);

            if (status == 401 && !refreshed)
            {
                refreshed = true;
                forceToken = true;
                continue;
            }
            if (status is 429 or 502 or 503 or 504 && attempt < MaxAttempts)
            {
                var delay = resp.Headers.RetryAfter?.Delta
                    ?? (resp.Headers.RetryAfter?.Date is { } when ? when - DateTimeOffset.UtcNow : (TimeSpan?)null)
                    ?? TimeSpan.FromSeconds(Math.Pow(2, attempt));
                await Task.Delay(delay < TimeSpan.Zero ? TimeSpan.Zero : delay > TimeSpan.FromSeconds(60) ? TimeSpan.FromSeconds(60) : delay, ct);
                continue;
            }
            return (status, body);
        }
    }

    private static JsonDocument ParseOrThrow(int status, string body, HttpMethod method, string url)
    {
        if (status is < 200 or >= 300)
            throw new HttpRequestException($"{status} for {method} {url}: {body}", null, (System.Net.HttpStatusCode)status);
        return JsonDocument.Parse(string.IsNullOrWhiteSpace(body) ? "{}" : body);
    }
}
