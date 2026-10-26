using SharePointSmartCopy.Localization;
using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Http.HttpClientLibrary;
using SharePointSmartCopy.Models;
using SpColumnDef  = SharePointSmartCopy.Models.ColumnDefinition;
using SpColumnMap  = SharePointSmartCopy.Models.ColumnMapping;

namespace SharePointSmartCopy.Services;

// Partial class SharePointService: Permissions. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
    // ── Permissions ───────────────────────────────────────────────────────────

    // Checks whether a single SP object has unique (broken) permissions.
    // apiPath examples: "web/lists('{guid}')", "web"
    // Returns null when the answer could not be read (throttling that outlasted the retries, a 401/403, a network
    // error) so the caller can tell "inherits" from "we could not find out". Reporting false there silently skipped
    // the permission copy under load. Cancellation is never swallowed.
    public async Task<bool?> TryGetHasUniqueRoleAssignmentsAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}?$select=HasUniqueRoleAssignments";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return null;
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("HasUniqueRoleAssignments", out var v) ? v.GetBoolean() : null;
        }
        catch (Exception ex) when (ex is not OperationCanceledException) { return null; }
    }

    // Kept for callers that only need a yes/no and treat "unknown" as "no".
    public async Task<bool> GetHasUniqueRoleAssignmentsAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
        => await TryGetHasUniqueRoleAssignmentsAsync(siteUrl, apiPath, ct) == true;

    // Returns null on a genuine READ failure (distinct from a real empty list) so the caller can
    // tell "this item truly has no unique role assignments" from "we couldn't find out" — the two
    // used to look identical (both `[]`), so a transient failure reading the SOURCE permissions
    // silently reported "0 applied, no error" instead of skipping the item to retry later, and the
    // target was left with only inherited permissions with nothing in the log to explain why.
    public async Task<List<RoleAssignmentInfo>?> GetRoleAssignmentsAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/roleassignments" +
                  "?$expand=Member,RoleDefinitionBindings&$select=Member/Id,Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return null;
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("value", out var vals)) return [];

            var result = new List<RoleAssignmentInfo>();
            foreach (var el in vals.EnumerateArray())
            {
                if (!el.TryGetProperty("Member", out var member)) continue;
                var pid        = member.TryGetProperty("Id",            out var idEl)   ? idEl.GetInt32()          : 0;
                var loginName  = member.TryGetProperty("LoginName",     out var lnEl)   ? lnEl.GetString() ?? ""   : "";
                var title      = member.TryGetProperty("Title",         out var tEl)    ? tEl.GetString()  ?? ""   : "";
                var ptype      = member.TryGetProperty("PrincipalType", out var ptEl)   ? ptEl.GetInt32()          : 0;
                var roleNames  = new List<string>();
                if (el.TryGetProperty("RoleDefinitionBindings", out var rdbs))
                    foreach (var rdb in rdbs.EnumerateArray())
                        if (rdb.TryGetProperty("Name", out var nm))
                            roleNames.Add(nm.GetString() ?? "");
                result.Add(new RoleAssignmentInfo(pid, ptype, loginName, title, roleNames));
            }
            return result;
        }
        catch { return null; }
    }

    public async Task BreakPermissionInheritanceAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)";
        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(string.Empty, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct);
        if (!resp.IsSuccessStatusCode)
        {
            var body = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException(Loc.T("Svc_Sp_BreakPermissionInheritanceFailed", (int)resp.StatusCode, body));
        }
    }

    public async Task AddRoleAssignmentAsync(
        string siteUrl, string apiPath, int principalId, int roleDefId, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/roleassignments/addroleassignment(principalid={principalId},roleDefId={roleDefId})";
        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(string.Empty, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct);
        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException(Loc.T("Svc_Sp_AddRoleAssignmentFailed", (int)resp.StatusCode, errBody));
        }
    }

    public async Task<int?> EnsureUserAsync(
        string siteUrl, string loginName, CancellationToken ct = default)
    {
        // Cache successes only — the set of distinct emails needing resolution across a library is
        // typically tens, so this collapses thousands of redundant POSTs to one per account. A
        // failure is NOT cached: it could be a transient throttle/error rather than "this account
        // doesn't exist on this site", and caching it would permanently misreport every later folder
        // sharing that author/editor as unresolvable (the same failure class as the lookup-value
        // cache poisoning bug this app has hit before).
        var cacheKey = $"{siteUrl}|{loginName}";
        if (_ensureUserCache.TryGetValue(cacheKey, out var cachedId)) return cachedId;

        var url  = $"{siteUrl.TrimEnd('/')}/_api/web/ensureuser";
        var body = JsonSerializer.Serialize(new { logonName = loginName });
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return null;
            var respBody = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(respBody);
            int? id = doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetInt32() : null;
            if (id != null) _ensureUserCache[cacheKey] = id;
            return id;
        }
        catch { return null; }
    }

    // "Not found" results, remembered briefly. EnsureUserAsync deliberately never caches a failure (it could
    // be a throttle), but across thousands of files the same unmapped author would cost one failing round
    // trip per file. A short memory collapses that; it is short so a transient failure heals quickly.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, DateTimeOffset> _missingUsersUntil
        = new(StringComparer.OrdinalIgnoreCase);
    private static readonly TimeSpan MissingUserMemory = TimeSpan.FromSeconds(30);

    // Does this account (email or claims login) exist in the TARGET tenant? Resolving it also registers it
    // on the target site's user list, which the copy needs anyway before it can stamp or grant it.
    internal async Task<bool> TargetUserExistsAsync(string siteUrl, string login, CancellationToken ct = default)
    {
        var key = $"{siteUrl}|{login}";
        if (_missingUsersUntil.TryGetValue(key, out var until) && until > DateTimeOffset.UtcNow) return false;
        if (await EnsureUserAsync(siteUrl, login, ct) != null)
        {
            _missingUsersUntil.TryRemove(key, out _);
            return true;
        }
        _missingUsersUntil[key] = DateTimeOffset.UtcNow + MissingUserMemory;
        return false;
    }

    // Cross-tenant: a person column naming someone who does not exist in the target makes the whole
    // ValidateUpdateListItem call fail (the commit is all-or-nothing), taking every other field and the dates
    // down with it. Check each person first: the missing ones are recorded for the report and left out, and
    // everything else still lands. Returns null when nobody is left. Same-tenant: returns the value untouched.
    private async Task<object?> DropUnresolvedPeopleAsync(string siteUrl, object value, CancellationToken ct)
    {
        if (!IsCrossTenant || value is not PersonFieldValue person) return value;
        var keep = new List<string>();
        foreach (var login in person.Logins)
        {
            var mapped = Identity.MapLogin(login);
            if (await TargetUserExistsAsync(siteUrl, mapped, ct)) keep.Add(login);
            else Unresolved.Record(IdentityMapper.DisplayIdentity(login, null), IdentityMapper.DisplayIdentity(mapped, null), "person column");
        }
        return keep.Count == 0 ? null : new PersonFieldValue([.. keep]);
    }

    public async Task<int?> GetSiteGroupIdAsync(
        string siteUrl, string groupTitle, CancellationToken ct = default)
    {
        var encoded = Uri.EscapeDataString(groupTitle.Replace("'", "''"));
        var url     = $"{siteUrl.TrimEnd('/')}/_api/web/sitegroups/getbyname('{encoded}')?$select=Id";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return null;
            var respBody = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(respBody);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetInt32() : null;
        }
        catch { return null; }
    }

    // Loads all role definitions for the site and returns a name→ID dictionary.
    public async Task<Dictionary<string, int>> GetAllRoleDefinitionsAsync(
        string siteUrl, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/roledefinitions?$select=Id,Name";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("value", out var vals)) return [];
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var el in vals.EnumerateArray())
            {
                var id   = el.TryGetProperty("Id",   out var idEl)   ? idEl.GetInt32()        : 0;
                var name = el.TryGetProperty("Name", out var nameEl) ? nameEl.GetString() ?? "" : "";
                if (id > 0 && !string.IsNullOrEmpty(name))
                    result[name] = id;
            }
            return result;
        }
        catch { return []; }
    }

    private sealed class TokenProvider : IAccessTokenProvider
    {
        private readonly TenantRouter _router;
        public TokenProvider(TenantRouter router) => _router = router;

        public AllowedHostsValidator AllowedHostsValidator { get; } =
            new AllowedHostsValidator(["graph.microsoft.com"]);

        public async Task<string> GetAuthorizationTokenAsync(
            Uri uri,
            Dictionary<string, object>? additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
            => await _router.ForGraphUri(uri).GetAccessTokenAsync(cancellationToken: cancellationToken);
    }
}
