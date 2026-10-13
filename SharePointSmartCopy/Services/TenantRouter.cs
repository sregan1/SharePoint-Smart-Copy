using System.Collections.Concurrent;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Services;

// Decides which AuthService (and therefore which tenant's token) a request should use.
//
// Single-tenant mode: nothing is registered, so every lookup returns the default AuthService —
// behavior is identical to a single shared AuthService.
//
// Multi-tenant mode: the source and target sides each have their own AuthService (own app
// registration, authority and account). The router maps a SharePoint host, or a Graph drive ID, to
// the AuthService of the tenant that owns it, so the existing Graph/REST call sites keep their
// signatures and the right token is picked per request from the URL.
public sealed class TenantRouter
{
    private readonly ConcurrentDictionary<string, AuthService> _hosts  = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, AuthService> _drives = new(StringComparer.Ordinal);

    // Set by Scope() so a Graph $batch call (whose URL carries no drive or site) still picks the
    // right tenant's token. AsyncLocal flows through the awaits inside the Graph request pipeline.
    private readonly AsyncLocal<AuthService?> _scoped = new();

    public TenantRouter(AuthService defaultAuth) => Default = defaultAuth;

    public AuthService Default { get; }

    // True once at least one host is registered. Until then every lookup falls back to Default.
    public bool IsMultiTenant => !_hosts.IsEmpty;

    public void RegisterHost(string siteUrlOrHost, AuthService auth)
        => _hosts[HostOf(siteUrlOrHost)] = auth;

    public void RegisterDrive(string driveId, AuthService auth)
    {
        if (!string.IsNullOrEmpty(driveId)) _drives[driveId] = auth;
    }

    // Registers a drive under whichever tenant owns the given site URL.
    public void RegisterDriveForSite(string driveId, string siteUrl)
    {
        if (IsMultiTenant) RegisterDrive(driveId, ForSiteUrl(siteUrl));
    }

    public void Clear()
    {
        _hosts.Clear();
        _drives.Clear();
    }

    public AuthService ForSiteUrl(string siteUrl)
    {
        var host = HostOf(siteUrl);
        if (_hosts.TryGetValue(host, out var auth)) return auth;
        // Once tenants are registered, an unknown host must not silently use the source sign-in.
        if (IsMultiTenant) throw new InvalidOperationException(Loc.T("Svc_Sp_HostNotConnected", host));
        return Default;
    }

    public AuthService ForDrive(string driveId)
    {
        if (_drives.TryGetValue(driveId, out var auth)) return auth;
        if (!IsMultiTenant) return Default;
        throw new InvalidOperationException(Loc.T("Svc_Sp_DriveNotConnected", driveId));
    }

    public AuthService ForGraphUri(Uri uri)
    {
        if (_scoped.Value is { } scoped) return scoped;
        if (!IsMultiTenant) return Default;

        var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
        // Path is "/v1.0/<collection>/<id>/..." — skip the version segment.
        for (var i = 0; i + 1 < segments.Length; i++)
        {
            var collection = segments[i];
            if (collection.Equals("drives", StringComparison.OrdinalIgnoreCase))
                return ForDrive(Uri.UnescapeDataString(segments[i + 1]));
            if (collection.Equals("sites", StringComparison.OrdinalIgnoreCase))
            {
                // Site keys are "host", "host:/path:", or "host,siteGuid,webGuid".
                var key  = Uri.UnescapeDataString(segments[i + 1]);
                var host = key.Split(',', ':')[0];
                if (_hosts.TryGetValue(host, out var auth)) return auth;
                // A real hostname that isn't registered: refuse rather than guess the tenant.
                // ("root" and bare ids carry no host, so they fall through to Default.)
                if (host.Contains('.')) throw new InvalidOperationException(Loc.T("Svc_Sp_HostNotConnected", host));
            }
        }
        return Default;
    }

    // Everything awaited inside the scope (e.g. a Graph $batch POST) uses this tenant's token.
    public IDisposable Scope(AuthService auth)
    {
        var previous = _scoped.Value;
        _scoped.Value = auth;
        return new ScopeToken(this, previous);
    }

    private static string HostOf(string siteUrlOrHost)
        => Uri.TryCreate(siteUrlOrHost, UriKind.Absolute, out var u) ? u.Host : siteUrlOrHost.Trim().TrimEnd('/');

    private sealed class ScopeToken(TenantRouter owner, AuthService? previous) : IDisposable
    {
        public void Dispose() => owner._scoped.Value = previous;
    }
}
