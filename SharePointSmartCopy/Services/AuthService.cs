using Microsoft.Identity.Client;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class AuthService
{
    private IPublicClientApplication? _app;
    private AuthenticationResult? _authResult;

    // In-memory access-token cache keyed by scope-set. MSAL's own cache makes AcquireTokenSilent
    // a no-network call, but GetAccountsAsync + AcquireTokenSilent still do real work (cache
    // deserialization, account enumeration, internal locking) on EVERY call — and at 23k files ×
    // many requests/file under high parallelism that lock contention serializes the workers.
    // Returning a still-valid cached token collapses the hot path to a dictionary read.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, AuthenticationResult> _tokenCache = new();
    private static readonly TimeSpan TokenRefreshSkew = TimeSpan.FromMinutes(5);

    // Single-flight gate for the cache-miss path. Every Graph/REST request funnels through the
    // token getters; at the 5-minute expiry boundary dozens of workers used to miss the cache
    // simultaneously and each run its own GetAccountsAsync + AcquireTokenSilent — and if silent
    // refresh failed mid-job (CA policy, revoked session), each one fell through to
    // AcquireTokenInteractive, spawning multiple browser prompts and MSAL concurrent-interactive
    // exceptions. One refresher runs; the rest wait and pick up its cached result.
    private readonly SemaphoreSlim _refreshLock = new(1, 1);

    private bool TryGetCachedToken(string key, out string accessToken)
    {
        if (_tokenCache.TryGetValue(key, out var r) && r.ExpiresOn > DateTimeOffset.UtcNow + TokenRefreshSkew)
        {
            accessToken = r.AccessToken;
            return true;
        }
        accessToken = string.Empty;
        return false;
    }

    private readonly string[] _scopes =
    [
        "Sites.ReadWrite.All",
        "Files.ReadWrite.All",
        "offline_access"
    ];

    public void Configure(AppSettings settings)
    {
        var reg = settings.ActiveRegistration;
        if (reg == null) return;
        Configure(reg);
    }

    public void Configure(AzureRegistration registration)
    {
        _app = PublicClientApplicationBuilder
            .Create(registration.ClientId)
            .WithAuthority($"https://login.microsoftonline.com/{(string.IsNullOrWhiteSpace(registration.TenantId) ? "common" : registration.TenantId)}")
            .WithRedirectUri("http://localhost")
            .Build();
        _authResult = null;
        _tokenCache.Clear();
    }

    public async Task<string> GetAccessTokenAsync(bool forceInteractive = false, CancellationToken cancellationToken = default)
    {
        if (_app == null)
            throw new InvalidOperationException("Auth service not configured. Please set Client ID in Settings.");

        var cacheKey = string.Join(' ', _scopes);
        if (!forceInteractive && TryGetCachedToken(cacheKey, out var fastToken))
            return fastToken;

        await _refreshLock.WaitAsync(cancellationToken);
        try
        {
            if (!forceInteractive)
            {
                // Re-check after acquiring: another waiter may have refreshed while we queued.
                if (TryGetCachedToken(cacheKey, out var cachedToken))
                    return cachedToken;

                var accounts = await _app.GetAccountsAsync();
                var account = accounts.FirstOrDefault();
                if (account != null)
                {
                    try
                    {
                        _authResult = await _app.AcquireTokenSilent(_scopes, account).ExecuteAsync(cancellationToken);
                        _tokenCache[cacheKey] = _authResult;
                        return _authResult.AccessToken;
                    }
                    catch (MsalUiRequiredException) { /* fall through to interactive */ }
                }
            }

            _authResult = await _app.AcquireTokenInteractive(_scopes)
                .WithPrompt(Prompt.SelectAccount)
                .ExecuteAsync(cancellationToken);
            _tokenCache[cacheKey] = _authResult;
            return _authResult.AccessToken;
        }
        finally { _refreshLock.Release(); }
    }

    // Returns a token scoped for the SharePoint REST API (audience = tenant.sharepoint.com).
    // The Graph token from GetAccessTokenAsync cannot be used for /_api/ endpoints.
    // spScope: SharePoint-specific permission name, e.g. "Sites.ReadWrite.All" or "AllSites.FullControl".
    // The Azure AD app must have the requested permission registered and admin-consented.
    public async Task<string> GetSharePointTokenAsync(string siteUrl, string spScope = "Sites.ReadWrite.All", CancellationToken cancellationToken = default, bool forceRefresh = false)
    {
        if (_app == null)
            throw new InvalidOperationException("Auth service not configured. Please sign in first.");

        var uri    = new Uri(siteUrl.TrimEnd('/'));
        var scopes = new[] { $"{uri.Scheme}://{uri.Host}/{spScope}" };
        var cacheKey = scopes[0];

        if (!forceRefresh && TryGetCachedToken(cacheKey, out var fastToken))
            return fastToken;

        await _refreshLock.WaitAsync(cancellationToken);
        try
        {
            // Re-check after acquiring: another waiter may have refreshed while we queued.
            if (!forceRefresh && TryGetCachedToken(cacheKey, out var cachedToken))
                return cachedToken;

            var accounts = await _app.GetAccountsAsync();
            var account  = accounts.FirstOrDefault();
            if (account != null)
            {
                try
                {
                    var result = await _app.AcquireTokenSilent(scopes, account)
                        .WithForceRefresh(forceRefresh)
                        .ExecuteAsync(cancellationToken);
                    _tokenCache[cacheKey] = result;
                    return result.AccessToken;
                }
                catch (MsalUiRequiredException) { /* fall through to interactive */ }
            }

            var authResult = await _app.AcquireTokenInteractive(scopes)
                .WithPrompt(Prompt.SelectAccount)
                .ExecuteAsync(cancellationToken);
            _tokenCache[cacheKey] = authResult;
            return authResult.AccessToken;
        }
        finally { _refreshLock.Release(); }
    }

    public async Task SignOutAsync()
    {
        if (_app == null) return;
        foreach (var account in await _app.GetAccountsAsync())
            await _app.RemoveAsync(account);
        _authResult = null;
        _tokenCache.Clear();
    }

    public bool IsAuthenticated => _authResult != null;
    public bool IsConfigured => _app != null;
    public string? UserName => _authResult?.Account?.Username;
}
