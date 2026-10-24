using SharePointSmartCopy.Localization;
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

    // The account that completed interactive sign-in. Silent acquisition is pinned to it so a second
    // account in the MSAL cache can never be picked by accident (matters once two tenants sign in).
    private string? _pinnedAccountId;

    private async Task<IAccount?> GetPinnedAccountAsync()
    {
        var accounts = (await _app!.GetAccountsAsync()).ToList();
        if (_pinnedAccountId != null)
            return accounts.FirstOrDefault(a => a.HomeAccountId.Identifier == _pinnedAccountId);
        return accounts.FirstOrDefault();
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
        _pinnedAccountId = null;
        FormsUserName = null;
        _tokenCache.Clear();
    }

    public async Task<string> GetAccessTokenAsync(bool forceInteractive = false, CancellationToken cancellationToken = default)
    {
        if (_app == null)
            throw new InvalidOperationException(Loc.T("Svc_Auth_NotConfiguredClientId"));

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

                var account = await GetPinnedAccountAsync();
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
            _pinnedAccountId = _authResult.Account?.HomeAccountId?.Identifier;
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
            throw new InvalidOperationException(Loc.T("Svc_Auth_NotConfiguredSignIn"));

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

            var account = await GetPinnedAccountAsync();
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
            // Re-pin to whoever just signed in (same as the Graph path) so a stale or removed pinned
            // account can't make every later refresh fall back to an interactive prompt.
            _pinnedAccountId = authResult.Account?.HomeAccountId?.Identifier ?? _pinnedAccountId;
            _tokenCache[cacheKey] = authResult;
            return authResult.AccessToken;
        }
        finally { _refreshLock.Release(); }
    }

    // Token for the (unofficial) Microsoft Forms web API, audience forms.office.com. Delegated only: the
    // API rejects app-only tokens. The app registration needs the Microsoft Forms delegated permission
    // consented (and Group.Read.All for group-owned forms); without it MSAL surfaces a consent error that
    // the Forms dialog translates for the user. Same single-flight / pinned-account behavior as the other
    // token getters; the access token is an encrypted JWE, so its claims can't be read client-side.
    private const string FormsScope = "https://forms.office.com/.default";

    // The account the Forms token was issued to (shown on the Forms target step); null until a token is acquired.
    public string? FormsUserName { get; private set; }

    public async Task<string> GetFormsTokenAsync(CancellationToken cancellationToken = default, bool forceRefresh = false)
    {
        if (_app == null)
            throw new InvalidOperationException(Loc.T("Svc_Auth_NotConfiguredSignIn"));

        var scopes   = new[] { FormsScope };
        var cacheKey = FormsScope;

        if (!forceRefresh && TryGetCachedToken(cacheKey, out var fastToken))
            return fastToken;

        await _refreshLock.WaitAsync(cancellationToken);
        try
        {
            if (!forceRefresh && TryGetCachedToken(cacheKey, out var cachedToken))
                return cachedToken;

            var account = await GetPinnedAccountAsync();
            if (account != null)
            {
                try
                {
                    var result = await _app.AcquireTokenSilent(scopes, account)
                        .WithForceRefresh(forceRefresh)
                        .ExecuteAsync(cancellationToken);
                    _tokenCache[cacheKey] = result;
                    FormsUserName = result.Account?.Username ?? FormsUserName;
                    return result.AccessToken;
                }
                catch (MsalUiRequiredException) { /* fall through to interactive */ }
            }

            // Silent acquisition fails the first time (new resource = new consent). Reuse the account already
            // signed in so the user sees at most a consent screen, not the account picker again.
            var interactive = _app.AcquireTokenInteractive(scopes);
            interactive = account != null
                ? interactive.WithAccount(account)
                : interactive.WithPrompt(Prompt.SelectAccount);
            var authResult = await interactive.ExecuteAsync(cancellationToken);
            // Pin only when nothing is pinned yet. Re-pinning here would let an account picked on this prompt
            // silently replace the one the SharePoint/Graph sign-in uses (the same AuthService in single-tenant mode).
            _pinnedAccountId ??= authResult.Account?.HomeAccountId?.Identifier;
            _tokenCache[cacheKey] = authResult;
            FormsUserName = authResult.Account?.Username ?? FormsUserName;
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
        _pinnedAccountId = null;
        FormsUserName = null;
        _tokenCache.Clear();
    }

    public bool IsAuthenticated => _authResult != null;
    public bool IsConfigured => _app != null;
    public string? UserName => _authResult?.Account?.Username;
}
