using Microsoft.Identity.Client;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class AuthService
{
    private IPublicClientApplication? _app;
    private AuthenticationResult? _authResult;

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
    }

    public async Task<string> GetAccessTokenAsync(bool forceInteractive = false, CancellationToken cancellationToken = default)
    {
        if (_app == null)
            throw new InvalidOperationException("Auth service not configured. Please set Client ID in Settings.");

        if (!forceInteractive)
        {
            var accounts = await _app.GetAccountsAsync();
            var account = accounts.FirstOrDefault();
            if (account != null)
            {
                try
                {
                    _authResult = await _app.AcquireTokenSilent(_scopes, account).ExecuteAsync(cancellationToken);
                    return _authResult.AccessToken;
                }
                catch (MsalUiRequiredException) { /* fall through to interactive */ }
            }
        }

        _authResult = await _app.AcquireTokenInteractive(_scopes)
            .WithPrompt(Prompt.SelectAccount)
            .ExecuteAsync(cancellationToken);
        return _authResult.AccessToken;
    }

    // Returns a token scoped for the SharePoint REST API (audience = tenant.sharepoint.com).
    // The Graph token from GetAccessTokenAsync cannot be used for /_api/ endpoints.
    public async Task<string> GetSharePointTokenAsync(string siteUrl, CancellationToken cancellationToken = default)
    {
        if (_app == null)
            throw new InvalidOperationException("Auth service not configured. Please sign in first.");

        var uri    = new Uri(siteUrl.TrimEnd('/'));
        var scopes = new[] { $"{uri.Scheme}://{uri.Host}/Sites.ReadWrite.All" };

        var accounts = await _app.GetAccountsAsync();
        var account  = accounts.FirstOrDefault();
        if (account != null)
        {
            try
            {
                var result = await _app.AcquireTokenSilent(scopes, account).ExecuteAsync(cancellationToken);
                return result.AccessToken;
            }
            catch (MsalUiRequiredException) { /* fall through to interactive */ }
        }

        var authResult = await _app.AcquireTokenInteractive(scopes)
            .WithPrompt(Prompt.SelectAccount)
            .ExecuteAsync(cancellationToken);
        return authResult.AccessToken;
    }

    public async Task SignOutAsync()
    {
        if (_app == null) return;
        foreach (var account in await _app.GetAccountsAsync())
            await _app.RemoveAsync(account);
        _authResult = null;
    }

    public bool IsAuthenticated => _authResult != null;
    public bool IsConfigured => _app != null;
    public string? UserName => _authResult?.Account?.Username;
}
