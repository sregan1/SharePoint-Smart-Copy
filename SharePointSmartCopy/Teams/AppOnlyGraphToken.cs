using Microsoft.Identity.Client;

namespace SharePointSmartCopy.Teams;

// An application-only (client credentials) Microsoft Graph token for the Teams message import. Microsoft allows the import
// API (startMigration, import message, completeMigration) only with the application permission Teamwork.Migrate.All, so
// the signed-in user's token cannot be used: this signs in as the app registration itself, with its client secret.
public sealed class AppOnlyGraphToken
{
    private readonly IConfidentialClientApplication _app;

    // tenantId must be the real directory ID of the TARGET tenant ("common" does not work for an application sign-in).
    private readonly string _role;

    // role names the side in the error text ("target" for the import, "source" for the administrator read).
    public AppOnlyGraphToken(string tenantId, string clientId, string clientSecret, string role = "target")
    {
        _role = role;
        _app = ConfidentialClientApplicationBuilder.Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
            .Build();
    }

    public async Task<string> GetAsync(CancellationToken ct, bool forceRefresh = false)
    {
        try
        {
            var result = await _app.AcquireTokenForClient(["https://graph.microsoft.com/.default"])
                .WithForceRefresh(forceRefresh)
                .ExecuteAsync(ct);
            return result.AccessToken;
        }
        catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS7000215", StringComparison.Ordinal) || ex.ErrorCode == "invalid_client")
        {
            // The most common slip: the secret's ID was pasted instead of its Value (the Value is shown only once, when it is created).
            throw new InvalidOperationException(
                $"Microsoft rejected the client secret saved for the {_role} registration (AADSTS7000215). In Settings, paste the secret's VALUE, " +
                "not its Secret ID, into that registration's client secret box. The Value is shown only when the secret is created, so if you " +
                "did not copy it then, create a new secret under Certificates & secrets (and check it has not expired). The secret must belong " +
                $"to the same registration that is chosen as the {_role}.", ex);
        }
    }

    // The application permissions (roles) the token actually carries: shows at a glance what the registration was granted.
    public async Task<string[]> GetRolesAsync(CancellationToken ct)
    {
        var parts = (await GetAsync(ct, true)).Split('.');
        if (parts.Length < 2) return [];
        var payload = parts[1].Replace('-', '+').Replace('_', '/');
        payload = payload.PadRight(payload.Length + (4 - payload.Length % 4) % 4, '=');
        using var doc = System.Text.Json.JsonDocument.Parse(Convert.FromBase64String(payload));
        return doc.RootElement.TryGetProperty("roles", out var r) ? r.EnumerateArray().Select(x => x.GetString() ?? "").ToArray() : [];
    }

    // The same text naming the target registration, so it is clear which one to change in the Azure portal.
    public static string AccessMessageFor(string registration) =>
        "Microsoft Teams refused the message import. The target app registration " + registration + " needs the Microsoft Graph APPLICATION permissions " +
        "Teamwork.Migrate.All (and, for chats, ChatMember.ReadWrite.All), added under API permissions as type Application, with admin consent granted. " +
        "These are separate from the delegated permissions. The client secret saved in Settings must belong to that same registration. " +
        "After granting consent, wait a minute and run the copy again: it continues into the same team.";

    public const string AccessMessage =
        "Microsoft Teams refused the message import. It needs the application permissions Teamwork.Migrate.All " +
        "(and, for chats, ChatMember.ReadWrite.All) (Microsoft Graph, Application type, with admin consent) on the target app " +
        "registration, and that registration's client secret saved in Settings. See the README for details.";
}
