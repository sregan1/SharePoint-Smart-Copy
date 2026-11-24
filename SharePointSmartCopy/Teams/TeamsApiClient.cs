using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;

namespace SharePointSmartCopy.Teams;

// Thin Microsoft Graph client for the Teams tool. Phase 1 is read-only discovery: teams, channels (with their
// membership type), members, channel files folders and chats. Writes (migration-mode import) are added later.
//
// The token provider is called as (ct, forceRefresh) so one 401 can be retried with a fresh token, the same
// shape FormsApiClient uses. Throttling (429) and transient server errors (503/504) honor Retry-After.
public sealed class TeamsApiClient
{
    public const string TeamsReadScope   = "Team.ReadBasic.All";
    public const string ChannelReadScope = "Channel.ReadBasic.All";
    public const string MemberReadScope  = "TeamMember.Read.All";
    public const string MessageReadScope = "ChannelMessage.Read.All";
    public const string ChatReadScope    = "Chat.Read";

    // Everything the discovery dialog asks for in one consent.
    public static readonly string[] ReadScopes = [TeamsReadScope, ChannelReadScope, MemberReadScope, ChatReadScope];

    // What copying a team additionally needs: create the team and its channels, add members, and set the team's settings.
    // Asked for only when a copy starts, so browsing stays a read-only consent.
    public static readonly string[] WriteScopes =
        ["Team.Create", "Channel.Create", "ChannelMember.ReadWrite.All", "TeamMember.ReadWrite.All", "TeamSettings.ReadWrite.All"];
    public static readonly string[] CopyScopes = [.. ReadScopes, .. WriteScopes];

    // Added only when the run needs them, so a plain team copy keeps its smaller consent.
    public const string ChannelMessageReadScope = "ChannelMessage.Read.All";   // read a channel's messages and replies
    public const string ChatMemberReadScope     = "ChatMember.Read";           // read who is in a chat
    public const string ChatCreateScope         = "Chat.Create";               // create the copy of a chat
    public const string UserLookupScope         = "User.ReadBasic.All";        // find the mapped person in the target tenant

    // What the SOURCE sign-in needs for a run, and what the TARGET sign-in needs.
    public static string[] SourceScopes(bool messages, bool chats) =>
        [.. ReadScopes, .. (messages ? new[] { ChannelMessageReadScope } : []), .. (chats || messages ? new[] { ChatMemberReadScope } : [])];
    public static string[] TargetScopes(bool crossTenant, bool chats) =>
        [.. CopyScopes, .. (crossTenant ? new[] { UserLookupScope } : []), .. (chats ? new[] { ChatCreateScope } : [])];

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };
    private const string Base = "https://graph.microsoft.com/v1.0";
    private const int MaxAttempts = 6;
    private const int MaxPages = 10_000;   // loop guard against a nextLink that never ends

    private readonly Func<CancellationToken, bool, Task<string>> _token;
    private readonly string? _accessMessage;

    // Microsoft's own wording of a refusal (its error message), appended to ours so the real reason is visible.
    private static string GraphSaid(string body)
    {
        try
        {
            var msg = JsonNode.Parse(body)?["error"]?["message"]?.GetValue<string>();
            return string.IsNullOrWhiteSpace(msg) ? string.Empty : "  Microsoft said: " + (msg.Length > 240 ? msg[..240] : msg);
        }
        catch (Exception) { return string.Empty; }
    }

    // accessMessage replaces the default "which permissions to add" text on a 401/403 (the message import is application-only,
    // so its hint is about Teamwork.Migrate.All and the client secret, not the delegated permissions).
    public TeamsApiClient(Func<CancellationToken, bool, Task<string>> tokenProvider, string? accessMessage = null)
    {
        _token = tokenProvider;
        _accessMessage = accessMessage;
    }

    // The permission is missing or consent was not granted: the message tells the user what to grant.
    public sealed class TeamsAccessException(string message) : Exception(message);

    // Microsoft treats reading other people's chats and channel messages with an application permission as a protected API: besides the
    // permission and admin consent, the tenant must be approved by Microsoft (a request form), which is not something this app can do.
    public const string ProtectedApiMessage =
        "Microsoft blocks reading messages and chats with an application permission until it approves the tenant for its Teams protected APIs " +
        "(admin consent alone is not enough). Request access with the form linked from \"Protected APIs in Microsoft Teams\" on learn.microsoft.com, " +
        "wait for the approval, then try again. Copying a team without its messages, and listing teams, channels, members and files, do not need it.";

    // The text for a 401/403 on the administrator (application) sign-in of the SOURCE tenant.
    public static string AdminAccessMessageFor(string registration) =>
        "Microsoft Teams refused the administrator read. The source app registration " + registration + " needs these Microsoft Graph APPLICATION " +
        "permissions (type Application, admin consent granted): Group.Read.All, Team.ReadBasic.All, Channel.ReadBasic.All, TeamMember.Read.All, " +
        "ChannelMember.Read.All, Files.Read.All and User.Read.All; to copy messages also ChannelMessage.Read.All, Chat.Read.All and " +
        "ChatMember.Read.All (those three also need Microsoft's protected-API approval). The client secret saved in Settings must belong to " +
        "that same registration. After granting consent, wait a minute and try again.";

    public enum ChannelKind { Standard, Private, Shared }

    public sealed record TeamInfo(string Id, string DisplayName, string? Description, string? Visibility);

    public sealed record ChannelInfo(string Id, string DisplayName, string? Description, ChannelKind Kind, DateTimeOffset? Created);

    public sealed record MemberInfo(string UserId, string DisplayName, string? Email, bool IsOwner);

    // The folder that holds a channel's files: the (drive, item) pair CopyJob uses, plus the web URL for logs.
    public sealed record FilesFolderInfo(string DriveId, string ItemId, string? WebUrl);

    // People: who is in the chat (names), filled when an administrator lists someone else's chats, where a chat with no topic would otherwise
    // be shown only as "oneOnOne".
    public sealed record ChatInfo(string Id, string ChatType, string? Topic, DateTimeOffset? Created, string? People = null);

    public sealed record UserInfo(string Id, string DisplayName, string? Email);

    // ── Discovery ────────────────────────────────────────────────────────────

    public async Task<List<TeamInfo>> ListJoinedTeamsAsync(CancellationToken ct = default)
    {
        var rows = await GetAllPagesAsync($"{Base}/me/joinedTeams", ct);
        return rows.Select(t => new TeamInfo(
            Str(t, "id"), Str(t, "displayName"), StrOrNull(t, "description"), StrOrNull(t, "visibility"))).ToList();
    }

    // Administrator view (application sign-in): every team in the tenant, not only the ones the signed-in person joined. A team is a
    // group that was provisioned as one.
    public async Task<List<TeamInfo>> ListAllTeamsAsync(CancellationToken ct = default)
    {
        var rows = await GetAllPagesAsync(
            $"{Base}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description,visibility&$top=100", ct);
        return rows.Select(t => new TeamInfo(
            Str(t, "id"), Str(t, "displayName"), StrOrNull(t, "description"), StrOrNull(t, "visibility"))).ToList();
    }

    // Administrator view: people whose name, email or sign-in name starts with the text (a short list to choose from).
    public async Task<List<UserInfo>> SearchUsersAsync(string text, CancellationToken ct = default)
    {
        var q = text.Trim().Replace("'", "''");
        if (q.Length == 0) return [];
        var filter = Uri.EscapeDataString($"startswith(displayName,'{q}') or startswith(mail,'{q}') or startswith(userPrincipalName,'{q}')");
        var json = await GetJsonAsync($"{Base}/users?$filter={filter}&$select=id,displayName,mail,userPrincipalName&$top=25", ct);
        return (json["value"] as JsonArray ?? []).Select(u => new UserInfo(
            Str(u!, "id"), Str(u!, "displayName"), StrOrNull(u!, "mail") ?? StrOrNull(u!, "userPrincipalName"))).ToList();
    }

    // Administrator view: everyone in the directory (pages of 999), for the "show all people" list.
    public async Task<List<UserInfo>> ListAllUsersAsync(CancellationToken ct = default)
    {
        var rows = await GetAllPagesAsync($"{Base}/users?$select=id,displayName,mail,userPrincipalName&$top=999", ct);
        return rows.Select(u => new UserInfo(
            Str(u, "id"), Str(u, "displayName"), StrOrNull(u, "mail") ?? StrOrNull(u, "userPrincipalName")))
            .OrderBy(u => u.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
    }

    // Administrator view: the chats one person takes part in, with who is in each (one request, members expanded).
    public async Task<List<ChatInfo>> ListUserChatsAsync(string userId, CancellationToken ct = default)
    {
        var rows = await GetAllPagesAsync($"{Base}/users/{userId}/chats?$expand=members&$top=50", ct);
        return rows.Select(c => new ChatInfo(
            Str(c, "id"), Str(c, "chatType"), StrOrNull(c, "topic"),
            DateTimeOffset.TryParse(StrOrNull(c, "createdDateTime"), out var d) ? d : null,
            string.Join(", ", (c["members"] as JsonArray ?? []).Select(m => StrOrNull(m!, "displayName")).Where(n => !string.IsNullOrWhiteSpace(n))))).ToList();
    }

    public async Task<List<ChannelInfo>> ListChannelsAsync(string teamId, CancellationToken ct = default)
    {
        // /channels returns standard channels plus the private and shared channels the signed-in user can see.
        var rows = await GetAllPagesAsync($"{Base}/teams/{teamId}/channels", ct);
        return rows.Select(c => new ChannelInfo(
            Str(c, "id"), Str(c, "displayName"), StrOrNull(c, "description"),
            ParseKind(StrOrNull(c, "membershipType")),
            DateTimeOffset.TryParse(StrOrNull(c, "createdDateTime"), out var d) ? d : null)).ToList();
    }

    // Private and shared channels have their own member list; for a standard channel this is the team's.
    public async Task<List<MemberInfo>> ListChannelMembersAsync(string teamId, string channelId, CancellationToken ct = default)
        => ParseMembers(await GetAllPagesAsync($"{Base}/teams/{teamId}/channels/{channelId}/members", ct));

    public async Task<List<MemberInfo>> ListTeamMembersAsync(string teamId, CancellationToken ct = default)
        => ParseMembers(await GetAllPagesAsync($"{Base}/teams/{teamId}/members", ct));

    public async Task<FilesFolderInfo?> GetFilesFolderAsync(string teamId, string channelId, CancellationToken ct = default)
    {
        try
        {
            var f = await GetJsonAsync($"{Base}/teams/{teamId}/channels/{channelId}/filesFolder", ct);
            var drive = f["parentReference"]?["driveId"]?.GetValue<string>();
            var item  = f["id"]?.GetValue<string>();
            return drive == null || item == null ? null : new FilesFolderInfo(drive, item, StrOrNull(f, "webUrl"));
        }
        catch (TeamsNotFoundException) { return null; }   // folder not provisioned yet, or the user is not a member
    }

    public async Task<List<ChatInfo>> ListChatsAsync(CancellationToken ct = default)
    {
        var rows = await GetAllPagesAsync($"{Base}/me/chats", ct);
        return rows.Select(c => new ChatInfo(
            Str(c, "id"), Str(c, "chatType"), StrOrNull(c, "topic"),
            DateTimeOffset.TryParse(StrOrNull(c, "createdDateTime"), out var d) ? d : null)).ToList();
    }

    // ── Writing (team copy) ──────────────────────────────────────────────────

    // A non-success answer from Graph, with the status so callers can treat "already a member" (409) differently.
    public class TeamsHttpException(int status, string message) : Exception(message)
    {
        public int Status { get; } = status;
    }

    public async Task<string> GetMyUserIdAsync(CancellationToken ct = default)
        => Str(await GetJsonAsync($"{Base}/me?$select=id", ct), "id");

    // The team's General channel. Its name is localized per team, so it is found by being the primary channel, not by name.
    public async Task<string?> GetPrimaryChannelIdAsync(string teamId, CancellationToken ct = default)
    {
        try { return StrOrNull(await GetJsonAsync($"{Base}/teams/{teamId}/primaryChannel?$select=id", ct), "id"); }
        catch (TeamsNotFoundException) { return null; }
    }

    // True when the team still exists (a copy resumed from the id map must not trust an id of a deleted team).
    public async Task<bool> TeamExistsAsync(string teamId, CancellationToken ct = default)
    {
        try { await GetJsonAsync($"{Base}/teams/{teamId}?$select=id", ct); return true; }
        catch (TeamsNotFoundException) { return false; }
    }

    // The address of the SharePoint site a drive belongs to (a private or shared channel's drive is on its own site).
    public async Task<string> GetDriveRootSiteAsync(string driveId, CancellationToken ct = default)
    {
        var root = await GetJsonAsync($"{Base}/drives/{driveId}/root?$select=webUrl,sharepointIds", ct);
        return root["sharepointIds"]?["siteUrl"]?.GetValue<string>()
            ?? throw new Exception("Microsoft Graph did not return the site address for drive " + driveId);
    }

    public async Task<JsonObject> GetTeamAsync(string teamId, CancellationToken ct = default)
        => (await GetJsonAsync($"{Base}/teams/{teamId}", ct)) as JsonObject ?? new JsonObject();

    // Creates a team from the standard template and waits for Graph's asynchronous provisioning to finish.
    public async Task<string> CreateTeamAsync(string displayName, string? description, string? visibility, CancellationToken ct = default)
    {
        var body = new JsonObject
        {
            ["template@odata.bind"] = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
            ["displayName"] = displayName,
        };
        if (!string.IsNullOrWhiteSpace(description)) body["description"] = description;
        if (!string.IsNullOrWhiteSpace(visibility))  body["visibility"]  = visibility;

        var r = await SendRawAsync(HttpMethod.Post, $"{Base}/teams", body, ct);
        var teamId = QuotedId(r.ContentLocation ?? r.Location, "teams")
            ?? throw new Exception("Microsoft Teams did not return the id of the new team.");
        if (r.Location != null) await WaitForOperationAsync(r.Location, ct);
        return teamId;
    }

    // Writes the four settings groups of a team (member, guest, messaging, fun). `settings` is the object read from the source team.
    public async Task UpdateTeamSettingsAsync(string teamId, JsonObject source, CancellationToken ct = default)
    {
        var body = new JsonObject();
        foreach (var key in new[] { "memberSettings", "guestSettings", "messagingSettings", "funSettings" })
            if (source[key] is JsonObject o) body[key] = JsonNode.Parse(o.ToJsonString());
        if (body.Count == 0) return;
        await SendAsync(HttpMethod.Patch, $"{Base}/teams/{teamId}", body, ct);
    }

    public async Task AddTeamMemberAsync(string teamId, string userId, bool owner, CancellationToken ct = default)
        => await SendAsync(HttpMethod.Post, $"{Base}/teams/{teamId}/members", MemberBody(userId, owner), ct);

    public async Task AddChannelMemberAsync(string teamId, string channelId, string userId, bool owner, CancellationToken ct = default)
        => await SendAsync(HttpMethod.Post, $"{Base}/teams/{teamId}/channels/{channelId}/members", MemberBody(userId, owner), ct);

    // Creates a channel and returns its id. A private or shared channel must be created with an owner; `ownerUserId` is
    // that owner (it has to be a member of the team already).
    public async Task<string> CreateChannelAsync(string teamId, string displayName, string? description, ChannelKind kind,
        string? ownerUserId, CancellationToken ct = default)
    {
        var body = new JsonObject
        {
            ["displayName"] = displayName,
            ["membershipType"] = kind switch { ChannelKind.Private => "private", ChannelKind.Shared => "shared", _ => "standard" },
        };
        if (!string.IsNullOrWhiteSpace(description)) body["description"] = description;
        if (kind != ChannelKind.Standard)
        {
            if (ownerUserId == null) throw new InvalidOperationException("A private or shared channel needs an owner.");
            body["members"] = new JsonArray(MemberBody(ownerUserId, owner: true));
        }

        var r = await SendRawAsync(HttpMethod.Post, $"{Base}/teams/{teamId}/channels", body, ct);
        var id = StrOrNull(r.Json, "id") ?? QuotedId(r.ContentLocation ?? r.Location, "channels")
            ?? throw new Exception("Microsoft Teams did not return the id of the new channel.");
        return id;
    }

    private static JsonObject MemberBody(string userId, bool owner, bool fullHistory = false)
    {
        var m = new JsonObject
        {
            ["@odata.type"] = "#microsoft.graph.aadUserConversationMember",
            ["roles"] = owner ? new JsonArray("owner") : new JsonArray(),
            ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{userId}')",
        };
        // Imported messages are older than the chat itself: a member only sees history from this date on, so ask for all of it.
        if (fullHistory) m["visibleHistoryStartDateTime"] = "0001-01-01T00:00:00Z";
        return m;
    }

    // Pulls the id out of a Graph location such as "/teams('abc')/operations('def')" for the segment "teams".
    private static string? QuotedId(string? location, string segment)
    {
        if (location == null) return null;
        var m = System.Text.RegularExpressions.Regex.Match(location, segment + @"\('([^']+)'\)");
        return m.Success ? m.Groups[1].Value : null;
    }

    // Polls a teamsAsyncOperation until it succeeds; throws with Graph's own error when it fails.
    private async Task WaitForOperationAsync(string location, CancellationToken ct)
    {
        var url = location.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? location : Base + location;
        var deadline = DateTime.UtcNow.AddMinutes(10);
        while (DateTime.UtcNow < deadline)
        {
            JsonNode op;
            try { op = await GetJsonAsync(url, ct); }
            catch (TeamsNotFoundException)
            {
                // The operation is not visible for a moment after the request is accepted: not an error, ask again.
                await Task.Delay(TimeSpan.FromSeconds(3), ct);
                continue;
            }
            switch (StrOrNull(op, "status")?.ToLowerInvariant())
            {
                case "succeeded": return;
                case "failed":
                    throw new Exception("Microsoft Teams could not finish creating it: " + (op["error"]?["message"]?.GetValue<string>() ?? "no reason given"));
            }
            await Task.Delay(TimeSpan.FromSeconds(3), ct);
        }
        throw new TimeoutException("Microsoft Teams took more than 10 minutes to finish creating it.");
    }

    // ── Cross-tenant, messages and chats ─────────────────────────────────────

    // The id of the person with this email or sign-in name in THIS tenant, or null when there is none.
    public async Task<string?> FindUserIdAsync(string emailOrUpn, CancellationToken ct = default)
    {
        var e = emailOrUpn.Trim().Replace("'", "''");
        var page = await GetJsonAsync($"{Base}/users?$select=id&$top=2&$filter=" +
            Uri.EscapeDataString($"mail eq '{e}' or userPrincipalName eq '{e}'"), ct);
        return page["value"] is JsonArray { Count: > 0 } arr ? StrOrNull(arr[0]!, "id") : null;
    }

    // A channel's top-level messages, oldest first. Replies are read per message (ListRepliesAsync).
    public async Task<List<JsonNode>> ListChannelMessagesAsync(string teamId, string channelId, CancellationToken ct = default)
        => Chronological(await GetAllPagesAsync($"{Base}/teams/{teamId}/channels/{channelId}/messages?$top=50", ct));

    public async Task<List<JsonNode>> ListRepliesAsync(string teamId, string channelId, string messageId, CancellationToken ct = default)
        => Chronological(await GetAllPagesAsync($"{Base}/teams/{teamId}/channels/{channelId}/messages/{messageId}/replies?$top=50", ct));

    public async Task<List<JsonNode>> ListChatMessagesAsync(string chatId, CancellationToken ct = default)
        => Chronological(await GetAllPagesAsync($"{Base}/chats/{chatId}/messages?$top=50", ct));

    public async Task<List<MemberInfo>> ListChatMembersAsync(string chatId, CancellationToken ct = default)
        => ParseMembers(await GetAllPagesAsync($"{Base}/chats/{chatId}/members", ct));

    // The bytes behind a Graph address (an inline image of a message); null when it cannot be read.
    public async Task<byte[]?> GetBytesAsync(string url, CancellationToken ct = default)
    {
        for (var attempt = 0; attempt <= 3; attempt++)
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await _token(ct, false));
            using var resp = await Http.SendAsync(req, ct);
            if (resp.IsSuccessStatusCode) return await resp.Content.ReadAsByteArrayAsync(ct);
            if ((int)resp.StatusCode is 429 or 502 or 503 or 504) { await Task.Delay(resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(2 * (attempt + 1)), ct); continue; }
            return null;
        }
        return null;
    }

    // Creates the copy of a chat with these members (the signed-in user must be one of them) and returns its id.
    public async Task<string> CreateChatAsync(string chatType, string? topic, IEnumerable<string> memberUserIds, CancellationToken ct = default)
    {
        var members = new JsonArray();
        foreach (var id in memberUserIds.Distinct(StringComparer.OrdinalIgnoreCase)) members.Add(MemberBody(id, owner: true));
        var body = new JsonObject { ["chatType"] = chatType, ["members"] = members };
        if (chatType == "group" && !string.IsNullOrWhiteSpace(topic)) body["topic"] = topic;
        var r = await SendRawAsync(HttpMethod.Post, $"{Base}/chats", body, ct);
        return StrOrNull(r.Json, "id") ?? throw new Exception("Microsoft Teams did not return the id of the new chat.");
    }

    // Imported messages are older than the chat itself, and a member only sees a chat from their own visibleHistoryStartDateTime
    // on. Graph cannot change that for a member who is already in, and refuses it when the chat is created, so each member is
    // removed and added back with all history shared. Needs the application permission ChatMember.ReadWrite.All, so call it on
    // the application-only client. Works for group chats (not one-to-one chats). One member at a time, so the chat is never empty.
    public async Task ShareChatHistoryAsync(string chatId, CancellationToken ct = default)
    {
        var rows = await GetAllPagesAsync($"{Base}/chats/{chatId}/members", ct);
        var people = rows.Select(m => (MembershipId: Str(m, "id"), UserId: Str(m, "userId"),
            Owner: m["roles"] is JsonArray r && r.Any(x => string.Equals(x?.GetValue<string>(), "owner", StringComparison.OrdinalIgnoreCase)))).ToList();
        foreach (var (membershipId, userId, owner) in people)
        {
            if (userId.Length == 0) continue;
            await SendAsync(HttpMethod.Delete, $"{Base}/chats/{chatId}/members/{membershipId}", null, ct);
            await ReAddMemberAsync(chatId, userId, owner, fullHistory: true);
        }
    }

    // Adds a removed member back, and keeps trying (even when the run is cancelled) so a failure never leaves someone out of the
    // chat. If full history cannot be set, the member is at least added back normally.
    private async Task ReAddMemberAsync(string chatId, string userId, bool owner, bool fullHistory)
    {
        Exception? last = null;
        for (var attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                await SendAsync(HttpMethod.Post, $"{Base}/chats/{chatId}/members", MemberBody(userId, owner, fullHistory && attempt < 2), CancellationToken.None);
                return;
            }
            catch (TeamsHttpException ex) when (ex.Status == 409) { return; }   // already back in
            catch (Exception ex) { last = ex; await Task.Delay(TimeSpan.FromSeconds(2), CancellationToken.None); }
        }
        throw new Exception("A chat member was removed and could not be added back (user id " + userId + "): " + last?.Message, last);
    }

    // Migration mode (application-only): start, import messages, complete. startFrom must be older than the thread's creation time.
    public Task StartChannelMigrationAsync(string teamId, string channelId, DateTimeOffset? startFrom, CancellationToken ct = default)
        => SendAsync(HttpMethod.Post, $"{Base}/teams/{teamId}/channels/{channelId}/startMigration", StartBody(startFrom), ct);
    public Task CompleteChannelMigrationAsync(string teamId, string channelId, CancellationToken ct = default)
        => SendAsync(HttpMethod.Post, $"{Base}/teams/{teamId}/channels/{channelId}/completeMigration", null, ct);
    public Task StartChatMigrationAsync(string chatId, DateTimeOffset? startFrom, CancellationToken ct = default)
        => SendAsync(HttpMethod.Post, $"{Base}/chats/{chatId}/startMigration", StartBody(startFrom), ct);
    public Task CompleteChatMigrationAsync(string chatId, CancellationToken ct = default)
        => SendAsync(HttpMethod.Post, $"{Base}/chats/{chatId}/completeMigration", null, ct);

    // Each returns the id of the imported message. A 409 (TeamsHttpException) means a message with that timestamp is already there.
    public async Task<string> ImportChannelMessageAsync(string teamId, string channelId, JsonObject message, CancellationToken ct = default)
        => Str((await SendRawAsync(HttpMethod.Post, $"{Base}/teams/{teamId}/channels/{channelId}/messages", message, ct)).Json, "id");
    public async Task<string> ImportReplyAsync(string teamId, string channelId, string messageId, JsonObject message, CancellationToken ct = default)
        => Str((await SendRawAsync(HttpMethod.Post, $"{Base}/teams/{teamId}/channels/{channelId}/messages/{messageId}/replies", message, ct)).Json, "id");
    public async Task<string> ImportChatMessageAsync(string chatId, JsonObject message, CancellationToken ct = default)
        => Str((await SendRawAsync(HttpMethod.Post, $"{Base}/chats/{chatId}/messages", message, ct)).Json, "id");

    private static JsonObject? StartBody(DateTimeOffset? startFrom)
        => startFrom is { } d ? new JsonObject { ["conversationCreationDateTime"] = d.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'") } : null;

    private static List<JsonNode> Chronological(List<JsonNode> rows)
        => rows.OrderBy(m => DateTimeOffset.TryParse(StrOrNull(m, "createdDateTime"), out var d) ? d : DateTimeOffset.MinValue).ToList();

    // ── Parsing helpers ──────────────────────────────────────────────────────

    private static ChannelKind ParseKind(string? membershipType) => membershipType?.ToLowerInvariant() switch
    {
        "private" => ChannelKind.Private,
        "shared"  => ChannelKind.Shared,
        _         => ChannelKind.Standard,
    };

    private static List<MemberInfo> ParseMembers(List<JsonNode> rows) => rows.Select(m => new MemberInfo(
        Str(m, "userId"), Str(m, "displayName"), StrOrNull(m, "email"),
        m["roles"] is JsonArray roles && roles.Any(r => string.Equals(r?.GetValue<string>(), "owner", StringComparison.OrdinalIgnoreCase)))).ToList();

    private static string Str(JsonNode n, string key) => n[key]?.GetValue<string>() ?? "";
    private static string? StrOrNull(JsonNode n, string key) => n[key]?.GetValue<string>();

    // ── HTTP ─────────────────────────────────────────────────────────────────

    // A 404. It is a TeamsHttpException too, so a write that gets one is handled like any other refused request.
    private sealed class TeamsNotFoundException(string message) : TeamsHttpException(404, message);

    private async Task<JsonNode> GetJsonAsync(string url, CancellationToken ct)
        => await SendAsync(HttpMethod.Get, url, null, ct);

    private async Task<List<JsonNode>> GetAllPagesAsync(string url, CancellationToken ct)
    {
        var all = new List<JsonNode>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? next = url;
        for (var pages = 0; next != null; pages++)
        {
            if (pages >= MaxPages || !seen.Add(next))
                throw new Exception("Microsoft Graph returned a paging loop for " + url);
            var page = await GetJsonAsync(next, ct);
            if (page["value"] is JsonArray arr)
                foreach (var n in arr) if (n != null) all.Add(n);
            next = page["@odata.nextLink"]?.GetValue<string>();
        }
        return all;
    }

    private async Task<JsonNode> SendAsync(HttpMethod method, string url, JsonNode? body, CancellationToken ct)
        => (await SendRawAsync(method, url, body, ct)).Json;

    private sealed record RawResponse(JsonNode Json, string? Location, string? ContentLocation);

    private async Task<RawResponse> SendRawAsync(HttpMethod method, string url, JsonNode? body, CancellationToken ct)
    {
        var refreshed = false;
        for (var attempt = 0; ; attempt++)
        {
            var token = await _token(ct, refreshed && attempt > 0);
            using var req = new HttpRequestMessage(method, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            if (body != null) req.Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json");
            using var resp = await Http.SendAsync(req, ct);
            var text = await resp.Content.ReadAsStringAsync(ct);
            var code = (int)resp.StatusCode;

            // 412 "Version mismatch for thread resource": Teams was still changing the thread (a team or channel just created); nothing was applied, so ask again
            if (code is 429 or 502 or 503 or 504 or 412 && attempt < MaxAttempts)
            {
                var wait = resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(2 * (attempt + 1));
                await Task.Delay(wait, ct);
                continue;
            }
            if (code == 401 && !refreshed)
            {
                refreshed = true;   // one retry with a freshly acquired token
                continue;
            }
            if (code == 404) throw new TeamsNotFoundException($"Teams HTTP 404: not found ({method.Method} {new Uri(url).AbsolutePath})");
            if (code is 401 or 403 && text.Contains("Protected API", StringComparison.OrdinalIgnoreCase))
                throw new TeamsAccessException($"[HTTP {code} {method.Method} {new Uri(url).AbsolutePath}] " + ProtectedApiMessage);
            if (code is 401 or 403)
                throw new TeamsAccessException($"[HTTP {code} {method.Method} {new Uri(url).AbsolutePath}] " + (_accessMessage ??
                    "Microsoft Teams refused the request. The app registration needs the delegated Microsoft Graph permissions " +
                    "Team.ReadBasic.All, Channel.ReadBasic.All, TeamMember.Read.All and Chat.Read, and to copy a team also " +
                    "Team.Create, Channel.Create, ChannelMember.ReadWrite.All, TeamMember.ReadWrite.All and TeamSettings.ReadWrite.All " +
                    "(with admin consent), and the signed-in account must be a member of the team or channel. See the README for details.") + GraphSaid(text));
            if (!resp.IsSuccessStatusCode)
                throw new TeamsHttpException(code, $"Teams HTTP {code}: {(text.Length > 300 ? text[..300] : text)}");
            var json = string.IsNullOrWhiteSpace(text) ? new JsonObject() : JsonNode.Parse(text) ?? new JsonObject();
            return new RawResponse(json, resp.Headers.Location?.ToString(),
                resp.Content.Headers.TryGetValues("Content-Location", out var cl) ? cl.FirstOrDefault()
                    : resp.Headers.TryGetValues("Content-Location", out var cl2) ? cl2.FirstOrDefault() : null);
        }
    }
}
