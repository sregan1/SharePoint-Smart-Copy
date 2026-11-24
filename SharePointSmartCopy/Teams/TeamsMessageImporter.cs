using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Teams;

// Copies the messages of one channel or chat into its counterpart on the target, using Microsoft's import (migration mode):
// start migration on the target thread, post each message with its ORIGINAL time and author, complete the migration.
// The import API is application-only, so `import` signs in as the app registration (AppOnlyGraphToken), while the messages
// are READ from the source with the user's own token.
//
// What carries over: text and HTML, subject, importance-free formatting, inline images, replies, the original timestamps and
// the author (mapped to the same person in the target tenant). What does not: file attachments (named in the text instead),
// reactions, edits, deleted messages, system messages (someone joined...), and cards. An @mention is kept as plain bold text,
// because a mention points at a person id of the source tenant.
//
// Safe to run again: a message's timestamp is unique in a thread, so one already imported answers 409 and is skipped, and the
// ids of imported messages are remembered so replies to an earlier message can still be added.
public sealed class TeamsMessageImporter
{
    public sealed class Result
    {
        public int Imported { get; set; }
        public int Skipped { get; set; }
        public int Failed { get; set; }
        public int Unmapped { get; set; }       // messages attributed to the copying user because the author has no match
        public int AttachmentsNamed { get; set; }
        public int AttachmentsLinked { get; set; }   // messages that link to the copy of an attached file
        public List<string> Warnings { get; } = [];
    }

    private static readonly TimeSpan PostGap = TimeSpan.FromMilliseconds(200);   // Microsoft's limit is 5 requests a second per thread
    private const int SaveEvery = 50;
    private const int MaxInlineImageBytes = 4 * 1024 * 1024;

    private readonly TeamsApiClient _source;
    private readonly TeamsApiClient _import;
    private readonly string _fallbackUserId;                                  // the copying user in the target tenant
    private readonly Func<string, CancellationToken, Task<string?>> _resolveAuthor;   // source user id -> target user id, or null
    private readonly IProgress<string> _log;

    public TeamsMessageImporter(TeamsApiClient source, TeamsApiClient import, string fallbackUserId,
        Func<string, CancellationToken, Task<string?>> resolveAuthor, IProgress<string> log)
    {
        _source = source;
        _import = import;
        _fallbackUserId = fallbackUserId;
        _resolveAuthor = resolveAuthor;
        _log = log;
    }

    // One thread's import. `save` persists the map (it is called every few messages and at the end).
    // mapAttachmentUrl turns the address of a file attached to a source message into the address of its copy in the target channel's
    // Files (null when it is not one of the channel's files): the new message then links to the copied file instead of naming it.
    public async Task<Result> ImportChannelAsync(string sourceTeamId, string sourceChannelId, string targetTeamId, string targetChannelId,
        TeamsIdMapStore.TeamMap map, Action save, CancellationToken ct, Func<string, string?>? mapAttachmentUrl = null)
        => await ImportAsync(sourceChannelId, map, save, ct, mapAttachmentUrl,
            list: () => _source.ListChannelMessagesAsync(sourceTeamId, sourceChannelId, ct),
            replies: id => _source.ListRepliesAsync(sourceTeamId, sourceChannelId, id, ct),
            start: from => _import.StartChannelMigrationAsync(targetTeamId, targetChannelId, from, ct),
            post: (m, ct2) => _import.ImportChannelMessageAsync(targetTeamId, targetChannelId, m, ct2),
            postReply: (parent, m, ct2) => _import.ImportReplyAsync(targetTeamId, targetChannelId, parent, m, ct2),
            complete: () => _import.CompleteChannelMigrationAsync(targetTeamId, targetChannelId, ct),
            listTarget: () => _import.ListChannelMessagesAsync(targetTeamId, targetChannelId, ct));

    public async Task<Result> ImportChatAsync(string sourceChatId, string targetChatId, TeamsIdMapStore.TeamMap map, Action save, CancellationToken ct)
        => await ImportAsync(sourceChatId, map, save, ct, null,
            list: () => _source.ListChatMessagesAsync(sourceChatId, ct),
            replies: null,
            start: from => _import.StartChatMigrationAsync(targetChatId, from, ct),
            post: (m, ct2) => _import.ImportChatMessageAsync(targetChatId, m, ct2),
            postReply: null,
            complete: () => _import.CompleteChatMigrationAsync(targetChatId, ct),
            listTarget: null);

    private async Task<Result> ImportAsync(string sourceThreadId, TeamsIdMapStore.TeamMap map, Action save, CancellationToken ct,
        Func<string, string?>? mapAttachmentUrl,
        Func<Task<List<JsonNode>>> list, Func<string, Task<List<JsonNode>>>? replies, Func<DateTimeOffset?, Task> start,
        Func<JsonObject, CancellationToken, Task<string>> post, Func<string, JsonObject, CancellationToken, Task<string>>? postReply,
        Func<Task> complete, Func<Task<List<JsonNode>>>? listTarget)
    {
        var result = new Result();
        var messages = (await list()).Where(IsImportable).ToList();
        if (messages.Count == 0) { map.MessagesDone[sourceThreadId] = true; save(); return result; }

        // The thread's start must be older than its first message (a message cannot predate its channel).
        var earliest = messages.Select(x => Created(x)).Min();
        var alreadyStarted = map.MessagesStarted.GetValueOrDefault(sourceThreadId);
        try
        {
            await start(alreadyStarted ? null : earliest.AddSeconds(-1));
            map.MessagesStarted[sourceThreadId] = true;
            save();
        }
        catch (TeamsApiClient.TeamsHttpException ex) when (ex.Status == 400 && alreadyStarted)
        {
            // an earlier run was interrupted before it completed: the thread is still in migration mode, carry on
        }
        catch (TeamsApiClient.TeamsHttpException ex) when (ex.Status == 400)
        {
            _log.Report(Loc.T("Teams_Msg_StartNote", ex.Message));   // possibly already in migration mode: the posts below will say
            map.MessagesStarted[sourceThreadId] = true;
        }

        // The replies of every message are read ahead, a few at a time, while the messages are being posted one after another
        // (one request per message made a long thread wait on each read in turn).
        using var prefetchCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        Dictionary<string, Task<List<JsonNode>>>? prefetched = null;
        if (replies != null)
        {
            var gate = new SemaphoreSlim(4);
            prefetched = [];
            foreach (var pm in messages)
            {
                var pid = Str(pm, "id");
                if (prefetched.ContainsKey(pid)) continue;
                prefetched[pid] = Task.Run(async () =>
                {
                    await gate.WaitAsync(prefetchCts.Token);
                    try { return await replies(pid); }
                    finally { gate.Release(); }
                }, prefetchCts.Token);
            }
        }
        var usedMs = new HashSet<long>();
        var sinceSave = 0;
        Dictionary<long, string>? targetByMs = null;   // loaded only when a parent is already there without a remembered id
        try
        {
        foreach (var m in messages)
        {
            ct.ThrowIfCancellationRequested();
            var id = Str(m, "id");
            string? targetId = map.Messages.GetValueOrDefault(Key(sourceThreadId, id));
            if (targetId == null)
            {
                var ok = await TryPostAsync(m, usedMs, result, ct, b => post(b, ct), thread: sourceThreadId, mapAttachmentUrl);
                targetId = ok.Id;
                if (targetId == null && ok.Conflict && replies != null && listTarget != null)
                {
                    // Already imported by an interrupted run whose ids were not saved: find it by its time, so its replies still go in.
                    try
                    {
                        targetByMs ??= (await listTarget()).Where(t => t["replyToId"] is null)
                            .GroupBy(t => Created(t).ToUnixTimeMilliseconds()).ToDictionary(g => g.Key, g => Str(g.First(), "id"));
                        if (targetByMs.TryGetValue(ok.Created.ToUnixTimeMilliseconds(), out var found)) targetId = found;
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException) { targetByMs = []; }
                    if (targetId == null && (prefetched != null && prefetched.TryGetValue(id, out var ahead2) ? await ahead2 : await replies(id)).Any(IsImportable))
                    {
                        result.Failed++;   // not marked done, so the replies are not silently lost
                        if (result.Warnings.Count < 20) result.Warnings.Add(Loc.T("Teams_Msg_RepliesSkipped", id));
                    }
                }
                if (targetId != null) map.Messages[Key(sourceThreadId, id)] = targetId;
            }
            else
            {
                Created(m, usedMs);   // keep the timestamps of the earlier run reserved so later ones stay unique and identical
                result.Skipped++;
            }

            if (replies != null && postReply != null && targetId != null)
            {
                List<JsonNode> replyList;
                try { replyList = prefetched != null && prefetched.TryGetValue(id, out var ahead) ? await ahead : await replies(id); }
                catch (TeamsApiClient.TeamsHttpException ex) when (ex.Status == 404) { replyList = []; }   // the message is gone
                foreach (var r in replyList.Where(IsImportable))
                {
                    ct.ThrowIfCancellationRequested();
                    var rid = Str(r, "id");
                    if (map.Messages.ContainsKey(Key(sourceThreadId, rid))) { Created(r, usedMs); result.Skipped++; continue; }
                    var rr = await TryPostAsync(r, usedMs, result, ct, b => postReply(targetId, b, ct), thread: sourceThreadId, mapAttachmentUrl);
                    if (rr.Id != null) map.Messages[Key(sourceThreadId, rid)] = rr.Id;
                    if (++sinceSave >= SaveEvery) { save(); sinceSave = 0; }
                }
            }
            if (++sinceSave >= SaveEvery) { save(); sinceSave = 0; }
        }
        }
        finally { prefetchCts.Cancel(); save(); }   // also on cancel or error: the ids of what was imported must survive, or the next run cannot place replies

        // Complete even when some messages failed, so the thread is usable (no import banner); a failed one is retried by the next run.
        try { await complete(); }
        catch (Exception ex) when (ex is not OperationCanceledException) { result.Warnings.Add(Loc.T("Teams_Msg_CompleteFailed", ex.Message)); }
        if (result.Failed == 0) map.MessagesDone[sourceThreadId] = true;
        save();
        return result;
    }

    // Builds and posts one message. A 409 means it is already there. When Teams refuses it (400/404) it is tried again without the file
    // links, then as the copying user (most often the author is not a valid person here).
    private async Task<(string? Id, bool Conflict, DateTimeOffset Created)> TryPostAsync(JsonNode m, HashSet<long> usedMs, Result result, CancellationToken ct,
        Func<JsonObject, Task<string>> send, string thread, Func<string, string?>? mapAttachmentUrl)
    {
        var created = Created(m, usedMs);
        var useFallback = false;
        var withLinks = true;
        for (var attempt = 0; attempt < 3; attempt++)
        {
            JsonObject body; bool unmapped, hadLinks;
            try { (body, unmapped, hadLinks) = await BuildAsync(m, created, useFallback, withLinks ? mapAttachmentUrl : null, result, ct); }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                result.Failed++;
                result.Warnings.Add(Loc.T("Teams_Msg_BuildFailed", Str(m, "id"), ex.Message));
                return (null, false, created);
            }
            try
            {
                var id = await send(body);
                result.Imported++;
                if (unmapped) result.Unmapped++;
                if (hadLinks) result.AttachmentsLinked++;
                await Task.Delay(PostGap, ct);
                return (id, false, created);
            }
            catch (TeamsApiClient.TeamsHttpException ex) when (ex.Status == 409)
            {
                result.Skipped++;
                return (null, true, created);
            }
            catch (TeamsApiClient.TeamsHttpException ex) when (ex.Status is 400 or 404 && attempt < 2 && (hadLinks || !useFallback))
            {
                // first a link the target does not accept (name the file instead), then the author (most often not a valid person here:
                // post it as the copying user instead)
                if (hadLinks) withLinks = false;
                else useFallback = true;
                continue;
            }
            catch (TeamsApiClient.TeamsHttpException ex)
            {
                result.Failed++;
                if (result.Warnings.Count < 20) result.Warnings.Add(Loc.T("Teams_Msg_PostFailed", Str(m, "id"), ex.Message));
                return (null, false, created);
            }
        }
        result.Failed++;
        return (null, false, created);
    }

    private async Task<(JsonObject Body, bool Unmapped, bool HadLinks)> BuildAsync(JsonNode m, DateTimeOffset created, bool forceFallback,
        Func<string, string?>? mapAttachmentUrl, Result result, CancellationToken ct)
    {
        var contentType = Str(m["body"] ?? new JsonObject(), "contentType");
        var html = !string.Equals(contentType, "text", StringComparison.OrdinalIgnoreCase);
        var content = Str(m["body"] ?? new JsonObject(), "content");
        var hosted = new JsonArray();

        if (html)
        {
            content = CleanHtml(content);
            content = await EmbedImagesAsync(content, hosted, ct);
        }

        // The original author, mapped to the same person here; otherwise the copying user, with the real name in the text.
        string authorId = _fallbackUserId, authorName = ""; var unmapped = false;
        var from = m["from"]?["user"];
        var sourceUserId = from?["id"]?.GetValue<string>();
        var displayName = from?["displayName"]?.GetValue<string>() ?? m["from"]?["application"]?["displayName"]?.GetValue<string>() ?? "";
        var mapped = !forceFallback && sourceUserId != null ? await _resolveAuthor(sourceUserId, ct) : null;
        if (mapped != null) { authorId = mapped; authorName = displayName; }
        else
        {
            unmapped = true;
            var note = string.IsNullOrWhiteSpace(displayName) ? Loc.T("Teams_Msg_UnknownAuthor") : Loc.T("Teams_Msg_PostedBy", displayName);
            content = html ? $"<p><i>{System.Net.WebUtility.HtmlEncode(note)}</i></p>{content}" : $"[{note}]\n{content}";
        }

        // A file attached to the message: when it is one of the channel's files (copied with them), the new message links to the copy;
        // otherwise (a chat, a file elsewhere, or a link Teams refused) it is only named in the text.
        var linked = new JsonArray();
        if (m["attachments"] is JsonArray atts)
            foreach (var a in atts)
            {
                var name = a?["name"]?.GetValue<string>();
                if (string.IsNullOrWhiteSpace(name)) continue;
                var type = a?["contentType"]?.GetValue<string>();
                var url = a?["contentUrl"]?.GetValue<string>();
                var targetUrl = html && mapAttachmentUrl != null && string.Equals(type, "reference", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(url)
                    ? mapAttachmentUrl(url) : null;
                if (targetUrl != null)
                {
                    var attachmentId = Guid.NewGuid().ToString();
                    linked.Add(new JsonObject { ["id"] = attachmentId, ["contentType"] = "reference", ["contentUrl"] = targetUrl, ["name"] = name });
                    content += $"<attachment id=\"{attachmentId}\"></attachment>";
                    continue;
                }
                var line = Loc.T("Teams_Msg_AttachmentNote", name);
                content += html ? $"<p><i>{System.Net.WebUtility.HtmlEncode(line)}</i></p>" : $"\n[{line}]";
                result.AttachmentsNamed++;
            }

        var body = new JsonObject
        {
            ["createdDateTime"] = created.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"),
            ["from"] = new JsonObject { ["user"] = new JsonObject { ["id"] = authorId, ["displayName"] = authorName.Length > 0 ? authorName : "Imported", ["userIdentityType"] = "aadUser" } },
            ["body"] = new JsonObject { ["contentType"] = html ? "html" : "text", ["content"] = content },
        };
        if (!string.IsNullOrWhiteSpace(Str(m, "subject"))) body["subject"] = Str(m, "subject");
        if (hosted.Count > 0) body["hostedContents"] = hosted;
        if (linked.Count > 0) body["attachments"] = linked;
        return (body, unmapped, linked.Count > 0);
    }

    // An @mention points at a person id of the source tenant, so it becomes plain bold text; an attachment placeholder (the
    // file itself is not copied) is dropped.
    public static string CleanHtml(string content)
    {
        content = Regex.Replace(content, @"<at id=""\d+"">(.*?)</at>", "<b>@$1</b>", RegexOptions.Singleline);
        return Regex.Replace(content, @"<attachment id=""[^""]*""></attachment>", string.Empty);
    }

    // Inline images live behind a source address that the target cannot read: fetch each and send it with the message.
    private async Task<string> EmbedImagesAsync(string html, JsonArray hosted, CancellationToken ct)
    {
        var sb = new StringBuilder();
        var last = 0;
        foreach (Match img in Regex.Matches(html, @"<img\b[^>]*?\bsrc=""([^""]*/hostedContents/[^/""]+/\$value)""[^>]*>", RegexOptions.IgnoreCase))
        {
            sb.Append(html, last, img.Index - last);
            last = img.Index + img.Length;
            var bytes = await _source.GetBytesAsync(img.Groups[1].Value, ct);
            if (bytes == null || bytes.Length > MaxInlineImageBytes) { sb.Append("<i>[image not copied]</i>"); continue; }
            var tempId = (hosted.Count + 1).ToString();
            hosted.Add(new JsonObject
            {
                ["@microsoft.graph.temporaryId"] = tempId,
                ["contentBytes"] = Convert.ToBase64String(bytes),
                ["contentType"] = ImageType(bytes),
            });
            sb.Append(img.Value.Replace(img.Groups[1].Value, $"../hostedContents/{tempId}/$value"));
        }
        sb.Append(html, last, html.Length - last);
        return sb.ToString();
    }

    private static string ImageType(byte[] b) =>
        b.Length > 3 && b[0] == 0xFF && b[1] == 0xD8 ? "image/jpeg" :
        b.Length > 3 && b[0] == 0x47 && b[1] == 0x49 && b[2] == 0x46 ? "image/gif" : "image/png";

    // Only real messages: not "someone was added" events, and not deleted ones.
    private static bool IsImportable(JsonNode m)
        => string.Equals(m["messageType"]?.GetValue<string>() ?? "message", "message", StringComparison.OrdinalIgnoreCase)
           && m["deletedDateTime"] is null;

    private static DateTimeOffset Created(JsonNode m)
        => DateTimeOffset.TryParse(Str(m, "createdDateTime"), System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.AssumeUniversal, out var d) ? d : DateTimeOffset.UtcNow;

    // The message's time, moved forward a millisecond at a time until it is unique in its thread (Microsoft rejects duplicates).
    // Deterministic for the same thread, so a re-run produces the same times and its already imported messages answer 409.
    private static DateTimeOffset Created(JsonNode m, HashSet<long> used)
    {
        var ms = Created(m).ToUnixTimeMilliseconds();
        while (!used.Add(ms)) ms++;
        return DateTimeOffset.FromUnixTimeMilliseconds(ms);
    }

    private static string Str(JsonNode n, string key) => n[key]?.GetValue<string>() ?? "";

    // A message id is derived from its timestamp, so two channels can hold the same id: remember it per thread.
    private static string Key(string threadId, string messageId) => threadId + "|" + messageId;
}
