using System.Collections.ObjectModel;
using System.Net.Http;
using System.Text.Json.Nodes;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Teams;

// Copies a team: the team and its settings, its members, every channel (standard, private and shared), each channel's files
// and, when asked, each channel's messages. Also copies a single chat with its messages.
//
// The source and the target may be two tenants. Everything is READ with the source client and WRITTEN with the target client;
// a person is found in the target by mapping their email through the identity mapping (IdentityMapper), the same one the
// SharePoint copies use. Someone with no match is left out and reported, never guessed.
//
// Messages are imported in Microsoft's migration mode (application-only), see TeamsMessageImporter. Without an import client
// (no client secret saved) messages and chats cannot be copied.
//
// A second run of the same copy resumes: the ids of the target team and channels are remembered (TeamsIdMapStore), a channel
// whose files or messages were already copied is not copied again.
public sealed class TeamsCopyService
{
    // ChannelIds: the source channels to copy (null = every channel). The General channel is always copied: every team has one.
    public sealed record Options(string NewName, bool CopySettings, bool CopyMembers, bool CopyFiles, bool CopyMessages = false, IReadOnlySet<string>? ChannelIds = null);

    // One source member and what the mapping made of them.
    public sealed record MemberPlan(TeamsApiClient.MemberInfo Member, string? MappedEmail, string? TargetUserId)
    {
        public bool Resolved => TargetUserId != null;
    }

    private readonly TeamsApiClient _src;
    private readonly TeamsApiClient _tgt;
    private readonly TeamsApiClient? _import;
    private readonly SharePointService _sp;
    private readonly CopyService _copy;
    private readonly string _targetAccount;
    private readonly int _maxParallel;
    private readonly IdentityMapper _mapper;
    private readonly bool _crossTenant;

    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _emailBySourceId;
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, string?> _targetIdByEmail;
    private string? _targetMe;

    // One tenant: the same client reads and writes. (Kept for callers that copy within a tenant without messages.)
    public TeamsCopyService(TeamsApiClient api, SharePointService sp, CopyService copy, string targetAccount, int maxParallel)
        : this(api, api, null, sp, copy, targetAccount, maxParallel, IdentityMapper.None, false) { }

    public TeamsCopyService(TeamsApiClient source, TeamsApiClient target, TeamsApiClient? import, SharePointService sp, CopyService copy,
        string targetAccount, int maxParallel, IdentityMapper mapper, bool crossTenant, TeamsPeopleCache? people = null)
    {
        people ??= new TeamsPeopleCache();
        _emailBySourceId = people.EmailBySourceId;
        _targetIdByEmail = people.TargetIdByEmail;
        _src = source;
        _tgt = target;
        _import = import;
        _sp = sp;
        _copy = copy;
        _targetAccount = targetAccount;
        _maxParallel = Math.Max(1, maxParallel);
        _mapper = mapper;
        _crossTenant = crossTenant;
    }

    public sealed class Outcome
    {
        public string TargetTeamId { get; set; } = "";
        public int ChannelsCreated { get; set; }
        public int ChannelsReused { get; set; }
        public int ChannelsWithFiles { get; set; }
        public int MembersAdded { get; set; }
        public int FilesCopied { get; set; }
        public int FilesFailed { get; set; }
        public int MessagesImported { get; set; }
        public int MessagesSkipped { get; set; }
        public int MessagesFailed { get; set; }
        public int MessagesUnmapped { get; set; }
        public List<string> UnresolvedPeople { get; } = [];
        public List<string> Warnings { get; } = [];
    }

    // ── Mapping step: who from the source team can be found in the target ──

    // Reads the team's members and maps each to a person in the target. Nothing is written. Within one tenant everyone resolves.
    public async Task<List<MemberPlan>> PlanMembersAsync(TeamsApiClient.TeamInfo source, CancellationToken ct)
    {
        var members = await _src.ListTeamMembersAsync(source.Id, ct);
        // Each person is looked up in the target by email: several at a time (a team of a hundred people was a hundred requests in a row)
        var plan = new MemberPlan[members.Count];
        await Parallel.ForEachAsync(Enumerable.Range(0, members.Count), new ParallelOptions { MaxDegreeOfParallelism = 6, CancellationToken = ct }, async (i, t) =>
        {
            var m = members[i];
            if (!string.IsNullOrWhiteSpace(m.Email)) _emailBySourceId[m.UserId] = m.Email;
            var mapped = _crossTenant && !string.IsNullOrWhiteSpace(m.Email) ? _mapper.MapEmail(m.Email) : m.Email;
            plan[i] = new MemberPlan(m, mapped, await ResolveAsync(m.UserId, m.Email, t));
        });
        return [.. plan];
    }

    // The target person for a source person: the same id within one tenant, otherwise found by mapped email (null = no match).
    private async Task<string?> ResolveAsync(string sourceUserId, string? email, CancellationToken ct)
    {
        if (!_crossTenant) return sourceUserId;
        email ??= _emailBySourceId.GetValueOrDefault(sourceUserId);
        if (string.IsNullOrWhiteSpace(email)) return null;
        var mapped = _mapper.MapEmail(email) ?? email;
        if (_targetIdByEmail.TryGetValue(mapped, out var cached)) return cached;
        string? id = null;
        try { id = await _tgt.FindUserIdAsync(mapped, ct); }
        catch (TeamsApiClient.TeamsHttpException) { /* not found / not readable: unresolved */ }
        _targetIdByEmail[mapped] = id;
        return id;
    }

    private Task<string?> ResolveAuthorAsync(string sourceUserId, CancellationToken ct) => ResolveAsync(sourceUserId, null, ct);

    private async Task<string> TargetMeAsync(CancellationToken ct) => _targetMe ??= await _tgt.GetMyUserIdAsync(ct);

    public async Task<Outcome> CopyAsync(TeamsApiClient.TeamInfo source, Options options, IProgress<string> log, CancellationToken ct)
    {
        var outcome = new Outcome();
        if (options.CopyMessages && _import == null)
            throw new InvalidOperationException(Loc.T("Teams_Msg_NeedSecret"));

        // ── Read everything from the source first, so a read problem stops before anything is created ──
        log.Report(Loc.T("Teams_LoadingTeam", source.DisplayName));
        var sourceTeam = await _src.GetTeamAsync(source.Id, ct);
        var sourceChannels = await _src.ListChannelsAsync(source.Id, ct);
        var plan = options.CopyMembers || options.CopyMessages ? await PlanMembersAsync(source, ct) : [];
        var sourcePrimary = await _src.GetPrimaryChannelIdAsync(source.Id, ct);
        var me = await TargetMeAsync(ct);
        if (options.ChannelIds != null)
            sourceChannels = sourceChannels.Where(c => c.Id == sourcePrimary || options.ChannelIds.Contains(c.Id)).ToList();

        // ── The target team: resume the earlier copy when it is still there ──
        var map = TeamsIdMapStore.Find(source.Id, _targetAccount);
        if (map != null && !await _tgt.TeamExistsAsync(map.TargetTeamId, ct))
        {
            log.Report(Loc.T("Teams_Step_MapStale"));
            map = null;
        }
        // A different name asks for a NEW team. Only the same name continues an earlier copy; without this the name typed in the
        // wizard was ignored and the copy went into the earlier team again.
        if (map != null && !string.Equals(map.TargetName, options.NewName, StringComparison.OrdinalIgnoreCase))
        {
            log.Report(Loc.T("Teams_Step_NewNameNewTeam", options.NewName, map.TargetName));
            map = null;
        }
        if (map == null)
        {
            log.Report(Loc.T("Teams_Step_CreatingTeam", options.NewName));
            var teamId = await _tgt.CreateTeamAsync(options.NewName, source.Description,
                sourceTeam["visibility"]?.GetValue<string>() ?? source.Visibility, ct);
            map = new TeamsIdMapStore.TeamMap { TargetTeamId = teamId, TargetName = options.NewName, CopiedAt = DateTimeOffset.Now };
            TeamsIdMapStore.Save(source.Id, _targetAccount, map);
        }
        else
        {
            log.Report(Loc.T("Teams_Step_Resuming", map.TargetName));
        }
        outcome.TargetTeamId = map.TargetTeamId;
        var targetTeamId = map.TargetTeamId;

        // ── Settings ──
        if (options.CopySettings)
        {
            log.Report(Loc.T("Teams_Step_Settings"));
            await Guarded(outcome, "settings", () => _tgt.UpdateTeamSettingsAsync(targetTeamId, sourceTeam, ct), ct);
        }

        // ── Members (before channels: a private or shared channel's members must already belong to the team) ──
        if (options.CopyMembers)
        {
            log.Report(Loc.T("Teams_Step_Members", plan.Count.ToString()));
            var toAdd = new List<MemberPlan>();
            foreach (var p in plan)
            {
                if (!p.Resolved)
                {
                    outcome.UnresolvedPeople.Add(PersonText(p.Member));
                    continue;
                }
                if (string.Equals(p.TargetUserId, me, StringComparison.OrdinalIgnoreCase)) continue;   // the creator is already an owner
                toAdd.Add(p);
            }
            // several at a time: each add is its own request, and a large team was one after another
            var added = 0;
            await Parallel.ForEachAsync(toAdd, new ParallelOptions { MaxDegreeOfParallelism = 4, CancellationToken = ct }, async (p, t) =>
                await Guarded(outcome, p.Member.DisplayName, async () =>
                {
                    await _tgt.AddTeamMemberAsync(targetTeamId, p.TargetUserId!, p.Member.IsOwner, t);
                    Interlocked.Increment(ref added);
                }, t, ignoreConflict: true));
            outcome.MembersAdded += added;
        }

        // ── Channels ──
        var pairs = new List<(TeamsApiClient.ChannelInfo Source, string TargetId)>();
        var targetPrimary = await _tgt.GetPrimaryChannelIdAsync(targetTeamId, ct);
        foreach (var ch in sourceChannels)
        {
            ct.ThrowIfCancellationRequested();
            if (map.Channels.TryGetValue(ch.Id, out var existing))
            {
                outcome.ChannelsReused++;
                pairs.Add((ch, existing));
                continue;
            }
            // The General channel exists in every new team: pair it instead of creating a second one.
            if (ch.Id == sourcePrimary && targetPrimary != null)
            {
                map.Channels[ch.Id] = targetPrimary;
                TeamsIdMapStore.Save(source.Id, _targetAccount, map);
                outcome.ChannelsReused++;
                pairs.Add((ch, targetPrimary));
                continue;
            }

            log.Report(Loc.T("Teams_Step_Channel", ch.DisplayName, KindText(ch.Kind)));
            string newId;
            try
            {
                newId = await _tgt.CreateChannelAsync(targetTeamId, ch.DisplayName, ch.Description, ch.Kind, me, ct);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                outcome.Warnings.Add(Loc.T("Teams_Warn_Channel", ch.DisplayName, ex.Message));
                continue;
            }
            map.Channels[ch.Id] = newId;
            TeamsIdMapStore.Save(source.Id, _targetAccount, map);
            outcome.ChannelsCreated++;
            pairs.Add((ch, newId));

            // A private or shared channel has its own member list: bring over the source's.
            if (options.CopyMembers && ch.Kind != TeamsApiClient.ChannelKind.Standard)
            {
                List<TeamsApiClient.MemberInfo> channelMembers;
                try { channelMembers = await _src.ListChannelMembersAsync(source.Id, ch.Id, ct); }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    outcome.Warnings.Add(Loc.T("Teams_Warn_ChannelMembers", ch.DisplayName, ex.Message));
                    continue;
                }
                foreach (var m in channelMembers)
                {
                    if (!string.IsNullOrWhiteSpace(m.Email)) _emailBySourceId[m.UserId] = m.Email;
                    var tid = await ResolveAsync(m.UserId, m.Email, ct);
                    if (tid == null) { var t = PersonText(m); if (!outcome.UnresolvedPeople.Contains(t)) outcome.UnresolvedPeople.Add(t); continue; }
                    if (string.Equals(tid, me, StringComparison.OrdinalIgnoreCase)) continue;
                    await Guarded(outcome, $"{ch.DisplayName}: {m.DisplayName}",
                        () => _tgt.AddChannelMemberAsync(targetTeamId, newId, tid, m.IsOwner, ct), ct, ignoreConflict: true);
                }
            }
        }

        // ── Files (before the messages, so a message can link to the copy of a file attached to it) ──
        if (options.CopyFiles)
            await CopyFilesAsync(source, map, pairs, outcome, log, ct);

        // ── Messages ──
        if (options.CopyMessages)
            await CopyMessagesAsync(source, map, pairs, me, outcome, log, ct);

        return outcome;
    }

    private async Task CopyMessagesAsync(TeamsApiClient.TeamInfo source, TeamsIdMapStore.TeamMap map,
        List<(TeamsApiClient.ChannelInfo Source, string TargetId)> pairs, string me, Outcome outcome, IProgress<string> log, CancellationToken ct)
    {
        var importer = new TeamsMessageImporter(_src, _import!, me, ResolveAuthorAsync, log);
        var todo = pairs.Where(p => !map.MessagesDone.GetValueOrDefault(p.Source.Id)).ToList();
        // Microsoft limits the import to 5 requests a second PER THREAD (channel), so several channels are imported at the same time
        await Parallel.ForEachAsync(todo, new ParallelOptions { MaxDegreeOfParallelism = 3, CancellationToken = ct }, async (pair, t) =>
        {
            var (ch, targetChannelId) = pair;
            log.Report(Loc.T("Teams_Step_Messages", ch.DisplayName));
            try
            {
                var r = await importer.ImportChannelAsync(source.Id, ch.Id, map.TargetTeamId, targetChannelId, map,
                    () => TeamsIdMapStore.Save(source.Id, _targetAccount, map), t, await AttachmentMapperAsync(source.Id, ch.Id, map.TargetTeamId, targetChannelId, t));
                lock (outcome) AddMessageResult(outcome, ch.DisplayName, r);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                lock (outcome) outcome.Warnings.Add(Loc.T("Teams_Warn_Messages", ch.DisplayName, ex.Message));
            }
        });
    }

    // Turns the address of a file in the source channel's Files into the address of its copy in the target channel's Files, so a message's
    // attachment can point at the copy. Null when either Files folder cannot be found (then the attachments are only named).
    private async Task<Func<string, string?>?> AttachmentMapperAsync(string sourceTeamId, string sourceChannelId, string targetTeamId, string targetChannelId, CancellationToken ct)
    {
        try
        {
            var from = await _src.GetFilesFolderAsync(sourceTeamId, sourceChannelId, ct);
            var to = await _tgt.GetFilesFolderAsync(targetTeamId, targetChannelId, ct);
            if (from?.WebUrl is not { Length: > 0 } fromUrl || to?.WebUrl is not { Length: > 0 } toUrl) return null;
            static string Plain(string u) => Uri.UnescapeDataString(u).TrimEnd('/');
            var fromPlain = Plain(fromUrl);
            return url =>
            {
                var plain = Plain(url);
                if (!plain.StartsWith(fromPlain + "/", StringComparison.OrdinalIgnoreCase)) return null;
                var rest = plain[fromPlain.Length..].Split('/', StringSplitOptions.RemoveEmptyEntries);
                return toUrl.TrimEnd('/') + string.Concat(rest.Select(seg => "/" + Uri.EscapeDataString(seg)));
            };
        }
        catch (Exception ex) when (ex is not OperationCanceledException) { return null; }
    }

    private static void AddMessageResult(Outcome outcome, string where, TeamsMessageImporter.Result r)
    {
        outcome.MessagesImported += r.Imported;
        outcome.MessagesSkipped += r.Skipped;
        outcome.MessagesFailed += r.Failed;
        outcome.MessagesUnmapped += r.Unmapped;
        foreach (var w in r.Warnings) outcome.Warnings.Add($"{where}: {w}");
        if (r.Failed > 0) outcome.Warnings.Add(Loc.T("Teams_Warn_MessagesFailed", where, r.Failed.ToString()));
    }

    // ── A single chat ──

    public sealed class ChatOutcome
    {
        public string TargetChatId { get; set; } = "";
        public int MembersMapped { get; set; }
        public int MessagesImported { get; set; }
        public int MessagesSkipped { get; set; }
        public int MessagesFailed { get; set; }
        public int MessagesUnmapped { get; set; }
        public List<string> UnresolvedPeople { get; } = [];
        public List<string> Warnings { get; } = [];
    }

    // Copies a chat into a NEW chat in the target (the same people, mapped) and imports its messages. Resumes on a second run.
    public async Task<ChatOutcome> CopyChatAsync(TeamsApiClient.ChatInfo chat, IProgress<string> log, CancellationToken ct)
    {
        if (_import == null) throw new InvalidOperationException(Loc.T("Teams_Msg_NeedSecret"));
        var outcome = new ChatOutcome();
        var me = await TargetMeAsync(ct);
        var label = string.IsNullOrWhiteSpace(chat.Topic) ? chat.ChatType : chat.Topic!;

        var map = TeamsIdMapStore.Find(chat.Id, _targetAccount);
        if (map != null && map.TargetTeamId.Length == 0) map = null;
        if (map == null)
        {
            log.Report(Loc.T("Teams_Chat_Members", label));
            var people = new List<string> { me };
            foreach (var m in await _src.ListChatMembersAsync(chat.Id, ct))
            {
                if (!string.IsNullOrWhiteSpace(m.Email)) _emailBySourceId[m.UserId] = m.Email;
                var tid = await ResolveAsync(m.UserId, m.Email, ct);
                if (tid == null) { outcome.UnresolvedPeople.Add(PersonText(m)); continue; }
                people.Add(tid);
            }
            var distinct = people.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            outcome.MembersMapped = distinct.Count;
            if (distinct.Count < 2)
            {
                outcome.Warnings.Add(Loc.T("Teams_Chat_NoOneElse", label));
                return outcome;
            }
            // Always a group chat, even for a one-to-one source: Graph can only share the imported history with the people of a group chat.
            log.Report(Loc.T("Teams_Chat_Creating", label));
            var id = await _tgt.CreateChatAsync("group", chat.Topic, distinct, ct);
            map = new TeamsIdMapStore.TeamMap { TargetTeamId = id, TargetName = label, CopiedAt = DateTimeOffset.Now };
            TeamsIdMapStore.Save(chat.Id, _targetAccount, map);
        }
        outcome.TargetChatId = map.TargetTeamId;

        if (map.MessagesDone.GetValueOrDefault(chat.Id)) { log.Report(Loc.T("Teams_Chat_AlreadyDone", label)); return outcome; }
        log.Report(Loc.T("Teams_Step_Messages", label));
        var importer = new TeamsMessageImporter(_src, _import, me, ResolveAuthorAsync, log);
        try
        {
            var r = await importer.ImportChatAsync(chat.Id, map.TargetTeamId, map, () => TeamsIdMapStore.Save(chat.Id, _targetAccount, map), ct);
            outcome.MessagesImported = r.Imported; outcome.MessagesSkipped = r.Skipped; outcome.MessagesFailed = r.Failed; outcome.MessagesUnmapped = r.Unmapped;
            outcome.Warnings.AddRange(r.Warnings);

            // The imported messages predate the chat, so the people in it cannot see them yet: share the history with each of them.
            if (r.Imported > 0)
            {
                log.Report(Loc.T("Teams_Chat_SharingHistory", label));
                try { await _import.ShareChatHistoryAsync(map.TargetTeamId, ct); }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    outcome.Warnings.Add(Loc.T("Teams_Chat_ShareFailed", label, ex.Message));
                }
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            outcome.Warnings.Add(Loc.T("Teams_Warn_Messages", label, ex.Message));
        }
        return outcome;
    }

    private static string PersonText(TeamsApiClient.MemberInfo m)
        => string.IsNullOrWhiteSpace(m.Email) ? m.DisplayName : $"{m.DisplayName} <{m.Email}>";

    // One folder job per channel: the contents of the source channel's files folder into the target channel's.
    private async Task CopyFilesAsync(TeamsApiClient.TeamInfo source, TeamsIdMapStore.TeamMap map,
        List<(TeamsApiClient.ChannelInfo Source, string TargetId)> pairs, Outcome outcome, IProgress<string> log, CancellationToken ct)
    {
        foreach (var (ch, targetChannelId) in pairs)
        {
            ct.ThrowIfCancellationRequested();
            if (map.FilesDone.GetValueOrDefault(ch.Id)) { outcome.ChannelsWithFiles++; continue; }

            log.Report(Loc.T("Teams_Step_Files", ch.DisplayName));
            var from = await _src.GetFilesFolderAsync(source.Id, ch.Id, ct);
            if (from == null)
            {
                outcome.Warnings.Add(Loc.T("Teams_Warn_NoSourceFiles", ch.DisplayName));
                continue;
            }
            // A new channel's folder is provisioned asynchronously (a private or shared channel gets a whole new site).
            var to = await WaitForFilesFolderAsync(map.TargetTeamId, targetChannelId, ct);
            if (to == null)
            {
                outcome.Warnings.Add(Loc.T("Teams_Warn_NoTargetFiles", ch.DisplayName));
                continue;
            }

            var job = await BuildFolderJobAsync(ch.DisplayName, from, to);
            var results = new ObservableCollection<CopyResult>();
            var metadataPasses = new MetadataPasses();
            await _copy.ExecuteAsync([job], results, new CopyRunOptions
            {
                OverwriteMode = OverwriteMode.Skip,
                CopyVersions = true,
                MaxParallel = _maxParallel,
                MaxVersions = 0,
                CopyMode = CopyMode.EnhancedRest,
                PreserveMetadata = true,
                ReapplyFolderMetadata = true,
                OnMetadataDone = metadataPasses.Progress,
            }, ct);
            await metadataPasses.WaitAsync(ct);   // the folder dates and authors are written after ExecuteAsync returns

            var failed = results.Count(r => r.Status == CopyStatus.Failed);
            outcome.FilesCopied += results.Count(r => r.Status == CopyStatus.Success);
            outcome.FilesFailed += failed;
            if (failed > 0)
                outcome.Warnings.Add(Loc.T("Teams_Warn_FilesFailed", ch.DisplayName, failed.ToString()));
            else
            {
                map.FilesDone[ch.Id] = true;   // only a clean channel is skipped on a resume; a failed one is retried
                TeamsIdMapStore.Save(source.Id, _targetAccount, map);
            }
            outcome.ChannelsWithFiles++;
        }
    }

    private async Task<TeamsApiClient.FilesFolderInfo?> WaitForFilesFolderAsync(string teamId, string channelId, CancellationToken ct)
    {
        for (var attempt = 0; attempt < 24; attempt++)   // up to about two minutes
        {
            if (await _tgt.GetFilesFolderAsync(teamId, channelId, ct) is { } f) return f;
            await Task.Delay(TimeSpan.FromSeconds(5), ct);
        }
        return null;
    }

    private async Task<CopyJob> BuildFolderJobAsync(string channelName, TeamsApiClient.FilesFolderInfo from, TeamsApiClient.FilesFolderInfo to)
    {
        var sourceSiteUrl = await _src.GetDriveRootSiteAsync(from.DriveId);
        var targetSiteUrl = await _tgt.GetDriveRootSiteAsync(to.DriveId);
        // Between two tenants the copy engine picks each request's sign-in by drive, and these drives came from Teams, not from a
        // site the engine browsed, so tell it which tenant owns each (a no-op in one tenant).
        _sp.Router.RegisterDriveForSite(from.DriveId, sourceSiteUrl);
        _sp.Router.RegisterDriveForSite(to.DriveId, targetSiteUrl);
        var targetSiteId = await _sp.GetSiteIdAsync(targetSiteUrl);
        return new CopyJob
        {
            SourceDriveId = from.DriveId,
            SourceItemId = from.ItemId,
            SourceName = channelName,
            SourceSiteUrl = sourceSiteUrl,
            SourceDisplayPath = from.WebUrl ?? channelName,
            TargetDriveId = to.DriveId,
            TargetParentItemId = to.ItemId,
            TargetSiteId = targetSiteId,
            TargetSiteUrl = targetSiteUrl,
            TargetDisplayPath = to.WebUrl ?? channelName,
            IsFolder = true,
            IsLibrary = true,   // copy the folder's contents into the target folder, not a folder of the same name inside it
        };
    }

    // Runs one step; a failure is recorded as a warning so one bad member or setting does not stop the whole copy.
    private async Task Guarded(Outcome outcome, string what, Func<Task> step, CancellationToken ct, bool ignoreConflict = false)
    {
        try { await step(); }
        catch (OperationCanceledException) { throw; }
        catch (TeamsApiClient.TeamsHttpException ex) when (ignoreConflict && ex.Status == 409) { /* already a member */ }
        catch (Exception ex) { lock (outcome.Warnings) outcome.Warnings.Add($"{what}: {ex.Message}"); }
        ct.ThrowIfCancellationRequested();
    }

    private static string KindText(TeamsApiClient.ChannelKind kind) => kind switch
    {
        TeamsApiClient.ChannelKind.Private => Loc.T("Teams_KindPrivate"),
        TeamsApiClient.ChannelKind.Shared  => Loc.T("Teams_KindShared"),
        _                                  => Loc.T("Teams_KindStandard"),
    };

    // CopyService.ExecuteAsync reports OnMetadataDone exactly once per call, on every exit: true when its (detached) folder
    // metadata pass finished or there was none, false when the call threw. Waiting for it keeps the dialog from reporting
    // "done" while folder dates and authors are still being written in the background.
    private sealed class MetadataPasses
    {
        private readonly TaskCompletionSource _done = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public IProgress<bool> Progress { get; }

        public MetadataPasses() => Progress = new Progress<bool>(_ => _done.TrySetResult());

        public Task WaitAsync(CancellationToken ct) => _done.Task.WaitAsync(ct);
    }
}
