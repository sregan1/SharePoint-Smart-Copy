using System.Text.RegularExpressions;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Workflows;

// A column of a list: its internal name (what flows store), display name, and SharePoint type (for display only).
// Required = the list marks the column as required (a "Create item" step that does not set it fails on import).
public sealed record FlowColumnInfo(string InternalName, string DisplayName, string TypeName = "", bool Required = false);

// A list or library on a site.
public sealed record FlowListInfo(string Id, string Title);

// What the resolver needs from SharePoint. The app implements this with SharePointService (and the
// multi-tenant sign-in routing); tests use a fake.
public interface IFlowSiteLookup
{
    // Title of a list in the SOURCE site, by GUID (null = not found).
    Task<string?> GetSourceListTitleAsync(string siteUrl, string listId, CancellationToken ct);
    // GUID of a list in the TARGET site, by title (null = not found).
    Task<string?> GetTargetListIdAsync(string siteUrl, string title, CancellationToken ct);
    // All lists and libraries on a TARGET site, so the user can pick one when the name differs.
    Task<IReadOnlyList<FlowListInfo>> GetTargetListsAsync(string siteUrl, CancellationToken ct);
    // The columns of a list (editable ones, built-in such as Title included), by GUID.
    Task<IReadOnlyList<FlowColumnInfo>> GetSourceColumnsAsync(string siteUrl, string listId, CancellationToken ct);
    Task<IReadOnlyList<FlowColumnInfo>> GetTargetColumnsAsync(string siteUrl, string listId, CancellationToken ct);
    // Signs in to (and registers) both sites before lookups run in parallel, so a first-time sign-in happens once,
    // up front, instead of from several workers at once.
    Task PrepareAsync(string sourceSite, string targetSite, CancellationToken ct);
}

// Manual = target chosen by the user. AlreadyOnTarget = the flow's list ID is not on the source site but IS a list on the
// target site, which is what a package this tool already wrote looks like: nothing left to map.
public enum ListMatchStatus { Matched, SiteNotMapped, SourceListNotFound, TargetListNotFound, Manual, AlreadyOnTarget }

// Same = identical internal name on the target; Renamed = same display name, different internal name;
// Similar = names match only after ignoring case, spaces, underscores and "_x0020_" (worth a second look);
// NotFoundInTarget = nothing suitable, the user must choose or leave it.
public enum ColumnMatchStatus { Same, Renamed, Similar, NotFoundInTarget }

public sealed record ListResolution(string SourceSite, string SourceListId, string? Title,
                                    string? TargetSite, string? TargetListId, ListMatchStatus Status);

public sealed record ColumnResolution(string SourceSite, string SourceListId, string? ListTitle, string SourceColumn,
                                      string? TargetColumn, ColumnMatchStatus Status);

// The user's (or the automapper's) final decision for one column on one list. Target null/empty = leave unchanged.
public sealed record ColumnChoice(string SourceListId, string SourceColumn, string? TargetColumn);

public sealed class FlowResolveResult
{
    public List<ListResolution> Lists { get; } = [];
    public List<ColumnResolution> Columns { get; } = [];
    public List<string> Notes { get; } = [];

    // Keyed by SOURCE list GUID: the target list's columns (the choices for each row), every column of the source list,
    // and the names of the source list's columns that the flows actually use.
    public Dictionary<string, IReadOnlyList<FlowColumnInfo>> TargetColumnsBySourceList { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, IReadOnlyList<FlowColumnInfo>> SourceColumnsBySourceList { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, HashSet<string>> UsedColumnsBySourceList { get; } = new(StringComparer.OrdinalIgnoreCase);

    // Every list and library on each target site (keyed by the normalized target site URL), for the target list dropdown.
    public Dictionary<string, IReadOnlyList<FlowListInfo>> TargetListsBySite { get; } = new(StringComparer.OrdinalIgnoreCase);

    // The automapper's choices: Renamed and Similar matches. Same / NotFound rows make no change.
    public IEnumerable<ColumnChoice> AutoChoices() => Columns
        .Where(c => c.Status is ColumnMatchStatus.Renamed or ColumnMatchStatus.Similar)
        .Select(c => new ColumnChoice(c.SourceListId, c.SourceColumn, c.TargetColumn));

    // Copies everything that resolved into the mapping (sites, lists, column renames). Existing non-empty targets in
    // the mapping are left alone so a user's manual edits win.
    public void ApplyTo(FlowMapping map, IReadOnlyDictionary<string, string> siteMap)
    {
        foreach (var (src, tgt) in siteMap)
            if (!string.IsNullOrWhiteSpace(tgt) && string.IsNullOrWhiteSpace(map.Sites.GetValueOrDefault(src)))
                map.Sites[src] = tgt;

        foreach (var l in Lists.Where(l => l.Status is ListMatchStatus.Matched or ListMatchStatus.Manual))
            if (string.IsNullOrWhiteSpace(map.Lists.GetValueOrDefault(l.SourceListId)))
                map.Lists[l.SourceListId] = l.TargetListId!;

        map.ApplyColumnChoices(AutoChoices());
    }
}

// Turns "the flow points at list {guid} on site X" into "that list is called 'Y'; the target site has a list
// called 'Y' with this GUID", and does the same for columns (internal name, then display name, then a normalized
// comparison).
public static partial class FlowMappingResolver
{
    private const int MaxParallelLookups = 4;

    private sealed record PairOutcome(ListResolution List, List<ColumnResolution> Columns, IReadOnlyList<FlowColumnInfo>? TargetColumns,
                                      IReadOnlyList<FlowColumnInfo>? SourceColumns, HashSet<string>? Used);

    public static string NormalizeSite(string url) => url.Trim().TrimEnd('/');

    // "Tax_x0020_ID", "Tax ID", "tax id", "TaxID" all normalize to "taxid".
    public static string NormalizeColumnName(string name) =>
        NonAlnum().Replace(name.Replace("_x0020_", "", StringComparison.OrdinalIgnoreCase), "").ToLowerInvariant();

    // The target list that goes with a source list title: same title, ignoring case, then the same once spaces,
    // punctuation and case are ignored, and only when exactly one list qualifies.
    public static FlowListInfo? MatchListByTitle(string title, IReadOnlyList<FlowListInfo> targetLists)
    {
        var exact = targetLists.FirstOrDefault(l => string.Equals(l.Title, title, StringComparison.OrdinalIgnoreCase));
        if (exact is not null) return exact;
        var wanted = NormalizeColumnName(title);
        if (wanted.Length == 0) return null;
        var similar = targetLists.Where(l => NormalizeColumnName(l.Title) == wanted).ToList();
        return similar.Count == 1 ? similar[0] : null;
    }

    // Finds the target column for one source column: same internal name, then same display name, then a similar name
    // (ignoring case, spaces, underscores and "_x0020_"; only a UNIQUE candidate is proposed). Used for the first
    // suggestion and again whenever the user picks a different target list.
    public static (string? Target, ColumnMatchStatus Status) MatchColumn(string name, FlowColumnInfo? source, IReadOnlyList<FlowColumnInfo> targetCols)
    {
        if (targetCols.Any(c => c.InternalName == name)) return (name, ColumnMatchStatus.Same);

        var byDisplay = source is null ? null
            : targetCols.FirstOrDefault(c => string.Equals(c.DisplayName, source.DisplayName, StringComparison.OrdinalIgnoreCase));
        if (byDisplay is not null) return (byDisplay.InternalName, ColumnMatchStatus.Renamed);

        var wanted = new[] { NormalizeColumnName(name), source is null ? "" : NormalizeColumnName(source.DisplayName) }
            .Where(w => w.Length > 0).ToHashSet();
        var similar = targetCols
            .Where(c => wanted.Contains(NormalizeColumnName(c.InternalName)) || wanted.Contains(NormalizeColumnName(c.DisplayName)))
            .ToList();
        return similar.Count == 1 ? (similar[0].InternalName, ColumnMatchStatus.Similar) : (null, ColumnMatchStatus.NotFoundInTarget);
    }

    public static async Task<FlowResolveResult> ResolveAsync(
        IEnumerable<FlowScan> scans,
        IReadOnlyDictionary<string, string> siteMap,
        IFlowSiteLookup lookup,
        CancellationToken ct = default,
        IReadOnlyDictionary<string, string>? listOverrides = null,   // source list GUID -> target list GUID chosen by the user
        Action<string>? progress = null,                             // plain-language "what is happening now" lines for the UI
        IReadOnlyDictionary<string, string>? sourceReadSites = null)  // flow site -> the site to READ its lists from on the source (default: the same)
    {
        var result = new FlowResolveResult();
        var sites = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in siteMap) sites[NormalizeSite(k)] = NormalizeSite(v);

        // The flows name a site; the tool can be told to read that site's lists from a different address on the source (the
        // flow's own address may be old or renamed). The flow's address stays the key for what gets replaced.
        var readSites = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (sourceReadSites is not null)
            foreach (var (flowSite, readFrom) in sourceReadSites)
                if (!string.IsNullOrWhiteSpace(readFrom)) readSites[NormalizeSite(flowSite)] = NormalizeSite(readFrom);
        string ReadSite(string flowSite) => readSites.GetValueOrDefault(flowSite) ?? flowSite;

        var scanList = scans.ToList();
        var findings = scanList.SelectMany(s => s.Findings).ToList();

        // A flow that uses exactly one SharePoint list (and says which one literally) reads that list's columns in its
        // expressions too: triggerOutputs()?['body/Status'] is the Status column of that list. Those names count as used
        // by the list, so they get a row in the column mapping. With two or more lists a name can't be tied to one of
        // them, so those flows contribute nothing here.
        var expressionColumns = new Dictionary<(string Site, string List), HashSet<string>>();
        foreach (var scan in scanList)
        {
            var listFindings = scan.Findings.Where(f => f.Kind == FindingKind.ListId).ToList();
            if (listFindings.Select(f => f.Value).Distinct().Count() != 1) continue;
            if (listFindings[0] is not { Site: not null, ListId: not null } only) continue;
            var key = (Site: NormalizeSite(only.Site), List: only.ListId);
            if (!expressionColumns.TryGetValue(key, out var set)) expressionColumns[key] = set = new HashSet<string>(StringComparer.Ordinal);
            foreach (var f in scan.Findings.Where(f => f.Kind == FindingKind.ColumnExpression)) set.Add(f.Value);
        }

        // Distinct (site, list) pairs the flows actually use with literal values.
        var pairs = findings.Where(f => f is { Kind: FindingKind.ListId, Site: not null, ListId: not null })
            .Select(f => (Site: NormalizeSite(f.Site!), List: f.ListId!))
            .Distinct().ToList();

        // Sign in once per site pair, in order, before the lookups fan out; and read each target site's lists once, so
        // the default match and the dropdown both come from that single listing.
        foreach (var (site, targetSite) in pairs.Select(p => p.Site).Distinct(StringComparer.OrdinalIgnoreCase)
                     .Where(s => sites.TryGetValue(s, out var t) && !string.IsNullOrWhiteSpace(t))
                     .Select(s => (s, sites[s])))
        {
            ct.ThrowIfCancellationRequested();
            progress?.Invoke(Loc.T("Flows_Prog_Connecting", ReadSite(site), targetSite));
            await lookup.PrepareAsync(ReadSite(site), targetSite, ct);
            progress?.Invoke(Loc.T("Flows_Prog_Connected", ReadSite(site), targetSite));

            if (!result.TargetListsBySite.ContainsKey(targetSite))
            {
                progress?.Invoke(Loc.T("Flows_Prog_ReadingTargetLists", targetSite));
                var targetLists = await lookup.GetTargetListsAsync(targetSite, ct);
                result.TargetListsBySite[targetSite] = targetLists;
                progress?.Invoke(Loc.T("Flows_Prog_FoundTargetLists", targetLists.Count, targetSite));
            }
        }

        // Each (site, list) pair is independent, so a handful are looked up at once (a lookup is several SharePoint
        // round trips). Outcomes are merged in the original order so the result is the same as a sequential run.
        using var gate = new SemaphoreSlim(MaxParallelLookups);
        var outcomes = await Task.WhenAll(pairs.Select(async p =>
        {
            await gate.WaitAsync(ct);
            try { return await ResolvePairAsync(p.Site, p.List); }
            finally { gate.Release(); }
        }));
        foreach (var o in outcomes)
        {
            result.Lists.Add(o.List);
            result.Columns.AddRange(o.Columns);
            if (o.TargetColumns is not null) result.TargetColumnsBySourceList[o.List.SourceListId] = o.TargetColumns;
            if (o.SourceColumns is not null) result.SourceColumnsBySourceList[o.List.SourceListId] = o.SourceColumns;
            if (o.Used is not null) result.UsedColumnsBySourceList[o.List.SourceListId] = o.Used;
        }

        async Task<PairOutcome> ResolvePairAsync(string site, string listId)
        {
            ct.ThrowIfCancellationRequested();
            if (!sites.TryGetValue(site, out var targetSite) || string.IsNullOrWhiteSpace(targetSite))
                return new(new(site, listId, null, null, null, ListMatchStatus.SiteNotMapped), [], null, null, UsedNames(site, listId, [], sourceKnown: false));

            var readSite = ReadSite(site);
            progress?.Invoke(Loc.T("Flows_Prog_ReadingSource", listId, readSite));
            var title = await lookup.GetSourceListTitleAsync(readSite, listId, ct);
            var manual = listOverrides is not null && listOverrides.TryGetValue(listId, out var chosen) && !string.IsNullOrWhiteSpace(chosen) ? chosen : null;
            if (title is null && manual is null)
            {
                // The ID may belong to the TARGET site: that is what a package this tool already wrote looks like (its list IDs
                // were replaced with the target's). Say so instead of reporting a plain failure.
                if (result.TargetListsBySite.TryGetValue(targetSite, out var onTarget)
                    && onTarget.FirstOrDefault(l => string.Equals(l.Id, listId, StringComparison.OrdinalIgnoreCase)) is { } here)
                {
                    progress?.Invoke(Loc.T("Flows_Prog_AlreadyOnTarget", listId, here.Title));
                    var hereCols = await lookup.GetTargetColumnsAsync(targetSite, here.Id, ct);
                    return new(new(site, listId, here.Title, targetSite, here.Id, ListMatchStatus.AlreadyOnTarget), [], hereCols, null,
                               UsedNames(site, listId, [], sourceKnown: false));
                }
                return new(new(site, listId, null, targetSite, null, ListMatchStatus.SourceListNotFound), [], null, null,
                           UsedNames(site, listId, [], sourceKnown: false));
            }

            // A target the user chose wins over name matching (it is how lists with different names are mapped).
            string? targetId = manual;
            if (targetId is null)
            {
                progress?.Invoke(Loc.T("Flows_Prog_LookingTarget", title, targetSite));
                if (result.TargetListsBySite.TryGetValue(targetSite, out var known) && MatchListByTitle(title!, known) is { } hit) targetId = hit.Id;
                targetId ??= await lookup.GetTargetListIdAsync(targetSite, title!, ct);
            }
            var sourceCols = title is null ? [] : await lookup.GetSourceColumnsAsync(readSite, listId, ct);
            var sourceKnown = title is not null;
            if (targetId is null)
                return new(new(site, listId, title, targetSite, null, ListMatchStatus.TargetListNotFound), [], null, sourceCols, UsedNames(site, listId, sourceCols, sourceKnown));

            var listResolution = new ListResolution(site, listId, title, targetSite, targetId, manual is not null ? ListMatchStatus.Manual : ListMatchStatus.Matched);
            var columns = new List<ColumnResolution>();

            // Columns this flow uses on this list: written fields, names inside $filter/$orderby/$select, and (for a
            // flow with exactly one list) columns read in expressions.
            progress?.Invoke(Loc.T("Flows_Prog_ReadingTargetColumns", title ?? listId));
            var targetCols = await lookup.GetTargetColumnsAsync(targetSite, targetId, ct);
            var used = UsedNames(site, listId, sourceCols, sourceKnown);
            var builtInOrUnlisted = 0;

            foreach (var name in used.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
            {
                var src = sourceCols.FirstOrDefault(c => c.InternalName == name);
                // With the source list readable, a name in neither list is a built-in field (ID, Author, Created...) or an
                // unsupported type: nothing to map, nothing to warn about. With the source unreadable we cannot tell, so the
                // name is kept and shown for the user to decide.
                if (sourceKnown && src is null && !targetCols.Any(c => c.InternalName == name)) { builtInOrUnlisted++; continue; }
                var (target, status) = MatchColumn(name, src, targetCols);
                columns.Add(new(site, listId, title, name, target, status));
            }

            // Say, in plain words, what was found for this list, so an empty column grid is never a mystery.
            var label = title ?? listId;
            int Count(ColumnMatchStatus s) => columns.Count(c => c.Status == s);
            if (sourceCols.Count == 0 && targetCols.Count == 0)
                progress?.Invoke(Loc.T("Flows_Prog_NoColumns", label));
            else if (used.Count == 0)
                progress?.Invoke(Loc.T("Flows_Prog_NoneUsed", label, sourceCols.Count));
            else
                progress?.Invoke(Loc.T("Flows_Prog_Summary", label, used.Count, Count(ColumnMatchStatus.Same), Count(ColumnMatchStatus.Renamed),
                    Count(ColumnMatchStatus.Similar), Count(ColumnMatchStatus.NotFoundInTarget), builtInOrUnlisted));
            return new(listResolution, columns, targetCols, sourceCols, used);
        }

        // The column names the flows use on one list. When the source list could not be read, names cannot be checked against
        // its columns: the names written by the flow's steps and read in its expressions are all kept (they come from the flow
        // itself), and OData query text, which can only be checked against real columns, is skipped.
        HashSet<string> UsedNames(string site, string listId, IReadOnlyList<FlowColumnInfo> sourceCols, bool sourceKnown)
        {
            var sourceNames = sourceCols.Select(c => c.InternalName).ToHashSet(StringComparer.Ordinal);
            var used = new HashSet<string>(StringComparer.Ordinal);
            foreach (var f in findings.Where(f => f.Site is not null && f.ListId == listId && NormalizeSite(f.Site) == site))
            {
                if (f.Kind == FindingKind.ColumnName) used.Add(f.Value);
                else if (sourceKnown && f.Kind == FindingKind.QueryFilter)
                    foreach (Match m in Identifier().Matches(f.Value))
                        if (sourceNames.Contains(m.Value)) used.Add(m.Value);
            }
            if (expressionColumns.TryGetValue((site, listId), out var fromExpressions))
                foreach (var name in fromExpressions.Where(n => !sourceKnown || sourceNames.Contains(n))) used.Add(name);
            return used;
        }

        foreach (var site in findings.Where(f => f.Kind == FindingKind.SiteUrl).Select(f => NormalizeSite(f.Value)).Distinct(StringComparer.OrdinalIgnoreCase))
            if (!sites.ContainsKey(site) || string.IsNullOrWhiteSpace(sites[site]))
                result.Notes.Add(Loc.T("Flows_Note_SiteNoTarget", site));

        return result;
    }

    [GeneratedRegex(@"[A-Za-z_][A-Za-z0-9_]*")] private static partial Regex Identifier();
    [GeneratedRegex(@"[^A-Za-z0-9]")] private static partial Regex NonAlnum();
}
