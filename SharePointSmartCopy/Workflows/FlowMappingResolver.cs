using System.Text.RegularExpressions;

namespace SharePointSmartCopy.Workflows;

public sealed record FlowColumnInfo(string InternalName, string DisplayName);

// What the resolver needs from SharePoint. The app implements this with SharePointService (and the
// multi-tenant sign-in routing); tests use a fake.
public interface IFlowSiteLookup
{
    // Title of a list in the SOURCE site, by GUID (null = not found).
    Task<string?> GetSourceListTitleAsync(string siteUrl, string listId, CancellationToken ct);
    // GUID of a list in the TARGET site, by title (null = not found).
    Task<string?> GetTargetListIdAsync(string siteUrl, string title, CancellationToken ct);
    Task<IReadOnlyList<FlowColumnInfo>> GetSourceColumnsAsync(string siteUrl, string listId, CancellationToken ct);
    Task<IReadOnlyList<FlowColumnInfo>> GetTargetColumnsAsync(string siteUrl, string listId, CancellationToken ct);
    // Signs in to (and registers) both sites before lookups run in parallel, so a first-time sign-in happens once,
    // up front, instead of from several workers at once.
    Task PrepareAsync(string sourceSite, string targetSite, CancellationToken ct);
}

public enum ListMatchStatus { Matched, SiteNotMapped, SourceListNotFound, TargetListNotFound, Manual }   // Manual = target chosen by the user

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

    // The target list's columns, keyed by SOURCE list GUID, so a UI can offer them as the choices for each row.
    public Dictionary<string, IReadOnlyList<FlowColumnInfo>> TargetColumnsBySourceList { get; } = new(StringComparer.OrdinalIgnoreCase);

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

    private sealed record PairOutcome(ListResolution List, List<ColumnResolution> Columns, IReadOnlyList<FlowColumnInfo>? TargetColumns);

    public static string NormalizeSite(string url) => url.Trim().TrimEnd('/');

    // "Tax_x0020_ID", "Tax ID", "tax id", "TaxID" all normalize to "taxid".
    public static string NormalizeColumnName(string name) =>
        NonAlnum().Replace(name.Replace("_x0020_", "", StringComparison.OrdinalIgnoreCase), "").ToLowerInvariant();

    public static async Task<FlowResolveResult> ResolveAsync(
        IEnumerable<FlowScan> scans,
        IReadOnlyDictionary<string, string> siteMap,
        IFlowSiteLookup lookup,
        CancellationToken ct = default,
        IReadOnlyDictionary<string, string>? listOverrides = null)   // source list GUID -> target list GUID chosen by the user
    {
        var result = new FlowResolveResult();
        var sites = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in siteMap) sites[NormalizeSite(k)] = NormalizeSite(v);

        var findings = scans.SelectMany(s => s.Findings).ToList();

        // Distinct (site, list) pairs the flows actually use with literal values.
        var pairs = findings.Where(f => f is { Kind: FindingKind.ListId, Site: not null, ListId: not null })
            .Select(f => (Site: NormalizeSite(f.Site!), List: f.ListId!))
            .Distinct().ToList();

        // Sign in once per site pair, in order, before the lookups fan out.
        foreach (var (site, targetSite) in pairs.Select(p => p.Site).Distinct(StringComparer.OrdinalIgnoreCase)
                     .Where(s => sites.TryGetValue(s, out var t) && !string.IsNullOrWhiteSpace(t))
                     .Select(s => (s, sites[s])))
        {
            ct.ThrowIfCancellationRequested();
            await lookup.PrepareAsync(site, targetSite, ct);
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
        }

        async Task<PairOutcome> ResolvePairAsync(string site, string listId)
        {
            ct.ThrowIfCancellationRequested();
            if (!sites.TryGetValue(site, out var targetSite) || string.IsNullOrWhiteSpace(targetSite))
                return new(new(site, listId, null, null, null, ListMatchStatus.SiteNotMapped), [], null);

            var title = await lookup.GetSourceListTitleAsync(site, listId, ct);
            var manual = listOverrides is not null && listOverrides.TryGetValue(listId, out var chosen) && !string.IsNullOrWhiteSpace(chosen) ? chosen : null;
            if (title is null && manual is null)
                return new(new(site, listId, null, targetSite, null, ListMatchStatus.SourceListNotFound), [], null);

            // A target the user chose wins over name matching (it is how lists with different names are mapped).
            var targetId = manual ?? await lookup.GetTargetListIdAsync(targetSite, title!, ct);
            if (targetId is null)
                return new(new(site, listId, title, targetSite, null, ListMatchStatus.TargetListNotFound), [], null);

            var listResolution = new ListResolution(site, listId, title, targetSite, targetId, manual is not null ? ListMatchStatus.Manual : ListMatchStatus.Matched);
            var columns = new List<ColumnResolution>();

            // Columns this flow uses on this list: written fields, plus names inside $filter/$orderby/$select.
            var sourceCols = await lookup.GetSourceColumnsAsync(site, listId, ct);
            var targetCols = await lookup.GetTargetColumnsAsync(targetSite, targetId, ct);
            var sourceNames = sourceCols.Select(c => c.InternalName).ToHashSet(StringComparer.Ordinal);

            var used = new HashSet<string>(StringComparer.Ordinal);
            foreach (var f in findings.Where(f => f.Site is not null && f.ListId == listId && NormalizeSite(f.Site) == site))
            {
                if (f.Kind == FindingKind.ColumnName) used.Add(f.Value);
                else if (f.Kind == FindingKind.QueryFilter)
                    foreach (Match m in Identifier().Matches(f.Value))
                        if (sourceNames.Contains(m.Value)) used.Add(m.Value);
            }

            foreach (var name in used.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
            {
                var src = sourceCols.FirstOrDefault(c => c.InternalName == name);
                var sameInternal = targetCols.FirstOrDefault(c => c.InternalName == name);
                // The app's column listing covers custom columns only. A name in neither list is a built-in
                // field (Title, ID, Author...) or an unsupported type: nothing to map, nothing to warn about.
                if (src is null && sameInternal is null) continue;
                if (sameInternal is not null)
                {
                    columns.Add(new(site, listId, title, name, name, ColumnMatchStatus.Same));
                    continue;
                }

                var byDisplay = src is null ? null
                    : targetCols.FirstOrDefault(c => string.Equals(c.DisplayName, src.DisplayName, StringComparison.OrdinalIgnoreCase));
                if (byDisplay is not null)
                {
                    columns.Add(new(site, listId, title, name, byDisplay.InternalName, ColumnMatchStatus.Renamed));
                    continue;
                }

                // Third tier: names that differ only in case, spaces, underscores or "_x0020_". Only a UNIQUE
                // candidate is proposed; an ambiguous one is left for the user.
                var wanted = new[] { NormalizeColumnName(name), src is null ? "" : NormalizeColumnName(src.DisplayName) }
                    .Where(w => w.Length > 0).ToHashSet();
                var similar = targetCols
                    .Where(c => wanted.Contains(NormalizeColumnName(c.InternalName)) || wanted.Contains(NormalizeColumnName(c.DisplayName)))
                    .ToList();
                columns.Add(similar.Count == 1
                    ? new(site, listId, title, name, similar[0].InternalName, ColumnMatchStatus.Similar)
                    : new(site, listId, title, name, null, ColumnMatchStatus.NotFoundInTarget));
            }
            return new(listResolution, columns, targetCols);
        }

        foreach (var site in findings.Where(f => f.Kind == FindingKind.SiteUrl).Select(f => NormalizeSite(f.Value)).Distinct(StringComparer.OrdinalIgnoreCase))
            if (!sites.ContainsKey(site) || string.IsNullOrWhiteSpace(sites[site]))
                result.Notes.Add($"Site '{site}' has no target; lists on it were not resolved.");

        return result;
    }

    [GeneratedRegex(@"[A-Za-z_][A-Za-z0-9_]*")] private static partial Regex Identifier();
    [GeneratedRegex(@"[^A-Za-z0-9]")] private static partial Regex NonAlnum();
}
