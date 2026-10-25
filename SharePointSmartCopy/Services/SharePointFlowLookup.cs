using SharePointSmartCopy.Workflows;

namespace SharePointSmartCopy.Services;

// Connects the Flows mapping resolver to the app's SharePoint access. In multi-tenant mode every site is first
// registered with the right side's sign-in (prepareSite), exactly like Compare and History do. The cancellation
// token reaches the sign-in itself, so Cancel can stop a browser sign-in that is waiting.
public sealed class SharePointFlowLookup : IFlowSiteLookup
{
    private readonly SharePointService _sp;
    private readonly Func<string, bool, CancellationToken, Task> _prepareSite;   // (siteUrl, isSource, ct)
    private readonly Action<string> _warn;

    public SharePointFlowLookup(SharePointService sp, Func<string, bool, CancellationToken, Task> prepareSite, Action<string> warn)
    {
        _sp = sp;
        _prepareSite = prepareSite;
        _warn = warn;
    }

    // The app's own error ("source and target are on the same tenant") doesn't say which site clashed or on which
    // side; add both so the Sites step can be corrected.
    private async Task PrepareSiteAsync(string siteUrl, bool isSource, CancellationToken ct)
    {
        try { await _prepareSite(siteUrl, isSource, ct); }
        catch (InvalidOperationException ex)
        {
            throw new InvalidOperationException(
                $"{ex.Message} ({(isSource ? "Source" : "Target")} site: {siteUrl})", ex);
        }
    }

    public async Task PrepareAsync(string sourceSite, string targetSite, CancellationToken ct)
    {
        await PrepareSiteAsync(sourceSite, true, ct);
        await PrepareSiteAsync(targetSite, false, ct);
    }

    public async Task<string?> GetSourceListTitleAsync(string siteUrl, string listId, CancellationToken ct)
    {
        await PrepareSiteAsync(siteUrl, true, ct);
        var title = await _sp.GetListTitleByIdAsync(siteUrl, listId, ct);
        // The service returns null for every failure (missing list, no access, throttling); say so instead of
        // letting it read as a plain "not found".
        if (title is null) _warn($"List {listId} on {siteUrl} could not be read (it may not exist, or this account may lack access).");
        return title;
    }

    public async Task<string?> GetTargetListIdAsync(string siteUrl, string title, CancellationToken ct)
    {
        await PrepareSiteAsync(siteUrl, false, ct);
        return await _sp.GetListIdByTitleAsync(siteUrl, title, ct);
    }

    public async Task<IReadOnlyList<FlowListInfo>> GetTargetListsAsync(string siteUrl, CancellationToken ct)
    {
        try
        {
            await PrepareSiteAsync(siteUrl, false, ct);
            var lists = await _sp.GetSiteListsAsync(siteUrl, ct);
            return lists.Select(l => new FlowListInfo(l.Id, l.Title)).ToList();
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            // Without the list of lists the dropdown is empty, but a list can still be matched by name or pasted by ID.
            _warn($"Could not read the lists on {siteUrl}: {ex.Message}");
            return [];
        }
    }
    public Task<IReadOnlyList<FlowColumnInfo>> GetSourceColumnsAsync(string siteUrl, string listId, CancellationToken ct) =>
        ColumnsAsync(siteUrl, listId, isSource: true, ct);

    public Task<IReadOnlyList<FlowColumnInfo>> GetTargetColumnsAsync(string siteUrl, string listId, CancellationToken ct) =>
        ColumnsAsync(siteUrl, listId, isSource: false, ct);

    private async Task<IReadOnlyList<FlowColumnInfo>> ColumnsAsync(string siteUrl, string listId, bool isSource, CancellationToken ct)
    {
        try
        {
            await PrepareSiteAsync(siteUrl, isSource, ct);
            var cols = await _sp.GetListFieldsAsync(siteUrl, listId, ct);
            return cols.Select(c => new FlowColumnInfo(c.InternalName, c.Title, c.Type, c.Required)).ToList();
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            // One unreadable list must not abort the whole resolution; its columns just can't be matched.
            _warn($"Could not read columns of list {listId} on {siteUrl}: {ex.Message}");
            return [];
        }
    }
}
