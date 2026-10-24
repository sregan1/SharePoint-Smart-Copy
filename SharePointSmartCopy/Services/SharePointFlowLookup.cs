using SharePointSmartCopy.Workflows;

namespace SharePointSmartCopy.Services;

// Connects the Flows mapping resolver to the app's SharePoint access. In multi-tenant mode every site is first
// registered with the right side's sign-in (prepareSite), exactly like Compare and History do.
public sealed class SharePointFlowLookup : IFlowSiteLookup
{
    private readonly SharePointService _sp;
    private readonly Func<string, bool, Task> _prepareSite;   // (siteUrl, isSource)
    private readonly Action<string> _warn;

    public SharePointFlowLookup(SharePointService sp, Func<string, bool, Task> prepareSite, Action<string> warn)
    {
        _sp = sp;
        _prepareSite = prepareSite;
        _warn = warn;
    }

    public async Task PrepareAsync(string sourceSite, string targetSite, CancellationToken ct)
    {
        await _prepareSite(sourceSite, true);
        await _prepareSite(targetSite, false);
    }

    public async Task<string?> GetSourceListTitleAsync(string siteUrl, string listId, CancellationToken ct)
    {
        await _prepareSite(siteUrl, true);
        var title = await _sp.GetListTitleByIdAsync(siteUrl, listId);
        // The service returns null for every failure (missing list, no access, throttling); say so instead of
        // letting it read as a plain "not found".
        if (title is null) _warn($"List {listId} on {siteUrl} could not be read (it may not exist, or this account may lack access).");
        return title;
    }

    public async Task<string?> GetTargetListIdAsync(string siteUrl, string title, CancellationToken ct)
    {
        await _prepareSite(siteUrl, false);
        return await _sp.GetListIdByTitleAsync(siteUrl, title);
    }

    public Task<IReadOnlyList<FlowColumnInfo>> GetSourceColumnsAsync(string siteUrl, string listId, CancellationToken ct) =>
        ColumnsAsync(siteUrl, listId, isSource: true);

    public Task<IReadOnlyList<FlowColumnInfo>> GetTargetColumnsAsync(string siteUrl, string listId, CancellationToken ct) =>
        ColumnsAsync(siteUrl, listId, isSource: false);

    private async Task<IReadOnlyList<FlowColumnInfo>> ColumnsAsync(string siteUrl, string listId, bool isSource)
    {
        try
        {
            await _prepareSite(siteUrl, isSource);
            var cols = await _sp.GetLibraryColumnsAsync(siteUrl, listId, skipCache: true);
            return cols.Select(c => new FlowColumnInfo(c.InternalName, c.DisplayName)).ToList();
        }
        catch (Exception ex)
        {
            // One unreadable list must not abort the whole resolution; its columns just can't be matched.
            _warn($"Could not read columns of list {listId} on {siteUrl}: {ex.Message}");
            return [];
        }
    }
}
