using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Http.HttpClientLibrary;
using SharePointSmartCopy.Models;
using SpColumnDef  = SharePointSmartCopy.Models.ColumnDefinition;
using SpColumnMap  = SharePointSmartCopy.Models.ColumnMapping;

namespace SharePointSmartCopy.Services;

public class SharePointService
{
    private readonly AuthService _authService;
    private GraphServiceClient? _graphClient;
    // 30-minute timeout: site copies involve large file downloads and long-running REST calls.
    private readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromMinutes(30) };
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (string siteUrl, string listId, string listItemId)> _spIdsCache = new();
    // Deduplicates concurrent folder-creation calls for the same path segment.
    // Lazy<Task<string>> ensures only one Graph call is made per "{driveId}|{parentId}|{segment}"
    // key even when multiple parallel tasks race — all share the same Task and await its result.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, Lazy<Task<string>>> _folderSegmentTasks = new(StringComparer.OrdinalIgnoreCase);
    // Limits concurrent Graph $batch calls across all parallel SPMI job pipelines.
    // Each batch call carries 20 sub-requests; 10 parallel jobs firing simultaneously
    // would produce 200 concurrent Graph ops, reliably hitting per-app throttle limits.
    // 3 permits → at most 60 concurrent Graph ops from batch calls at any instant.
    private readonly System.Threading.SemaphoreSlim _batchGate = new(3, 3);

    public SharePointService(AuthService authService)
    {
        _authService = authService;
    }

    public Task<string> GetSharePointTokenAsync(string siteUrl, string scope = "Sites.ReadWrite.All")
        => _authService.GetSharePointTokenAsync(siteUrl, scope);

    // Returns the root item ID for a drive (needed to target uploads to a library root).
    public async Task<string?> GetLibraryRootItemIdAsync(string driveId)
    {
        try
        {
            var root = await Graph.Drives[driveId].Root.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id"]);
            return root?.Id;
        }
        catch { return null; }
    }

    // Returns the SharePoint folder UniqueId (GUID) for the library root folder.
    // Required for the SPMI manifest SPFolder entry — without it SPMI cannot resolve the
    // parent folder for files in newly created empty libraries ("Missing file info" error).
    public async Task<string?> GetLibraryRootFolderUniqueIdAsync(string siteUrl, string libraryServerRelativeUrl)
    {
        // Use GetFolderByServerRelativePath(decodedurl=...) instead of ...ByServerRelativeUrl(...):
        // the Url variant mishandles special characters in folder names (e.g. '#' in "SOW#3", '%',
        // '&') and returns 404 for them, which then failed the whole batch's manifest with the
        // "Could not resolve target subfolder ID" guard. The Path/decodedurl variant resolves them
        // correctly. Single quotes are OData-escaped (doubled), then the literal is URL-encoded.
        var odataLiteral = libraryServerRelativeUrl.Replace("'", "''");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetFolderByServerRelativePath(decodedurl='{Uri.EscapeDataString(odataLiteral)}')?$select=UniqueId";
        try
        {
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);
            if (!response.IsSuccessStatusCode) return null;
            var body = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("UniqueId", out var uid) ? uid.GetString() : null;
        }
        catch { return null; }
    }

    public void Initialize()
    {
        var provider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(_authService));
        var handlers = KiotaClientFactory.CreateDefaultHandlers();
        // CreateDefaultHandlers() order: [0]=UriReplacementHandler [1]=RetryHandler [2]=RedirectHandler ...
        // Replace the RetryHandler (index 1) with one configured for large migrations.
        // Default MaxRetry=3; with 10 parallel SPMI jobs the Graph per-app rate limit is
        // hit frequently enough that 3 retries is insufficient — bump to 10 (library max).
        handlers[1] = new Microsoft.Kiota.Http.HttpClientLibrary.Middleware.RetryHandler(
            new Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options.RetryHandlerOption
            {
                MaxRetry = 10,
            });
        // Insert our throttle-notify handler after RetryHandler (now at index 2, after RedirectHandler)
        // so it sits inside Kiota's retry loop and sees raw 429/503 responses before they are absorbed.
        // This makes Graph throttles visible to the adaptive parallelism controller.
        handlers.Insert(3, new GraphThrottleNotifyHandler(
            (delay, attempt, max, reason) => Throttled?.Invoke(delay, attempt, max, reason)));
        // Configure the underlying socket handler: gzip so metadata/JSON responses transfer
        // compressed (binary downloads are already-compressed and unaffected), and recycle pooled
        // connections every few minutes for stability across multi-hour 20k+ file runs.
        // EnableMultipleHttp2Connections matters at our concurrency levels: SocketsHttpHandler
        // multiplexes ALL requests to a host over a SINGLE HTTP/2 connection by default, so up to
        // 16 concurrent Graph calls (the app's max parallel-copies setting) would otherwise share
        // one TCP connection's congestion window and frame-processing loop instead of spreading
        // across several — a client-side ceiling independent of (and easy to mistake for) Graph's
        // own throttling.
        var finalHandler = new System.Net.Http.SocketsHttpHandler
        {
            AutomaticDecompression         = System.Net.DecompressionMethods.All,
            PooledConnectionLifetime       = TimeSpan.FromMinutes(2),
            EnableMultipleHttp2Connections = true,
        };
        var httpClient = KiotaClientFactory.Create(handlers, finalHandler);
        httpClient.Timeout = TimeSpan.FromMinutes(30);
        var adapter  = new HttpClientRequestAdapter(provider, httpClient: httpClient);
        _graphClient = new GraphServiceClient(adapter);
    }

    private GraphServiceClient Graph => _graphClient
        ?? throw new InvalidOperationException("Not initialized. Please sign in first.");

    // ── Site ──────────────────────────────────────────────────────────────────

    public async Task<string> GetSiteIdAsync(string siteUrl)
    {
        var uri = new Uri(siteUrl.TrimEnd('/'));
        var hostname = uri.Host;
        var path = uri.AbsolutePath.TrimStart('/');
        var key = string.IsNullOrEmpty(path) ? hostname : $"{hostname}:/{path}";

        var site = await Graph.Sites[key].GetAsync();
        return site?.Id ?? throw new Exception($"Could not find site at {siteUrl}");
    }

    // ── Libraries ─────────────────────────────────────────────────────────────

    public async Task<List<SharePointNode>> GetLibrariesAsync(string siteId, string siteUrl)
    {
        // No $select — the Kiota SDK does not reliably populate webUrl when explicitly selected
        // on the drives collection; the default response includes it.
        var drives = await Graph.Sites[siteId].Drives.GetAsync();
        var result = new List<SharePointNode>();

        foreach (var d in drives?.Value ?? [])
        {
            if (d.Id == null || d.Name == null) continue;

            string? serverRelUrl = null;
            if (d.WebUrl != null)
            {
                var uri = new Uri(d.WebUrl);
                var path = Uri.UnescapeDataString(uri.AbsolutePath.TrimEnd('/'));
                // Exclude view URLs like /Forms/AllItems.aspx — these aren't the library root
                if (!path.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
                    serverRelUrl = path;
            }

            var node = new SharePointNode
            {
                Id                 = "root",
                DriveId            = d.Id,
                SiteId             = siteId,
                SiteUrl            = siteUrl,
                Name               = d.Name,
                Type               = NodeType.Library,
                HasChildren        = true,
                ServerRelativePath = serverRelUrl
            };
            node.Children.Add(Placeholder());
            result.Add(node);
        }
        return result;
    }

    // ── Children ──────────────────────────────────────────────────────────────

    public async Task<List<SharePointNode>> GetChildrenAsync(
        string driveId, string itemId, string siteId, string siteUrl = "", bool foldersOnly = false)
    {
        var resolvedId = itemId == "root" ? "root" : itemId;
        var page = await Graph.Drives[driveId].Items[resolvedId].Children
            .GetAsync(cfg => cfg.QueryParameters.Top = 1000);

        var items = new List<DriveItem>();
        while (page != null)
        {
            items.AddRange(page.Value ?? []);
            if (page.OdataNextLink == null) break;
            page = await Graph.Drives[driveId].Items[resolvedId].Children
                .WithUrl(page.OdataNextLink).GetAsync();
        }

        var result = new List<SharePointNode>();
        foreach (var item in items)
        {
            if (item.Id == null || item.Name == null) continue;
            if (foldersOnly && item.Folder == null) continue;

            bool isFolder = item.Folder != null;
            var node = new SharePointNode
            {
                Id          = item.Id,
                DriveId     = driveId,
                SiteId      = siteId,
                SiteUrl     = siteUrl,
                Name        = item.Name,
                Type        = isFolder ? NodeType.Folder : NodeType.File,
                Size        = item.Size,
                WebUrl      = item.WebUrl,
                HasChildren = isFolder && (item.Folder?.ChildCount ?? 0) > 0
            };

            if (node.HasChildren)
                node.Children.Add(Placeholder());

            result.Add(node);
        }
        return result;
    }

    // ── Enumerate all files under a folder (for copy) ─────────────────────────

    // One discovered source file, with the modified date the Children listing already carries —
    // captured here so the If Newer overwrite decision never has to re-fetch it per file later.
    public readonly record struct SourceFileEntry(
        string DriveId, string ItemId, string Name, string RelativePath, DateTimeOffset? Modified);

    // Concurrent replacement for the old sequential recursive walk (same channel + sibling-fan-out
    // pattern as EnumerateFilesWithMetadataAsync below). The sequential version issued ONE Graph
    // call at a time — on a 3,000+-folder library that alone took ~30 minutes before the copy could
    // even start, with nothing logged. The controller bounds concurrent listings and backs off on
    // throttle; $select trims 100k+ full DriveItem payloads down to the five fields the copy needs.
    internal async IAsyncEnumerable<SourceFileEntry> EnumerateFilesForCopyAsync(
        string driveId, string rootItemId, string basePath,
        AdaptiveParallelismController controller,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
    {
        var channel = System.Threading.Channels.Channel.CreateUnbounded<SourceFileEntry>();
        var walkTask = Task.Run(async () =>
        {
            try { await WalkFilesForCopyAsync(driveId, rootItemId, basePath, controller, channel.Writer, ct); }
            finally { channel.Writer.Complete(); }
        }, ct);

        await foreach (var f in channel.Reader.ReadAllAsync(ct))
            yield return f;

        await walkTask; // propagate walk exceptions (channel completion alone would swallow them)
    }

    private async Task WalkFilesForCopyAsync(
        string driveId, string itemId, string basePath, AdaptiveParallelismController controller,
        System.Threading.Channels.ChannelWriter<SourceFileEntry> writer, CancellationToken ct)
    {
        List<DriveItem> items;
        await controller.WaitAsync(ct);
        try { items = await GetChildrenWithMetadataAsync(driveId, itemId, ct); }
        finally { controller.Release(); }

        var subfolderWalks = new List<Task>();
        foreach (var item in items)
        {
            ct.ThrowIfCancellationRequested();
            if (item.Id == null || item.Name == null) continue;
            var childPath = string.IsNullOrEmpty(basePath) ? item.Name : $"{basePath}/{item.Name}";
            if (item.Folder != null)
                subfolderWalks.Add(WalkFilesForCopyAsync(driveId, item.Id, childPath, controller, writer, ct));
            else
                await writer.WriteAsync(new SourceFileEntry(
                    driveId, item.Id, item.Name, childPath, item.LastModifiedDateTime), ct);
        }
        await Task.WhenAll(subfolderWalks);
    }

    // ── Enumerate all sub-folders under a folder ──────────────────────────────

    public async IAsyncEnumerable<(string driveId, string itemId, string relativePath)>
        EnumerateFoldersAsync(string driveId, string rootItemId, string basePath = "")
    {
        var children = await GetChildrenAsync(driveId, rootItemId, string.Empty);
        foreach (var child in children.Where(c => c.Type == NodeType.Folder))
        {
            var childPath = string.IsNullOrEmpty(basePath) ? child.Name : $"{basePath}/{child.Name}";
            yield return (driveId, child.Id, childPath);
            await foreach (var item in EnumerateFoldersAsync(driveId, child.Id, childPath))
                yield return item;
        }
    }

    // ── Enumerate all files with metadata, for independent post-copy verification ──────────
    // Separate from EnumerateFilesAsync/GetChildrenAsync (used by the copy engine) so a
    // verification re-scan never shares state/behavior with the copy path, and so its Graph
    // payload can be trimmed with an explicit $select for 100k+ file scale.
    //
    // Sibling subfolders are walked CONCURRENTLY (fanned out as tasks, funneled through a
    // channel), not one at a time — a plain recursive `await foreach` would serialize the entire
    // folder-tree walk to one Graph call in flight at a time regardless of how much concurrency
    // the controller allows, since each subfolder's whole subtree would have to finish before its
    // sibling starts. The controller still bounds how many of those concurrent walks are actually
    // making Graph calls at once (extra tasks just queue on controller.WaitAsync), so this doesn't
    // bypass throttle protection — it just lets the walk actually use the concurrency budget the
    // controller already grants, the same way the file-copy pipeline does with Parallel.ForEachAsync.
    internal async IAsyncEnumerable<ScannedFile> EnumerateFilesWithMetadataAsync(
        string driveId, string rootItemId, string basePath,
        AdaptiveParallelismController controller,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
    {
        var channel = System.Threading.Channels.Channel.CreateUnbounded<ScannedFile>();
        var walkTask = Task.Run(async () =>
        {
            try { await WalkFilesWithMetadataAsync(driveId, rootItemId, basePath, controller, channel.Writer, ct); }
            finally { channel.Writer.Complete(); }
        }, ct);

        await foreach (var f in channel.Reader.ReadAllAsync(ct))
            yield return f;

        await walkTask; // propagate any exception from the walk (channel completion alone would swallow it)
    }

    private async Task WalkFilesWithMetadataAsync(
        string driveId, string itemId, string basePath, AdaptiveParallelismController controller,
        System.Threading.Channels.ChannelWriter<ScannedFile> writer, CancellationToken ct)
    {
        List<DriveItem> items;
        await controller.WaitAsync(ct);
        try { items = await GetChildrenWithMetadataAsync(driveId, itemId, ct); }
        finally { controller.Release(); }

        var subfolderWalks = new List<Task>();
        foreach (var item in items)
        {
            ct.ThrowIfCancellationRequested();
            if (item.Id == null || item.Name == null) continue;
            var childPath = string.IsNullOrEmpty(basePath) ? item.Name : $"{basePath}/{item.Name}";
            if (item.Folder != null)
            {
                subfolderWalks.Add(WalkFilesWithMetadataAsync(driveId, item.Id, childPath, controller, writer, ct));
            }
            else
            {
                await writer.WriteAsync(new ScannedFile(driveId, item.Id, item.Name, childPath,
                    item.LastModifiedDateTime), ct);
            }
        }
        await Task.WhenAll(subfolderWalks);
    }

    // Read-only counterpart to GetOrCreateFolderPathAsync: walks an existing path segment-by-segment
    // (same per-segment escaping, never creates anything) to resolve the item id a relative path
    // points at — used by verification to navigate to the actual copied folder/file under a
    // library root, rather than assuming the path IS the item id. Returns null if any segment is
    // missing (e.g. the folder was deleted or renamed since the original copy).
    internal async Task<string?> ResolveItemIdByPathAsync(string driveId, string parentItemId, string relativePath)
    {
        var current = parentItemId;
        foreach (var part in relativePath.Split('/').Where(p => !string.IsNullOrEmpty(p)))
        {
            try
            {
                var item = await Graph.Drives[driveId].Items[current]
                    .ItemWithPath(Uri.EscapeDataString(part))
                    .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
                if (item?.Id == null) return null;
                current = item.Id;
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
            {
                return null;
            }
        }
        return current;
    }

    // Fetches a single item directly (no recursion) for verification of a plain single-file root.
    internal async Task<ScannedFile?> GetFileForVerificationAsync(string driveId, string itemId, string relativePath)
    {
        var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "name", "file", "lastModifiedDateTime"]);
        if (item?.Id == null || item.Name == null || item.File == null) return null;
        return new ScannedFile(driveId, item.Id, item.Name, relativePath, item.LastModifiedDateTime);
    }

    // $select keeps the payload small at 100k+ files: only fields the comparison needs. Size and
    // content hash were tried as match signals and dropped (see ComparisonStatus) — verification
    // is existence-only now, so only identity/type/date fields are requested.
    private async Task<List<DriveItem>> GetChildrenWithMetadataAsync(
        string driveId, string itemId, CancellationToken ct)
    {
        var resolvedId = itemId == "root" ? "root" : itemId;
        var page = await Graph.Drives[driveId].Items[resolvedId].Children.GetAsync(cfg =>
        {
            cfg.QueryParameters.Top    = 1000;
            cfg.QueryParameters.Select = ["id", "name", "file", "folder", "lastModifiedDateTime"];
        }, ct);

        var items = new List<DriveItem>();
        while (page != null)
        {
            items.AddRange(page.Value ?? []);
            if (page.OdataNextLink == null) break;
            page = await Graph.Drives[driveId].Items[resolvedId].Children
                .WithUrl(page.OdataNextLink).GetAsync(cancellationToken: ct);
        }
        return items;
    }

    // ── File content ──────────────────────────────────────────────────────────

    public async Task<Stream> DownloadFileAsync(string driveId, string itemId)
    {
        var stream = await Graph.Drives[driveId].Items[itemId].Content.GetAsync();
        return stream ?? throw new Exception("Empty response downloading file.");
    }

    public async Task<Stream> DownloadVersionAsync(string driveId, string itemId, string versionId)
    {
        var stream = await Graph.Drives[driveId].Items[itemId].Versions[versionId].Content.GetAsync();
        return stream ?? throw new Exception("Empty response downloading version content.");
    }

    public async Task<List<DriveItemVersion>> GetVersionsAsync(string driveId, string itemId)
    {
        var all  = new List<DriveItemVersion>();
        var page = await Graph.Drives[driveId].Items[itemId].Versions.GetAsync();
        while (page != null)
        {
            all.AddRange(page.Value ?? []);
            if (page.OdataNextLink == null) break;
            page = await Graph.Drives[driveId].Items[itemId].Versions
                .WithUrl(page.OdataNextLink).GetAsync();
        }
        return SortVersions(all);
    }

    // ── Graph $batch metadata + versions ─────────────────────────────────────

    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    // Fetches metadata and the full version list for each item using Graph $batch
    // (up to 20 sub-requests per call: 1 metadata + 1 versions list per item, 10 items/batch).
    // Items absent from the result failed their sub-request; callers should fall back
    // to individual GetFileMetadataAsync / GetVersionsAsync calls for those items.
    //
    // multiVersionItemIds: when supplied, the per-file /versions sub-request is issued ONLY for items
    // in this set (those an upstream analyze pass already counted as having >1 version). Single-version
    // files — the bulk of typical libraries — get a metadata-only fetch (1 sub-request instead of 2)
    // and a synthetic single current-version entry, roughly HALVING Graph request volume here and
    // sharply reducing throttling. Pass null to fetch versions for every item (the fallback path).
    public async Task<Dictionary<string, (FileMetadata Metadata, List<DriveItemVersion> Versions)>>
        BatchFetchMetadataAndVersionsAsync(
            IReadOnlyList<(string driveId, string itemId)> items,
            CancellationToken ct = default,
            IReadOnlySet<string>? multiVersionItemIds = null)
    {
        var result = new Dictionary<string, (FileMetadata, List<DriveItemVersion>)>();
        if (items.Count == 0) return result;

        // Up to 2 Graph requests per file (metadata + versions); Graph batch cap is 20 → 10 files/call.
        const int ItemsPerBatch = 10;

        foreach (var chunk in items.Chunk(ItemsPerBatch))
        {
            ct.ThrowIfCancellationRequested();
            await _batchGate.WaitAsync(ct);
            try { await FetchBatchChunkAsync(chunk, result, ct, multiVersionItemIds); }
            catch (OperationCanceledException) { throw; }
            catch { /* chunk failed; items absent from result, caller falls back */ }
            finally { _batchGate.Release(); }
        }

        return result;
    }

    // Fetches ONLY the version count per item (capped at maxVersionsCap), in PARALLEL and
    // memory-light: the heavy version objects are discarded immediately after counting, so only one
    // small int per file persists. Issues just the /versions sub-request (no metadata round-trip —
    // this pass doesn't use metadata), halving its Graph footprint vs a full metadata+versions fetch.
    // Used to size version-aware migration batches over a whole library (tens/hundreds of thousands of
    // files) without holding all metadata in memory. Items whose sub-request fails default to a count of 1.
    public async Task<Dictionary<string, int>> FetchVersionCountsAsync(
        IReadOnlyList<(string driveId, string itemId)> items,
        int maxVersionsCap,
        int maxConcurrency,
        IProgress<int>? progress = null,
        CancellationToken ct = default)
    {
        var counts = new System.Collections.Concurrent.ConcurrentDictionary<string, int>();
        if (items.Count == 0) return new Dictionary<string, int>();

        // Versions-only here → 1 sub-request per file, so the 20-sub-request $batch cap allows 20 files.
        const int ItemsPerBatch = 20;
        var chunks = items.Chunk(ItemsPerBatch).ToList();

        int done = 0, lastReported = 0;
        var reportLock = new object();

        await Parallel.ForEachAsync(chunks,
            new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, maxConcurrency), CancellationToken = ct },
            async (chunk, c) =>
            {
                var local = new Dictionary<string, (FileMetadata, List<DriveItemVersion>)>();
                try { await FetchBatchChunkAsync(chunk, local, c, versionsOnly: true); }
                catch (OperationCanceledException) { throw; }
                catch { /* items absent → counted as 1 by the caller's fallback */ }
                foreach (var kv in local)
                {
                    int v = kv.Value.Item2?.Count ?? 1;
                    counts[kv.Key] = maxVersionsCap > 0 ? Math.Min(v, maxVersionsCap) : v;
                }
                int d = Interlocked.Add(ref done, chunk.Length);
                lock (reportLock)
                {
                    if (d - lastReported >= 500 || d == items.Count) { lastReported = d; progress?.Report(d); }
                }
            });

        return new Dictionary<string, int>(counts);
    }

    // Bulk-fetches ONLY each item's LastModifiedDateTime — 1 sub-request/file (vs. 2 for the full
    // metadata+version fetch below), so twice as many files fit per $batch call. Used by If Newer
    // mode's pre-filter to decide skip-vs-copy for files that already exist at the target WITHOUT
    // paying for the full metadata+version fetch those files may never need if they turn out
    // skippable — the same waste the Skip-mode pre-filter avoids, just for a mode that still needs
    // *something* from Graph (the date) to make its decision, unlike Skip's pure by-name check.
    public async Task<Dictionary<string, DateTimeOffset?>> FetchModifiedDatesAsync(
        IReadOnlyList<(string driveId, string itemId)> items,
        int maxConcurrency,
        IProgress<int>? progress = null,
        CancellationToken ct = default)
    {
        var result = new System.Collections.Concurrent.ConcurrentDictionary<string, DateTimeOffset?>();
        if (items.Count == 0) return new Dictionary<string, DateTimeOffset?>();

        const int ItemsPerBatch = 20; // metadata-only = 1 sub-request/file → 20 files per $batch

        // Adaptive gate scoped to this call: shrinks concurrency on throttle and re-probes upward,
        // instead of every retry round re-bursting at the same fixed width straight back into the
        // still-depleted budget. VERIFIED (114k-file run, 2026-07-01): under sustained throttling the
        // old fixed-width 3-round retry left ~111k of 114k items unresolved, and the caller's
        // "undetermined → treat as needing a copy" fallback then misclassified nearly an entire
        // up-to-date library as needing re-copy — hours of wasted download/upload/import for files
        // that didn't need touching. The gate makes the retry rounds actually converge instead.
        using var gate = new AdaptiveParallelismController(maxConcurrency);
        void onThrottle(TimeSpan delay, int __, int ___, string? ____) => gate.StepDown(delay);
        Throttled += onThrottle;

        int lastReported = 0;
        var reportLock = new object();
        void ReportProgress()
        {
            int resolved = result.Count;
            lock (reportLock)
            {
                if (resolved - lastReported < 500 && resolved != items.Count) return;
                lastReported = resolved;
            }
            progress?.Report(resolved);
        }

        async Task RunPassAsync(IReadOnlyList<(string driveId, string itemId)> toFetch)
        {
            await Parallel.ForEachAsync(toFetch.Chunk(ItemsPerBatch),
                new ParallelOptions { MaxDegreeOfParallelism = maxConcurrency, CancellationToken = ct },
                async (chunk, c) =>
                {
                    await gate.WaitAsync(c);
                    try
                    {
                        var batch = new BatchRequestContentCollection(Graph);
                        var ids = new string?[chunk.Length];
                        for (int i = 0; i < chunk.Length; i++)
                        {
                            var (driveId, itemId) = chunk[i];
                            var req = new HttpRequestMessage(HttpMethod.Get,
                                $"{GraphBaseUrl}/drives/{driveId}/items/{itemId}?$select=lastModifiedDateTime");
                            ids[i] = batch.AddBatchRequestStep(req);
                        }

                        BatchResponseContentCollection response;
                        try { response = await Graph.Batch.PostAsync(batch, c); }
                        catch (OperationCanceledException) { throw; }
                        catch { return; } // whole batch call failed — retry rounds below cover these items

                        for (int i = 0; i < chunk.Length; i++)
                        {
                            if (ids[i] == null) continue;
                            try
                            {
                                using var http = await response.GetResponseByIdAsync(ids[i]!);
                                if (!http.IsSuccessStatusCode) continue;
                                using var doc = JsonDocument.Parse(await http.Content.ReadAsStringAsync(c));
                                result[chunk[i].itemId] = TryGetBatchDateTimeOffset(doc.RootElement, "lastModifiedDateTime");
                            }
                            catch { /* leave absent; retried below, then caller falls back if still missing */ }
                        }
                    }
                    finally { gate.Release(); }

                    ReportProgress();
                });
        }

        try
        {
            await RunPassAsync(items);

            // Retry rounds for anything still missing (transient throttle). The adaptive gate above
            // keeps concurrency below the throttle threshold so this converges rather than needing a
            // large fixed round count; capped at 8 as a backstop against a genuinely unresolvable item
            // (e.g. a source file deleted mid-run) looping forever. Progress is reported every round,
            // not just the first pass — a prior version went silent here, which under heavy throttling
            // looked identical to a hang for many minutes at a time.
            for (int round = 0; round < 8; round++)
            {
                var missing = items.Where(i => !result.ContainsKey(i.itemId)).ToList();
                if (missing.Count == 0) break;
                await RunPassAsync(missing);
            }
        }
        finally { Throttled -= onThrottle; }

        return new Dictionary<string, DateTimeOffset?>(result);
    }

    // Fetches metadata + versions for every item ONCE, in PARALLEL, into a reusable cache the download
    // producer consumes directly — so the producer makes zero Graph metadata calls during the copy
    // (those per-batch $batch calls were being silently throttled by Kiota's retry handler and stalling
    // the pipeline). Memory-light at scale: single-version files (the bulk of a typical library) keep
    // their metadata but DROP the version objects (Versions stored as null; the producer synthesizes a
    // current-version entry), so only multi-version files retain version lists. Items whose sub-request
    // failed are absent → the producer falls back to individual Graph calls for them. Misses are
    // RETRIED (a transient throttle can drop a sub-request); an incomplete cache would under-count a
    // file's versions, and that file's batch could then exceed the SPMI entry ceiling and fail import.
    public async Task<Dictionary<string, (FileMetadata Metadata, List<DriveItemVersion>? Versions)>>
        FetchMetadataAndVersionCacheAsync(
            IReadOnlyList<(string driveId, string itemId)> items,
            int maxConcurrency,
            IProgress<int>? progress = null,
            bool includeVersions = true,
            CancellationToken ct = default)
    {
        var cache = new System.Collections.Concurrent.ConcurrentDictionary<string, (FileMetadata, List<DriveItemVersion>?)>();
        if (items.Count == 0) return new Dictionary<string, (FileMetadata, List<DriveItemVersion>?)>();

        // includeVersions=false (copy runs with Versions: Off, maxVersions=1) skips the /versions
        // sub-request entirely — the version list would be sliced to the current version anyway, so
        // fetching it is pure waste. Metadata-only = 1 sub-request/file → 20 files per $batch call,
        // halving the Graph call count for the whole analysis phase.
        int itemsPerBatch = includeVersions ? 10 : 20;
        var skipAllVersions = includeVersions ? null : new HashSet<string>(); // empty set = "no item is multi-version"

        // Adaptive gate + progress-on-every-round: parity with FetchModifiedDatesAsync above. The prior
        // fixed-width retry rounds re-burst into a depleted throttle budget AND went silent after the
        // first pass — on the 114k-file run (2026-07-01) that meant 29 minutes of no log output while
        // retries churned, indistinguishable from a hang.
        using var gate = new AdaptiveParallelismController(maxConcurrency);
        void onThrottle(TimeSpan delay, int __, int ___, string? ____) => gate.StepDown(delay);
        Throttled += onThrottle;

        int lastReported = 0;
        var reportLock = new object();
        void ReportProgress()
        {
            int resolved = cache.Count;
            lock (reportLock)
            {
                if (resolved - lastReported < 500 && resolved != items.Count) return;
                lastReported = resolved;
            }
            progress?.Report(resolved);
        }

        async Task RunPassAsync(IReadOnlyList<(string driveId, string itemId)> toFetch)
        {
            await Parallel.ForEachAsync(toFetch.Chunk(itemsPerBatch),
                new ParallelOptions { MaxDegreeOfParallelism = maxConcurrency, CancellationToken = ct },
                async (chunk, c) =>
                {
                    var local = new Dictionary<string, (FileMetadata, List<DriveItemVersion>)>();
                    await gate.WaitAsync(c);
                    try { await FetchBatchChunkAsync(chunk, local, c, multiVersionItemIds: skipAllVersions); }
                    catch (OperationCanceledException) { throw; }
                    catch { /* items absent → retried below, then producer falls back if still missing */ }
                    finally { gate.Release(); }
                    foreach (var kv in local)
                    {
                        var versions = kv.Value.Item2;
                        // Single-version → keep metadata, drop the version object (producer synthesizes it).
                        cache[kv.Key] = versions.Count <= 1
                            ? (kv.Value.Item1, (List<DriveItemVersion>?)null)
                            : (kv.Value.Item1, versions);
                    }
                    ReportProgress();
                });
        }

        try
        {
            await RunPassAsync(items);

            // Retry rounds for any item the first pass missed (transient throttle). The adaptive gate
            // keeps concurrency below the throttle threshold so rounds converge; capped as a backstop
            // against genuinely unresolvable items (e.g. a file deleted mid-run).
            for (int round = 0; round < 8; round++)
            {
                var missing = items.Where(i => !cache.ContainsKey(i.itemId)).ToList();
                if (missing.Count == 0) break;
                await RunPassAsync(missing);
            }
        }
        finally { Throttled -= onThrottle; }

        return new Dictionary<string, (FileMetadata, List<DriveItemVersion>?)>(cache);
    }

    private async Task FetchBatchChunkAsync(
        (string driveId, string itemId)[] chunk,
        Dictionary<string, (FileMetadata, List<DriveItemVersion>)> result,
        CancellationToken ct,
        IReadOnlySet<string>? multiVersionItemIds = null,
        bool versionsOnly = false)
    {
        // BatchRequestContentCollection is the current recommended API (BatchRequestContent is obsolete).
        // Graph enforces a 20-sub-request cap; we send 1 or 2 sub-requests per file and cap chunks at 10
        // files, so a chunk is always ≤ 20 sub-requests.
        //   versionsOnly = true  → only the /versions call (the analyze/sizing pass needs counts, not
        //                          metadata it would discard); skips the metadata round-trip per file.
        //   multiVersionItemIds  → in the download pass, only fetch /versions for items known to have
        //                          >1 version; single-version files get metadata-only.
        var batch = new BatchRequestContentCollection(Graph);

        // AddBatchRequestStep(HttpRequestMessage) auto-assigns a request ID and returns it.
        // A null id means that sub-request was deliberately not issued for this item.
        var metaIds = new string?[chunk.Length];
        var versIds = new string?[chunk.Length];

        for (int i = 0; i < chunk.Length; i++)
        {
            var (driveId, itemId) = chunk[i];

            if (!versionsOnly)
            {
                var metaReq = new HttpRequestMessage(HttpMethod.Get,
                    $"{GraphBaseUrl}/drives/{driveId}/items/{itemId}" +
                    "?$select=createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,sharepointIds,size");
                metaIds[i] = batch.AddBatchRequestStep(metaReq);

                // Skip the versions round-trip for files an upstream pass already counted as single-version.
                if (multiVersionItemIds != null && !multiVersionItemIds.Contains(itemId))
                    continue;
            }

            var versReq = new HttpRequestMessage(HttpMethod.Get,
                $"{GraphBaseUrl}/drives/{driveId}/items/{itemId}/versions?$top=500");
            versIds[i] = batch.AddBatchRequestStep(versReq);
        }

        var response = await Graph.Batch.PostAsync(batch, ct);

        for (int i = 0; i < chunk.Length; i++)
        {
            var (driveId, itemId) = chunk[i];

            // Use nullable sentinels: only add to result when the needed sub-request(s) succeeded.
            // A failed sub-request (429, 404, etc.) leaves the item absent so the caller falls back
            // to individual Graph calls rather than using empty/stale data. In versionsOnly mode no
            // metadata is fetched, so a placeholder stands in (the caller uses only the version count).
            FileMetadata? metadata = versionsOnly ? new FileMetadata() : null;
            if (metaIds[i] != null)
            {
                try
                {
                    using var metaHttp = await response.GetResponseByIdAsync(metaIds[i]!);
                    if (metaHttp.IsSuccessStatusCode)
                    {
                        var metaJson = await metaHttp.Content.ReadAsStringAsync(ct);
                        metadata = ParseBatchFileMetadata(metaJson);
                        TryCacheBatchSpIds(driveId, itemId, metaJson);
                    }
                }
                catch { }
            }

            List<DriveItemVersion>? versions = null;
            if (versIds[i] == null)
            {
                // Versions deliberately skipped (single-version file). Synthesize one current-version
                // entry; the manifest builder fills its label/date/author from the file metadata, and
                // the downloader fetches the file's current content for this single entry.
                versions = new List<DriveItemVersion> { new DriveItemVersion() };
            }
            else
            {
                try
                {
                    using var versHttp = await response.GetResponseByIdAsync(versIds[i]!);
                    if (versHttp.IsSuccessStatusCode)
                    {
                        var (parsed, nextLink) = ParseBatchVersions(await versHttp.Content.ReadAsStringAsync(ct));
                        var vList = parsed;
                        while (nextLink != null)
                        {
                            ct.ThrowIfCancellationRequested();
                            var page = await Graph.Drives[driveId].Items[itemId].Versions
                                .WithUrl(nextLink).GetAsync(cancellationToken: ct);
                            vList.AddRange(page?.Value ?? []);
                            nextLink = page?.OdataNextLink;
                        }
                        versions = SortVersions(vList);
                    }
                }
                catch { }
            }

            if (metadata != null && versions != null)
                result[itemId] = (metadata, versions);
        }
    }

    private static FileMetadata ParseBatchFileMetadata(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        return new FileMetadata
        {
            CreatedDateTime  = TryGetBatchDateTimeOffset(root, "createdDateTime"),
            ModifiedDateTime = TryGetBatchDateTimeOffset(root, "lastModifiedDateTime"),
            CreatedByEmail   = ParseBatchIdentityEmail(root, "createdBy"),
            ModifiedByEmail  = ParseBatchIdentityEmail(root, "lastModifiedBy"),
            Size             = root.TryGetProperty("size", out var sz) && sz.ValueKind == JsonValueKind.Number
                ? sz.GetInt64() : null,
        };
    }

    private static DateTimeOffset? TryGetBatchDateTimeOffset(JsonElement root, string property)
    {
        if (!root.TryGetProperty(property, out var el) || el.ValueKind == JsonValueKind.Null)
            return null;
        return el.TryGetDateTimeOffset(out var dt) ? dt : null;
    }

    private static string? ParseBatchIdentityEmail(JsonElement root, string identitySetName)
    {
        if (!root.TryGetProperty(identitySetName, out var set)) return null;
        if (!set.TryGetProperty("user", out var user)) return null;
        if (user.TryGetProperty("email", out var email) && email.ValueKind == JsonValueKind.String)
            return email.GetString();
        if (user.TryGetProperty("userPrincipalName", out var upn) && upn.ValueKind == JsonValueKind.String)
            return upn.GetString();
        return null;
    }

    private void TryCacheBatchSpIds(string driveId, string itemId, string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("sharepointIds", out var sp)) return;
            if (!sp.TryGetProperty("siteUrl",    out var su)  || su.ValueKind  != JsonValueKind.String) return;
            if (!sp.TryGetProperty("listId",     out var li)  || li.ValueKind  != JsonValueKind.String) return;
            if (!sp.TryGetProperty("listItemId", out var lii) || lii.ValueKind != JsonValueKind.String) return;
            _spIdsCache[$"{driveId}|{itemId}"] = (su.GetString()!, li.GetString()!, lii.GetString()!);
        }
        catch { }
    }

    private static (List<DriveItemVersion> versions, string? nextLink) ParseBatchVersions(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root     = doc.RootElement;
        var versions = new List<DriveItemVersion>();

        if (root.TryGetProperty("value", out var arr))
        {
            foreach (var el in arr.EnumerateArray())
            {
                var v = new DriveItemVersion
                {
                    Id = el.TryGetProperty("id", out var id) ? id.GetString() : null,
                    LastModifiedDateTime = TryGetBatchDateTimeOffset(el, "lastModifiedDateTime"),
                };

                if (el.TryGetProperty("lastModifiedBy", out var lmb) &&
                    lmb.TryGetProperty("user", out var user))
                {
                    var extra = new Dictionary<string, object?>();
                    if (user.TryGetProperty("email", out var em) && em.ValueKind == JsonValueKind.String)
                        extra["email"] = em.GetString();
                    else if (user.TryGetProperty("userPrincipalName", out var upn) && upn.ValueKind == JsonValueKind.String)
                        extra["userPrincipalName"] = upn.GetString();
                    v.LastModifiedBy = new IdentitySet
                    {
                        User = new Microsoft.Graph.Models.Identity { AdditionalData = extra }
                    };
                }

                versions.Add(v);
            }
        }

        string? nextLink = root.TryGetProperty("@odata.nextLink", out var nl) && nl.ValueKind == JsonValueKind.String
            ? nl.GetString() : null;

        return (versions, nextLink);
    }

    // Sort oldest-first by numeric version label ("2.0" → 2.0) rather than timestamp:
    // versions saved within the same second would otherwise keep Graph's newest-first
    // order and be replayed out of sequence.
    private static List<DriveItemVersion> SortVersions(List<DriveItemVersion> versions) =>
        versions.OrderBy(v =>
        {
            var parts = (v.Id ?? "0").Split('.');
            int major = parts.Length > 0 && int.TryParse(parts[0], out var mj) ? mj : 0;
            int minor = parts.Length > 1 && int.TryParse(parts[1], out var mn) ? mn : 0;
            return (major, minor);
        }).ToList();

    // ── Metadata ──────────────────────────────────────────────────────────────

    public async Task<FileMetadata> GetFileMetadataAsync(string driveId, string itemId)
    {
        var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
            cfg.QueryParameters.Select = ["createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy", "size"]);

        return new FileMetadata
        {
            CreatedDateTime  = item?.CreatedDateTime,
            CreatedByEmail   = GetIdentityEmail(item?.CreatedBy?.User),
            ModifiedDateTime = item?.LastModifiedDateTime,
            ModifiedByEmail  = GetIdentityEmail(item?.LastModifiedBy?.User),
            Size             = item?.Size,
        };
    }

    // Fetches each SOURCE folder's created/modified metadata for the migration manifest (folders
    // otherwise get a hardcoded placeholder date, which surfaced as a wrong "1999" timestamp on the
    // target). `folders` maps a caller key (the folder's relative path) to a sample FILE item id inside
    // that folder — the file's parentReference IS the folder, so this needs no path-mapping assumptions.
    // The library root ("" key) is read from the drive root. Parallel + the Graph client's own throttle
    // retries; a folder that can't be read is simply absent (the builder keeps the placeholder for it).
    public async Task<Dictionary<string, FileMetadata>> FetchFolderMetadataAsync(
        string rootDriveId,
        IReadOnlyList<(string folderKey, string driveId, string sampleFileItemId)> folders,
        int maxConcurrency,
        IProgress<int>? progress = null,
        CancellationToken ct = default)
    {
        int completed = 0;
        var result = new System.Collections.Concurrent.ConcurrentDictionary<string, FileMetadata>();

        // Library root folder.
        try
        {
            var root = await Graph.Drives[rootDriveId].Root.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy"],
                cancellationToken: ct);
            if (root != null)
                result[string.Empty] = new FileMetadata
                {
                    CreatedDateTime  = root.CreatedDateTime,
                    CreatedByEmail   = GetIdentityEmail(root.CreatedBy?.User),
                    ModifiedDateTime = root.LastModifiedDateTime,
                    ModifiedByEmail  = GetIdentityEmail(root.LastModifiedBy?.User),
                };
        }
        catch { /* keep placeholder */ }

        await Parallel.ForEachAsync(folders,
            new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, maxConcurrency), CancellationToken = ct },
            async (f, c) =>
            {
                try
                {
                    var file = await Graph.Drives[f.driveId].Items[f.sampleFileItemId].GetAsync(cfg =>
                        cfg.QueryParameters.Select = ["parentReference"], cancellationToken: c);
                    var parentId = file?.ParentReference?.Id;
                    if (string.IsNullOrEmpty(parentId)) return;
                    result[f.folderKey] = await GetFileMetadataAsync(f.driveId, parentId);
                }
                catch { /* keep placeholder */ }
                finally { progress?.Report(Interlocked.Increment(ref completed)); }
            });

        return new Dictionary<string, FileMetadata>(result);
    }

    public static string? GetIdentityEmail(Microsoft.Graph.Models.Identity? identity)
    {
        if (identity?.AdditionalData == null) return null;
        if (identity.AdditionalData.TryGetValue("email", out var email) && email != null)
            return email.ToString();
        if (identity.AdditionalData.TryGetValue("userPrincipalName", out var upn) && upn != null)
            return upn.ToString();
        return null;
    }

    public async Task<(string siteUrl, string listId, string listItemId)?> GetSharePointIdsAsync(
        string driveId, string itemId)
    {
        var key = $"{driveId}|{itemId}";
        if (_spIdsCache.TryGetValue(key, out var cached))
        {
            System.Diagnostics.Debug.WriteLine($"[GetSPIds] cache hit: itemId={itemId}");
            return cached;
        }

        for (int attempt = 0; attempt < 3; attempt++)
        {
            if (attempt > 0) await Task.Delay(attempt * 1500);
            try
            {
                System.Diagnostics.Debug.WriteLine($"[GetSPIds] attempt {attempt + 1}/3 for itemId={itemId}");
                var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
                    cfg.QueryParameters.Select = ["sharepointIds"]);

                var ids = item?.SharepointIds;
                if (ids?.SiteUrl == null || ids.ListId == null || ids.ListItemId == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[GetSPIds] attempt {attempt + 1}: sharepointIds not yet populated for itemId={itemId}, will retry");
                    continue;
                }

                System.Diagnostics.Debug.WriteLine($"[GetSPIds] success: listItemId={ids.ListItemId} siteUrl={ids.SiteUrl}");
                var result = (ids.SiteUrl, ids.ListId, ids.ListItemId);
                _spIdsCache[key] = result;
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetSPIds] attempt {attempt + 1} threw {ex.GetType().Name}: {ex.Message}");
            }
        }
        System.Diagnostics.Debug.WriteLine($"[GetSPIds] all 3 attempts failed for itemId={itemId}");
        return null;
    }

    private async Task<string?> PatchTimestampsViaRestAsync(
        string driveId, string itemId, DateTimeOffset? modified, DateTimeOffset? created,
        string? createdByEmail = null, string? modifiedByEmail = null)
    {
        var ids = await GetSharePointIdsAsync(driveId, itemId);
        if (ids == null) return "SP IDs unavailable — item not found or sharepointIds not propagated";

        var (siteUrl, listId, listItemId) = ids.Value;

        var formValues = new List<object>();
        if (modified.HasValue)
            formValues.Add(new { FieldName = "Modified", FieldValue = modified.Value.ToUniversalTime().ToString("M/d/yyyy H:mm:ss", System.Globalization.CultureInfo.InvariantCulture) });
        if (created.HasValue)
            formValues.Add(new { FieldName = "Created",  FieldValue = created.Value.ToUniversalTime().ToString("M/d/yyyy H:mm:ss", System.Globalization.CultureInfo.InvariantCulture) });
        if (!string.IsNullOrEmpty(modifiedByEmail))
            formValues.Add(new { FieldName = "Editor", FieldValue = $"i:0#.f|membership|{modifiedByEmail}" });
        if (!string.IsNullOrEmpty(createdByEmail))
            formValues.Add(new { FieldName = "Author", FieldValue = $"i:0#.f|membership|{createdByEmail}" });

        if (formValues.Count == 0) return null;

        try
        {
            var token = await _authService.GetSharePointTokenAsync(siteUrl);
            var url   = $"{siteUrl}/_api/web/lists('{listId}')/items({listItemId})/ValidateUpdateListItem()";
            using var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new System.Net.Http.StringContent(
                    JsonSerializer.Serialize(new { formValues, bNewDocumentUpdate = true }),
                    System.Text.Encoding.UTF8, "application/json")
            };
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            req.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            var response = await _httpClient.SendAsync(req);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var preview = body.Length > 200 ? body[..200] : body;
                return $"ValidateUpdateListItem HTTP {(int)response.StatusCode}: {preview}";
            }

            try
            {
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("value", out var arr))
                {
                    var fieldErrors = new List<string>();
                    foreach (var field in arr.EnumerateArray())
                    {
                        bool hasEx = field.TryGetProperty("HasException", out var he) && he.GetBoolean();
                        int  ec    = field.TryGetProperty("ErrorCode",    out var ecEl) ? ecEl.GetInt32() : 0;
                        if (hasEx || ec != 0)
                        {
                            var fn = field.TryGetProperty("FieldName",    out var fnEl) ? fnEl.GetString() ?? "" : "";
                            var em = field.TryGetProperty("ErrorMessage", out var emEl) ? emEl.GetString() ?? "" : "";
                            fieldErrors.Add($"{fn}: {em} (code {ec})");
                        }
                    }
                    if (fieldErrors.Count > 0)
                        return $"Field errors: {string.Join("; ", fieldErrors)}";
                }
            }
            catch { }

            return null;
        }
        catch (Exception ex)
        {
            return $"ValidateUpdateListItem exception: {ex.Message}";
        }
    }

    public async Task<string?> ApplyFileMetadataAsync(
        string driveId, string itemId, string siteId, FileMetadata metadata)
    {
        return await PatchTimestampsViaRestAsync(driveId, itemId,
            metadata.ModifiedDateTime, metadata.CreatedDateTime,
            metadata.CreatedByEmail, metadata.ModifiedByEmail);
    }

    public async Task<string?> GetCurrentVersionIdAsync(string driveId, string itemId)
    {
        try
        {
            var page = await Graph.Drives[driveId].Items[itemId].Versions.GetAsync();
            return page?.Value?.FirstOrDefault()?.Id;
        }
        catch { return null; }
    }

    // Patches fileSystemInfo.lastModifiedDateTime — the field shown in SharePoint version history.
    // Always creates a phantom version in a versioned library.
    public async Task<string?> PatchFileSystemDateAsync(string driveId, string itemId,
        DateTimeOffset lastModifiedDateTime, DateTimeOffset? createdDateTime = null)
    {
        System.Diagnostics.Debug.WriteLine($"[PatchFSDate] itemId={itemId} modified={lastModifiedDateTime:u}");
        try
        {
            await Graph.Drives[driveId].Items[itemId].PatchAsync(new DriveItem
            {
                FileSystemInfo = new Microsoft.Graph.Models.FileSystemInfo
                {
                    LastModifiedDateTime = lastModifiedDateTime,
                    CreatedDateTime      = createdDateTime
                }
            });
            System.Diagnostics.Debug.WriteLine($"[PatchFSDate] success for itemId={itemId}");
            return null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PatchFSDate] FAILED for itemId={itemId}: {ex.GetType().Name}: {ex.Message}");
            return $"PatchFileSystemDate exception: {ex.Message}";
        }
    }

    public async Task<string?> DeleteItemVersionAsync(string driveId, string itemId, string versionId)
    {
        try
        {
            await Graph.Drives[driveId].Items[itemId].Versions[versionId].DeleteAsync();
            return null;
        }
        catch (Exception ex)
        {
            return $"DeleteItemVersion exception: {ex.Message}";
        }
    }

    // ── Migration API helpers ─────────────────────────────────────────────────

    // Calls /_api/site/ProvisionMigrationContainers on the target site.
    // Returns (dataContainerUri, metadataContainerUri, encryptionKey).
    public async Task<(string dataUri, string metadataUri, byte[] encryptionKey)>
        ProvisionMigrationContainersAsync(string siteUrl)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/site/ProvisionMigrationContainers";
        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"ProvisionMigrationContainers HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;
        var dataUri     = root.GetProperty("DataContainerUri").GetString()!;
        var metadataUri = root.GetProperty("MetadataContainerUri").GetString()!;
        var keyBase64   = root.GetProperty("EncryptionKey").GetString()!;
        return (dataUri, metadataUri, Convert.FromBase64String(keyBase64));
    }

    // Returns the signed-in user's display name, email, and IsSiteAdmin flag for siteUrl.
    public async Task<(string title, string email, bool isSiteAdmin)> GetCurrentUserInfoAsync(string siteUrl)
    {
        var token = await _authService.GetSharePointTokenAsync(siteUrl, "AllSites.FullControl");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/currentuser?$select=Title,Email,LoginName,IsSiteAdmin";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.ParseAdd("application/json;odata=nometadata");

        var response = await _httpClient.SendAsync(req);
        if (!response.IsSuccessStatusCode)
            return ("(unknown)", "(HTTP error)", false);

        var body = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;
        var title   = root.TryGetProperty("Title",      out var t) ? t.GetString() ?? "" : "";
        var email   = root.TryGetProperty("Email",      out var e) ? e.GetString() ?? "" : "";
        var login   = root.TryGetProperty("LoginName",  out var l) ? l.GetString() ?? "" : "";
        var isAdmin = root.TryGetProperty("IsSiteAdmin", out var a) && a.GetBoolean();
        return (title, string.IsNullOrEmpty(email) ? login : email, isAdmin);
    }

    // Returns (webId, serverRelativeUrl) for the root web of the given site URL.
    public async Task<(string webId, string serverRelativeUrl)> GetWebInfoAsync(string siteUrl)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web?$select=Id,ServerRelativeUrl";
        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"GetWebInfo HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        var webId  = doc.RootElement.GetProperty("Id").GetString()!;
        var relUrl = doc.RootElement.GetProperty("ServerRelativeUrl").GetString()!;
        return (webId, relUrl);
    }

    // Returns the GUID of a document library given its server-relative URL.
    public async Task<string> GetListIdByServerRelativeUrlAsync(string siteUrl, string serverRelativeUrl)
    {
        var encodedUrl = Uri.EscapeDataString($"'{serverRelativeUrl}'");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetList({encodedUrl})?$select=Id";
        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"GetListId HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        return doc.RootElement.GetProperty("Id").GetString()!;
    }

    // Gets the SharePoint list GUID for a document library via the Graph drive's associated list.
    // More reliable than GetListIdByServerRelativeUrlAsync when the server-relative URL is uncertain.
    public async Task<string?> GetListIdFromDriveAsync(string driveId)
    {
        try
        {
            var list = await Graph.Drives[driveId].List
                .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
            return list?.Id;
        }
        catch
        {
            return null;
        }
    }

    // Gets the server-relative URL of a document library by fetching the drive's root item webUrl.
    // More reliable than deriving it from drive.WebUrl which the Kiota SDK may not populate.
    public async Task<string> GetLibraryServerRelativeUrlAsync(string driveId)
    {
        var root = await Graph.Drives[driveId].Root
            .GetAsync(cfg => cfg.QueryParameters.Select = ["webUrl"]);
        if (root?.WebUrl == null)
            throw new InvalidOperationException($"Cannot determine library path for drive {driveId}.");
        return Uri.UnescapeDataString(new Uri(root.WebUrl).AbsolutePath.TrimEnd('/'));
    }

    // Submits a migration job using SP-provided encrypted containers.
    // Returns the job ID GUID.
    //
    // Uses raw CSOM ProcessQuery XML rather than the CSOM client library because
    // the library's ExecutingWebRequest token injection is unreliable on .NET 8
    // (compiled for .NET 4.0; HttpWebRequest behaviour changed under the compat layer).
    public async Task<string> CreateMigrationJobEncryptedAsync(
        string siteUrl, string gWebId,
        string dataContainerUri, string metadataContainerUri,
        byte[] encryptionKey)
    {
        // AllSites.FullControl is required: Sites.ReadWrite.All caps effective privilege below
        // site-collection-admin level in SP's OAuth permission model, causing CreateMigrationJobEncrypted
        // to reject even explicit SCAs.  The Azure AD app must have AllSites.FullControl delegated
        // permission registered and admin-consented; the user will be prompted if not yet granted.
        var keyB64 = Convert.ToBase64String(encryptionKey);

        // Build the ProcessQuery XML.  EncryptionOption TypeId is {85614ad4-7a40-49e0-b272-6d1807dbfcc6}.
        // AES256CBCKey is a byte[] serialised as Base64Binary.
        // SP reads the per-blob IV from the first 16 bytes of each encrypted blob,
        // so no IV is needed here — only the key.
        var ns = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/sharepoint/clientquery/2009");
        var requestXml = new System.Xml.Linq.XDocument(
            new System.Xml.Linq.XElement(ns + "Request",
                new System.Xml.Linq.XAttribute("SchemaVersion",   "15.0.0.0"),
                new System.Xml.Linq.XAttribute("LibraryVersion",  "16.0.0.0"),
                new System.Xml.Linq.XAttribute("ApplicationName", "SharePointSmartCopy"),
                new System.Xml.Linq.XElement(ns + "Actions",
                    new System.Xml.Linq.XElement(ns + "Method",
                        new System.Xml.Linq.XAttribute("Name",         "CreateMigrationJobEncrypted"),
                        new System.Xml.Linq.XAttribute("Id",           "1"),
                        new System.Xml.Linq.XAttribute("ObjectPathId", "2"),
                        new System.Xml.Linq.XElement(ns + "Parameters",
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("Type", "Guid"), gWebId),
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("Type", "String"), dataContainerUri),
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("Type", "String"), metadataContainerUri),
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("Type", "String"), ""),
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("TypeId", "{85614ad4-7a40-49e0-b272-6d1807dbfcc6}"),
                                new System.Xml.Linq.XElement(ns + "Property",
                                    new System.Xml.Linq.XAttribute("Name", "AES256CBCKey"),
                                    new System.Xml.Linq.XAttribute("Type", "Base64Binary"),
                                    keyB64))))),
                new System.Xml.Linq.XElement(ns + "ObjectPaths",
                    new System.Xml.Linq.XElement(ns + "Property",
                        new System.Xml.Linq.XAttribute("Id",       "2"),
                        new System.Xml.Linq.XAttribute("ParentId", "0"),
                        new System.Xml.Linq.XAttribute("Name",     "Site")),
                    new System.Xml.Linq.XElement(ns + "StaticProperty",
                        new System.Xml.Linq.XAttribute("Id",     "0"),
                        new System.Xml.Linq.XAttribute("TypeId", "{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}"),
                        new System.Xml.Linq.XAttribute("Name",   "Current")))));

        var xmlBody = requestXml.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        var url     = $"{siteUrl.TrimEnd('/')}/_vti_bin/client.svc/ProcessQuery";

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(xmlBody, System.Text.Encoding.UTF8, "text/xml")
            };
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl, "AllSites.FullControl");
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"CreateMigrationJobEncrypted HTTP {(int)response.StatusCode}: {body}");

        // ProcessQuery returns a JSON array.  Index 0 is the header with ErrorInfo.
        // On success the job GUID appears as a bare string element later in the array.
        using var doc = JsonDocument.Parse(body);
        var arr = doc.RootElement;

        if (arr.GetArrayLength() > 0 &&
            arr[0].TryGetProperty("ErrorInfo", out var errInfo) &&
            errInfo.ValueKind != JsonValueKind.Null)
        {
            var msg = errInfo.TryGetProperty("ErrorMessage", out var m) ? m.GetString() : "Unknown CSOM error";
            var code = errInfo.TryGetProperty("ErrorCode", out var c) ? c.GetInt32() : 0;
            if (code == -2147024891) // E_ACCESSDENIED
            {
                var (uTitle, uEmail, uIsAdmin) = await GetCurrentUserInfoAsync(siteUrl);
                throw new UnauthorizedAccessException(
                    $"The Migration API requires explicit Site Collection Administrator membership on {siteUrl}.\n\n" +
                    $"SP sees you as: {uTitle} ({uEmail}), IsSiteAdmin={uIsAdmin}\n\n" +
                    "Note: Global Admin / SharePoint Admin does not automatically populate this list.\n" +
                    "Fix: go to the target site → Site Settings → Site Collection Administrators → add your account.",
                    new Exception(msg));
            }
            throw new Exception($"CreateMigrationJobEncrypted CSOM error {code}: {msg}");
        }

        for (int i = 1; i < arr.GetArrayLength(); i++)
        {
            var el = arr[i];
            if (el.ValueKind != JsonValueKind.String) continue;
            var s = el.GetString() ?? "";
            // CSOM returns Guids as "\/Guid(xxxxxxxx-...)\/"; also handle bare GUIDs.
            if (s.StartsWith("/Guid(", StringComparison.OrdinalIgnoreCase) && s.EndsWith(")/"))
                s = s[6..^2];
            if (Guid.TryParse(s, out var g))
                return g.ToString("D");
        }

        throw new Exception($"Could not find job ID GUID in ProcessQuery response: {body[..Math.Min(body.Length, 500)]}");
    }

    // Polls the migration job using the paging-based GetMigrationJobProgress endpoint.
    // Yields each event JSON string. Caller should stop when a JobEnd event is received.
    public async IAsyncEnumerable<JsonElement> PollMigrationJobAsync(
        string siteUrl, string jobId, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var token = await _authService.GetSharePointTokenAsync(siteUrl);
        int nextToken  = 0;
        int pollCount  = 0;
        var pollDelay  = TimeSpan.FromSeconds(2);

        while (!cancellationToken.IsCancellationRequested)
        {
            // Poll immediately on the first pass (the previous fixed 3 s pre-delay added latency to
            // every batch); then back off for long-running jobs to reduce API noise:
            // 2 s for the first ~1 min → 10 s for the next ~6 min → 30 s thereafter.
            if (pollCount > 0)
                await Task.Delay(pollDelay, cancellationToken);
            pollCount++;
            if (pollCount == 40)  pollDelay = TimeSpan.FromSeconds(10);
            if (pollCount == 100) pollDelay = TimeSpan.FromSeconds(30);

            // Guid parameters in SP REST require the guid'...' syntax, not just '...'.
            // AllSites.FullControl is needed — the same reason CreateMigrationJobEncrypted needs it.
            var url = $"{siteUrl.TrimEnd('/')}/_api/site/GetMigrationJobProgress(jobId=guid'{jobId}',nextToken={nextToken})";
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer",
                await _authService.GetSharePointTokenAsync(siteUrl, "AllSites.FullControl", cancellationToken));
            req.Headers.Accept.ParseAdd("application/json;odata=nometadata");

            HttpResponseMessage response;
            try
            {
                response = await _httpClient.SendAsync(req, cancellationToken);
            }
            catch (TaskCanceledException) { yield break; }

            bool hitJobEnd = false;
            using (response)
            {
            bool isThrottle = response.StatusCode is System.Net.HttpStatusCode.TooManyRequests
                                                  or System.Net.HttpStatusCode.ServiceUnavailable;
            bool isTransient = !isThrottle && (int)response.StatusCode >= 500;
            if (isThrottle || isTransient)
            {
                // Throttle: honour Retry-After; 5xx: SP sometimes returns 500 when internally
                // throttled — treat as retriable rather than fatal.
                var delay = isThrottle
                    ? (response.Headers.RetryAfter?.Delta
                       ?? (response.Headers.RetryAfter?.Date is { } when
                               ? when - DateTimeOffset.UtcNow
                               : TimeSpan.FromSeconds(30)))
                    : TimeSpan.FromSeconds(30);
                if (delay < TimeSpan.Zero) delay = TimeSpan.FromSeconds(1);
                if (delay > TimeSpan.FromSeconds(120)) delay = TimeSpan.FromSeconds(120);
                await Task.Delay(delay, cancellationToken);
                continue;
            }
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync(cancellationToken);
                throw new Exception($"GetMigrationJobProgress HTTP {(int)response.StatusCode}: {err[..Math.Min(err.Length, 300)]}");
            }

            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            System.Diagnostics.Debug.WriteLine($"[Poll] raw response ({body.Length} bytes): {body[..Math.Min(body.Length, 3000)]}");
            JsonDocument doc;
            try { doc = JsonDocument.Parse(body); }
            catch { continue; }
            using (doc)
            {
                // SP GetMigrationJobProgress returns { "Logs": [...], "nextToken": N }
                // where each element of Logs is a JSON-encoded string (not an object).
                // Some older docs describe "value" or a bare array — handle all variants.
                JsonElement arr = default;
                if (doc.RootElement.ValueKind == JsonValueKind.Array)
                    arr = doc.RootElement;
                else if (doc.RootElement.TryGetProperty("Logs", out var logs))
                    arr = logs;
                else if (doc.RootElement.TryGetProperty("value", out var v))
                    arr = v;

                if (arr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var rawEvt in arr.EnumerateArray())
                    {
                        JsonElement evtObj;
                        if (rawEvt.ValueKind == JsonValueKind.String)
                        {
                            try
                            {
                                using var inner = JsonDocument.Parse(rawEvt.GetString()!);
                                evtObj = inner.RootElement.Clone();
                            }
                            catch { continue; }
                        }
                        else if (rawEvt.ValueKind == JsonValueKind.Object)
                            evtObj = rawEvt.Clone();
                        else
                            continue;

                        if (!evtObj.TryGetProperty("Event", out var evtName)) continue;
                        yield return evtObj;
                        if (evtName.GetString() == "JobEnd")
                        {
                            hitJobEnd = true;
                            break;
                        }
                    }
                }

                // nextToken may be a string or number, and may be null
                if (doc.RootElement.TryGetProperty("nextToken", out var nt))
                {
                    if (nt.ValueKind == JsonValueKind.Number) nextToken = nt.GetInt32();
                    else if (nt.ValueKind == JsonValueKind.String && int.TryParse(nt.GetString(), out var p)) nextToken = p;
                }
                else if (doc.RootElement.TryGetProperty("NextToken", out var nt2))
                {
                    if (nt2.ValueKind == JsonValueKind.Number) nextToken = nt2.GetInt32();
                    else if (nt2.ValueKind == JsonValueKind.String && int.TryParse(nt2.GetString(), out var p2)) nextToken = p2;
                }
            }

            } // end using (response)

            if (hitJobEnd) yield break;
        }
    }

    // ── Upload ────────────────────────────────────────────────────────────────

    // Fetches all file children of a folder in one paginated pass.
    // Returns filename → (itemId, lastModifiedDateTime), case-insensitive.
    // Used by the migration pre-flight to avoid one Graph call per file.
    public async Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>> FetchFolderItemsAsync(
        string driveId, string folderId)
    {
        // Retries the WHOLE paginated walk on a transient transport failure (see
        // IsTransientTransportError) rather than leaving a mid-pagination blip to propagate out of the
        // whole drive-group scan and fail every pending file. Re-walking from page 1 is cheap relative
        // to the cost of failing the entire run.
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                var result = new Dictionary<string, (string, DateTimeOffset?)>(StringComparer.OrdinalIgnoreCase);
                var page = await Graph.Drives[driveId].Items[folderId].Children
                    .GetAsync(cfg =>
                    {
                        cfg.QueryParameters.Top    = 1000;
                        cfg.QueryParameters.Select = ["id", "name", "lastModifiedDateTime", "file"];
                    });

                while (page != null)
                {
                    foreach (var item in page.Value ?? [])
                    {
                        if (item.File == null || item.Name == null || item.Id == null) continue;
                        result[item.Name] = (item.Id, item.LastModifiedDateTime);
                    }
                    if (page.OdataNextLink == null) break;
                    page = await Graph.Drives[driveId].Items[folderId].Children
                        .WithUrl(page.OdataNextLink).GetAsync();
                }
                return result;
            }
            catch (Exception ex) when (attempt < 4 && IsTransientTransportError(ex))
            {
                await Task.Delay(TimeSpan.FromSeconds((attempt + 1) * 3));
            }
        }
    }

    // Lists a folder's files via SP REST (GetFolderByServerRelativeUrl/Files) — the same AllDocs
    // store the Migration API validates imports against. Graph's Children listing does NOT return
    // rows left behind by a previous SPMI job that fatal-aborted mid-import, so an overwrite
    // pre-flight that trusts Graph alone reports "0 already exist" while SPMI then rejects those
    // same files with "already exists" (observed 2026-07-02: 100+ per batch). Returns name →
    // TimeLastModified; a missing folder (404) returns empty. Follows odata.nextLink paging.
    public async Task<Dictionary<string, DateTimeOffset?>> FetchFolderFileNamesRestAsync(
        string siteUrl, string folderServerRelativeUrl)
    {
        var result  = new Dictionary<string, DateTimeOffset?>(StringComparer.OrdinalIgnoreCase);
        var baseUrl = siteUrl.TrimEnd('/');
        // *Path (not *Url) API: folder names with #, %, or + (e.g. "A+C 091522") are unresolvable
        // via GetFolderByServerRelativeUrl, so its /Files listing 404s and the pre-flight wrongly
        // reports the folder empty. See ServerRelativePathArg.
        string? nextUrl = $"{baseUrl}/_api/web/GetFolderByServerRelativePath({ServerRelativePathArg(folderServerRelativeUrl)})/Files" +
                          "?$select=Name,TimeLastModified&$top=5000";

        for (int attempt = 0; nextUrl != null; )
        {
            try
            {
                var requestUrl = nextUrl;
                using var response = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl);

                if (!response.IsSuccessStatusCode)
                {
                    // 404 = folder doesn't exist yet (fresh target) — genuinely empty. Anything else is
                    // logged but still returns what we have: the caller merges this with the Graph
                    // listing, so a degraded REST view falls back to Graph-only behavior, not failure.
                    System.Diagnostics.Debug.WriteLine(
                        $"[RestFolderList] HTTP {(int)response.StatusCode} for {folderServerRelativeUrl}");
                    return result;
                }

                var body = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("value", out var items))
                {
                    foreach (var item in items.EnumerateArray())
                    {
                        if (!item.TryGetProperty("Name", out var n) || n.ValueKind != JsonValueKind.String) continue;
                        DateTimeOffset? modified = null;
                        if (item.TryGetProperty("TimeLastModified", out var tm) &&
                            tm.ValueKind == JsonValueKind.String &&
                            DateTimeOffset.TryParse(tm.GetString(), out var dt))
                            modified = dt;
                        result[n.GetString()!] = modified;
                    }
                }
                nextUrl = doc.RootElement.TryGetProperty("odata.nextLink", out var nl) &&
                          nl.ValueKind == JsonValueKind.String
                    ? nl.GetString()
                    : null;
                attempt = 0; // reset the retry budget per page
            }
            catch (Exception ex) when (attempt < 4 && IsTransientTransportError(ex))
            {
                attempt++;
                await Task.Delay(TimeSpan.FromSeconds(attempt * 3));
            }
        }
        return result;
    }

    // Cheap "is this folder completely empty?" check (files OR subfolders), via a single $top=1
    // request. Used to detect a fresh migration target so the per-folder existing-file scan can be
    // skipped entirely — on an empty target nothing can already exist.
    public async Task<bool> IsFolderEmptyAsync(string driveId, string folderId)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                var page = await Graph.Drives[driveId].Items[folderId].Children
                    .GetAsync(cfg =>
                    {
                        cfg.QueryParameters.Top    = 1;
                        cfg.QueryParameters.Select = ["id"];
                    });
                return (page?.Value?.Count ?? 0) == 0;
            }
            catch (Exception ex) when (attempt < 4 && IsTransientTransportError(ex))
            {
                await Task.Delay(TimeSpan.FromSeconds((attempt + 1) * 3));
            }
        }
    }

    public async Task<bool> FileExistsAsync(string driveId, string parentItemId, string fileName)
    {
        try
        {
            await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName)).GetAsync();
            System.Diagnostics.Debug.WriteLine($"[FileExists] EXISTS: {fileName}");
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[FileExists] NOT FOUND (expected 404): {fileName} — {ex.GetType().Name}: {ex.Message}");
            return false;
        }
    }

    public async Task DeleteFileIfExistsAsync(string driveId, string parentItemId, string fileName)
    {
        try
        {
            await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName)).DeleteAsync();
        }
        catch { }
    }

    // Returns the file's Graph item ID and last-modified timestamp, or null if it does
    // not exist. One call serves the existence check, the Copy-if-newer comparison, and
    // (via the ID) permission refresh on files that end up skipped as up to date.
    public async Task<(string ItemId, DateTimeOffset? Modified)?> GetFileInfoAsync(
        string driveId, string parentItemId, string fileName)
    {
        try
        {
            var item = await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName))
                .GetAsync(cfg => cfg.QueryParameters.Select = ["id", "lastModifiedDateTime"]);
            return item?.Id is { } id ? (id, item.LastModifiedDateTime) : null;
        }
        catch
        {
            return null;
        }
    }

    // Builds the OData "...ByServerRelativePath(decodedurl='…')" argument for a REST call.
    // The *Path (not *Url) API family is REQUIRED for any path containing #, %, or + — the
    // GetFileByServerRelativeUrl/GetFolderByServerRelativeUrl variants cannot resolve those even
    // when percent-encoded, so listings 404 (→ pre-flight misses the file) and deletes silently
    // fail (→ the file survives into the SPMI import as an "already exists" conflict). Observed
    // 2026-07-02 on folders like "A+C 091522" (+) and files like "…SOW#3-MUC16.docx" (#).
    // decodedurl takes the LITERAL path: OData-escape embedded quotes (' → ''), then percent-encode
    // the whole thing for transport; SharePoint decodes it back to the literal server-relative path.
    private static string ServerRelativePathArg(string serverRelativeUrl) =>
        $"decodedurl='{Uri.EscapeDataString(serverRelativeUrl.Replace("'", "''"))}'";

    // Returns the SharePoint UniqueId (AllDocs GUID) for a file by its server-relative URL.
    // Works via REST (not Graph) so it finds zombie files — SPFile blobs without a list item
    // that Graph returns 404 for. Returns null if the file doesn't exist.
    public async Task<string?> GetFileUniqueIdAsync(string siteUrl, string serverRelativeUrl)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetFileByServerRelativePath({ServerRelativePathArg(serverRelativeUrl)})/UniqueId";
        try
        {
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!response.IsSuccessStatusCode) return null;
            var body = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.ValueKind == JsonValueKind.String
                ? doc.RootElement.GetString()
                : doc.RootElement.TryGetProperty("value", out var v) ? v.GetString() : null;
        }
        catch { return null; }
    }

    // Permanently deletes a file by recycling it then immediately purging the recycle bin entry.
    // Graph DeleteAsync only soft-deletes (recycle bin); soft-deleted list item records interfere
    // with SPMI imports at the same URL, producing "Missing file info for list item" errors.
    //
    // Retries transient failures at each step and returns false (rather than silently swallowing
    // the error) if the file can't be confirmed gone. Previously a single dropped call here — under
    // sustained throttling on a large batch — left the old item in place while the caller proceeded
    // as if it had been deleted, so SPMI rejected the re-import with "already exists" even though
    // overwrite was requested (observed on a 5,000-file overwrite run, 2026-07-02). Callers must
    // check the result and fail the item cleanly rather than queue a doomed import.
    // onFail, when provided, receives a short human-readable reason (HTTP status + trimmed SharePoint
    // error body) whenever this returns false — so the caller can surface WHY a delete failed into the
    // activity log instead of only "still present after delete." SharePoint's REST error body usually
    // names the real cause (locked/checked-out/access-denied/not-found).
    // onFail receives a short reason when this returns false. onTrace (when provided) ALWAYS receives
    // a one-line trace of what each REST step actually returned — recycle status, bin id, purge status —
    // even on success. That trace is the only way to see these steps in a Release build (the old
    // Debug.WriteLine calls are compiled out unless a debugger is attached), and it's what finally
    // showed the delete "succeeding" while the file survived.
    public async Task<bool> PermanentlyDeleteFileAsync(
        string siteUrl, string serverRelativeUrl, Action<string>? onFail = null, Action<string>? onTrace = null)
    {
        const int maxAttempts = 4;
        var baseUrl = siteUrl.TrimEnd('/');
        string lastReason = "unknown";
        var trace = new System.Text.StringBuilder();
        void Trace(string s) { trace.Append(s).Append(' '); }
        bool Done(bool ok, string? reason = null)
        {
            onTrace?.Invoke(trace.ToString().TrimEnd());
            if (!ok) onFail?.Invoke(reason ?? lastReason);
            return ok;
        }

        var uniqueId = await GetFileUniqueIdAsync(siteUrl, serverRelativeUrl);
        if (string.IsNullOrEmpty(uniqueId)) { Trace("resolve=null(gone)"); return Done(true); }
        Trace($"id={uniqueId}");

        // Step 1: recycle by ID. GetFileById('{guid}') keeps the path (and its '#'/'+'/'%' hazards) out
        // of the URL entirely. A recycle 404 is NOT treated as "already gone" here — the caller's verify
        // loop is the real gone-check; a false 404 would otherwise mask a delete that never happened.
        var recycleUrl = $"{baseUrl}/_api/web/GetFileById('{uniqueId}')/recycleObject";
        string? recycleBinGuid = null;
        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
            bool lastAttempt = attempt == maxAttempts - 1;
            try
            {
                using var recycleResponse = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Post, recycleUrl);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                    return r;
                }, siteUrl);

                var body = await recycleResponse.Content.ReadAsStringAsync();
                Trace($"recycle={(int)recycleResponse.StatusCode}");

                if (!recycleResponse.IsSuccessStatusCode)
                {
                    lastReason = $"recycle HTTP {(int)recycleResponse.StatusCode}: {Trim(body)}";
                    // recycleObject fails for a file with no list item (zombie AllDocs row) — fall back to
                    // deleteObject by ID, which removes the row at the AllDocs level directly.
                    if (lastAttempt)
                    {
                        var (delOk, delReason) = await TryDeleteObjectByIdAsync(siteUrl, uniqueId, Trace);
                        return Done(delOk, delReason ?? lastReason);
                    }
                    await Task.Delay(TimeSpan.FromSeconds(3 * (attempt + 1)));
                    continue;
                }

                using var doc = JsonDocument.Parse(body);
                recycleBinGuid = doc.RootElement.ValueKind == JsonValueKind.String
                    ? doc.RootElement.GetString()
                    : doc.RootElement.TryGetProperty("value", out var v) ? v.GetString() : null;
                break;
            }
            catch (Exception ex)
            {
                lastReason = $"recycle exception: {ex.Message}";
                if (lastAttempt) return Done(false);
                await Task.Delay(TimeSpan.FromSeconds(3 * (attempt + 1)));
            }
        }

        Trace($"bin={recycleBinGuid ?? "null"}");
        if (string.IsNullOrEmpty(recycleBinGuid)) return Done(false, $"recycle returned no bin id ({lastReason})");

        // Step 2: purge. recycleObject drops the item in the WEB (first-stage) recycle bin, so purge from
        // _api/web/RecycleBin; the previous code purged _api/site/RecycleBin (site collection / second
        // stage), where a web-bin id isn't found — that 404 was then wrongly treated as success, so the
        // file was only recycled, never purged (and the item lingered). Try web first, site as fallback.
        // NOTE: a purge failure still leaves the file recycled (out of the folder), which is enough for
        // SPMI to re-import — so we return true once recycle succeeded even if the purge is imperfect.
        foreach (var scope in new[] { "web", "site" })
        {
            var purgeUrl = $"{baseUrl}/_api/{scope}/RecycleBin('{recycleBinGuid}')/DeleteObject";
            try
            {
                using var purgeResponse = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Post, purgeUrl);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                    return r;
                }, siteUrl);
                Trace($"purge[{scope}]={(int)purgeResponse.StatusCode}");
                if (purgeResponse.IsSuccessStatusCode) return Done(true);
            }
            catch (Exception ex)
            {
                Trace($"purge[{scope}]=ex");
                lastReason = $"purge exception: {ex.Message}";
            }
        }

        // Recycle succeeded (file left the folder) but neither bin purge confirmed. Still good enough for
        // the re-import; report success but leave the trace so a lingering-file case is diagnosable.
        return Done(true);
    }

    // Trims a SharePoint REST error body to a short single line for logging.
    private static string Trim(string body) =>
        System.Text.RegularExpressions.Regex.Replace(body ?? "", @"\s+", " ").Trim() is { Length: > 0 } s
            ? s[..Math.Min(s.Length, 200)] : "(empty body)";

    private async Task<(bool ok, string? reason)> TryDeleteObjectByIdAsync(
        string siteUrl, string uniqueId, Action<string>? trace = null)
    {
        const int maxAttempts = 4;
        var deleteUrl = $"{siteUrl.TrimEnd('/')}/_api/web/GetFileById('{uniqueId}')/deleteObject";
        string? lastReason = null;
        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
            try
            {
                using var response = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Post, deleteUrl);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                    return r;
                }, siteUrl);
                var body = await response.Content.ReadAsStringAsync();
                trace?.Invoke($"deleteObject={(int)response.StatusCode}");
                if (response.IsSuccessStatusCode || response.StatusCode == System.Net.HttpStatusCode.NotFound)
                    return (true, null);
                lastReason = $"deleteObject HTTP {(int)response.StatusCode}: {Trim(body)}";
            }
            catch (Exception ex)
            {
                lastReason = $"deleteObject exception: {ex.Message}";
            }
            if (attempt < maxAttempts - 1)
                await Task.Delay(TimeSpan.FromSeconds(3 * (attempt + 1)));
        }
        return (false, lastReason);
    }

    public async Task<string> UploadFileAsync(
        string targetDriveId,
        string targetParentItemId,
        string fileName,
        Stream content,
        bool overwrite,
        IProgress<int>? progress = null)
    {
        content.Position = 0;
        long size = content.Length;
        var conflictBehavior = overwrite ? "replace" : "fail";

        if (size < 4 * 1024 * 1024)
        {
            var item = await Graph.Drives[targetDriveId].Items[targetParentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName)).Content.PutAsync(content);
            progress?.Report(100);
            return item?.Id ?? string.Empty;
        }

        var sessionBody = new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody
        {
            Item = new DriveItemUploadableProperties
            {
                AdditionalData = new Dictionary<string, object>
                {
                    { "@microsoft.graph.conflictBehavior", conflictBehavior }
                }
            }
        };

        var session = await Graph.Drives[targetDriveId].Items[targetParentItemId]
            .ItemWithPath(Uri.EscapeDataString(fileName)).CreateUploadSession.PostAsync(sessionBody);

        var uploadTask = new Microsoft.Graph.LargeFileUploadTask<DriveItem>(
            session!, content, 320 * 1024, Graph.RequestAdapter);

        var uploadResult = await uploadTask.UploadAsync(new Progress<long>(uploaded =>
        {
            if (size > 0) progress?.Report((int)(uploaded * 100 / size));
        }));

        if (!uploadResult.UploadSucceeded)
            throw new Exception("Large file upload failed.");

        progress?.Report(100);
        return uploadResult.ItemResponse?.Id ?? string.Empty;
    }

    // Creates a modern Site Page via the SitePages REST API and returns the Graph item ID
    // and the SitePages integer ID. The SitePages ID is needed for SavePage / Publish calls.
    // AddTemplateFile(templateFileType=0) creates a classic page without CanvasContent1 support.
    public async Task<(string graphItemId, int sitePagesId)> CreatePageStubAsync(
        string siteUrl, string targetFolderRelUrl,
        string driveId, string parentItemId,
        string fileName, bool overwrite)
    {
        System.Diagnostics.Debug.WriteLine($"[CreatePageStub] {fileName} → {siteUrl}");

        // Delete existing page first — SitePages API returns 409 if the file already exists.
        if (overwrite)
        {
            try
            {
                var existing = await Graph.Drives[driveId].Items[parentItemId]
                    .ItemWithPath(Uri.EscapeDataString(fileName))
                    .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
                if (existing?.Id != null)
                {
                    await Graph.Drives[driveId].Items[existing.Id].DeleteAsync();
                    System.Diagnostics.Debug.WriteLine($"[CreatePageStub] deleted existing: {existing.Id}");
                }
            }
            catch { /* file doesn't exist — nothing to delete */ }
        }

        var title = System.IO.Path.GetFileNameWithoutExtension(fileName)
                        .Replace('-', ' ').Replace('_', ' ');

        var body = System.Text.Json.JsonSerializer.Serialize(new
        {
            __metadata     = new { type = "SP.Publishing.SitePage" },
            FileName       = fileName,
            Title          = title,
            PageLayoutType = "Article",
        });

        using var createResp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post,
                $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages");
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=verbose");
            r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
            r.Content.Headers.ContentType =
                System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            return r;
        }, siteUrl);

        if (!createResp.IsSuccessStatusCode)
        {
            var err = await createResp.Content.ReadAsStringAsync();
            throw new Exception(
                $"Page create HTTP {(int)createResp.StatusCode}: {err[..Math.Min(400, err.Length)]}");
        }

        // Parse the SitePages integer ID from the response — needed for SavePage and Publish.
        var respBody = await createResp.Content.ReadAsStringAsync();
        int sitePagesId = 0;
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(respBody);
            var root = doc.RootElement.TryGetProperty("d", out var d) ? d : doc.RootElement;
            if (root.TryGetProperty("Id", out var idProp))
                sitePagesId = idProp.GetInt32();
        }
        catch { }
        System.Diagnostics.Debug.WriteLine($"[CreatePageStub] sitePagesId={sitePagesId}");

        // Resolve Graph item ID with retry (Graph can lag 1–2 s after SitePages creation).
        for (int attempt = 0; attempt < 3; attempt++)
        {
            if (attempt > 0) await Task.Delay(1000);
            try
            {
                var item = await Graph.Drives[driveId].Items[parentItemId]
                    .ItemWithPath(Uri.EscapeDataString(fileName))
                    .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
                if (item?.Id != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[CreatePageStub] graphItemId={item.Id}");
                    return (item.Id, sitePagesId);
                }
            }
            catch { }
        }
        System.Diagnostics.Debug.WriteLine($"[CreatePageStub] could not resolve Graph item ID after 3 attempts");
        return (string.Empty, sitePagesId);
    }

    // Writes page content to a draft modern page via the SitePages SavePage API.
    // ValidateUpdateListItem writes the list item fields but the modern page renderer reads
    // content from the SitePages checkout state — SavePage is the correct write path.
    public async Task<string?> SavePageContentAsync(
        string siteUrl, int sitePagesId, PageMetadata pageMeta, string sourceSiteUrl)
    {
        if (sitePagesId == 0) return "SitePages ID unknown — cannot save content";

        string? SubstUrl(string? json) =>
            json == null ? null : SubstituteUrls(json, sourceSiteUrl, siteUrl);

        // Build a save body with only the fields that have values
        var fields = new Dictionary<string, object>
        {
            ["__metadata"] = (object)new { type = "SP.Publishing.SitePage" },
        };
        var canvas  = SubstUrl(pageMeta.CanvasContent1);
        var layout  = SubstUrl(pageMeta.LayoutWebpartsContent);
        var banner  = SubstUrl(pageMeta.BannerImageUrl);
        if (canvas  != null) fields["CanvasContent1"]        = canvas;
        if (layout  != null) fields["LayoutWebpartsContent"] = layout;
        if (banner  != null) fields["BannerImageUrl"]        = banner;
        if (pageMeta.Description    != null) fields["Description"]    = pageMeta.Description;
        if (pageMeta.PageLayoutType != null) fields["PageLayoutType"]  = pageMeta.PageLayoutType;
        fields["PromotedState"] = (object)pageMeta.PromotedState;

        var saveBody = System.Text.Json.JsonSerializer.Serialize(fields);
        var saveUrl  = $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages({sitePagesId})/SavePage";

        System.Diagnostics.Debug.WriteLine($"[SavePageContent] POST {saveUrl}");
        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, saveUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=verbose");
            r.Content = new StringContent(saveBody, System.Text.Encoding.UTF8, "application/json");
            r.Content.Headers.ContentType =
                System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            return r;
        }, siteUrl);

        System.Diagnostics.Debug.WriteLine($"[SavePageContent] status={resp.StatusCode}");
        if (!resp.IsSuccessStatusCode)
        {
            var err = await resp.Content.ReadAsStringAsync();
            return $"SavePage HTTP {(int)resp.StatusCode}: {err[..Math.Min(300, err.Length)]}";
        }
        return null;
    }

    // Publishes a draft modern page via the SitePages Publish endpoint.
    // Must be called after SavePageContentAsync so the published version contains the content.
    public async Task<string?> PublishPageAsync(string siteUrl, int sitePagesId)
    {
        if (sitePagesId == 0) return "SitePages ID unknown — cannot publish";
        var url = $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages({sitePagesId})/Publish";
        try
        {
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
            System.Diagnostics.Debug.WriteLine($"[PublishPage] status={response.StatusCode}");
            return response.IsSuccessStatusCode ? null : $"Publish HTTP {(int)response.StatusCode}";
        }
        catch (Exception ex)
        {
            return $"Publish: {ex.Message}";
        }
    }

    // ── Folder creation ───────────────────────────────────────────────────────

    public async Task<string> GetOrCreateFolderPathAsync(string driveId, string parentItemId, string relativePath)
    {
        var current = parentItemId;
        foreach (var part in relativePath.Split('/').Where(p => !string.IsNullOrEmpty(p)))
        {
            // Capture loop variables for the lambda — the Lazy factory runs on first .Value access.
            var cacheKey = $"{driveId}|{current}|{part}";
            var d = driveId; var p = current; var n = part;
            // GetOrAdd may construct multiple Lazy instances under contention but only stores one.
            // Because Lazy.Value is called after GetOrAdd returns, the discarded instances never
            // invoke their factories — so only one Graph call is made per unique folder segment.
            var lazy = _folderSegmentTasks.GetOrAdd(cacheKey,
                _ => new Lazy<Task<string>>(
                    () => GetOrCreateFolderAsync(d, p, n),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication));
            current = await lazy.Value;
        }
        return current;
    }

    // True for exceptions that represent a dropped/failed connection rather than a real Graph response
    // (including a proper throttle response, which Kiota's own retry handler already resolves before
    // it would ever reach this code). VERIFIED (114k-file run, 2026-07-01): under sustained throttling
    // a single "An error occurred while sending the request" transport blip during pre-flight folder
    // provisioning/scanning went uncaught here, propagated out of the whole drive-group loop, and hit
    // MigrationJobService's top-level catch-all — which marked EVERY file still pending (i.e. nearly
    // the entire 114k-file job) as Failed. The download/upload paths elsewhere in this codebase already
    // retry on exactly these exception types; pre-flight Graph calls need the same treatment so one
    // blip doesn't take down an entire run.
    private static bool IsTransientTransportError(Exception ex) =>
        ex is System.Net.Http.HttpRequestException || ex is IOException;

    private async Task<string> GetOrCreateFolderAsync(string driveId, string parentItemId, string folderName)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                return await GetOrCreateFolderCoreAsync(driveId, parentItemId, folderName);
            }
            catch (Exception ex) when (attempt < 4 && IsTransientTransportError(ex))
            {
                await Task.Delay(TimeSpan.FromSeconds((attempt + 1) * 3));
            }
        }
    }

    private async Task<string> GetOrCreateFolderCoreAsync(string driveId, string parentItemId, string folderName)
    {
        try
        {
            var existing = await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(folderName)).GetAsync();
            System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] EXISTS: {folderName}");
            return existing?.Id ?? throw new Exception("Null ID from existing folder.");
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] not found, creating: {folderName}");
            try
            {
                var created = await Graph.Drives[driveId].Items[parentItemId].Children.PostAsync(new DriveItem
                {
                    Name   = folderName,
                    Folder = new Folder()
                });
                System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] created: {folderName}");
                return created?.Id ?? throw new Exception("Null ID from created folder.");
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError conflict) when (conflict.ResponseStatusCode == 409)
            {
                System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] 409 conflict (race), re-fetching: {folderName}");
                var existing = await Graph.Drives[driveId].Items[parentItemId]
                    .ItemWithPath(Uri.EscapeDataString(folderName)).GetAsync();
                return existing?.Id ?? throw new Exception("Null ID after concurrent folder creation.");
            }
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    // Raised when a request is throttled and a retry is scheduled — lets the UI show
    // "Throttled — retrying in Ns…" instead of appearing stalled.
    public event Action<TimeSpan, int, int, string?>? Throttled;

    // Sends a SharePoint REST request with resilience:
    //  - 401: retried once with a force-refreshed token.
    //  - 429/503 (throttling): retried up to 5 times, honoring the Retry-After header,
    //    with capped exponential backoff + jitter when the header is absent.
    // buildRequest receives the bearer token and must return a fresh HttpRequestMessage each call.
    internal async Task<HttpResponseMessage> SendSharePointRequestAsync(
        Func<string, HttpRequestMessage> buildRequest,
        string siteUrl,
        string spScope = "Sites.ReadWrite.All",
        CancellationToken cancellationToken = default)
    {
        const int maxThrottleRetries = 8;
        var token = await _authService.GetSharePointTokenAsync(siteUrl, spScope, cancellationToken);
        HttpResponseMessage response;
        bool refreshedToken = false;

        for (int attempt = 0; ; attempt++)
        {
            using var req = buildRequest(token);
            response = await _httpClient.SendAsync(req, cancellationToken);

            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized && !refreshedToken)
            {
                response.Dispose();
                token = await _authService.GetSharePointTokenAsync(siteUrl, spScope, cancellationToken, forceRefresh: true);
                refreshedToken = true;
                continue;
            }

            bool throttled = response.StatusCode is System.Net.HttpStatusCode.TooManyRequests
                                                 or System.Net.HttpStatusCode.ServiceUnavailable;
            if (throttled && attempt < maxThrottleRetries)
            {
                // Retry-After can arrive as a delta or an absolute date.
                var delay = response.Headers.RetryAfter?.Delta
                    ?? (response.Headers.RetryAfter?.Date is { } when
                            ? when - DateTimeOffset.UtcNow
                            : TimeSpan.FromSeconds(Math.Pow(2, attempt + 1))
                              + TimeSpan.FromMilliseconds(Random.Shared.Next(0, 1000))); // jitter
                if (delay < TimeSpan.Zero) delay = TimeSpan.FromSeconds(1);
                if (delay > TimeSpan.FromSeconds(120)) delay = TimeSpan.FromSeconds(120);

                System.Diagnostics.Debug.WriteLine(
                    $"[SP-REST] {(int)response.StatusCode} throttled — retrying in {delay.TotalSeconds:N0}s (attempt {attempt + 1}/{maxThrottleRetries})");
                Throttled?.Invoke(delay, attempt + 1, maxThrottleRetries, null);
                response.Dispose();
                await Task.Delay(delay, cancellationToken);
                continue;
            }

            return response;
        }
    }

    private static SharePointNode Placeholder() =>
        new() { Name = "__placeholder__" };

    // SP REST pagination link: "odata.nextLink" under odata=nometadata (classic), but
    // "@odata.nextLink" under OData v4 / minimalmetadata responses. Check both so
    // pagination never silently truncates at the first page.
    private static string? GetNextLink(JsonElement root) =>
        root.TryGetProperty("odata.nextLink",  out var nl1) ? nl1.GetString() :
        root.TryGetProperty("@odata.nextLink", out var nl2) ? nl2.GetString() : null;

    // ── Column / Custom Field Methods ─────────────────────────────────────────

    // Built-in field internal names that should never be treated as custom columns.
    private static readonly HashSet<string> _builtInFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "ID", "Title", "Author", "Editor", "Created", "Modified", "ContentType",
        "FileLeafRef", "FileDirRef", "FileRef", "_UIVersionString", "MetaInfo",
        "_ModerationStatus", "_ModerationComments", "Edit", "SelectTitle", "InstanceID",
        "Order", "GUID", "WorkflowVersion", "_HasCopyDestinations", "_CopySource",
        "owshiddenversion", "WorkflowInstanceID", "ParentLeafName", "UniqueId",
        "ProgId", "ScopeId", "HTML_x0020_File_x0020_Type", "_EditMenuTableStart",
        "_EditMenuTableEnd", "LinkFilenameNoMenu", "LinkFilename", "DocConcurrencyToken",
        "SelectFilename", "ItemChildCount", "FolderChildCount", "Restricted",
        "OriginatorId", "NoExecute", "ContentVersion", "_dlc_DocId", "_dlc_DocIdUrl",
        "_dlc_DocIdPersistId", "TemplateUrl", "xd_ProgID", "xd_Signature",
        "AppAuthor", "AppEditor", "_IsCurrentVersion", "SMTotalSize",
        "SMLastModifiedDate", "_Level", "_IsCurrentVersion", "CheckedOutUserId",
        "IsCheckedoutToLocal", "CheckoutUser", "SMTotalFileStreamSize",
        "ComplianceAssetId", "_ComplianceFlags", "_ComplianceTag",
        "_ComplianceTagWrittenTime", "_ComplianceTagUserId",
        "AccessPolicy", "BSN", "MicroBlogging",
    };

    private static readonly Dictionary<int, SupportedFieldType> _fieldTypeMap = new()
    {
        { 2,  SupportedFieldType.Text },
        { 3,  SupportedFieldType.Note },
        { 4,  SupportedFieldType.DateTime },
        { 6,  SupportedFieldType.Choice },
        { 8,  SupportedFieldType.Boolean },
        { 9,  SupportedFieldType.Number },
        { 15, SupportedFieldType.MultiChoice },
        { 20, SupportedFieldType.User },
    };

    // Types that FieldTypeKind alone cannot distinguish (taxonomy fields report kind 0,
    // multi-user fields report kind 20 like single-user).
    private static SupportedFieldType? ResolveFieldType(int typeKind, string? typeAsString) => typeAsString switch
    {
        "UserMulti"              => SupportedFieldType.UserMulti,
        "TaxonomyFieldType"      => SupportedFieldType.Taxonomy,
        "TaxonomyFieldTypeMulti" => SupportedFieldType.TaxonomyMulti,
        "Lookup"                 => SupportedFieldType.Lookup,
        "LookupMulti"            => SupportedFieldType.LookupMulti,
        _ => _fieldTypeMap.TryGetValue(typeKind, out var t) ? t : null,
    };

    // Keyed by listId. Concurrent — read/written from parallel copy tasks.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, List<SpColumnDef>> _columnCache = new();

    // Keyed by "{listId}|{showField}|{displayValue}" — caches resolved target lookup item IDs.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int?> _lookupValueCache = new();

    // Returns the custom (non-built-in) columns for a library.
    public async Task<List<SpColumnDef>> GetLibraryColumnsAsync(string siteUrl, string listId, bool skipCache = false)
    {
        if (!skipCache && _columnCache.TryGetValue(listId, out var cached))
            return cached;

        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/fields" +
                  "?$filter=Hidden eq false and ReadOnlyField eq false and FromBaseType eq false" +
                  "&$select=InternalName,Title,FieldTypeKind,TypeAsString,Choices,SchemaXml";

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl);

        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"GetLibraryColumns HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        var result = new List<SpColumnDef>();

        if (doc.RootElement.TryGetProperty("value", out var values))
        {
            foreach (var field in values.EnumerateArray())
            {
                var internalName = field.GetProperty("InternalName").GetString() ?? "";
                if (_builtInFields.Contains(internalName)) continue;
                var typeKind = field.GetProperty("FieldTypeKind").GetInt32();
                var typeAsString = field.TryGetProperty("TypeAsString", out var tas) ? tas.GetString() : null;
                if (ResolveFieldType(typeKind, typeAsString) is not { } fieldType) continue;

                string[]? choices = null;
                if (field.TryGetProperty("Choices", out var choicesProp) &&
                    choicesProp.ValueKind == JsonValueKind.Array)
                {
                    choices = choicesProp.EnumerateArray()
                        .Select(c => c.GetString() ?? "")
                        .Where(s => s.Length > 0)
                        .ToArray();
                }
                // Also handle OData verbose choice format: { "results": [...] }
                if (choices == null &&
                    field.TryGetProperty("Choices", out var choicesVerbose) &&
                    choicesVerbose.ValueKind == JsonValueKind.Object &&
                    choicesVerbose.TryGetProperty("results", out var results))
                {
                    choices = results.EnumerateArray()
                        .Select(c => c.GetString() ?? "")
                        .Where(s => s.Length > 0)
                        .ToArray();
                }

                var schemaXml = field.TryGetProperty("SchemaXml", out var sx) ? sx.GetString() : null;

                string? lookupListId    = null;
                string? lookupShowField = null;
                if (SpColumnDef.IsLookupType(fieldType) && schemaXml != null)
                {
                    var listMatch = System.Text.RegularExpressions.Regex.Match(schemaXml, @"List=""(\{?[0-9A-Fa-f\-]+\}?)""");
                    if (listMatch.Success)
                        lookupListId = listMatch.Groups[1].Value.Trim('{', '}');
                    var showMatch = System.Text.RegularExpressions.Regex.Match(schemaXml, @"ShowField=""([^""]+)""");
                    if (showMatch.Success)
                        lookupShowField = showMatch.Groups[1].Value;
                }

                result.Add(new SpColumnDef
                {
                    InternalName    = internalName,
                    DisplayName     = field.GetProperty("Title").GetString() ?? internalName,
                    FieldType       = fieldType,
                    Choices         = choices,
                    SchemaXml       = schemaXml,
                    LookupListId    = lookupListId,
                    LookupShowField = lookupShowField,
                });
            }
        }

        _columnCache[listId] = result;
        return result;
    }

    // Bulk-reads custom field values for all items in a library in one paginated pass.
    // Returns a dictionary keyed by SP list item integer ID (as string, e.g. "42").
    // This key matches the listItemId returned by GetSharePointIdsAsync, enabling O(1) cache lookup per file.
    public async Task<Dictionary<string, Dictionary<string, object?>>> BulkReadCustomFieldsAsync(
        string siteUrl, string listId, IEnumerable<SpColumnDef> columns,
        IProgress<int>? progress = null, CancellationToken ct = default)
    {
        var cols = columns.ToList();
        if (cols.Count == 0) return [];

        // Chunk columns into groups of 20 to stay within URL length limits.
        var result = new Dictionary<string, Dictionary<string, object?>>();
        var chunks  = cols.Chunk(20).ToList();

        for (int chunk = 0; chunk < chunks.Count; chunk++)
        {
            // User fields are lookups: they must be $expand-ed and read via {name}/Name
            // (the claims login). Other field types are selected directly.
            var selectParts = new List<string> { "ID" };
            var expandParts = new List<string>();
            foreach (var c in chunks[chunk])
            {
                if (SpColumnDef.IsUserType(c.FieldType))
                {
                    selectParts.Add($"{c.InternalName}/Name");
                    expandParts.Add(c.InternalName);
                }
                else if (SpColumnDef.IsLookupType(c.FieldType))
                {
                    selectParts.Add($"{c.InternalName}/LookupId");
                    selectParts.Add($"{c.InternalName}/LookupValue");
                    expandParts.Add(c.InternalName);
                }
                else
                {
                    selectParts.Add(c.InternalName);
                }
            }
            var nextUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                          $"?$select={Uri.EscapeDataString(string.Join(",", selectParts))}" +
                          (expandParts.Count > 0 ? $"&$expand={Uri.EscapeDataString(string.Join(",", expandParts))}" : "") +
                          "&$top=1000";

            while (nextUrl != null)
            {
                ct.ThrowIfCancellationRequested();
                using var response = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, nextUrl);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl, cancellationToken: ct);

                var body = await response.Content.ReadAsStringAsync(ct);
                if (!response.IsSuccessStatusCode) break;

                using var doc = JsonDocument.Parse(body);
                var root = doc.RootElement;

                if (root.TryGetProperty("value", out var items))
                {
                    foreach (var item in items.EnumerateArray())
                    {
                        if (!item.TryGetProperty("ID", out var idProp)) continue;
                        var itemId = idProp.ValueKind == JsonValueKind.Number
                            ? idProp.GetInt32().ToString()
                            : idProp.GetString() ?? "";
                        if (string.IsNullOrEmpty(itemId)) continue;

                        if (!result.TryGetValue(itemId, out var fields))
                        {
                            fields = new Dictionary<string, object?>();
                            result[itemId] = fields;
                        }

                        foreach (var col in chunks[chunk])
                        {
                            if (!item.TryGetProperty(col.InternalName, out var valProp)) continue;
                            fields[col.InternalName] = ParseFieldValue(valProp, col.FieldType);
                        }
                    }
                }

                nextUrl = GetNextLink(root);
            }

            progress?.Report((chunk + 1) * 100 / chunks.Count);
        }

        return result;
    }

    private static object? ParseFieldValue(JsonElement el, SupportedFieldType type)
    {
        if (el.ValueKind == JsonValueKind.Null || el.ValueKind == JsonValueKind.Undefined)
            return null;

        if (SpColumnDef.IsUserType(type))
        {
            var logins = EnumerateComplexValues(el)
                .Select(u => u.TryGetProperty("Name", out var n) ? n.GetString() : null)
                .Where(s => !string.IsNullOrEmpty(s))
                .Select(s => s!)
                .ToArray();
            return logins.Length > 0 ? new PersonFieldValue(logins) : null;
        }

        if (SpColumnDef.IsTaxonomyType(type))
        {
            var terms = EnumerateComplexValues(el)
                .Select(t => (
                    Label:    t.TryGetProperty("Label",    out var l) ? l.GetString() ?? "" : "",
                    TermGuid: t.TryGetProperty("TermGuid", out var g) ? g.GetString() ?? "" : ""))
                .Where(t => t.TermGuid.Length > 0)
                .ToArray();
            return terms.Length > 0 ? new TaxonomyFieldValue(terms) : null;
        }

        if (SpColumnDef.IsLookupType(type))
        {
            var entries = EnumerateComplexValues(el)
                .Select(e => (
                    Id:           e.TryGetProperty("LookupId",    out var id) && id.ValueKind == JsonValueKind.Number ? id.GetInt32() : 0,
                    DisplayValue: e.TryGetProperty("LookupValue", out var dv) ? dv.GetString() ?? "" : ""))
                .Where(e => e.Id > 0)
                .ToArray();
            return entries.Length > 0 ? new LookupFieldValue(entries) : null;
        }

        return type switch
        {
            SupportedFieldType.MultiChoice when el.ValueKind == JsonValueKind.Array =>
                string.Join(";#", el.EnumerateArray().Select(v => v.GetString() ?? "")),
            SupportedFieldType.MultiChoice when el.ValueKind == JsonValueKind.Object &&
                el.TryGetProperty("results", out var r) =>
                string.Join(";#", r.EnumerateArray().Select(v => v.GetString() ?? "")),
            SupportedFieldType.Boolean =>
                el.ValueKind == JsonValueKind.True,
            SupportedFieldType.DateTime when el.ValueKind == JsonValueKind.String =>
                el.GetString(),
            _ => el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString()
        };
    }

    // Complex field values (user, taxonomy) arrive as a single object, a bare array,
    // or a verbose { "results": [...] } wrapper depending on OData mode and multiplicity.
    private static IEnumerable<JsonElement> EnumerateComplexValues(JsonElement el)
    {
        if (el.ValueKind == JsonValueKind.Array)
            return el.EnumerateArray();
        if (el.ValueKind == JsonValueKind.Object && el.TryGetProperty("results", out var results) &&
            results.ValueKind == JsonValueKind.Array)
            return results.EnumerateArray();
        if (el.ValueKind == JsonValueKind.Object)
            return [el];
        return [];
    }

    // Finds the ID of an item in a lookup list by its display value.
    // ShowField values "LinkTitle"/"LinkTitleNoMenu" are queried as "Title".
    // Returns null when the display value cannot be matched; result is cached per run.
    private async Task<int?> ResolveLookupValueAsync(
        string siteUrl, string lookupListId, string showField, string displayValue,
        CancellationToken ct = default)
    {
        var queryField = showField is "LinkTitle" or "LinkTitleNoMenu" ? "Title" : showField;
        var cacheKey   = $"{lookupListId}|{queryField}|{displayValue}";
        if (_lookupValueCache.TryGetValue(cacheKey, out var cached)) return cached;

        var escaped = Uri.EscapeDataString(displayValue.Replace("'", "''"));
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{lookupListId}')/items" +
                  $"?$select=ID&$filter={Uri.EscapeDataString(queryField)} eq '{escaped}'&$top=1";

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl, cancellationToken: ct);

        int? result = null;
        if (response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var items) &&
                items.GetArrayLength() > 0 &&
                items[0].TryGetProperty("ID", out var id))
            {
                result = id.GetInt32();
            }
        }

        _lookupValueCache[cacheKey] = result;
        return result;
    }

    // Applies custom field values to a list item via ValidateUpdateListItem.
    // mappings translates source InternalName → target InternalName.
    public async Task<string?> ApplyFileCustomFieldsAsync(
        string driveId, string itemId,
        Dictionary<string, object?> fields,
        IEnumerable<SpColumnMap> mappings,
        CancellationToken ct = default)
    {
        if (fields.Count == 0) return null;

        var mappingLookup = SpColumnMap.BuildTargetNameMap(mappings);

        // Resolve target SharePoint IDs first so we can look up target column definitions
        // (needed to find the target lookup list GUID for Lookup/LookupMulti fields).
        var ids = await GetSharePointIdsAsync(driveId, itemId)
            ?? throw new Exception($"Could not resolve SharePoint IDs for {driveId}/{itemId}");

        // Fetch target column defs (cached) to resolve lookup list GUIDs.
        List<SpColumnDef> targetCols;
        try { targetCols = await GetLibraryColumnsAsync(ids.siteUrl, ids.listId); }
        catch { targetCols = []; }
        var targetColsByName = targetCols.ToDictionary(c => c.InternalName, StringComparer.OrdinalIgnoreCase);

        var formValues      = new List<object>();
        var submittedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var lookupErrors    = new List<string>();
        foreach (var (srcName, value) in fields)
        {
            if (value == null) continue;
            // Resolve target name via mapping, or use source name directly
            string targetName;
            if (mappingLookup.TryGetValue(srcName, out var mapped))
            {
                if (mapped == null) continue; // explicitly skipped
                targetName = mapped;
            }
            else
            {
                targetName = srcName;
            }

            // Lookup fields: resolve each display value to a target item ID.
            if (value is LookupFieldValue lookup)
            {
                if (!targetColsByName.TryGetValue(targetName, out var tgtCol) ||
                    string.IsNullOrEmpty(tgtCol.LookupListId))
                {
                    lookupErrors.Add(targetName);
                    continue;
                }

                var showField  = string.IsNullOrEmpty(tgtCol.LookupShowField) ? "Title" : tgtCol.LookupShowField;
                var resolvedIds = new List<int>();
                foreach (var (_, display) in lookup.Entries)
                {
                    var targetId = await ResolveLookupValueAsync(ids.siteUrl, tgtCol.LookupListId, showField, display, ct);
                    if (targetId.HasValue) resolvedIds.Add(targetId.Value);
                }

                if (resolvedIds.Count == 0) continue; // nothing resolved — skip field
                // Single: "3", Multi: "3;#;#5;#" (SP lookup wire format with ID only)
                var formatted = resolvedIds.Count == 1
                    ? resolvedIds[0].ToString()
                    : string.Join(";#;#", resolvedIds.Select(id => id.ToString())) + ";#";
                formValues.Add(new { FieldName = targetName, FieldValue = formatted });
                submittedFields.Add(targetName);
                continue;
            }

            var formattedValue = FormatFieldValueForValidate(value);
            formValues.Add(new { FieldName = targetName, FieldValue = formattedValue });
            submittedFields.Add(targetName);
        }

        if (formValues.Count == 0) return lookupErrors.Count > 0 ? $"Lookup unresolved: {string.Join(", ", lookupErrors)}" : null;

        var url = $"{ids.siteUrl.TrimEnd('/')}/_api/web/lists('{ids.listId}')/items({ids.listItemId})/ValidateUpdateListItem()";
        var payload = JsonSerializer.Serialize(new
        {
            formValues,
            bNewDocumentUpdate = true,
        });

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, ids.siteUrl);

        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            return $"ApplyCustomFields HTTP {(int)response.StatusCode}";

        // ValidateUpdateListItem returns an entry for every field in the list definition,
        // including read-only system fields (Created, Modified, etc.) that always report
        // HasException=true.  Only surface errors for fields we actually submitted.
        var fieldErrors = new List<string>(lookupErrors.Select(n => $"{n} (lookup unresolved)"));
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var vals))
            {
                foreach (var v in vals.EnumerateArray())
                {
                    if (!v.TryGetProperty("HasException", out var ex) || !ex.GetBoolean()) continue;
                    if (!v.TryGetProperty("FieldName", out var fn)) continue;
                    var name = fn.GetString() ?? "";
                    if (submittedFields.Contains(name))
                        fieldErrors.Add(name);
                }
            }
        }
        catch { /* ignore parse errors */ }

        return fieldErrors.Count > 0 ? $"Custom field errors: {string.Join(", ", fieldErrors)}" : null;
    }

    private static string FormatFieldValueForValidate(object value)
    {
        // ValidateUpdateListItem person format: JSON array of claims keys.
        if (value is PersonFieldValue person)
            return JsonSerializer.Serialize(person.Logins.Select(l => new { Key = l }));

        // ValidateUpdateListItem taxonomy format: "Label|guid" (";"-joined for multi).
        // SharePoint resolves WssId and the hidden note field server-side.
        if (value is TaxonomyFieldValue taxonomy)
            return string.Join(";", taxonomy.Terms.Select(t => $"{t.Label}|{t.TermGuid}"));

        if (value is bool b)
            return b ? "1" : "0";

        var str = value.ToString() ?? "";
        // MultiChoice: ensure ";#val1;#val2;#" format
        if (str.Contains(";#") && !str.StartsWith(";#"))
            str = ";#" + str + ";#";
        return str;
    }

    // ── System libraries ──────────────────────────────────────────────────────

    // Returns a SharePointNode for a document library identified by its display title,
    // or null if the library does not exist on the site.
    // Useful for hidden/system libraries (Site Assets, Style Library) that the Graph
    // Drives API does not always enumerate.
    public async Task<SharePointNode?> GetLibraryNodeByTitleAsync(string siteId, string siteUrl, string title)
    {
        try
        {
            var encoded  = Uri.EscapeDataString(title.Replace("'", "''"));
            var endpoint = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{encoded}')?$select=Id";

            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, endpoint);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) return null;

            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("Id", out var idProp)) return null;
            var listGuid = idProp.GetString();
            if (string.IsNullOrEmpty(listGuid)) return null;

            var drive = await Graph.Sites[siteId].Lists[listGuid].Drive.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "webUrl"]);

            var driveId = drive?.Id;
            if (string.IsNullOrEmpty(driveId)) return null;

            string? serverRelUrl = null;
            if (drive?.WebUrl != null)
            {
                var uri  = new Uri(drive.WebUrl);
                var path = Uri.UnescapeDataString(uri.AbsolutePath.TrimEnd('/'));
                if (!path.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
                    serverRelUrl = path;
            }

            var node = new SharePointNode
            {
                Id                 = "root",
                DriveId            = driveId,
                SiteId             = siteId,
                SiteUrl            = siteUrl,
                Name               = title,
                Type               = NodeType.Library,
                HasChildren        = true,
                ServerRelativePath = serverRelUrl,
            };
            node.Children.Add(Placeholder());
            return node;
        }
        catch
        {
            return null;
        }
    }

    // ── Pages ─────────────────────────────────────────────────────────────────

    // Returns the Site Pages library node, or null if not found.
    // Site Pages uses BaseTemplate=119 and is NOT returned by the Drives API on most tenants.
    // Strategy: fast-path via Drives, then reliable fallback via SharePoint REST + Graph list drive.
    public async Task<SharePointNode?> GetSitePagesLibraryAsync(string siteId, string siteUrl)
    {
        // Fast path: some tenants expose Site Pages as a Graph drive
        var libraries = await GetLibrariesAsync(siteId, siteUrl);
        var byDrive = libraries.FirstOrDefault(l =>
            l.Name.Equals("Site Pages", StringComparison.OrdinalIgnoreCase));
        if (byDrive != null) return byDrive;

        // Reliable fallback: use SharePoint REST to get the list GUID, then Graph to get its drive.
        // This is more reliable than expanding 'drive' on the Lists collection response,
        // which may not populate in some Graph SDK versions.
        try
        {
            // Step 1: get the Site Pages list GUID via REST
            var listGuid = await GetSitePagesListIdViaRestAsync(siteUrl);
            if (string.IsNullOrEmpty(listGuid)) return null;

            // Step 2: get the drive for this specific list via Graph
            var drive = await Graph.Sites[siteId].Lists[listGuid].Drive.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "webUrl"]);

            var driveId = drive?.Id;
            if (string.IsNullOrEmpty(driveId)) return null;

            // Step 3: get the root item so we have a Graph item ID for tree expansion
            var root = await Graph.Drives[driveId].Root.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "webUrl"]);

            if (root?.Id == null) return null;

            string? serverRelUrl = null;
            if (root.WebUrl != null)
                serverRelUrl = Uri.UnescapeDataString(new Uri(root.WebUrl).AbsolutePath.TrimEnd('/'));

            var node = new SharePointNode
            {
                Id                 = root.Id,
                DriveId            = driveId,
                SiteId             = siteId,
                SiteUrl            = siteUrl,
                Name               = "Site Pages",
                Type               = NodeType.Library,
                HasChildren        = true,
                ServerRelativePath = serverRelUrl,
            };
            node.Children.Add(Placeholder());
            return node;
        }
        catch
        {
            return null;
        }
    }

    // Uses SharePoint REST to find the Site Pages list GUID — more reliable than Graph Lists API
    // because it matches by BaseTemplate=119 regardless of display name locale.
    private async Task<string?> GetSitePagesListIdViaRestAsync(string siteUrl)
    {
        // Try by BaseTemplate filter first (locale-independent)
        var filterUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists?$filter=BaseTemplate eq 119&$select=Id";
        try
        {
            using var filterResp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, filterUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (filterResp.IsSuccessStatusCode)
            {
                var body = await filterResp.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(body);
                var val = doc.RootElement.TryGetProperty("value", out var v) ? v : doc.RootElement;
                var first = val.EnumerateArray().FirstOrDefault();
                if (first.TryGetProperty("Id", out var id))
                    return id.GetString();
            }
        }
        catch { /* fall through to title-based lookup */ }

        // Fallback: look up by title "Site Pages" (English)
        var titleUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('Site Pages')?$select=Id";
        try
        {
            using var titleResp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, titleUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (titleResp.IsSuccessStatusCode)
            {
                var body = await titleResp.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("Id", out var id))
                    return id.GetString();
            }
        }
        catch { }

        return null;
    }

    // Reads page metadata from a SharePoint modern page using a two-step SitePages API lookup.
    // Step 1: find the page's integer ID by filename ($filter avoids any path encoding issues).
    // Step 2: fetch the full content by that ID (/_api/sitepages/pages/{id}).
    // Returns (metadata, null) on success, (null, errorMessage) on failure.
    public async Task<(PageMetadata? meta, string? error)> GetPageMetadataAsync(
        string siteUrl, string pageServerRelativeUrl)
    {
        var fileName    = System.IO.Path.GetFileName(pageServerRelativeUrl.Replace('\\', '/'));
        var escapedName = fileName.Replace("'", "''");

        // Step 1: find the SitePages integer ID by filename
        int pageId = 0;
        var listUrl = $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages" +
                      $"?$filter=FileName eq '{escapedName}'&$select=Id";
        System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] step1 GET {listUrl}");
        try
        {
            using var listResp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, listUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!listResp.IsSuccessStatusCode)
            {
                var err = await listResp.Content.ReadAsStringAsync();
                var msg = $"SitePages list HTTP {(int)listResp.StatusCode}: {err[..Math.Min(200, err.Length)]}";
                System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] {msg}");
                return (null, msg);
            }

            var body = await listResp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var arr))
            {
                foreach (var item in arr.EnumerateArray())
                {
                    if (item.TryGetProperty("Id", out var idProp))
                    {
                        pageId = idProp.GetInt32();
                        break;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] step1 exception: {ex.Message}");
            return (null, $"Page lookup exception: {ex.Message}");
        }

        if (pageId == 0)
        {
            System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] page '{fileName}' not found in SitePages list");
            return (null, $"Page '{fileName}' not found in source Site Pages library");
        }

        // Step 2: fetch full content by integer ID
        var pageUrl = $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages({pageId})" +
                      "?$select=CanvasContent1,LayoutWebpartsContent,BannerImageUrl,Description,PageLayoutType,PromotedState";
        System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] step2 GET {pageUrl}");
        try
        {
            using var pageResp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, pageUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!pageResp.IsSuccessStatusCode)
            {
                var err = await pageResp.Content.ReadAsStringAsync();
                var msg = $"SitePages content HTTP {(int)pageResp.StatusCode}: {err[..Math.Min(200, err.Length)]}";
                System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] {msg}");
                return (null, msg);
            }

            var body = await pageResp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;

            var meta = new PageMetadata
            {
                CanvasContent1        = root.TryGetProperty("CanvasContent1",        out var c1) ? c1.GetString() : null,
                LayoutWebpartsContent = root.TryGetProperty("LayoutWebpartsContent", out var lw) ? lw.GetString() : null,
                BannerImageUrl        = root.TryGetProperty("BannerImageUrl",        out var bi) ? bi.GetString() : null,
                Description           = root.TryGetProperty("Description",           out var d)  ? d.GetString()  : null,
                PageLayoutType        = root.TryGetProperty("PageLayoutType",        out var pl) ? pl.GetString() : null,
                PromotedState         = root.TryGetProperty("PromotedState",         out var ps) && ps.TryGetInt32(out var psv) ? psv : 0,
            };
            System.Diagnostics.Debug.WriteLine(
                $"[GetPageMetadata] OK — CanvasContent1={meta.CanvasContent1?.Length ?? 0} chars");
            return (meta, null);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] step2 exception: {ex.Message}");
            return (null, $"Page content fetch exception: {ex.Message}");
        }
    }

    // Applies page metadata to a copied page, substituting source URLs with target URLs.
    // Returns a warning string if web parts referencing source-specific list IDs are detected.
    public async Task<string?> ApplyPageMetadataAsync(
        string targetSiteUrl, string targetDriveId, string targetItemId,
        PageMetadata metadata, string sourceSiteUrl)
    {
        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] START targetItemId={targetItemId}");
        var formValues = new List<object>();

        string? SubstUrl(string? json) =>
            json == null ? null : SubstituteUrls(json, sourceSiteUrl, targetSiteUrl);

        var canvas  = SubstUrl(metadata.CanvasContent1);
        var layout  = SubstUrl(metadata.LayoutWebpartsContent);
        var banner  = SubstUrl(metadata.BannerImageUrl);

        if (canvas  != null) formValues.Add(new { FieldName = "CanvasContent1",        FieldValue = canvas });
        if (layout  != null) formValues.Add(new { FieldName = "LayoutWebpartsContent", FieldValue = layout });
        if (banner  != null) formValues.Add(new { FieldName = "BannerImageUrl",        FieldValue = banner });
        if (metadata.Description   != null) formValues.Add(new { FieldName = "Description",   FieldValue = metadata.Description });
        if (metadata.PageLayoutType != null) formValues.Add(new { FieldName = "PageLayoutType", FieldValue = metadata.PageLayoutType });
        formValues.Add(new { FieldName = "PromotedState", FieldValue = metadata.PromotedState.ToString() });

        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] {formValues.Count} fields to write, resolving SP IDs…");
        if (formValues.Count == 0) return null;

        var ids = await GetSharePointIdsAsync(targetDriveId, targetItemId)
            ?? throw new Exception($"Could not resolve SharePoint IDs for {targetDriveId}/{targetItemId}");
        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] SP IDs resolved, posting ValidateUpdateListItem…");

        var url = $"{ids.siteUrl.TrimEnd('/')}/_api/web/lists('{ids.listId}')/items({ids.listItemId})/ValidateUpdateListItem()";
        var payload = JsonSerializer.Serialize(new { formValues, bNewDocumentUpdate = true });

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, ids.siteUrl);

        if (!response.IsSuccessStatusCode)
        {
            var errBody = await response.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] ValidateUpdateListItem HTTP {(int)response.StatusCode}: {errBody[..Math.Min(errBody.Length, 300)]}");
            return $"ApplyPageMetadata HTTP {(int)response.StatusCode}";
        }

        // Detect web parts referencing source-specific list GUIDs
        bool hasListIdRefs = canvas != null &&
            System.Text.RegularExpressions.Regex.IsMatch(canvas,
                @"""[Ll]ist[Ii]d""\s*:\s*""[{]?[0-9a-fA-F\-]{36}[}]?""");

        var result = hasListIdRefs
            ? "Some web parts reference list IDs from the source site and may need manual review."
            : null;
        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] DONE, result={(result ?? "OK")}");
        return result;
    }

    internal static string SubstituteUrls(string json, string sourceUrl, string targetUrl)
    {
        var src     = sourceUrl.TrimEnd('/');
        var tgt     = targetUrl.TrimEnd('/');
        var srcEnc  = Uri.EscapeDataString(src);
        var tgtEnc  = Uri.EscapeDataString(tgt);
        var srcEnc2 = src.Replace(":", "%3A").Replace("/", "%2F");
        var tgtEnc2 = tgt.Replace(":", "%3A").Replace("/", "%2F");

        return json
            .Replace(srcEnc2, tgtEnc2, StringComparison.OrdinalIgnoreCase)
            .Replace(srcEnc,  tgtEnc,  StringComparison.OrdinalIgnoreCase)
            .Replace(src,     tgt,     StringComparison.OrdinalIgnoreCase);
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    public record NavigationNode(int Id, string Title, string Url, bool IsExternal, List<NavigationNode> Children);

    // Reads Quick Launch (quickLaunch=true) or Top Navigation Bar (quickLaunch=false) nodes.
    public async Task<List<NavigationNode>> GetNavigationNodesAsync(string siteUrl, bool quickLaunch)
    {
        var section  = quickLaunch ? "quicklaunch" : "topnavigationbar";
        var endpoint = $"{siteUrl.TrimEnd('/')}/_api/web/navigation/{section}?$expand=Children&$select=Id,Title,Url,IsExternal";

        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, endpoint);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("value", out var arr)
                ? ParseNavigationNodes(arr)
                : [];
        }
        catch { return []; }
    }

    // Copies Quick Launch and/or Top Navigation from source to target, remapping site URLs.
    public async Task CopyNavigationAsync(string sourceSiteUrl, string targetSiteUrl, bool quickLaunch)
    {
        var nodes   = await GetNavigationNodesAsync(sourceSiteUrl, quickLaunch);
        var section = quickLaunch ? "quicklaunch" : "topnavigationbar";

        // Clear existing nodes at target first
        var existing = await GetNavigationNodesAsync(targetSiteUrl, quickLaunch);
        foreach (var node in existing)
            await DeleteNavigationNodeAsync(targetSiteUrl, section, node.Id);

        // Recreate from source with URL substitution
        foreach (var node in nodes)
        {
            var mappedUrl = SubstituteUrls(node.Url, sourceSiteUrl, targetSiteUrl);
            var newId     = await CreateNavigationNodeAsync(targetSiteUrl, section, null, node.Title, mappedUrl, node.IsExternal);

            foreach (var child in node.Children)
            {
                var childUrl = SubstituteUrls(child.Url, sourceSiteUrl, targetSiteUrl);
                await CreateNavigationNodeAsync(targetSiteUrl, section, newId, child.Title, childUrl, child.IsExternal);
            }
        }
    }

    private static List<NavigationNode> ParseNavigationNodes(JsonElement array)
    {
        var nodes = new List<NavigationNode>();
        foreach (var el in array.EnumerateArray())
        {
            var id    = el.TryGetProperty("Id",         out var ip)  ? ip.GetInt32()    : 0;
            var title = el.TryGetProperty("Title",      out var tp)  ? tp.GetString()   ?? "" : "";
            var url   = el.TryGetProperty("Url",        out var up)  ? up.GetString()   ?? "" : "";
            var ext   = el.TryGetProperty("IsExternal", out var ep)  && ep.GetBoolean();

            var children = new List<NavigationNode>();
            if (el.TryGetProperty("Children", out var cp))
            {
                if (cp.ValueKind == JsonValueKind.Array)
                    children = ParseNavigationNodes(cp);
                else if (cp.TryGetProperty("value", out var cv))
                    children = ParseNavigationNodes(cv);
            }
            nodes.Add(new NavigationNode(id, title, url, ext, children));
        }
        return nodes;
    }

    private async Task DeleteNavigationNodeAsync(string siteUrl, string section, int nodeId)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/navigation/{section}/GetById({nodeId})";
        try
        {
            using var _ = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Delete, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Headers.Add("IF-MATCH", "*");
                return r;
            }, siteUrl);
        }
        catch { }
    }

    private async Task<int> CreateNavigationNodeAsync(
        string siteUrl, string section, int? parentId,
        string title, string url, bool isExternal)
    {
        var endpoint = parentId.HasValue
            ? $"{siteUrl.TrimEnd('/')}/_api/web/navigation/{section}/GetById({parentId.Value})/Children"
            : $"{siteUrl.TrimEnd('/')}/_api/web/navigation/{section}";

        var body = JsonSerializer.Serialize(new { Title = title, Url = url, IsExternal = isExternal });

        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, endpoint);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) return 0;
            var respBody = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(respBody);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetInt32() : 0;
        }
        catch { return 0; }
    }

    // ── Custom lists ──────────────────────────────────────────────────────────

    // System list templates to exclude from custom list enumeration.
    private static readonly HashSet<int> _systemListBaseTemplates = [101, 119, 171, 544, 890];

    // Returns non-hidden, non-catalog, non-library lists (Tasks, Calendars, Announcements, etc.).
    public async Task<List<(string Id, string Title, int BaseTemplate)>> GetCustomListsAsync(string siteUrl)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists" +
                  "?$filter=Hidden eq false and IsCatalog eq false" +
                  " and BaseTemplate ne 101 and BaseTemplate ne 119" +
                  "&$select=Id,Title,BaseTemplate&$orderby=Title";

        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);

            var result = new List<(string, string, int)>();
            if (!doc.RootElement.TryGetProperty("value", out var values)) return result;

            foreach (var el in values.EnumerateArray())
            {
                var id       = el.TryGetProperty("Id",           out var ip) ? ip.GetString()  ?? "" : "";
                var title    = el.TryGetProperty("Title",        out var tp) ? tp.GetString()  ?? "" : "";
                var baseTmpl = el.TryGetProperty("BaseTemplate", out var bp) ? bp.GetInt32()        : 0;
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(title)) continue;
                if (_systemListBaseTemplates.Contains(baseTmpl)) continue;
                result.Add((id, title, baseTmpl));
            }
            return result;
        }
        catch { return []; }
    }

    // Returns the list GUID for a list looked up by title, or null if not found.
    public async Task<string?> GetListIdByTitleAsync(string siteUrl, string title)
    {
        var encoded = Uri.EscapeDataString(title.Replace("'", "''"));
        var url     = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{encoded}')?$select=Id";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);
            if (!resp.IsSuccessStatusCode) return null;
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetString() : null;
        }
        catch { return null; }
    }

    // Reads all items from a list with the given custom field names plus Created/Modified.
    // Returns each item as a string-keyed dictionary. Handles pagination automatically.
    public async Task<List<Dictionary<string, object?>>> GetListItemsAsync(
        string siteUrl, string listId,
        IEnumerable<SpColumnDef> customColumns,
        IEnumerable<string>? extraFieldNames = null,
        CancellationToken ct = default)
    {
        var cols       = customColumns.ToList();
        var colsByName = cols.ToDictionary(c => c.InternalName, StringComparer.OrdinalIgnoreCase);

        // User fields are lookups: $expand them and read {name}/Name (claims login).
        var selectParts = new List<string> { "Id", "Title", "Created", "Modified" };
        var expandParts = new List<string>();
        foreach (var c in cols)
        {
            if (SpColumnDef.IsUserType(c.FieldType))
            {
                selectParts.Add($"{c.InternalName}/Name");
                expandParts.Add(c.InternalName);
            }
            else
            {
                selectParts.Add(c.InternalName);
            }
        }
        if (extraFieldNames != null)
            selectParts.AddRange(extraFieldNames);

        var baseUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                      $"?$select={Uri.EscapeDataString(string.Join(",", selectParts))}" +
                      (expandParts.Count > 0 ? $"&$expand={Uri.EscapeDataString(string.Join(",", expandParts))}" : "") +
                      "&$top=5000";

        var result = new List<Dictionary<string, object?>>();
        string? next = baseUrl;

        while (next != null)
        {
            ct.ThrowIfCancellationRequested();
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, next);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) break;
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);

            if (doc.RootElement.TryGetProperty("value", out var values))
                foreach (var el in values.EnumerateArray())
                {
                    var item = new Dictionary<string, object?>();
                    foreach (var prop in el.EnumerateObject())
                        item[prop.Name] = colsByName.TryGetValue(prop.Name, out var col)
                            ? ParseFieldValue(prop.Value, col.FieldType)
                            : ExtractJsonValue(prop.Value);
                    result.Add(item);
                }

            next = GetNextLink(doc.RootElement);
        }
        return result;
    }

    // Returns (Id, Title, Modified) for each item in a list — lightweight query used for
    // tree display and for skip / copy-if-newer decisions during list item copies.
    public async Task<List<(string Id, string Title, DateTimeOffset? Modified)>> GetListItemTitlesAsync(
        string siteUrl, string listId, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                  "?$select=Id,Title,Modified&$top=5000&$orderby=Id";

        var result = new List<(string, string, DateTimeOffset?)>();
        string? next = url;

        while (next != null)
        {
            ct.ThrowIfCancellationRequested();
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, next);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) break;
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);

            if (doc.RootElement.TryGetProperty("value", out var values))
                foreach (var el in values.EnumerateArray())
                {
                    var id    = el.TryGetProperty("Id",    out var ip) ? ip.GetInt32().ToString() : string.Empty;
                    var title = el.TryGetProperty("Title", out var tp) && tp.ValueKind == JsonValueKind.String
                                    ? tp.GetString() ?? string.Empty
                                    : string.Empty;
                    DateTimeOffset? modified = el.TryGetProperty("Modified", out var mp) &&
                                               mp.ValueKind == JsonValueKind.String &&
                                               DateTimeOffset.TryParse(mp.GetString(), out var m)
                                                   ? m : null;
                    if (!string.IsNullOrEmpty(id))
                        result.Add((id, title, modified));
                }

            next = GetNextLink(doc.RootElement);
        }
        return result;
    }

    // Creates a list item and optionally back-fills Created/Modified timestamps.
    // Returns (new item ID or null, field-write error or null).
    public async Task<(string? Id, string? FieldError)> CreateListItemAsync(
        string siteUrl, string listId,
        Dictionary<string, object?> fields,
        string? createdDate, string? modifiedDate,
        CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();

        // Person/taxonomy values cannot be set through the plain item POST body — they
        // go through the ValidateUpdateListItem pass below alongside the date back-fill.
        var (simpleFields, complexFields) = SplitComplexFields(fields);

        var createUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items";
        var body      = JsonSerializer.Serialize(simpleFields);

        string? newItemId = null;
        using (var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, createUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl))
        {
            if (!resp.IsSuccessStatusCode)
            {
                var errBody = await resp.Content.ReadAsStringAsync();
                throw new HttpRequestException($"Create item failed: {(int)resp.StatusCode} {resp.ReasonPhrase} — {errBody}");
            }
            var respBody = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(respBody);
            newItemId = doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetRawText() : null;
        }

        if (newItemId == null) return (null, null);
        var fieldError = await ValidateUpdateItemFieldsAsync(siteUrl, listId, newItemId, complexFields, createdDate, modifiedDate);
        return (newItemId, fieldError);
    }

    // Splits person/taxonomy values (which require ValidateUpdateListItem) from fields
    // that can be written through a plain REST POST/MERGE body.
    private static (Dictionary<string, object?> Simple, List<(string Name, object Value)> Complex)
        SplitComplexFields(Dictionary<string, object?> fields)
    {
        var simple  = new Dictionary<string, object?>();
        var complex = new List<(string, object)>();
        foreach (var (name, value) in fields)
        {
            if (value == null) continue;
            if (value is PersonFieldValue or TaxonomyFieldValue)
                complex.Add((name, value));
            else
                simple[name] = value;
        }
        return (simple, complex);
    }

    // Applies complex field values and/or Created/Modified back-fill in one
    // ValidateUpdateListItem call (bNewDocumentUpdate avoids creating a new version).
    // Returns an error string if the write failed, or null on success.
    private async Task<string?> ValidateUpdateItemFieldsAsync(
        string siteUrl, string listId, string itemId,
        List<(string Name, object Value)> complexFields,
        string? createdDate, string? modifiedDate)
    {
        var formValues = complexFields
            .Select(f => (object)new { FieldName = f.Name, FieldValue = FormatFieldValueForValidate(f.Value) })
            .ToList();
        if (createdDate  != null) formValues.Add(new { FieldName = "Created",  FieldValue = createdDate });
        if (modifiedDate != null) formValues.Add(new { FieldName = "Modified", FieldValue = modifiedDate });
        if (formValues.Count == 0) return null;

        var payload = JsonSerializer.Serialize(new { formValues, bNewDocumentUpdate = true });
        var url     = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({itemId})/ValidateUpdateListItem()";

        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl);

        var body = await resp.Content.ReadAsStringAsync();
        if (!resp.IsSuccessStatusCode)
            return $"Field write failed ({(int)resp.StatusCode}): {body[..Math.Min(200, body.Length)]}";

        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var vals))
            {
                var errors = vals.EnumerateArray()
                    .Where(v => v.TryGetProperty("HasException", out var ex) && ex.GetBoolean())
                    .Select(v => v.TryGetProperty("FieldName", out var fn) ? fn.GetString() : "?")
                    .ToList();
                if (errors.Count > 0)
                    return $"Field errors on: {string.Join(", ", errors)}";
            }
        }
        catch { }
        return null;
    }

    // Updates an existing list item's fields via MERGE/PATCH and optionally back-fills Created/Modified.
    // Returns a field-write error string if person/taxonomy fields could not be applied, or null on success.
    public async Task<string?> UpdateListItemAsync(
        string siteUrl, string listId, string itemId,
        Dictionary<string, object?> fields,
        string? createdDate, string? modifiedDate,
        CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();

        var (simpleFields, complexFields) = SplitComplexFields(fields);

        var updateUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({itemId})";
        var body      = JsonSerializer.Serialize(simpleFields);

        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, updateUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Headers.Add("X-HTTP-Method", "MERGE");
            r.Headers.Add("If-Match", "*");
            r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl);

        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync();
            throw new HttpRequestException($"Update item failed: {(int)resp.StatusCode} {resp.ReasonPhrase} — {errBody}");
        }

        return await ValidateUpdateItemFieldsAsync(siteUrl, listId, itemId, complexFields, createdDate, modifiedDate);
    }

    private static object? ExtractJsonValue(JsonElement el) => el.ValueKind switch
    {
        JsonValueKind.String  => el.GetString(),
        JsonValueKind.Number  => el.TryGetInt64(out var l) ? (object?)l : el.GetDouble(),
        JsonValueKind.True    => true,
        JsonValueKind.False   => false,
        JsonValueKind.Null    => null,
        _                     => null,
    };

    // Reads ID + HasUniqueRoleAssignments for all items in a library/list in one paginated request.
    // Returns a dictionary keyed by "{listId}:{itemId}" (see PermissionFlagKey), value =
    // hasUniquePermissions. The composite key lets flags from multiple source libraries
    // be merged into one dictionary without item-ID collisions.
    public async Task<Dictionary<string, bool>> BulkReadPermissionFlagsAsync(
        string siteUrl, string listId, CancellationToken ct = default)
    {
        var result  = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        string? next = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                       "?$select=Id,HasUniqueRoleAssignments&$top=5000";
        while (next != null)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var resp = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, next);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl, cancellationToken: ct);
                if (!resp.IsSuccessStatusCode) break;
                var body = await resp.Content.ReadAsStringAsync(ct);
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("value", out var vals))
                    foreach (var el in vals.EnumerateArray())
                    {
                        var id      = el.TryGetProperty("Id",                    out var ip) ? ip.GetInt32().ToString() : null;
                        var unique  = el.TryGetProperty("HasUniqueRoleAssignments", out var up) && up.GetBoolean();
                        if (id != null) result[PermissionFlagKey(listId, id)] = unique;
                    }
                next = GetNextLink(doc.RootElement);
            }
            catch { break; }
        }
        return result;
    }

    // Composite key for permission-flag dictionaries. GUID braces are stripped so keys
    // match whether the list ID came from SP REST or Graph sharepointIds.
    public static string PermissionFlagKey(string listId, string itemId) =>
        $"{listId.Trim('{', '}')}:{itemId}";

    // ── Permissions ───────────────────────────────────────────────────────────

    // Checks whether a single SP object has unique (broken) permissions.
    // apiPath examples: "web/lists('{guid}')", "web"
    public async Task<bool> GetHasUniqueRoleAssignmentsAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}?$select=HasUniqueRoleAssignments";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return false;
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("HasUniqueRoleAssignments", out var v) && v.GetBoolean();
        }
        catch { return false; }
    }

    public async Task<List<RoleAssignmentInfo>> GetRoleAssignmentsAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/roleassignments" +
                  "?$expand=Member,RoleDefinitionBindings&$select=Member/Id,Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("value", out var vals)) return [];

            var result = new List<RoleAssignmentInfo>();
            foreach (var el in vals.EnumerateArray())
            {
                if (!el.TryGetProperty("Member", out var member)) continue;
                var pid        = member.TryGetProperty("Id",            out var idEl)   ? idEl.GetInt32()          : 0;
                var loginName  = member.TryGetProperty("LoginName",     out var lnEl)   ? lnEl.GetString() ?? ""   : "";
                var title      = member.TryGetProperty("Title",         out var tEl)    ? tEl.GetString()  ?? ""   : "";
                var ptype      = member.TryGetProperty("PrincipalType", out var ptEl)   ? ptEl.GetInt32()          : 0;
                var roleNames  = new List<string>();
                if (el.TryGetProperty("RoleDefinitionBindings", out var rdbs))
                    foreach (var rdb in rdbs.EnumerateArray())
                        if (rdb.TryGetProperty("Name", out var nm))
                            roleNames.Add(nm.GetString() ?? "");
                result.Add(new RoleAssignmentInfo(pid, ptype, loginName, title, roleNames));
            }
            return result;
        }
        catch { return []; }
    }

    public async Task BreakPermissionInheritanceAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)";
        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(string.Empty, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct);
        if (!resp.IsSuccessStatusCode)
        {
            var body = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException($"BreakPermissionInheritance failed: {(int)resp.StatusCode} — {body}");
        }
    }

    public async Task AddRoleAssignmentAsync(
        string siteUrl, string apiPath, int principalId, int roleDefId, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/roleassignments/addroleassignment(principalid={principalId},roleDefId={roleDefId})";
        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(string.Empty, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct);
        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException($"AddRoleAssignment failed: {(int)resp.StatusCode} — {errBody}");
        }
    }

    public async Task<int?> EnsureUserAsync(
        string siteUrl, string loginName, CancellationToken ct = default)
    {
        var url  = $"{siteUrl.TrimEnd('/')}/_api/web/ensureuser";
        var body = JsonSerializer.Serialize(new { logonName = loginName });
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return null;
            var respBody = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(respBody);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetInt32() : null;
        }
        catch { return null; }
    }

    public async Task<int?> GetSiteGroupIdAsync(
        string siteUrl, string groupTitle, CancellationToken ct = default)
    {
        var encoded = Uri.EscapeDataString(groupTitle.Replace("'", "''"));
        var url     = $"{siteUrl.TrimEnd('/')}/_api/web/sitegroups/getbyname('{encoded}')?$select=Id";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return null;
            var respBody = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(respBody);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetInt32() : null;
        }
        catch { return null; }
    }

    // Loads all role definitions for the site and returns a name→ID dictionary.
    public async Task<Dictionary<string, int>> GetAllRoleDefinitionsAsync(
        string siteUrl, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/roledefinitions?$select=Id,Name";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("value", out var vals)) return [];
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var el in vals.EnumerateArray())
            {
                var id   = el.TryGetProperty("Id",   out var idEl)   ? idEl.GetInt32()        : 0;
                var name = el.TryGetProperty("Name", out var nameEl) ? nameEl.GetString() ?? "" : "";
                if (id > 0 && !string.IsNullOrEmpty(name))
                    result[name] = id;
            }
            return result;
        }
        catch { return []; }
    }

    private sealed class TokenProvider : IAccessTokenProvider
    {
        private readonly AuthService _auth;
        public TokenProvider(AuthService auth) => _auth = auth;

        public AllowedHostsValidator AllowedHostsValidator { get; } =
            new AllowedHostsValidator(["graph.microsoft.com"]);

        public async Task<string> GetAuthorizationTokenAsync(
            Uri uri,
            Dictionary<string, object>? additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
            => await _auth.GetAccessTokenAsync();
    }
}
