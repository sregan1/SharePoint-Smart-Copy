using SharePointSmartCopy.Localization;
using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
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

// Partial class SharePointService: Content. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
    // ── File content ──────────────────────────────────────────────────────────

    public async Task<Stream> DownloadFileAsync(string driveId, string itemId, CancellationToken ct = default)
    {
        var stream = await Graph.Drives[driveId].Items[itemId].Content.GetAsync(cancellationToken: ct);
        return stream ?? throw new Exception(Loc.T("Svc_Sp_EmptyResponseDownloadingFile"));
    }

    public async Task<Stream> DownloadVersionAsync(string driveId, string itemId, string versionId, CancellationToken ct = default)
    {
        var stream = await Graph.Drives[driveId].Items[itemId].Versions[versionId].Content.GetAsync(cancellationToken: ct);
        return stream ?? throw new Exception(Loc.T("Svc_Sp_EmptyResponseDownloadingVersion"));
    }

    // Result of a ranged content fetch: whether the server actually honored the Range request
    // (206 + a Content-Range starting at the byte we asked for) vs. silently ignoring it and
    // sending the WHOLE file back from byte 0 (200) — the caller must know which happened, since
    // appending a 200 response to already-buffered bytes would corrupt the file.
    public readonly record struct RangedContent(Stream Content, bool IsPartial, long StartOffset, string? ETag = null);

    // Fetches file/version content starting at rangeStart, bypassing the Kiota-generated Content
    // accessor so a Range header can be attached to the request. Graph's content endpoint redirects
    // (302) to a pre-authenticated SharePoint/OneDrive blob URL; HttpClient's default redirect
    // handling forwards the Range header across that redirect, and the blob endpoint honors it with
    // a 206 + Content-Range — this is what lets a connection reset on a multi-GB file resume from
    // where it left off instead of re-downloading the whole thing (the single biggest cost observed
    // on a run dominated by multi-GB scan files hitting frequent mid-transfer resets).
    // versionId null = current version; otherwise a specific historical version.
    // Thrown by DownloadContentRangeAsync on a 429/503 so the download retry loop can distinguish a
    // throttle (wait out the server's Retry-After) from a connection reset (short fixed backoff).
    internal sealed class DownloadThrottledException(TimeSpan retryAfter) : Exception
    {
        public TimeSpan RetryAfter { get; } = retryAfter;
    }

    public async Task<RangedContent> DownloadContentRangeAsync(
        string driveId, string itemId, string? versionId, long rangeStart, CancellationToken ct, string? ifMatchETag = null)
    {
        var token = await _router.ForDrive(driveId).GetAccessTokenAsync(cancellationToken: ct);
        var url = versionId == null
            ? $"{GraphBaseUrl}/drives/{driveId}/items/{itemId}/content"
            : $"{GraphBaseUrl}/drives/{driveId}/items/{itemId}/versions/{versionId}/content";

        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (rangeStart > 0)
        {
            req.Headers.Range = new RangeHeaderValue(rangeStart, null);
            // Only continue if the file is still the one the first bytes came from; otherwise the buffer would hold old bytes
            // followed by new ones (a 412 comes back and the caller restarts from byte 0).
            if (ifMatchETag is not null) req.Headers.TryAddWithoutValidation("If-Match", ifMatchETag);
        }

        var response = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        // Content downloads use the raw HttpClient (needed for Range/resume), which bypasses the Graph
        // client's GraphThrottleNotifyHandler — so a 429/503 here is otherwise INVISIBLE to the adaptive
        // download controller, which then keeps hammering a throttled endpoint at full width (observed
        // 2026-07-20: 20+ min download stalls with no "Downloads: throttle backoff" line). Surface it via
        // the Throttled event (so the controller steps down) and throw a distinct exception carrying the
        // Retry-After so the caller's retry loop can wait the real window out instead of a blind 5s.
        if (response.StatusCode is System.Net.HttpStatusCode.TooManyRequests or System.Net.HttpStatusCode.ServiceUnavailable)
        {
            var retryAfter = response.Headers.RetryAfter?.Delta
                ?? (response.Headers.RetryAfter?.Date is { } when ? when - DateTimeOffset.UtcNow : (TimeSpan?)null)
                ?? TimeSpan.FromSeconds(10);
            if (retryAfter < TimeSpan.Zero) retryAfter = TimeSpan.FromSeconds(10);
            response.Dispose();
            Throttled?.Invoke(retryAfter, 1, 1, "content download");
            throw new DownloadThrottledException(retryAfter);
        }
        // Until the stream is handed to the caller the response is ours: a 4xx/5xx or a failed read must not leave the
        // connection checked out until the GC finalizes it (retried failures piled these up on runs with many 404s).
        HttpResponseMessage? owned = response;
        try
        {
            response.EnsureSuccessStatusCode(); // other 4xx/5xx -> HttpRequestException, same as the old Kiota path threw

            bool isPartial = response.StatusCode == System.Net.HttpStatusCode.PartialContent;
            long startOffset = isPartial && response.Content.Headers.ContentRange?.From is { } from ? from : 0;
            var stream = await response.Content.ReadAsStreamAsync(ct);
            owned = null;   // the stream (and with it the connection) now belongs to the caller
            return new RangedContent(stream, isPartial, startOffset, response.Headers.ETag?.ToString());
        }
        finally { owned?.Dispose(); }
    }

    public async Task<List<DriveItemVersion>> GetVersionsAsync(string driveId, string itemId)
    {
        // Only id/size/lastModifiedDateTime/lastModifiedBy are ever read off these (version replay in
        // CopyService/MigrationJobService) — $select trims the rest of the full DriveItemVersion
        // payload per version, per file.
        var all  = new List<DriveItemVersion>();
        var page = await Graph.Drives[driveId].Items[itemId].Versions.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "size", "lastModifiedDateTime", "lastModifiedBy"]);
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
    //
    // CURRENTLY UNCALLED anywhere in this codebase. Before wiring it up: it lacks the retry rounds,
    // adaptive throttle-aware gate, and Throttled subscription its siblings (FetchBatchChunkAsync's
    // callers) have, and the "default to 1 on failure" behavior is exactly the under-count
    // BatchDownloadMetadataAsync's own comment warns about — "an incomplete cache would under-count
    // a file's versions, and that file's batch could then exceed the SPMI entry ceiling and fail
    // import." Harden it to match those siblings before any caller depends on it.
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
        // Soft-start + inherit any active throttle window (see CreateThrottleAwareGate) so a phase that
        // follows an already-throttled one starts low instead of bursting into the same limit.
        using var gate = CreateThrottleAwareGate(maxConcurrency, Math.Min(maxConcurrency, 2));
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
                        var chunkAuth = _router.ForDrive(chunk[0].driveId); // resolved outside the try below so a routing error is loud
                        for (int i = 0; i < chunk.Length; i++)
                        {
                            var (driveId, itemId) = chunk[i];
                            var req = new HttpRequestMessage(HttpMethod.Get,
                                $"{GraphBaseUrl}/drives/{driveId}/items/{itemId}?$select=lastModifiedDateTime");
                            ids[i] = batch.AddBatchRequestStep(req);
                        }

                        BatchResponseContentCollection response;
                        try { using (_router.Scope(chunkAuth)) response = await Graph.Batch.PostAsync(batch, c); }
                        catch (OperationCanceledException) { throw; }
                        catch { return; } // whole batch call failed — retry rounds below cover these items

                        // Same sub-response throttle surfacing as FetchBatchChunkAsync: 429s inside a
                        // 200-OK $batch envelope are invisible to the pipeline handlers, so without
                        // this the retry rounds re-burst instantly into the depleted budget.
                        TimeSpan? subThrottleDelay = null;
                        for (int i = 0; i < chunk.Length; i++)
                        {
                            if (ids[i] == null) continue;
                            try
                            {
                                using var http = await response.GetResponseByIdAsync(ids[i]!);
                                if (!http.IsSuccessStatusCode)
                                {
                                    if (GetBatchSubResponseThrottleDelay(http) is { } d &&
                                        (subThrottleDelay == null || d > subThrottleDelay))
                                        subThrottleDelay = d;
                                    continue;
                                }
                                using var doc = JsonDocument.Parse(await http.Content.ReadAsStringAsync(c));
                                result[chunk[i].itemId] = TryGetBatchDateTimeOffset(doc.RootElement, "lastModifiedDateTime");
                            }
                            catch { /* leave absent; retried below, then caller falls back if still missing */ }
                        }
                        if (subThrottleDelay is { } throttle)
                            Throttled?.Invoke(throttle, 1, 1, "inside $batch");
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

    // Supplementary pass for the delta-based scan (see EnumerateFilesForCopyViaDeltaAsync / C2): Graph's
    // /delta does NOT support $expand=listItem (confirmed by a Microsoft maintainer — 400 "One of the
    // provided arguments is not acceptable"), so Document Set detection via listItem.contentType (see
    // IsSpecialContainer) can't ride along on the delta call the way it does on the per-folder /children
    // listing. Batched here (20 folders/call, same shape as FetchModifiedDatesAsync) rather than issued
    // one call per folder — a per-folder call here would erase delta's whole advantage, since delta's
    // win specifically comes from replacing "1 round trip per folder" with "1 round trip per ~thousand
    // items regardless of folder count". Only called for PLAIN folders (no package facet already seen)
    // that survive the package-facet-based special-container filter — OneNote-style containers never
    // reach this method at all.
    private async Task<Dictionary<string, string?>> BatchFetchContentTypeIdsAsync(
        IReadOnlyList<(string driveId, string itemId)> folders,
        int maxConcurrency,
        IProgress<int>? progress = null,
        CancellationToken ct = default)
    {
        var result = new System.Collections.Concurrent.ConcurrentDictionary<string, string?>();
        if (folders.Count == 0) return new Dictionary<string, string?>();

        const int ItemsPerBatch = 20; // 1 sub-request per folder

        using var gate = CreateThrottleAwareGate(maxConcurrency, Math.Min(maxConcurrency, 2));
        void onThrottle(TimeSpan delay, int __, int ___, string? ____) => gate.StepDown(delay);
        Throttled += onThrottle;

        int lastReported = 0;
        var reportLock = new object();
        void ReportProgress()
        {
            int resolved = result.Count;
            lock (reportLock)
            {
                if (resolved - lastReported < 500 && resolved != folders.Count) return;
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
                        var chunkAuth = _router.ForDrive(chunk[0].driveId); // resolved outside the try below so a routing error is loud
                        for (int i = 0; i < chunk.Length; i++)
                        {
                            var (driveId, itemId) = chunk[i];
                            var req = new HttpRequestMessage(HttpMethod.Get,
                                $"{GraphBaseUrl}/drives/{driveId}/items/{itemId}?$select=id&$expand=listItem($select=contentType)");
                            ids[i] = batch.AddBatchRequestStep(req);
                        }

                        BatchResponseContentCollection response;
                        try { using (_router.Scope(chunkAuth)) response = await Graph.Batch.PostAsync(batch, c); }
                        catch (OperationCanceledException) { throw; }
                        catch { return; } // whole batch call failed — retry rounds below cover these items

                        TimeSpan? subThrottleDelay = null;
                        for (int i = 0; i < chunk.Length; i++)
                        {
                            if (ids[i] == null) continue;
                            try
                            {
                                using var http = await response.GetResponseByIdAsync(ids[i]!);
                                if (!http.IsSuccessStatusCode)
                                {
                                    if (GetBatchSubResponseThrottleDelay(http) is { } d &&
                                        (subThrottleDelay == null || d > subThrottleDelay))
                                        subThrottleDelay = d;
                                    continue;
                                }
                                using var doc = JsonDocument.Parse(await http.Content.ReadAsStringAsync(c));
                                string? contentTypeId = null;
                                if (doc.RootElement.TryGetProperty("listItem", out var li) &&
                                    li.ValueKind == JsonValueKind.Object &&
                                    li.TryGetProperty("contentType", out var ct2) &&
                                    ct2.ValueKind == JsonValueKind.Object &&
                                    ct2.TryGetProperty("id", out var idProp))
                                    contentTypeId = idProp.GetString();
                                result[chunk[i].itemId] = contentTypeId;
                            }
                            catch { /* leave absent; retried below */ }
                        }
                        if (subThrottleDelay is { } throttle)
                            Throttled?.Invoke(throttle, 1, 1, "inside $batch");
                    }
                    finally { gate.Release(); }

                    ReportProgress();
                });
        }

        try
        {
            await RunPassAsync(folders);
            for (int round = 0; round < 8; round++)
            {
                var missing = folders.Where(f => !result.ContainsKey(f.itemId)).ToList();
                if (missing.Count == 0) break;
                await RunPassAsync(missing);
            }
        }
        finally { Throttled -= onThrottle; }

        return new Dictionary<string, string?>(result);
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
        // Soft-start + inherit any active throttle window (see CreateThrottleAwareGate).
        using var gate = CreateThrottleAwareGate(maxConcurrency, Math.Min(maxConcurrency, 2));
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

    // Extracts the Retry-After delay from a throttled (429/503) $batch SUB-response, or null when
    // the sub-response wasn't throttled. Sub-request throttles never pass through the HTTP pipeline
    // handlers — the $batch call itself returns 200 — so neither Kiota's retry handler nor
    // GraphThrottleNotifyHandler ever sees them. Before these were surfaced via the Throttled event,
    // the bulk-fetch retry rounds re-burst at full width straight back into the depleted budget and
    // exhausted every round in seconds (2026-07-18 run: 4,301 of 5,295 files left metadata-less in
    // ~100s, which then bypassed the large-file memory gate downstream).
    private static TimeSpan? GetBatchSubResponseThrottleDelay(HttpResponseMessage http)
    {
        if ((int)http.StatusCode is not (429 or 503)) return null;
        var delay = http.Headers.RetryAfter?.Delta
            ?? (http.Headers.RetryAfter?.Date is { } when
                    ? when - DateTimeOffset.UtcNow
                    : TimeSpan.FromSeconds(15));
        if (delay <= TimeSpan.Zero) delay = TimeSpan.FromSeconds(1);
        if (delay > TimeSpan.FromSeconds(120)) delay = TimeSpan.FromSeconds(120);
        return delay;
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

        BatchResponseContentCollection response;
        using (_router.Scope(_router.ForDrive(chunk[0].driveId)))
            response = await Graph.Batch.PostAsync(batch, ct);

        // Max Retry-After across this chunk's throttled sub-responses — raised once (below) via the
        // Throttled event so subscribed adaptive gates step down and hold their quiet window, making
        // the caller's retry rounds actually wait out the throttle instead of re-bursting into it.
        TimeSpan? subThrottleDelay = null;
        void NoteSubThrottle(HttpResponseMessage http)
        {
            if (GetBatchSubResponseThrottleDelay(http) is { } d &&
                (subThrottleDelay == null || d > subThrottleDelay))
                subThrottleDelay = d;
        }

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
                    else NoteSubThrottle(metaHttp);
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
                        // A thrown page is already handled by the outer catch. A NULL page (Kiota's
                        // representation of an empty/204 body) previously exited this loop silently
                        // with whatever partial list had accumulated so far, and `versions` below
                        // was still set from that partial list — the caller's own comment on
                        // multiVersionItemIds warns exactly what this causes: an under-counted file
                        // can then exceed the SPMI batch entry ceiling and fail import. Treat it the
                        // same as a failure: leave `versions` null so this item is skipped below and
                        // the caller falls back to an individual, complete fetch.
                        bool pageMissing = false;
                        while (nextLink != null)
                        {
                            ct.ThrowIfCancellationRequested();
                            var page = await Graph.Drives[driveId].Items[itemId].Versions
                                .WithUrl(nextLink).GetAsync(cancellationToken: ct);
                            if (page == null) { pageMissing = true; break; }
                            vList.AddRange(page.Value ?? []);
                            nextLink = page.OdataNextLink;
                        }
                        if (!pageMissing) versions = SortVersions(vList);
                    }
                    else NoteSubThrottle(versHttp);
                }
                catch { }
            }

            if (metadata != null && versions != null)
                result[itemId] = (metadata, versions);
        }

        if (subThrottleDelay is { } throttle)
            Throttled?.Invoke(throttle, 1, 1, "inside $batch");
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

}
