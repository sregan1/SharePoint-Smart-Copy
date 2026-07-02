using System.IO;
using System.Text.Json;
using System.Threading.Channels;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using SharePointSmartCopy.Models;
using CopyStatus = SharePointSmartCopy.Models.CopyStatus;

namespace SharePointSmartCopy.Services;

// Orchestrates the full Migration API pipeline for a batch of files:
// ProvisionContainers → build package XML → encrypt blobs → upload → submit job → poll.
// With maxParallel > 1: files are round-robin partitioned into min(maxParallel, 5) concurrent SPMI jobs;
// blob uploads within each job also run at up to maxParallel concurrency.
public class MigrationJobService(SharePointService spService)
{
    public async Task ExecuteAsync(
        IList<(CopyJob job, CopyResult result)> fileTasks,
        OverwriteMode overwriteMode,
        int maxVersions,
        int maxParallel,
        CancellationToken cancellationToken,
        bool copyCustomColumns = false,
        List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null,
        IProgress<(int completed, int total)>? preflightProgress = null,
        IProgress<string>? activityLog = null,
        IProgress<int>? onFilePacked = null)
    {
        if (fileTasks.Count == 0) return;

        int   preflightTotal   = fileTasks.Count;
        int[] preflightCounter = { 0 };
        preflightProgress?.Report((0, preflightTotal));

        var targetSiteUrl = fileTasks[0].job.TargetSiteUrl;
        if (string.IsNullOrEmpty(targetSiteUrl))
            throw new InvalidOperationException("TargetSiteUrl must be set on CopyJob for Migration API mode.");

        foreach (var (_, result) in fileTasks)
            result.Status = CopyStatus.Copying;

        // Adaptive gate — limits total concurrent Graph content downloads across all SPMI batches
        // and steps down concurrency when Graph throttles, then restores on a 5-second heartbeat.
        // The StepDown cooldown (2 s) prevents a burst of simultaneous 429s from cascading all the
        // way to 1 slot; in practice the controller finds an equilibrium just below the bandwidth
        // threshold and stays there, yielding much higher sustained throughput than a fixed gate
        // that repeatedly bursts → throttles → waits 6 s → bursts again.
        // Soft-start at 8 slots and ramp up to maxParallel via the restore heartbeat. Raised from 6
        // now that per-file metadata Graph calls are fetched once upfront (not per-batch during the
        // copy), so throttling is far milder and the opening burst tolerates a higher floor.
        int migrationSoftStart = Math.Min(maxParallel, 8);
        using var downloadController = new AdaptiveParallelismController(maxParallel, migrationSoftStart);
        // Registered only during the download/upload phase — NOT during analysis (metadata fetches,
        // folder enumeration). Analysis throttles are transient and shouldn't pre-damage the download
        // slot count before any file transfers have started.
        void onMigrationThrottle(TimeSpan delay, int __, int ___, string? ____) =>
            downloadController.StepDown(delay);

        // Separate adaptive gate for the pre-flight analysis loops (subfolder provisioning, existing-
        // file scan) — same backoff mechanism as downloadController but kept fully isolated from it,
        // per the isolation rationale above. Without this, those loops ran a plain fixed-width
        // Parallel.ForEachAsync(8): each worker individually waited out its own Retry-After but then
        // the full 8-wide burst resumed immediately, walking straight back into the same depleted
        // throttle budget (observed as repeated 60-120s waits back to back). This settles the width
        // below the threshold instead, same as the download gate does.
        const int AnalysisMaxParallelism = 8;
        using var analysisController = new AdaptiveParallelismController(AnalysisMaxParallelism);
        void onAnalysisThrottle(TimeSpan delay, int __, int ___, string? ____) =>
            analysisController.StepDown(delay);
        if (activityLog != null)
        {
            int lastAnalysisLimit = AnalysisMaxParallelism;
            analysisController.LimitChanged += n =>
            {
                bool down = n < lastAnalysisLimit;
                lastAnalysisLimit = n;
                activityLog.Report(down
                    ? $"↓ Analysis: {n}/{AnalysisMaxParallelism} slots (throttle backoff)"
                    : $"⬆ Analysis: {n}/{AnalysisMaxParallelism} slots (recovering)");
            };
        }
        if (activityLog != null)
        {
            int lastDlLimit = maxParallel;
            downloadController.LimitChanged += n =>
            {
                bool down = n < lastDlLimit;
                lastDlLimit = n;
                activityLog.Report(down
                    ? $"↓ Downloads: {n}/{maxParallel} slots (throttle backoff)"
                    : $"⬆ Downloads: {n}/{maxParallel} slots (recovering)");
            };
        }

        try
        {
            // Fetch shared pre-flight info concurrently — both calls are independent.
            var webInfoTask = spService.GetWebInfoAsync(targetSiteUrl);
            var siteIdTask  = spService.GetSiteIdAsync(targetSiteUrl);
            await Task.WhenAll(webInfoTask, siteIdTask);
            var (webId, webRelUrl) = webInfoTask.Result;
            var rawSiteId          = siteIdTask.Result;
            var siteId             = rawSiteId.Contains(',') ? rawSiteId.Split(',')[1] : rawSiteId;

            // Group by target drive — the manifest is built per library, so a batch that
            // spans multiple target libraries must run one pipeline per library.
            var driveGroups = fileTasks.GroupBy(t => t.job.TargetDriveId).ToList();

            foreach (var driveGroup in driveGroups)
            {
                var groupTasks = driveGroup.ToList();
                var firstJob   = groupTasks[0].job;

                var libraryServerRelUrl = firstJob.TargetLibraryServerRelativeUrl;
                if (string.IsNullOrEmpty(libraryServerRelUrl))
                    libraryServerRelUrl = await spService.GetLibraryServerRelativeUrlAsync(firstJob.TargetDriveId);

                string listId;
                try
                {
                    listId = await spService.GetListIdByServerRelativeUrlAsync(targetSiteUrl, libraryServerRelUrl);
                }
                catch when (!string.IsNullOrEmpty(firstJob.TargetDriveId))
                {
                    // URL-based lookup can fail when the library's actual server-relative URL
                    // differs from what was stored (e.g. "Shared Documents" vs "Documents").
                    // Fall back to resolving the list ID directly from the drive via Graph.
                    var fallbackId = await spService.GetListIdFromDriveAsync(firstJob.TargetDriveId);
                    listId = fallbackId
                        ?? throw new Exception($"Cannot resolve list ID for library at '{libraryServerRelUrl}'");
                }
                var libraryTitle = libraryServerRelUrl.Split('/').Last();

                // Pre-create subfolders before parallel split — concurrent creates for the same path conflict.
                // Capture the returned IDs so each batch can skip re-resolving them independently.
                var subFolderPaths = groupTasks
                    .Select(t => t.job.TargetSubFolderPath)
                    .Where(p => !string.IsNullOrEmpty(p))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var sharedFolderIdCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                sharedFolderIdCache[string.Empty] = firstJob.TargetParentItemId;

                // Folder SharePoint UniqueIds (server-relative URL → GUID), resolved at most once each
                // and shared across ALL batches. The deep folder tree repeats every batch; resolving it
                // per batch multiplied the throttle exposure that was returning null GUIDs and breaking
                // the manifest. Resolved resiliently (retry-until-success) so throttling slows, not fails.
                var sharedFolderUniqueIdCache =
                    new System.Collections.Concurrent.ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // Fresh-target fast path: check (cheaply, BEFORE we create any subfolders) whether the
                // destination is completely empty. If so, nothing can already exist, so the per-folder
                // existing-file scan below is skipped entirely — one $top=1 call instead of O(folders)
                // listings. Must be checked before creating subfolders, which would make it non-empty.
                var targetIsEmpty = await spService.IsFolderEmptyAsync(
                    firstJob.TargetDriveId, firstJob.TargetParentItemId);

                // Adaptive gate scoped tightly to just these two pre-flight loops — subscribed only
                // for their duration so throttles here don't bleed into the download-phase controller
                // (and vice versa), matching the isolation already used for onMigrationThrottle below.
                var sharedExistingByFolder = new System.Collections.Concurrent.ConcurrentDictionary<
                    string, Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>(
                    StringComparer.OrdinalIgnoreCase);
                spService.Throttled += onAnalysisThrottle;
                try
                {
                    if (subFolderPaths.Count > 0)
                    {
                        activityLog?.Report($"Provisioning {subFolderPaths.Count} target subfolder{(subFolderPaths.Count == 1 ? "" : "s")}...");
                        var cacheLock = new object();
                        // Outer MaxDegreeOfParallelism just provides enough lanes; analysisController's
                        // own semaphore (capped at AnalysisMaxParallelism, shrunk on throttle) is what
                        // actually governs live concurrency — the segment cache
                        // (SharePointService._folderSegmentCache) makes concurrent creates safe: races on
                        // shared path prefixes resolve via 409 conflict recovery, and the cache ensures
                        // each unique segment is created at most once.
                        await Parallel.ForEachAsync(subFolderPaths,
                            new ParallelOptions { MaxDegreeOfParallelism = AnalysisMaxParallelism, CancellationToken = cancellationToken },
                            async (folderPath, ct) =>
                            {
                                await analysisController.WaitAsync(ct);
                                try
                                {
                                    var id = await spService.GetOrCreateFolderPathAsync(
                                        firstJob.TargetDriveId, firstJob.TargetParentItemId, folderPath);
                                    lock (cacheLock) sharedFolderIdCache[folderPath] = id;
                                }
                                finally { analysisController.Release(); }
                            });
                    }

                    // Build the target-folder file listing once and share across all SPMI batches.
                    // Without sharing, N parallel batches each independently fetch all M folders = N×M calls.
                    // With sharing: M calls total (parallel at up to 8, adaptively), then all batches read
                    // the snapshot.
                    if (targetIsEmpty)
                    {
                        // Nothing exists — populate empty listings with no Graph calls at all. The per-batch
                        // pre-flight fast-path then treats every file as new (no scan, no zombie checks).
                        activityLog?.Report("Target is empty — skipping the existing-file scan.");
                        foreach (var key in sharedFolderIdCache.Keys)
                            sharedExistingByFolder[key] = new Dictionary<string, (string ItemId, DateTimeOffset? Modified)>(StringComparer.OrdinalIgnoreCase);
                    }
                    else
                    {
                        activityLog?.Report($"Scanning {sharedFolderIdCache.Count} target folder{(sharedFolderIdCache.Count == 1 ? "" : "s")} for existing files...");
                        await Parallel.ForEachAsync(sharedFolderIdCache,
                            new ParallelOptions { MaxDegreeOfParallelism = AnalysisMaxParallelism, CancellationToken = cancellationToken },
                            async (kvp, ct) =>
                            {
                                await analysisController.WaitAsync(ct);
                                try
                                {
                                    sharedExistingByFolder[kvp.Key] = await spService.FetchFolderItemsAsync(
                                        firstJob.TargetDriveId, kvp.Value);
                                }
                                finally { analysisController.Release(); }
                            });
                    }
                }
                finally
                {
                    spService.Throttled -= onAnalysisThrottle;
                }

                // A file matched by name here will be re-evaluated (and, for Skip, always skipped) later
                // in the per-batch pre-flight anyway, so fetching its metadata/versions for batching below
                // is wasted work for anything that turns out skippable. On a mostly-already-copied library
                // that waste dominates — thousands of metadata calls for files that were never going to be
                // copied, which is what exhausts the tenant's Graph rate limit and causes throttling to
                // bleed into the (much smaller) real download work. Overwrite always needs every file's
                // metadata (it re-copies everything unconditionally), so this only applies to Skip and If
                // Newer — and If Newer additionally needs the source's modified date to decide, so it gets
                // a cheap metadata-ONLY bulk fetch (1 sub-request/file) for just the existing subset,
                // rather than the full metadata+version fetch (2 sub-requests/file) below.
                if (overwriteMode is OverwriteMode.Skip or OverwriteMode.IfNewer)
                {
                    var existingAtTarget = new List<(CopyJob job, CopyResult result, DateTimeOffset? TargetModified)>();
                    var newAtTarget = new List<(CopyJob job, CopyResult result)>();
                    foreach (var t in groupTasks)
                    {
                        var subPath = t.job.TargetSubFolderPath ?? string.Empty;
                        if (sharedExistingByFolder.TryGetValue(subPath, out var existing) &&
                            existing.TryGetValue(t.job.SourceName, out var existingEntry))
                            existingAtTarget.Add((t.job, t.result, existingEntry.Modified));
                        else
                            newAtTarget.Add(t);
                    }

                    List<(CopyJob job, CopyResult result)> toProcess;
                    int preskipped;
                    if (overwriteMode == OverwriteMode.Skip)
                    {
                        foreach (var t in existingAtTarget) t.result.Status = CopyStatus.Skipped;
                        preskipped = existingAtTarget.Count;
                        toProcess = newAtTarget;
                    }
                    else // IfNewer
                    {
                        // Source modified dates captured during the scan walk (job.SourceModified) decide
                        // skip-vs-copy with ZERO Graph calls. Only jobs built outside the walk (single-file
                        // picks → SourceModified null) still need the bulk fetch. VERIFIED (114k-file run,
                        // 2026-07-01): fetching all 114k dates here took 22 minutes under throttling and
                        // still left 110k unresolved — which the "undetermined → needs copy" fallback then
                        // misrouted into hours of per-file re-checking. Dates in hand make that impossible.
                        var needsFetch = existingAtTarget
                            .Where(t => t.job.SourceModified == null)
                            .ToList();
                        Dictionary<string, DateTimeOffset?> modDates;
                        if (needsFetch.Count > 0)
                        {
                            activityLog?.Report($"Checking modified dates for {needsFetch.Count:N0} existing file(s)...");
                            var modDateProgress = new Progress<int>(d =>
                                activityLog?.Report($"Checking modified dates: {d:N0} / {needsFetch.Count:N0}"));
                            modDates = await spService.FetchModifiedDatesAsync(
                                needsFetch.Select(t => (t.job.SourceDriveId, t.job.SourceItemId)).ToList(),
                                maxConcurrency: 6, progress: modDateProgress, ct: cancellationToken);
                        }
                        else
                        {
                            modDates = [];
                        }

                        var stillNeedsCopy = new List<(CopyJob job, CopyResult result)>();
                        preskipped = 0;
                        int undetermined = 0;
                        foreach (var t in existingAtTarget)
                        {
                            // Undetermined (no scan-captured date AND fetch failed even after retries) is
                            // treated as needing a copy, not silently skipped — the full pass below will
                            // resolve it for real, same fallback direction the cache-miss handling
                            // elsewhere in this file already uses.
                            var srcModified = t.job.SourceModified
                                ?? (modDates.TryGetValue(t.job.SourceItemId, out var fetched) ? fetched : null);
                            if (srcModified.HasValue && srcModified.Value <= t.TargetModified)
                            {
                                t.result.Status       = CopyStatus.Skipped;
                                t.result.ErrorMessage  = CopyResult.UpToDate;
                                preskipped++;
                            }
                            else
                            {
                                if (!srcModified.HasValue) undetermined++;
                                stillNeedsCopy.Add((t.job, t.result));
                            }
                        }
                        // Surfaced so a repeat of the 114k-file throttle storm (2026-07-01, see
                        // FetchModifiedDatesAsync) is visible immediately instead of only showing up as
                        // an inexplicably low skip rate discovered many minutes later.
                        if (undetermined > 0)
                            activityLog?.Report($"⚠ {undetermined:N0} file(s) had no confirmed modified date after retries — treated as needing a copy");
                        toProcess = stillNeedsCopy.Concat(newAtTarget).ToList();
                    }

                    if (preskipped > 0)
                    {
                        int done = Interlocked.Add(ref preflightCounter[0], preskipped);
                        preflightProgress?.Report((done, preflightTotal));
                        activityLog?.Report($"{preskipped:N0} of {groupTasks.Count:N0} already up to date — skipping metadata fetch for those");
                    }
                    groupTasks = toProcess;
                }
                if (groupTasks.Count == 0) continue; // whole drive group already copied

                // Split into SEQUENTIAL jobs small enough that SharePoint imports them cleanly. Above a
                // ceiling SP fails every SPListItem with "Missing file info for list item" → the
                // 100-failure threshold → whole job cancels.
                //
                // VERIFIED: the ceiling is on the total number of manifest VERSION ENTRIES, NOT file
                // count. A file with N versions contributes ~N <File> entries, so a 120-file batch of
                // 3-version documents (~480 entries) fails while 120 single-version files (~120 entries)
                // succeed. So batch by total version count, not file count — otherwise version-heavy
                // regions (e.g. edited documents) blow the limit at far fewer files.
                //
                // This requires version counts up front: a one-time parallel "analyze" pass (memory-
                // light — counts only). Batches then stay large in single-version regions and shrink
                // automatically where files have many versions.
                // CALIBRATION RESULTS (2026-06-28): the SPMI "Missing file info" ceiling is DATA-DEPENDENT
                // (subfolder depth lowers it). On the deeply-nested SOW set (351 subfolders): 200 and 120
                // BOTH fail every batch; the 590-file SOW subset imported clean at ~100. So 100 is the
                // proven-safe value for deep structures. (120 was fine on flatter data — don't assume a
                // single global number.) Keep at 100 unless re-verified on the exact target structure.
                // The "Missing file info" failures came from the redundant standalone SPListItem objects
                // (now omitted — see MigrationPackageBuilder.EmitStandaloneListItems), which also tripped
                // the 100-error job-cancel threshold and forced tiny batches. With SPFile-only manifests
                // there's no per-item list-item error, so the old 120/133/200 "ceilings" (which were that
                // threshold, not a real SP limit) no longer apply. Raising to 200 to cut the job count for
                // 120k; verify a clean run, then try higher (250) if it holds.
                const int MaxVersionsPerJob = 250;
                const int MaxItemsPerJob    = 250;

                activityLog?.Report($"Analyzing {groupTasks.Count:N0} files for version-aware batching...");
                var sizingProgress = new Progress<int>(d =>
                    activityLog?.Report($"Analyzing files for batching: {d:N0} / {groupTasks.Count:N0}"));
                // Fetch metadata + versions for the whole group ONCE, upfront, into a cache the download
                // producer reuses — so the producer makes no Graph metadata calls during the copy (those
                // per-batch $batch calls were being silently throttled and stalling the pipeline). This
                // 6-wide: leaves more SP API quota headroom for container provisioning immediately after.
                // At 12-wide the analysis exhausts the tenant's rate limit, causing 120s Retry-After
                // waits when the first prep tasks try to provision Azure containers.
                // Versions: Off (maxVersions == 1) → the version list would be sliced to the current
                // version anyway, so skip the /versions sub-request per file (halves the Graph calls).
                var metaCache = await spService.FetchMetadataAndVersionCacheAsync(
                    groupTasks.Select(t => (t.job.SourceDriveId, t.job.SourceItemId)).ToList(),
                    maxConcurrency: 6, progress: sizingProgress,
                    includeVersions: maxVersions != 1, ct: cancellationToken);

                // Any files still missing from the cache after its retry rounds get version-counted as 1
                // for batching, but the producer re-fetches their real versions at download time — a
                // mismatch that could push a batch over the entry ceiling. Surface the count so it's never
                // silent; if non-zero on a failing run, that's the lead to chase.
                int cacheMisses = groupTasks.Count(t => !metaCache.ContainsKey(t.job.SourceItemId));
                if (cacheMisses > 0)
                    activityLog?.Report($"⚠ {cacheMisses:N0} file(s) missing metadata after retries — version counts may be approximate for those");

                // Source folder created/modified metadata (relative path → metadata; "" = library root),
                // so the manifest's SPFolder objects carry real dates/authors instead of a placeholder.
                // One sample file per distinct folder identifies the source folder (its parent). Shared
                // across all batches in this drive group.
                // Key by the slash-trimmed path so it matches the SPFolder `relPath` keys the manifest
                // builder uses (which come from TargetSubFolderPath.Trim('/')).
                var folderMetaInput = groupTasks
                    .Where(t => !string.IsNullOrEmpty(t.job.TargetSubFolderPath))
                    .GroupBy(t => t.job.TargetSubFolderPath!.Trim('/'))
                    .Where(g => g.Key.Length > 0)
                    .Select(g => (folderKey: g.Key, driveId: g.First().job.SourceDriveId, sampleFileItemId: g.First().job.SourceItemId))
                    .ToList();
                // Previously silent — for a deep/wide folder structure this is 2 Graph calls per
                // distinct folder, each subject to Kiota's own throttle retries, and could run for a
                // long time right after the metadata-analysis pass above with zero visible progress,
                // making the app look hung. Report it the same way as that pass.
                if (folderMetaInput.Count > 0)
                    activityLog?.Report($"Fetching metadata for {folderMetaInput.Count:N0} folders...");
                var folderProgress = new Progress<int>(d =>
                {
                    if (folderMetaInput.Count > 20 && (d % 100 == 0 || d == folderMetaInput.Count))
                        activityLog?.Report($"Fetching folder metadata: {d:N0} / {folderMetaInput.Count:N0}");
                });
                var folderMetadata = await spService.FetchFolderMetadataAsync(
                    groupTasks[0].job.SourceDriveId, folderMetaInput, maxConcurrency: 6,
                    progress: folderProgress, ct: cancellationToken);

                int VersionsOf((CopyJob job, CopyResult result) t)
                {
                    if (!metaCache.TryGetValue(t.job.SourceItemId, out var c)) return 1;
                    int raw = Math.Max(1, c.Versions?.Count ?? 1);
                    return maxVersions > 0 ? Math.Min(raw, maxVersions) : raw;
                }

                var batches      = new List<List<(CopyJob job, CopyResult result)>>();
                var currentBatch = new List<(CopyJob job, CopyResult result)>();
                int currentVersions = 0;
                foreach (var t in groupTasks)
                {
                    int v = VersionsOf(t);
                    if (currentBatch.Count > 0 &&
                        (currentVersions + v > MaxVersionsPerJob || currentBatch.Count >= MaxItemsPerJob))
                    {
                        batches.Add(currentBatch);
                        currentBatch = [];
                        currentVersions = 0;
                    }
                    currentBatch.Add(t);
                    currentVersions += v;
                }
                if (currentBatch.Count > 0) batches.Add(currentBatch);

                // Pipeline: a single prep producer (download → encrypt → upload blobs + manifest) feeds
                // a bounded channel; import workers drain it. Prep always overlaps imports, hiding
                // app-side prep time behind SharePoint import time.
                //
                // ── CONCURRENT IMPORTS ────────────────────────────────────────────────────────────
                // Concurrent import jobs DO work (the earlier "concurrency fails" result was confounded
                // by pre-version-aware batching). Measured: 2 and 4 give the SAME wall-clock (4 slightly
                // slower), so ~2 is the practical sweet spot. WHY 4 doesn't help is UNDETERMINED — either
                // SharePoint's active-processing limit (~2 jobs, queueing the rest) OR our SINGLE
                // sequential prep producer (one pipeline can't package fast enough to feed >2 imports).
                // To tell them apart, watch the activity log at =4: if ~4 jobs import concurrently it's
                // prep-bound (→ multiple prep producers would help); if only ~2 progress it's SP-bound.
                const int MaxConcurrentPrep    = 3;
                const int MaxConcurrentImports = 6;

                Task<PreparedBatch> Prep(int idx) => PrepareBatchAsync(
                    batches[idx], overwriteMode, maxVersions, maxParallel,
                    targetSiteUrl, webId, webRelUrl, siteId,
                    libraryServerRelUrl, listId, libraryTitle, cancellationToken,
                    copyCustomColumns, columnMappings, bulkFieldCache,
                    preflightCounter, preflightTotal, preflightProgress,
                    sharedFolderIdCache, sharedExistingByFolder,
                    batchLabel: batches.Count > 1 ? $"{idx + 1}/{batches.Count}" : "",
                    activityLog: activityLog,
                    onFilePacked: onFilePacked,
                    downloadController: downloadController,
                    folderUniqueIdCache: sharedFolderUniqueIdCache,
                    folderMetadata: folderMetadata,
                    metaCache: metaCache);

                // Connect throttle → download-slot step-down now that downloads are about to begin.
                // Analysis throttles (above) are isolated; only download-phase throttles shrink the gate.
                spService.Throttled += onMigrationThrottle;

                // Bounded so prep runs at most MaxConcurrentImports+1 batches ahead — keeps memory and
                // the per-batch SAS tokens bounded (a prepared batch is just blobs-in-Azure + refs).
                var prepChannel = Channel.CreateBounded<PreparedBatch>(
                    new BoundedChannelOptions(MaxConcurrentImports + 1) { FullMode = BoundedChannelFullMode.Wait });

                // Producer: prepare batches sequentially (downloads already share the global gate) and
                // hand each off to the import workers.
                var prepProducer = Task.Run(async () =>
                {
                    try
                    {
                        await Parallel.ForEachAsync(
                            Enumerable.Range(0, batches.Count),
                            new ParallelOptions { MaxDegreeOfParallelism = MaxConcurrentPrep, CancellationToken = cancellationToken },
                            async (idx, ct) =>
                            {
                                var prepared = await Prep(idx);
                                await prepChannel.Writer.WriteAsync(prepared, ct);
                            });
                    }
                    finally { prepChannel.Writer.Complete(); }
                }, cancellationToken);

                // Import workers: each pulls a prepared batch and runs its SP import to completion.
                // With MaxConcurrentImports=1 this is strict sequential import (the proven path).
                var importWorkers = Enumerable.Range(0, MaxConcurrentImports).Select(_ => Task.Run(async () =>
                {
                    await foreach (var prepared in prepChannel.Reader.ReadAllAsync(cancellationToken))
                        await SubmitAndPollBatchAsync(prepared, webId, targetSiteUrl, activityLog, cancellationToken);
                }, cancellationToken)).ToArray();

                await Task.WhenAll(importWorkers);
                await prepProducer;
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                result.Status       = CopyStatus.Failed;
                result.ErrorMessage = "Cancelled";
            }
        }
        catch (Exception ex)
        {
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                result.Status       = CopyStatus.Failed;
                result.ErrorMessage = ex.Message;
            }
        }
        finally
        {
            spService.Throttled -= onMigrationThrottle;
        }
    }

    // A batch whose blobs + manifest have been uploaded and is ready to import. Carries only
    // references (URIs + key); the file payloads live in Azure, so a staged batch is cheap to hold
    // while a previous batch is still importing.
    private sealed record PreparedBatch(
        string BatchLabel,
        IList<(CopyJob job, CopyResult result)> FileTasks,
        int CopyingCount,
        string DataUri,
        string MetadataUri,
        byte[] EncryptionKey,
        // Maps each manifest list-item GUID → its CopyResult, so a per-item SP import error
        // ("Missing file info for list item with id <guid>") can be attributed to the exact file
        // and marked Failed — instead of blanket-marking the whole batch Success.
        IReadOnlyDictionary<string, CopyResult> ListItemMap);

    // Phase 1 of a batch: provision containers, pre-flight, download → encrypt → upload blobs, and
    // upload the manifest. Does NOT submit the import job, so this can safely run (for batch N+1)
    // while a previous batch (N) is still importing — only the import itself must be serialized.
    private async Task<PreparedBatch> PrepareBatchAsync(
        IList<(CopyJob job, CopyResult result)> fileTasks,
        OverwriteMode overwriteMode,
        int maxVersions,
        int maxParallel,
        string targetSiteUrl,
        string webId,
        string webRelUrl,
        string siteId,
        string libraryServerRelUrl,
        string listId,
        string libraryTitle,
        CancellationToken cancellationToken,
        bool copyCustomColumns = false,
        List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null,
        int[]? preflightCounter = null,
        int preflightTotal = 0,
        IProgress<(int, int)>? preflightProgress = null,
        Dictionary<string, string>? folderIdCache = null,
        System.Collections.Concurrent.ConcurrentDictionary<
            string, Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>? prebuiltExistingByFolder = null,
        string batchLabel = "",
        IProgress<string>? activityLog = null,
        IProgress<int>? onFilePacked = null,
        AdaptiveParallelismController? downloadController = null,
        IReadOnlyDictionary<string, (FileMetadata Metadata, List<Microsoft.Graph.Models.DriveItemVersion>? Versions)>? metaCache = null,
        System.Collections.Concurrent.ConcurrentDictionary<string, string>? folderUniqueIdCache = null,
        IReadOnlyDictionary<string, FileMetadata>? folderMetadata = null)
    {
        var pfx = string.IsNullOrEmpty(batchLabel) ? "" : $"[{batchLabel}] ";
        try
        {
            // Step 2: pre-flight — determine what needs copying (and delete stale targets for
            // Overwrite/IfNewer) BEFORE provisioning containers.  Containers are only needed when
            // we actually have files to upload; skipping this for all-skipped batches avoids
            // hundreds of Azure provisioning calls in large "copy-if-newer" runs.
            //
            // Step 2a: use the pre-built folder ID cache passed in from ExecuteAsync, which already
            // created all subfolders (at MaxDegree=4 with segment-cache conflict resolution) and captured
            // the resulting item IDs. This eliminates redundant per-batch folder resolution.
            folderIdCache ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                { [string.Empty] = fileTasks[0].job.TargetParentItemId };

            // Step 2b: get the existing-file listing per target folder.
            // Reuse the pre-built snapshot from ExecuteAsync when available — eliminates redundant
            // per-batch Graph calls when multiple SPMI batches run in parallel.
            System.Collections.Concurrent.ConcurrentDictionary<
                string, Dictionary<string, (string ItemId, DateTimeOffset? Modified)>> existingByFolder;
            if (prebuiltExistingByFolder != null)
            {
                existingByFolder = prebuiltExistingByFolder;
            }
            else
            {
                existingByFolder = new System.Collections.Concurrent.ConcurrentDictionary<
                    string, Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>(
                    StringComparer.OrdinalIgnoreCase);
                await Parallel.ForEachAsync(folderIdCache,
                    new ParallelOptions { MaxDegreeOfParallelism = 8, CancellationToken = cancellationToken },
                    async (kvp, ct) =>
                    {
                        existingByFolder[kvp.Key] = await spService.FetchFolderItemsAsync(
                            fileTasks[0].job.TargetDriveId, kvp.Value);
                    });
            }

            // Step 2c: compare against cached folder contents; SP REST only for zombie detection.
            // Fast path: if every target folder is empty the check is trivially a no-op — all files
            // are new, so skip the per-file parallel scan and advance the preflight counter in bulk.
            var existingFileIds = new System.Collections.Concurrent.ConcurrentDictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            if (existingByFolder.Values.All(v => v.Count == 0))
            {
                if (preflightCounter != null)
                {
                    int done = Interlocked.Add(ref preflightCounter[0], fileTasks.Count);
                    preflightProgress?.Report((done, preflightTotal));
                }
            }
            else
            {
            await Parallel.ForEachAsync(fileTasks,
                new ParallelOptions { MaxDegreeOfParallelism = maxParallel, CancellationToken = cancellationToken },
                async (task, ct) =>
                {
                    var (job, result) = task;
                    var subPath          = job.TargetSubFolderPath ?? string.Empty;
                    var fileServerRelUrl = string.IsNullOrEmpty(subPath)
                        ? $"{libraryServerRelUrl}/{job.SourceName}"
                        : $"{libraryServerRelUrl}/{subPath}/{job.SourceName}";

                    if (overwriteMode == OverwriteMode.Overwrite)
                    {
                        if (existingByFolder[subPath].ContainsKey(job.SourceName))
                        {
                            // Real file: delete first so SPMI does a fresh INSERT.
                            // SPMI UPDATE (with existing GUID) appends imported versions to the
                            // existing version history instead of replacing it, causing duplication.
                            await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                        else if (existingByFolder[subPath].Count > 0 &&
                                 await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) != null)
                        {
                            // Zombie blob (AllDocs row without SPListItem) — delete via deleteObject,
                            // which bypasses the SPListItem requirement that recycleObject needs.
                            // Guard: if the folder is empty, no prior SPMI import ran here so no zombies exist.
                            await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                        existingFileIds[fileServerRelUrl] = null;
                    }
                    else if (overwriteMode == OverwriteMode.IfNewer)
                    {
                        if (existingByFolder[subPath].TryGetValue(job.SourceName, out var existing))
                        {
                            // Date sources in preference order: upfront metadata cache → the scan-captured
                            // date on the job → a per-file Graph call as last resort. VERIFIED (114k-file
                            // run, 2026-07-01): when throttling starved the metadata cache, this fallback
                            // made one live Graph call per file — 109k calls across 442 batches, hours of
                            // wall-clock. The scan-captured date makes the live call unreachable for any
                            // job built by the folder walk.
                            DateTimeOffset? srcModifiedDate;
                            if (metaCache != null && metaCache.TryGetValue(job.SourceItemId, out var ifNewerCached))
                                srcModifiedDate = ifNewerCached.Metadata.ModifiedDateTime;
                            else if (job.SourceModified != null)
                                srcModifiedDate = job.SourceModified;
                            else
                                srcModifiedDate = (await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId)).ModifiedDateTime;
                            if (srcModifiedDate is { } srcModified && srcModified <= existing.Modified)
                            {
                                result.Status       = CopyStatus.Skipped;
                                result.ErrorMessage = CopyResult.UpToDate;
                            }
                            else
                            {
                                // Source is newer — delete so SPMI does a fresh INSERT
                                // (same reasoning as overwrite mode).
                                await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                            }
                        }
                        else if (existingByFolder[subPath].Count > 0 &&
                                 await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) != null)
                        {
                            // Zombie SPFile blob — purge so SPMI can import cleanly.
                            await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                    }
                    else // Skip
                    {
                        if (existingByFolder[subPath].ContainsKey(job.SourceName))
                            result.Status = CopyStatus.Skipped;
                        else if (existingByFolder[subPath].Count > 0 &&
                                 await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) != null)
                        {
                            // Zombie SPFile blob (AllDocs row exists but Graph returns 404).
                            // Guard: only check when the folder has prior Graph-visible files,
                            // meaning a previous import ran and could have left orphaned blobs.
                            await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                    }

                    if (preflightCounter != null)
                    {
                        int done = Interlocked.Increment(ref preflightCounter[0]);
                        preflightProgress?.Report((done, preflightTotal));
                    }
                });
            } // end else (non-empty folders)

            var copyingCount = fileTasks.Count(t => t.result.Status == CopyStatus.Copying);
            var skippedCount = fileTasks.Count - copyingCount;
            activityLog?.Report($"{pfx}Pre-flight: {copyingCount} to copy, {skippedCount} already exist");

            if (copyingCount == 0)
            {
                activityLog?.Report($"{pfx}All files already exist — nothing to copy");
                return new PreparedBatch(batchLabel, fileTasks, 0, string.Empty, string.Empty, Array.Empty<byte>(),
                    new Dictionary<string, CopyResult>());
            }

            // Step 1: provision SP-provided encrypted containers — deferred until after preflight
            // so all-skipped batches never touch Azure at all.
            activityLog?.Report($"{pfx}Provisioning Azure migration containers...");
            var (dataUri, metadataUri, encryptionKey) =
                await spService.ProvisionMigrationContainersAsync(targetSiteUrl);

            System.Diagnostics.Debug.WriteLine($"[Migration] encryptionKey length={encryptionKey.Length}");
            System.Diagnostics.Debug.WriteLine($"[Migration] dataUri prefix={dataUri[..Math.Min(dataUri.Length,80)]}");
            System.Diagnostics.Debug.WriteLine($"[Migration] metaUri prefix={metadataUri[..Math.Min(metadataUri.Length,80)]}");

            // Steps 3+4 interleaved: for each file — download versions, encrypt, upload its
            // blobs, then free the encrypted bytes. Uploading per file instead of buffering
            // the whole batch keeps peak memory at one file's versions rather than the
            // entire library (which OOMs on multi-GB copies).
            // NetworkTimeout extends the per-request timeout from the 60-second Azure default
            // so large files don't cancel mid-upload on slow connections.
            var blobOptions = new BlobClientOptions();
            blobOptions.Retry.NetworkTimeout = TimeSpan.FromMinutes(30);
            var dataClient = new BlobContainerClient(new Uri(dataUri), blobOptions);
            var builder    = new MigrationPackageBuilder(encryptionKey);

            // parallelFileDownloads: how many files can queue simultaneously against the global
            // downloadController gate within one SPMI batch's Parallel.ForEachAsync.
            // versionParallelism=1: the global gate at maxParallel slots already controls total
            // concurrency. vp>1 multiplies Graph calls per slot and causes throttle cascades.
            int parallelFileDownloads = Math.Max(2, Math.Min(10, maxParallel));
            int versionParallelism    = 1;

            // Channel capacity = parallelFileDownloads + 2 so simultaneous chunk completions don't
            // immediately block the writers while the consumer is mid-encrypt.
            var pipe = Channel.CreateBounded<DownloadedFile>(
                new BoundedChannelOptions(parallelFileDownloads + 2) { FullMode = BoundedChannelFullMode.Wait });

            // For small batches log every file; for large ones log milestones only to avoid flooding the feed.
            bool verbosePerFile = copyingCount <= 20;
            int  milestoneStep  = Math.Max(1, copyingCount / 10);

            activityLog?.Report($"{pfx}Downloading {copyingCount:N0} files ({maxParallel} concurrent, {versionParallelism} version stream{(versionParallelism > 1 ? "s" : "")} each)...");

            // Files absent from the upfront cache (metadata fetch failed under throttling) that we copy
            // current-version-only to stay within the batch's entry budget. Reported in the batch summary.
            int currentOnlyMisses = 0;

            var producerTask = Task.Run(async () =>
            {
                try
                {
                    // Metadata + versions were fetched ONCE upfront into metaCache, so the producer makes
                    // ZERO Graph metadata calls here — only content downloads. This removes the per-batch
                    // $batch round-trips that were being silently throttled (Kiota retries the Retry-After
                    // invisibly), which stalled a batch's last chunk and starved the import workers.
                    var copyingTasks = fileTasks.Where(t => t.result.Status == CopyStatus.Copying).ToList();

                    await Parallel.ForEachAsync(copyingTasks,
                        new ParallelOptions { MaxDegreeOfParallelism = parallelFileDownloads, CancellationToken = cancellationToken },
                        async (taskPair, ct) =>
                        {
                            var (job, result) = taskPair;
                            if (result.Status != CopyStatus.Copying) return;

                            var subPath = job.TargetSubFolderPath ?? string.Empty;
                            var fileServerRelUrl = string.IsNullOrEmpty(subPath)
                                ? $"{libraryServerRelUrl}/{job.SourceName}"
                                : $"{libraryServerRelUrl}/{subPath}/{job.SourceName}";
                            existingFileIds.TryGetValue(fileServerRelUrl, out var existingFileId);

                            // Pull metadata/versions from the upfront cache. Single-version files cached
                            // their Versions as null → synthesize the current-version entry here.
                            // CACHE MISS GUARD: batching counted a missing file as 1 version, so we MUST
                            // emit only its current version — otherwise its (uncounted) extra versions would
                            // push the batch past the SPMI entry ceiling and fail the whole import. So a
                            // miss also gets a synthetic single-version entry (current-only); metadata is
                            // left null and fetched individually in DownloadFileDataAsync. This is lossless
                            // for single-version data; for a throttle-missed multi-version file it copies
                            // current-only (counted below and surfaced in the batch summary).
                            FileMetadata? cachedMeta = null;
                            List<Microsoft.Graph.Models.DriveItemVersion>? cachedVersions;
                            if (metaCache != null && metaCache.TryGetValue(job.SourceItemId, out var cached))
                            {
                                cachedMeta     = cached.Metadata;
                                cachedVersions = cached.Versions
                                    ?? new List<Microsoft.Graph.Models.DriveItemVersion> { new Microsoft.Graph.Models.DriveItemVersion() };
                            }
                            else
                            {
                                cachedVersions = new List<Microsoft.Graph.Models.DriveItemVersion> { new Microsoft.Graph.Models.DriveItemVersion() };
                                Interlocked.Increment(ref currentOnlyMisses);
                            }

                            // Acquire a global download slot — limits total concurrent Graph
                            // content downloads across all SPMI batches to maxParallel.
                            if (downloadController != null)
                                await downloadController.WaitAsync(ct);
                            try
                            {
                                if (verbosePerFile) activityLog?.Report($"{pfx}↓ {job.SourceName}");
                                var data = await DownloadFileDataAsync(
                                    job, result, maxVersions, versionParallelism, existingFileId, ct,
                                    cachedMeta, cachedVersions,
                                    copyCustomColumns, columnMappings, bulkFieldCache,
                                    pfxLabel: pfx, activityLog: activityLog);
                                await pipe.Writer.WriteAsync(data, ct);
                            }
                            catch (OperationCanceledException) { throw; }
                            catch (Exception ex)
                            {
                                result.Status       = CopyStatus.Failed;
                                result.ErrorMessage = $"Download failed: {ex.Message}";
                            }
                            finally
                            {
                                downloadController?.Release();
                            }
                        });
                }
                finally { pipe.Writer.Complete(); }
            }, cancellationToken);

            int packedInBatch = 0;
            // Upload blobs in parallel ACROSS files (bounded), not one file at a time. Awaiting each
            // file's upload before pulling the next made the upload stage serial across files — the
            // dominant prep bottleneck for libraries of many small files. Encryption stays sequential
            // in this consumer (builder.Files isn't thread-safe and AES is cheap); only the network
            // upload is parallelized. The gate bounds how many files' encrypted bytes are held at once.
            int uploadConcurrency = Math.Max(2, Math.Min(16, maxParallel));
            using var uploadGate  = new SemaphoreSlim(uploadConcurrency);
            var uploadTasks       = new List<Task>();

            // GUID → CopyResult, built as files are added, so per-item SP import errors can be
            // attributed back to the exact file. Written only in this sequential consumer loop.
            var listItemMap = new Dictionary<string, CopyResult>(StringComparer.OrdinalIgnoreCase);

            await foreach (var data in pipe.Reader.ReadAllAsync(cancellationToken))
            {
                int filesBefore = builder.Files.Count;
                try
                {
                    var versionStreams = data.Slots
                        .Select((ms, i) => (version: data.Versions[i], content: (Stream?)ms))
                        .Where(t => t.content != null)
                        .Select(t => (t.version, content: t.content!))
                        .ToList();

                    await builder.AddFileAsync(data.Job.SourceName, data.FolderRelPath, data.Metadata,
                        versionStreams, data.ExistingFileId, data.CustomFields);

                    // Hand this file's blob upload off to run concurrently with other files' uploads.
                    var entry      = builder.Files[^1];
                    var dataResult = data.Result;
                    listItemMap[entry.ListItemId] = dataResult; // for per-item import-error attribution
                    var versCount  = data.Versions.Count;
                    var fileName   = data.Job.SourceName;
                    await uploadGate.WaitAsync(cancellationToken); // back-pressure: cap in-flight uploads
                    uploadTasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            // Versions within a file upload sequentially; cross-file concurrency comes
                            // from the gate, so total concurrent blob PUTs stay ~uploadConcurrency.
                            foreach (var version in entry.Versions)
                            {
                                var content = version.EncryptedContent
                                    ?? throw new InvalidOperationException("Encrypted content already released.");
                                var ivB64 = Convert.ToBase64String(content[..16]);
                                var opts  = new BlobUploadOptions { Metadata = new Dictionary<string, string> { ["IV"] = ivB64 } };
                                var blob = dataClient.GetBlobClient(version.StreamId);

                                // Azure Storage's own internal retry (ClientOptions.Retry) already exhausts
                                // itself on sustained network blips ("Retry failed after N tries") — retry
                                // again here with backoff rather than failing the file on one bad window.
                                // A fresh MemoryStream is needed each attempt since UploadAsync consumes it.
                                for (int attempt = 0; ; attempt++)
                                {
                                    try
                                    {
                                        using var ms = new MemoryStream(content, 16, content.Length - 16);
                                        await blob.UploadAsync(ms, opts, cancellationToken);
                                        await blob.CreateSnapshotAsync(cancellationToken: cancellationToken);
                                        break;
                                    }
                                    catch (OperationCanceledException) { throw; }
                                    catch (Exception ex) when (attempt < 3 && (ex is IOException || ex is System.Net.Http.HttpRequestException || ex is Azure.RequestFailedException))
                                    {
                                        int waitsecs = (attempt + 1) * 5;
                                        activityLog?.Report($"⚠ {pfx}{fileName} — upload connection reset, retrying in {waitsecs}s ({attempt + 1}/3)");
                                        await Task.Delay(TimeSpan.FromSeconds(waitsecs), cancellationToken);
                                    }
                                }
                                version.EncryptedContent = null;
                            }
                            dataResult.VersionsCopied = versCount;
                            onFilePacked?.Report(1);
                            int n = Interlocked.Increment(ref packedInBatch);
                            if (!verbosePerFile && (n == 1 || n % milestoneStep == 0 || n == copyingCount))
                                activityLog?.Report($"{pfx}{n:N0} / {copyingCount:N0} files packaged");
                        }
                        catch (Exception ex) when (ex is not OperationCanceledException)
                        {
                            entry.Failed = true; // exclude from manifest — its blobs weren't all uploaded
                            dataResult.Status       = CopyStatus.Failed;
                            dataResult.ErrorMessage = $"Upload failed ({fileName}): {ex.Message}";
                        }
                        finally { uploadGate.Release(); }
                    }, cancellationToken));
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    // AddFileAsync (encrypt) failed — the entry is the last added, so RemoveLastFile is
                    // still correct here (this runs serially, before any upload task for this file).
                    if (builder.Files.Count > filesBefore)
                        builder.RemoveLastFile();
                    data.Result.Status       = CopyStatus.Failed;
                    data.Result.ErrorMessage = $"Package build failed: {ex.Message}";
                }
                finally
                {
                    // Encryption is done (or failed) — free the original download buffers now.
                    foreach (var ms in data.Slots) ms?.Dispose();
                }
            }

            await producerTask;
            await Task.WhenAll(uploadTasks); // all blobs uploaded before the manifest is built

            // Step 5: upload manifest XML blobs to the metadata container.
            // Fetch the root folder GUID so the manifest can include an explicit SPFolder entry.
            // This is required for newly created empty libraries — without it SPMI cannot resolve
            // the parent folder for files and fails with "Missing file info for list item".
            var rootFolderGuid = await GetCachedFolderUniqueIdAsync(
                folderUniqueIdCache, targetSiteUrl, libraryServerRelUrl, cancellationToken);
            // A null root-folder GUID makes SPMI reject EVERY list item in the batch with the cryptic
            // "Missing file info for list item". Under heavy throttling the underlying REST call can
            // exhaust its retries and return null — so fail the batch CLEARLY and re-queueably rather
            // than submitting a manifest that's guaranteed to fail 100% with an opaque error.
            if (string.IsNullOrEmpty(rootFolderGuid))
                throw new Exception("Could not resolve target library root folder ID — SharePoint throttling exhausted retries. Batch not submitted; re-run (ideally off-peak).");

            // Resolve the real target GUID of every nested subfolder (and all ancestor folders) so the
            // manifest can declare an SPFolder object for each. SP requires every SPFile to be preceded
            // by its parent SPFolder; for files in subfolders, omitting these objects makes the list
            // item fail to resolve its parent → "Missing file info for list item" for every nested file.
            // The folders already exist on the target (pre-created in ExecuteAsync), so we fetch their
            // actual UniqueIds rather than inventing GUIDs (which would conflict with the live folders).
            var folderGuids   = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var allFolderPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (job, _) in fileTasks)
            {
                var p = job.TargetSubFolderPath?.Trim('/');
                if (string.IsNullOrEmpty(p)) continue;
                var segs = p.Split('/');
                for (int i = 1; i <= segs.Length; i++)
                    allFolderPaths.Add(string.Join('/', segs[..i]));
            }
            if (allFolderPaths.Count > 0)
            {
                var libRel       = libraryServerRelUrl.TrimEnd('/');
                var guidLock     = new object();
                var unresolved   = new List<string>();
                await Parallel.ForEachAsync(allFolderPaths,
                    new ParallelOptions { MaxDegreeOfParallelism = 6, CancellationToken = cancellationToken },
                    async (path, ct) =>
                    {
                        var guid = await GetCachedFolderUniqueIdAsync(folderUniqueIdCache, targetSiteUrl, $"{libRel}/{path}", ct);
                        if (!string.IsNullOrEmpty(guid))
                            lock (guidLock) folderGuids[path] = guid!;
                        else
                            lock (guidLock) unresolved.Add(path);
                    });

                // A subfolder whose GUID we couldn't resolve makes SPMI reject every file under it with
                // "Missing file info for list item". Rather than submit a manifest that's partly broken,
                // fail the whole batch clearly so it can be re-run (off-peak) intact.
                if (unresolved.Count > 0)
                    throw new Exception($"Could not resolve {unresolved.Count} target subfolder ID(s) — SharePoint throttling exhausted retries (e.g. '{unresolved[0]}'). Batch not submitted; re-run (ideally off-peak).");
            }

            // IfNewer behaves like overwrite from SPMI's perspective: stale targets were
            // already deleted in step 2b, so surviving imports must be allowed to replace.
            var spmiOverwrite = overwriteMode != OverwriteMode.Skip;
            System.Diagnostics.Debug.WriteLine(
                $"[Migration] manifest params: siteId={siteId} webId={webId} listId={listId}" +
                $" webRelUrl={webRelUrl} libraryTitle={libraryTitle} libraryServerRelUrl={libraryServerRelUrl}" +
                $" rootFolderGuid={rootFolderGuid ?? "(null)"} overwrite={spmiOverwrite}");

            activityLog?.Report($"{pfx}Building and uploading SPMI manifest...");
            var metadataClient = new BlobContainerClient(new Uri(metadataUri));
            var manifests = builder.BuildManifestXml(
                siteId, webId, listId,
                targetSiteUrl, webRelUrl, libraryTitle, libraryServerRelUrl,
                spmiOverwrite, rootFolderGuid, folderGuids, folderMetadata);

            foreach (var (blobName, data) in manifests)
            {
                cancellationToken.ThrowIfCancellationRequested();
                System.Diagnostics.Debug.WriteLine($"[Migration] uploading manifest blob: {blobName} ({data.Length} bytes)");
                var ivB64m = Convert.ToBase64String(data[..16]);
                var optsM  = new BlobUploadOptions { Metadata = new Dictionary<string, string> { ["IV"] = ivB64m } };
                using var ms = new MemoryStream(data, 16, data.Length - 16);
                var metaBlob = metadataClient.GetBlobClient(blobName);
                await metaBlob.UploadAsync(ms, optsM, cancellationToken);
                await metaBlob.CreateSnapshotAsync(cancellationToken: cancellationToken);
            }

            if (currentOnlyMisses > 0)
                activityLog?.Report($"⚠ {pfx}{currentOnlyMisses:N0} file(s) copied current-version-only (metadata unavailable under throttling; lossless for single-version files)");
            activityLog?.Report($"{pfx}Packaged {copyingCount:N0} file{(copyingCount == 1 ? "" : "s")} — ready to import");
            return new PreparedBatch(batchLabel, fileTasks, copyingCount, dataUri, metadataUri, encryptionKey, listItemMap);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                result.Status       = CopyStatus.Failed;
                result.ErrorMessage = "Cancelled";
            }
        }
        catch (Exception ex)
        {
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                result.Status       = CopyStatus.Failed;
                result.ErrorMessage = ex.Message;
            }
        }
        // Prep failed (or was cancelled): return a sentinel with nothing to submit so the pipeline
        // skips the import for this batch and continues with the rest.
        return new PreparedBatch(batchLabel, fileTasks, 0, string.Empty, string.Empty, Array.Empty<byte>(),
            new Dictionary<string, CopyResult>());
    }

    // Phase 2 of a batch: submit the import job and poll to completion. This is the ONLY phase that
    // must be serialized across batches — SharePoint soft-cancels concurrent import jobs. No-op when
    // the prepared batch has nothing to import (all files skipped, or prep failed).
    private async Task SubmitAndPollBatchAsync(
        PreparedBatch batch, string webId, string targetSiteUrl,
        IProgress<string>? activityLog, CancellationToken cancellationToken)
    {
        if (batch.CopyingCount == 0) return;

        var pfx       = string.IsNullOrEmpty(batch.BatchLabel) ? "" : $"[{batch.BatchLabel}] ";
        var fileTasks = batch.FileTasks;
        try
        {
            var metadataClient = new BlobContainerClient(new Uri(batch.MetadataUri));

            // Step 6: submit the migration job
            var jobId = await spService.CreateMigrationJobEncryptedAsync(
                targetSiteUrl, webId, batch.DataUri, batch.MetadataUri, batch.EncryptionKey);
            System.Diagnostics.Debug.WriteLine($"[Migration] submitted job: {jobId}");
            activityLog?.Report($"{pfx}Submitted to SharePoint — waiting for import...");

            // Step 7: poll until JobEnd. Attribute each per-item error to its exact file via the
            // list-item GUID in the message, so we can mark only the truly-failed files (not the
            // whole batch) and report an honest imported/failed count.
            bool   fatal          = false;
            string? fatalMsg      = null;
            bool   seenJobStart   = false;
            int    liveErrorCount = 0; // per-item errors surfaced live so failures show as they happen
            int    totalErrorsReported = 0;
            await foreach (var evt in spService.PollMigrationJobAsync(targetSiteUrl, jobId, cancellationToken))
            {
                if (evt.TryGetProperty("Event", out var evtName))
                {
                    var name = evtName.GetString();
                    if (name == "JobStart" && !seenJobStart)
                    {
                        seenJobStart = true;
                        activityLog?.Report($"{pfx}SharePoint import started");
                    }
                    else if (name == "JobProgress")
                    {
                        if (evt.TryGetProperty("ObjectsProcessed", out var proc) &&
                            proc.ValueKind == System.Text.Json.JsonValueKind.Number)
                            activityLog?.Report($"{pfx}SP importing: {proc.GetInt32():N0} / {batch.CopyingCount:N0} files");
                    }
                    if (name == "JobEnd")
                    {
                        if (evt.TryGetProperty("TotalErrors", out var te) &&
                            te.ValueKind == JsonValueKind.Number)
                            totalErrorsReported = te.GetInt32();
                        break;
                    }
                    if (name == "JobFatalError")
                    {
                        fatalMsg = evt.TryGetProperty("Message", out var m) ? m.GetString() : "Unknown error";
                        activityLog?.Report($"⚠ {pfx}Fatal error: {fatalMsg}");
                        // Log the full SP event — may contain an error code or richer reason
                        // (e.g. "Operation canceled" can mean concurrent-job limit, SAS expiry, bad manifest).
                        activityLog?.Report($"  SP event: {evt}");
                        fatal = true;
                    }
                    else if (name == "JobError")
                    {
                        var msg = evt.TryGetProperty("Message", out var m) ? m.GetString() : "Unknown error";
                        liveErrorCount++;
                        // Attribute to the exact file when the message carries its list-item GUID, and
                        // mark that file Failed so it shows in the Failed filter and can be re-copied.
                        var guid = ExtractGuid(msg);
                        if (guid != null && batch.ListItemMap.TryGetValue(guid, out var failedRes)
                            && failedRes.Status == CopyStatus.Copying)
                        {
                            failedRes.Status       = CopyStatus.Failed;
                            failedRes.ErrorMessage = msg;
                        }
                        // Surface the first handful LIVE (as SP reports them) so a failing batch is
                        // visible immediately; suppress the rest to avoid flooding the feed.
                        if (liveErrorCount <= 5)
                            activityLog?.Report($"⚠ {pfx}import error {liveErrorCount}: {msg}");
                        else if (liveErrorCount == 6)
                            activityLog?.Report($"⚠ {pfx}…more errors this batch (suppressing; see batch summary)");
                        System.Diagnostics.Debug.WriteLine($"[Migration] JobError #{liveErrorCount}: {msg}");
                    }
                }
            }

            // Step 8: mark results accurately. Files already Failed (attributed above, or upload failures)
            // stay Failed; on a fatal abort every not-yet-succeeded file fails; otherwise the rest succeeded.
            foreach (var (_, result) in fileTasks)
            {
                if (result.Status != CopyStatus.Copying) continue; // already Failed/Skipped
                if (fatal)
                {
                    result.Status       = CopyStatus.Failed;
                    result.ErrorMessage ??= fatalMsg ?? "Migration job fatal error";
                }
                else
                {
                    result.Status = CopyStatus.Success;
                }
            }

            int failedCount   = fileTasks.Count(t => t.result.Status == CopyStatus.Failed);
            int importedCount = Math.Max(0, batch.CopyingCount - failedCount);

            if (fatal)
            {
                activityLog?.Report($"✗ {pfx}Import FAILED (fatal abort) — {importedCount} of {batch.CopyingCount} imported, {failedCount} failed");
                await TryLogMigrationReportAsync(metadataClient, jobId, batch.EncryptionKey, activityLog, pfx);
            }
            else if (failedCount > 0 || liveErrorCount > 0 || totalErrorsReported > 0)
            {
                // Reconcile counts: if SP/JobError reported more errors than we could attribute to a
                // specific GUID, surface the discrepancy so a silent shortfall is never hidden.
                int reported = Math.Max(failedCount, Math.Max(liveErrorCount, totalErrorsReported));
                var extra = reported > failedCount ? $" ({reported} errors reported, {failedCount} attributed)" : "";
                activityLog?.Report($"⚠ {pfx}Import finished with errors — {importedCount} of {batch.CopyingCount} imported, {failedCount} failed{extra}");
                await TryLogMigrationReportAsync(metadataClient, jobId, batch.EncryptionKey, activityLog, pfx);
            }
            else
            {
                activityLog?.Report($"{pfx}✓ Import complete — {importedCount} file{(importedCount == 1 ? "" : "s")} imported");
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                result.Status       = CopyStatus.Failed;
                result.ErrorMessage = "Cancelled";
            }
        }
        catch (Exception ex)
        {
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                result.Status       = CopyStatus.Failed;
                result.ErrorMessage = ex.Message;
            }
        }
    }

    private sealed record DownloadedFile(
        CopyJob Job,
        CopyResult Result,
        string FolderRelPath,
        string? ExistingFileId,
        FileMetadata Metadata,
        List<Microsoft.Graph.Models.DriveItemVersion> Versions,
        MemoryStream?[] Slots,
        Dictionary<string, string>? CustomFields);

    private async Task<DownloadedFile> DownloadFileDataAsync(
        CopyJob job,
        CopyResult result,
        int maxVersions,
        int maxParallel,
        string? existingFileId,
        CancellationToken ct,
        FileMetadata? prefetchedMetadata = null,
        List<Microsoft.Graph.Models.DriveItemVersion>? prefetchedVersions = null,
        bool copyCustomColumns = false,
        List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null,
        string pfxLabel = "",
        IProgress<string>? activityLog = null)
    {
        FileMetadata metadata;
        List<Microsoft.Graph.Models.DriveItemVersion> allVersions;

        // Metadata and versions are resolved INDEPENDENTLY: the producer may supply versions (even a
        // synthetic current-only list for a cache-miss file, to keep the batch within its entry budget)
        // while leaving metadata null so it's fetched here. Only fetch what wasn't supplied.
        if (prefetchedMetadata != null && prefetchedVersions != null)
        {
            metadata    = prefetchedMetadata;
            allVersions = prefetchedVersions;
        }
        else if (prefetchedVersions != null)
        {
            // Versions supplied (e.g. current-only for a cache miss); fetch just the metadata.
            metadata    = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
            allVersions = prefetchedVersions;
        }
        else
        {
            // Nothing prefetched — fall back to individual Graph calls for both.
            var metaTask     = spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
            var versionsTask = spService.GetVersionsAsync(job.SourceDriveId, job.SourceItemId);
            await Task.WhenAll(metaTask, versionsTask);
            metadata    = metaTask.Result;
            allVersions = versionsTask.Result;
        }
        var versions    = maxVersions > 0 && allVersions.Count > maxVersions
            ? allVersions.TakeLast(maxVersions).ToList()
            : allVersions;

        result.VersionsTotal = versions.Count;

        // Download all version content concurrently, buffering each stream immediately so the
        // HTTP connection is consumed before it can go stale. Index-keyed array preserves order.
        var slots = new MemoryStream?[versions.Count];
        await Parallel.ForEachAsync(
            Enumerable.Range(0, versions.Count),
            new ParallelOptions { MaxDegreeOfParallelism = maxParallel, CancellationToken = ct },
            async (idx, _) =>
            {
                var version = versions[idx];
                bool isLast = idx == versions.Count - 1;
                // Historical (non-current) versions need their own id to download. The current version
                // (isLast) is fetched from the file's current content and needs no id — so a null id is
                // fine there. This is the case for the synthetic single-version entry produced when the
                // /versions list is skipped for a known single-version file; skipping it here (as the old
                // unconditional guard did) left its content slot null → empty Versions → IndexOutOfRange.
                if (!isLast && version.Id == null) return;
                var ms = new MemoryStream();
                for (int attempt = 0; ; attempt++)
                {
                    ms.SetLength(0);
                    try
                    {
                        var networkStream = isLast
                            ? await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId)
                            : await spService.DownloadVersionAsync(job.SourceDriveId, job.SourceItemId, version.Id!);
                        await using (networkStream)
                            await networkStream.CopyToAsync(ms, ct);
                        ms.Position = 0;
                        slots[idx] = ms;
                        break;
                    }
                    catch (OperationCanceledException) { throw; }
                    // HTTP/2 stream resets surface as HttpRequestException, not IOException — must be
                    // caught alongside it or a single mid-stream RST_STREAM fails the file outright.
                    catch (Exception ex) when (attempt < 3 && (ex is System.IO.IOException || ex is System.Net.Http.HttpRequestException))
                    {
                        int waitsecs = (attempt + 1) * 5;
                        activityLog?.Report($"⚠ {pfxLabel}{job.SourceName} — connection reset, retrying in {waitsecs}s ({attempt + 1}/3)");
                        await Task.Delay(TimeSpan.FromSeconds(waitsecs), ct);
                    }
                }
            });

        // Custom field lookup from bulk cache.
        Dictionary<string, string>? customFields = null;
        if (copyCustomColumns && bulkFieldCache != null && columnMappings != null)
        {
            var spIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
            if (spIds.HasValue && bulkFieldCache.TryGetValue($"{spIds.Value.listId}:{spIds.Value.listItemId}", out var rawFields))
            {
                var mappingLookup = ColumnMapping.BuildTargetNameMap(columnMappings);
                customFields = new Dictionary<string, string>();
                foreach (var (srcName, value) in rawFields)
                {
                    if (value == null) continue;
                    string targetName;
                    if (mappingLookup.TryGetValue(srcName, out var mapped))
                    {
                        if (mapped == null) continue;
                        targetName = mapped;
                    }
                    else
                    {
                        targetName = srcName;
                    }
                    // SPMI lookup-value encoding: "-1;#value" asks SP to resolve at import.
                    customFields[targetName] = value switch
                    {
                        PersonFieldValue p   => string.Join(";#", p.Logins.Select(l => $"-1;#{l}")),
                        TaxonomyFieldValue t => string.Join(";#", t.Terms.Select(x => $"-1;#{x.Label}|{x.TermGuid}")),
                        LookupFieldValue l   => string.Join(";#", l.Entries.Select(e => $"-1;#{e.DisplayValue}")),
                        _ => value.ToString() ?? "",
                    };
                }
            }
        }

        var folderRelPath = string.IsNullOrEmpty(job.TargetSubFolderPath)
            ? ""
            : job.TargetSubFolderPath.TrimStart('/');

        return new DownloadedFile(job, result, folderRelPath, existingFileId,
            metadata, versions, slots, customFields);
    }

    // Resolves a folder's SharePoint UniqueId, retrying when the call comes back empty. The underlying
    // REST sender already retries 429s a few times, but under heavy sustained throttling it can exhaust
    // those and return null — and a null folder GUID silently breaks the manifest's file→folder linkage
    // (SPMI then rejects every affected list item with "Missing file info"). These GUIDs are
    // manifest-critical, so retry harder (riding out throttle windows) before giving up.
    private async Task<string?> ResolveFolderUniqueIdResilientAsync(
        string siteUrl, string folderServerRelativeUrl, CancellationToken ct)
    {
        // The folder was already provisioned, so a null result here is a transient throttle, not a real
        // 404 — keep retrying (with backoff, on top of the REST sender's own 429 waits) until it returns.
        const int maxAttempts = 10;
        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
            ct.ThrowIfCancellationRequested();
            var guid = await spService.GetLibraryRootFolderUniqueIdAsync(siteUrl, folderServerRelativeUrl);
            if (!string.IsNullOrEmpty(guid)) return guid;
            if (attempt < maxAttempts - 1)
                await Task.Delay(TimeSpan.FromSeconds(Math.Min(30, 5 * (attempt + 1))), ct);
        }
        return null;
    }

    // Resolves a folder's UniqueId through a cross-batch cache so each folder is resolved at most once
    // for the whole copy (the deep folder tree repeats across all 40+ batches; re-resolving it per batch
    // multiplied the throttle exposure). cacheKey is the folder's server-relative URL.
    private async Task<string?> GetCachedFolderUniqueIdAsync(
        System.Collections.Concurrent.ConcurrentDictionary<string, string>? cache,
        string siteUrl, string folderServerRelativeUrl, CancellationToken ct)
    {
        if (cache != null && cache.TryGetValue(folderServerRelativeUrl, out var cached))
            return cached;
        var guid = await ResolveFolderUniqueIdResilientAsync(siteUrl, folderServerRelativeUrl, ct);
        if (!string.IsNullOrEmpty(guid) && cache != null)
            cache[folderServerRelativeUrl] = guid!;
        return guid;
    }

    // Matches a standard GUID anywhere in a string (e.g. the list-item id in an SP import error).
    private static readonly System.Text.RegularExpressions.Regex GuidRegex = new(
        @"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
        System.Text.RegularExpressions.RegexOptions.Compiled);

    private static string? ExtractGuid(string? text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        var m = GuidRegex.Match(text);
        return m.Success ? m.Value : null;
    }

    private static async Task TryLogMigrationReportAsync(
        Azure.Storage.Blobs.BlobContainerClient metadataClient, string jobId, byte[] key,
        IProgress<string>? activityLog = null, string? pfx = null)
    {
        foreach (var suffix in new[] { ".err", ".log" })
        {
            var name = $"Import-{jobId}-1{suffix}";
            try
            {
                var blob = metadataClient.GetBlobClient(name);

                byte[]? iv = null;
                try
                {
                    var props = await blob.GetPropertiesAsync();
                    if (props.Value.Metadata.TryGetValue("IV", out var ivB64) && !string.IsNullOrEmpty(ivB64))
                        iv = Convert.FromBase64String(ivB64);
                }
                catch (Azure.RequestFailedException ex) when (ex.Status == 404)
                {
                    // Blob not written by SP — skip rather than falling through to DownloadToAsync
                    System.Diagnostics.Debug.WriteLine($"[SP-{suffix[1..]}] {name} not present (SP did not write it)");
                    continue;
                }
                catch { /* 403: metadata read not permitted by SAS; still attempt download */ }

                using var ms = new MemoryStream();
                await blob.DownloadToAsync(ms);
                var cipherBytes = ms.ToArray();

                byte[] plain;
                if (iv != null)
                {
                    plain = AesDecrypt(cipherBytes, key, iv);
                }
                else if (cipherBytes.Length > 16)
                {
                    plain = AesDecrypt(cipherBytes[16..], key, cipherBytes[..16]);
                }
                else
                {
                    plain = cipherBytes;
                }

                var reportText = System.Text.Encoding.UTF8.GetString(plain);
                System.Diagnostics.Debug.WriteLine($"[SP-{suffix[1..]}] {reportText}");

                // Surface the first few lines of SP's .err report into the activity feed so the actual
                // failure reason is visible to the user immediately (not just in VS Debug Output).
                if (activityLog != null && suffix == ".err")
                {
                    var lines = reportText.Split('\n')
                        .Select(l => l.Trim()).Where(l => l.Length > 0).Take(5).ToList();
                    foreach (var line in lines)
                        activityLog.Report($"   {pfx}SP: {line}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SP-{suffix[1..]}] Cannot read {name}: {ex.Message}");
            }
        }
    }

    private static byte[] AesDecrypt(byte[] ciphertext, byte[] key, byte[] iv)
    {
        using var aes = System.Security.Cryptography.Aes.Create();
        aes.Key     = key;
        aes.IV      = iv;
        aes.Mode    = System.Security.Cryptography.CipherMode.CBC;
        aes.Padding = System.Security.Cryptography.PaddingMode.PKCS7;
        using var output = new MemoryStream();
        using var cs = new System.Security.Cryptography.CryptoStream(
            new MemoryStream(ciphertext), aes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Read);
        cs.CopyTo(output);
        return output.ToArray();
    }
}
