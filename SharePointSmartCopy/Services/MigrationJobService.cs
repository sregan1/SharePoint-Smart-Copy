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
    // Files at or above this size get funneled through largeFileGate below, on top of the normal
    // maxParallel download slots. Each in-flight download holds the WHOLE file in memory twice —
    // once as the raw MemoryStream, again as the AES-encrypted byte[] handed to the uploader (see
    // DownloadFileDataAsync / AddFileAsync) — so several multi-GB files downloading at once can
    // exhaust the process working set. Observed 2026-07-02: OutOfMemoryException on .fastq/.czi
    // files (multi-GB genomics/microscopy) mid-download and mid-package-build on a 5,000-file run.
    private const long LargeFileThresholdBytes = 500L * 1024 * 1024; // 500 MB
    // Only this many large files may be IN MEMORY concurrently — the slot is held from download
    // start until the file's encrypted blobs finish uploading (see DownloadedFile.HoldsLargeSlot),
    // independent of maxParallel, so oversized files can't stack up multiple full in-memory copies
    // at once anywhere in the pipeline. Small files are unaffected — they never touch this gate.
    private const int MaxConcurrentLargeFiles = 2;

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
        using var largeFileGate = new SemaphoreSlim(MaxConcurrentLargeFiles);
        // Global BYTE budget for in-flight file payloads, shared across every concurrent batch —
        // see TransferMemoryBudget for why the count-based gates alone aren't enough (a library of
        // ~300-490 MB files slides under the large-file threshold and the per-batch pipe/upload
        // queues stack up 15+ GB of live buffers; observed 2026-07-18, 19 GB heap on a 32 GB
        // machine at Parallel Copies 5). Each file charges ~2× its total version bytes (raw
        // download + AES-encrypted copy coexist) from download start until its blobs finish
        // uploading. Sized to ~40% of machine RAM, clamped to [2 GB, 16 GB].
        //
        // This budget — not the download controller's slot count — was the real throughput limit
        // on the large-file run: at 8 GB with the 2× charge only ~4 multi-hundred-MB scans could be
        // in flight, far below the 16 download slots, yielding ~1 file/min (2026-07-18). Raising it
        // lets the download controller actually use its slots (up to where Graph throttling becomes
        // the limit). Safe to raise now that the smaller byte-per-job budget makes the batch-boundary
        // LOH compaction run regularly, so freed buffers return to the OS instead of piling into a
        // fragmented multi-GB floor. Still bounded: 40% of RAM leaves the rest for the OS, the .NET
        // runtime, WPF, and in-flight garbage between compactions.
        long machineRamBytes = GC.GetGCMemoryInfo().TotalAvailableMemoryBytes;
        var memoryBudget = new TransferMemoryBudget(
            Math.Clamp((long)(machineRamBytes * 0.40), 2L * 1024 * 1024 * 1024, 16L * 1024 * 1024 * 1024));
        activityLog?.Report(
            $"Transfer memory budget: {memoryBudget.Capacity / (1024.0 * 1024 * 1024):F1} GB (bounds concurrent file buffers)");
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

        // LimitChanged above only fires when the adaptive gate's width actually changes — a throttle
        // landing inside the StepDown cooldown window (or one Kiota's own RetryHandler quietly retries
        // before GraphThrottleNotifyHandler reports it) produces no log line at all, showing up only
        // as an unexplained stall. Log every throttle event directly, same pattern as CopyService's
        // onThrottleLog, rate-limited so a throttle storm doesn't flood the feed.
        Action<TimeSpan, int, int, string?>? onThrottleLog = null;
        if (activityLog != null)
        {
            var throttleLogLock = new object();
            var lastThrottleLog = DateTimeOffset.MinValue;
            onThrottleLog = (delay, attempt, max, reason) =>
            {
                lock (throttleLogLock)
                {
                    var now = DateTimeOffset.UtcNow;
                    if (now - lastThrottleLog < TimeSpan.FromSeconds(5)) return;
                    lastThrottleLog = now;
                }
                activityLog.Report($"⚠ Graph throttled — waiting {delay.TotalSeconds:0}s"
                    + (string.IsNullOrEmpty(reason) ? "" : $" [{reason}]"));
            };
            spService.Throttled += onThrottleLog;
        }

        // Periodic resource snapshot so a UCEERR_RENDERTHREADFAILURE crash (a native WPF failure
        // that bypasses every managed exception handler) still leaves a trail in the activity log
        // of what the process looked like in the minutes before it died. See ProcessDiagnostics.
        using var heartbeatCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var heartbeatTask = activityLog is null
            ? Task.CompletedTask
            : RunHeartbeatAsync(activityLog, heartbeatCts.Token);

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
                // Full pre-skip task list: folder metadata (dates + Author/Editor) must cover EVERY
                // folder in the selection, not just folders of files that still need copying —
                // otherwise an up-to-date re-run can never repair folder metadata (the exact
                // scenario after a completed 123k-file migration whose folder people fields are
                // wrong: the fix must not require re-transferring anything).
                var allGroupTasks = groupTasks;
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

                // The Graph-based "is the target empty?" fast path was REMOVED here. Graph's folder
                // listing lags behind SharePoint's content database after an SPMI import — for a while
                // after a migration writes files, Graph still reports the folder empty. On an overwrite
                // re-run that false "empty" skipped the entire existing-file scan, so nothing got
                // deleted and every file bounced off SPMI as "already exists" (reproduced 2026-07-02:
                // fresh target imported 100%, an immediate second pass failed 100%). Always run the scan
                // below — it reads the content DB via SP REST (no lag) and is cheap on empty folders.
                var targetIsEmpty = false;

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
                                    // Graph + SP REST merged: REST sees leftovers from aborted SPMI
                                    // imports that Graph omits (see FetchTargetFolderSnapshotAsync).
                                    var folderRelUrl = string.IsNullOrEmpty(kvp.Key)
                                        ? libraryServerRelUrl
                                        : $"{libraryServerRelUrl}/{kvp.Key}";
                                    sharedExistingByFolder[kvp.Key] = await FetchTargetFolderSnapshotAsync(
                                        firstJob.TargetDriveId, kvp.Value, targetSiteUrl, folderRelUrl);
                                }
                                finally { analysisController.Release(); }
                            });

                        // Diagnostic tally: how many existing files the scan actually found, and via
                        // which source. REST-supplied entries carry an empty ItemId (Graph entries carry
                        // a real one), so we can split the two without extra bookkeeping. If this reports
                        // 0 on an overwrite re-run whose target clearly has files, the scan itself is the
                        // fault (not the delete) — and the Graph/REST split says which listing went blind.
                        int totalExisting = sharedExistingByFolder.Values.Sum(d => d.Count);
                        int restOnly      = sharedExistingByFolder.Values.Sum(d => d.Count(e => string.IsNullOrEmpty(e.Value.ItemId)));
                        activityLog?.Report(
                            $"Existing-file scan: {totalExisting:N0} file(s) already in target " +
                            $"(Graph saw {totalExisting - restOnly:N0}, REST-only {restOnly:N0}).");
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
                            if (srcModified.HasValue && t.TargetModified.HasValue &&
                                TimestampComparer.IsUpToDate(srcModified.Value, t.TargetModified.Value))
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
                // Source folder created/modified metadata (relative path → metadata; "" = library root),
                // so the manifest's SPFolder objects carry real dates/authors instead of a placeholder,
                // and the post-import REST correction pass can set folder Author/Editor (which SPMI's
                // <Folder> element does not honor). Built from allGroupTasks — the FULL pre-skip
                // selection — not the needs-copy subset: folder metadata repair must cover every
                // folder even on a fully-up-to-date re-run where nothing transfers (the repair path
                // for an already-completed migration with wrong folder people fields).
                // Key by the slash-trimmed path so it matches the SPFolder `relPath` keys the manifest
                // builder uses (which come from TargetSubFolderPath.Trim('/')).
                var directFolderGroups = allGroupTasks
                    .Where(t => !string.IsNullOrEmpty(t.job.TargetSubFolderPath))
                    .GroupBy(t => t.job.TargetSubFolderPath!.Trim('/'), StringComparer.OrdinalIgnoreCase)
                    .Where(g => g.Key.Length > 0)
                    .ToDictionary(g => g.Key,
                        g => (driveId: g.First().job.SourceDriveId, sampleFileItemId: g.First().job.SourceItemId),
                        StringComparer.OrdinalIgnoreCase);

                // BUG FIX (2026-07-07): a folder containing ONLY subfolders — no file directly inside
                // it — never appeared as a key above, so it silently kept MigrationPackageBuilder's
                // hardcoded placeholder date. This surfaced as the two largest folders in a real run
                // (114k and 5k files) showing a bogus 2000-01-01 date on the folder itself, because both
                // organize everything into dated/categorized subfolders with nothing loose at their own
                // top level — large, well-organized trees hit this far more often than small flat ones.
                // Every ancestor segment of every file's path needs an entry (same expansion used for
                // folderGuids below), not just paths that happen to hold a file directly. For an ancestor
                // with no direct files, borrow the shallowest descendant that HAS one and walk up the
                // extra levels via repeated parentReference hops — still ID-based, so folder/file names
                // containing '#'/'%'/'+' are unaffected (unlike path-based Graph addressing elsewhere in
                // this codebase, which specifically has to avoid those).
                var folderMetaInput = new List<(string folderKey, string driveId, string sampleFileItemId, int hopsUp)>();
                var allAncestorFolderPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var key in directFolderGroups.Keys)
                {
                    var segs = key.Split('/');
                    for (int i = 1; i <= segs.Length; i++)
                        allAncestorFolderPaths.Add(string.Join('/', segs[..i]));
                }
                foreach (var path in allAncestorFolderPaths)
                {
                    if (directFolderGroups.TryGetValue(path, out var direct))
                    {
                        folderMetaInput.Add((path, direct.driveId, direct.sampleFileItemId, 0));
                        continue;
                    }
                    var pathDepth = path.Count(c => c == '/') + 1;
                    var prefix    = path + "/";
                    int bestDepth = int.MaxValue;
                    (string driveId, string sampleFileItemId)? best = null;
                    foreach (var (leafPath, leafInfo) in directFolderGroups)
                    {
                        if (!leafPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                        var leafDepth = leafPath.Count(c => c == '/') + 1;
                        if (leafDepth < bestDepth) { bestDepth = leafDepth; best = leafInfo; }
                    }
                    // Every ancestor path was derived FROM some direct-file leaf above, so a matching
                    // descendant always exists — best is never null here, but guard defensively anyway.
                    if (best.HasValue)
                        folderMetaInput.Add((path, best.Value.driveId, best.Value.sampleFileItemId, bestDepth - pathDepth));
                }
                // Previously silent — for a deep/wide folder structure this is 2+ Graph calls per
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
                    allGroupTasks[0].job.SourceDriveId, folderMetaInput, maxConcurrency: 6,
                    progress: folderProgress, ct: cancellationToken);
                // Surfaced so a repeat of the "silently placeholder-dated folder" symptom is visible
                // immediately in the log instead of only being discovered by eye in SharePoint later.
                int foldersMissingMetadata = folderMetaInput.Count - folderMetadata.Count;
                if (foldersMissingMetadata > 0)
                    activityLog?.Report($"⚠ {foldersMissingMetadata:N0} folder(s) could not be dated (Graph lookup failed after retries) — they will show a 2000-01-01 placeholder date instead of their real source date");

                // Post-import correction: SPMI's <Folder> manifest element doesn't honor
                // Author/ModifiedBy (confirmed 2026-07-08 — even a brand-new folder creation still
                // attributes to the importing account), unlike <File>, which sets both correctly via
                // the same GetUserId/UserGroup.xml mechanism. Patch every folder we have source
                // metadata for directly via REST — the WHOLE selection (root and every nested level),
                // regardless of which batch (or whether ANY batch) touched which folder. Re-sends
                // date alongside Author/Editor in the same call (see PatchFolderMetadataAsync) since
                // a separate author-only write would otherwise reset Modified to "now".
                async Task CorrectFolderMetadataAsync()
                {
                    var foldersNeedingAuthorFix = folderMetadata
                        .Where(kv => !string.IsNullOrEmpty(kv.Value.CreatedByEmail) || !string.IsNullOrEmpty(kv.Value.ModifiedByEmail))
                        .ToList();
                    // Distinguish "we never had a source email to try" (folderMetadata's CreatedByEmail/
                    // ModifiedByEmail came back null — GetIdentityEmail found no email/UPN on the Graph
                    // identity for that folder) from "we tried and SharePoint rejected it" — these look
                    // identical to the user (folder still shows the importing account) but have completely
                    // different causes and fixes, and conflating them was making this hard to diagnose.
                    int foldersMissingSourceEmail = folderMetadata.Count - foldersNeedingAuthorFix.Count;
                    if (foldersMissingSourceEmail > 0)
                        activityLog?.Report($"⚠ {foldersMissingSourceEmail:N0} folder(s) had no source Author/Modified-By email available — cannot correct those, they will show the importing account");
                    if (foldersNeedingAuthorFix.Count == 0) return;

                    activityLog?.Report($"Correcting folder metadata (dates + authorship) for {foldersNeedingAuthorFix.Count:N0} of {folderMetadata.Count:N0} folder(s)...");
                    int authorFixFailures = 0;
                    int authorFixDone = 0;
                    var sampleErrors = new System.Collections.Concurrent.ConcurrentBag<string>();
                    await Parallel.ForEachAsync(foldersNeedingAuthorFix,
                        new ParallelOptions { MaxDegreeOfParallelism = 6, CancellationToken = cancellationToken },
                        async (kv, ct) =>
                        {
                            var (relKey, meta) = kv;
                            var folderRelUrl = string.IsNullOrEmpty(relKey) ? libraryServerRelUrl : $"{libraryServerRelUrl}/{relKey}";
                            var guid = await GetCachedFolderUniqueIdAsync(sharedFolderUniqueIdCache, targetSiteUrl, folderRelUrl, ct);
                            if (string.IsNullOrEmpty(guid))
                            {
                                Interlocked.Increment(ref authorFixFailures);
                                sampleErrors.Add($"{(relKey.Length == 0 ? "(library root)" : relKey)}: could not resolve target folder GUID");
                            }
                            else
                            {
                                // Date and author go in the SAME call: a separate author-only write
                                // after SPMI already set the correct date would bump Modified back to
                                // "now" as a side effect (see PatchFolderMetadataAsync's doc comment).
                                var err = await spService.PatchFolderMetadataAsync(
                                    targetSiteUrl, listId, guid!, meta.CreatedDateTime, meta.ModifiedDateTime,
                                    meta.CreatedByEmail, meta.ModifiedByEmail);
                                if (err != null)
                                {
                                    Interlocked.Increment(ref authorFixFailures);
                                    sampleErrors.Add($"{(relKey.Length == 0 ? "(library root)" : relKey)}: {err}");
                                }
                            }
                            // Previously silent for the whole pass — on a large tree under sustained
                            // throttling this ran for tens of minutes with zero visible progress,
                            // indistinguishable from a hang (a real run: 39 minutes of nothing but
                            // throttle-wait lines). Mirrors the "Fetching folder metadata: N / M"
                            // reporting already used for the fetch phase just before this one.
                            int done = Interlocked.Increment(ref authorFixDone);
                            if (foldersNeedingAuthorFix.Count > 20 && (done % 100 == 0 || done == foldersNeedingAuthorFix.Count))
                                activityLog?.Report($"Correcting folder metadata: {done:N0} / {foldersNeedingAuthorFix.Count:N0}");
                        });
                    // Never silent: a failure here leaves that folder attributed to the importing
                    // account rather than crashing the run — surfaced with actual error text (not
                    // just a count) so a real cause (e.g. the source's real author/editor not
                    // existing as a resolvable account on the target — expected in a cross-tenant
                    // migration, since SharePoint literally cannot attribute a Person field to an
                    // account that doesn't exist there) is visible immediately instead of triggering
                    // another guess-and-check round.
                    if (authorFixFailures > 0)
                    {
                        var examples = string.Join(" | ", sampleErrors.Distinct().Take(3));
                        activityLog?.Report($"⚠ Could not correct metadata for {authorFixFailures:N0} folder(s). Example(s): {examples}");
                    }
                    else
                    {
                        activityLog?.Report($"✓ Folder metadata verified for {foldersNeedingAuthorFix.Count:N0} folder(s)");
                    }
                }

                // Post-import correction: ProgId (e.g. "OneNote.Notebook" for OneNote notebook
                // folders) is what tells SharePoint's UI a folder is a special container rather
                // than a plain folder — without it, a copied OneNote notebook shows up as an
                // ordinary folder full of .onetoc2/.one files instead of opening in the OneNote
                // web app (2026-07-10 investigation). MigrationPackageBuilder now emits the
                // source's real ProgId in the manifest, but going by the same "SPMI's <Folder>
                // element silently drops fields not in {TimeCreated, TimeLastModified}" pattern
                // already confirmed for Author/ModifiedBy, that's unlikely to be honored on
                // import — so it's patched via the same kind of REST write afterward, unverified
                // against a live tenant until this run.
                async Task CorrectProgIdAsync()
                {
                    var foldersNeedingProgIdFix = folderMetadata
                        .Where(kv => !string.IsNullOrEmpty(kv.Value.ProgId))
                        .ToList();
                    if (foldersNeedingProgIdFix.Count == 0) return;

                    activityLog?.Report($"Correcting {foldersNeedingProgIdFix.Count:N0} special folder(s) (e.g. OneNote notebooks)...");
                    int progIdFixFailures = 0;
                    var progIdErrors = new System.Collections.Concurrent.ConcurrentBag<string>();
                    await Parallel.ForEachAsync(foldersNeedingProgIdFix,
                        new ParallelOptions { MaxDegreeOfParallelism = 6, CancellationToken = cancellationToken },
                        async (kv, ct) =>
                        {
                            var (relKey, meta) = kv;
                            var folderRelUrl = string.IsNullOrEmpty(relKey) ? libraryServerRelUrl : $"{libraryServerRelUrl}/{relKey}";
                            var guid = await GetCachedFolderUniqueIdAsync(sharedFolderUniqueIdCache, targetSiteUrl, folderRelUrl, ct);
                            if (string.IsNullOrEmpty(guid))
                            {
                                Interlocked.Increment(ref progIdFixFailures);
                                progIdErrors.Add($"{(relKey.Length == 0 ? "(library root)" : relKey)}: could not resolve target folder GUID");
                                return;
                            }
                            var err = await spService.PatchFolderProgIdAsync(targetSiteUrl, listId, guid!, meta.ProgId!);
                            if (err != null)
                            {
                                Interlocked.Increment(ref progIdFixFailures);
                                progIdErrors.Add($"{(relKey.Length == 0 ? "(library root)" : relKey)}: {err}");
                            }
                        });
                    if (progIdFixFailures > 0)
                    {
                        var examples = string.Join(" | ", progIdErrors.Distinct().Take(3));
                        activityLog?.Report($"⚠ Could not correct ProgId for {progIdFixFailures:N0} folder(s). Example(s): {examples}");
                    }
                    else
                    {
                        activityLog?.Report($"✓ Special folder association verified for {foldersNeedingProgIdFix.Count:N0} folder(s)");
                    }
                }

                if (groupTasks.Count == 0)
                {
                    // Whole drive group already copied — nothing to import, but the folder metadata
                    // correction still runs. This is the cheap repair path for a completed migration
                    // whose folder dates/people are wrong: re-run the same selection in Copy-If-Newer,
                    // every file skips, and this pass patches every folder without re-transferring
                    // anything.
                    await CorrectFolderMetadataAsync();
                    await CorrectProgIdAsync();
                    continue;
                }

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
                // Byte budget alongside the entry budgets. This is what actually shapes batches for a
                // large-file library, since the 250-item cap never bites when files are big. It is the
                // key latency lever: an import fires only when a WHOLE batch finishes packaging, so a
                // large byte budget means the first import waits for that many GB of scans to download
                // → encrypt → upload before ANY of it starts importing. On a library of multi-hundred-MB
                // scan files a 10 GB budget produced ~34-40-file batches that took ~2 HOURS to
                // package before the first import even began (2026-07-18: 100+ min in, zero imports).
                // 2 GB → ~5-7 such scans per batch → first import in ~20-25 min, and imports then flow
                // continuously overlapping ongoing packaging instead of arriving in ~2-hour waves. It
                // also means the batch-boundary LOH compaction (see CompactLargeObjectHeap) actually
                // runs regularly instead of never, and a fatal abort costs a handful of files, not 40.
                // Small-file regions are unaffected — they still fill to the 250-item cap well under 2 GB.
                const long MaxBytesPerJob   = 2L * 1024 * 1024 * 1024;

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

                int VersionsOf((CopyJob job, CopyResult result) t)
                {
                    if (!metaCache.TryGetValue(t.job.SourceItemId, out var c)) return 1;
                    int raw = Math.Max(1, c.Versions?.Count ?? 1);
                    return maxVersions > 0 ? Math.Min(raw, maxVersions) : raw;
                }

                long BytesOf((CopyJob job, CopyResult result) t)
                {
                    // Cache miss (metadata fetch failed under throttling) → fall back to the size the
                    // scan walk captured, NOT 0: with 0 the byte budget can't see multi-GB files, so a
                    // throttled run packed batches of them unbounded (2026-07-18: 44 GB heap).
                    if (!metaCache.TryGetValue(t.job.SourceItemId, out var c)) return t.job.SourceSize ?? 0;
                    long current = c.Metadata.Size ?? t.job.SourceSize ?? 0;
                    if (c.Versions == null) return current;
                    var counted = maxVersions > 0 ? c.Versions.TakeLast(maxVersions) : c.Versions;
                    long total = 0;
                    foreach (var v in counted) total += v.Size ?? current;
                    return Math.Max(total, current);
                }

                // Warm-up ramp: the first two batches are deliberately tiny so the FIRST SharePoint
                // import fires within a few minutes — an early, unambiguous confirmation that the
                // import half of the pipeline works, on every run — then the caps grow to full size
                // for efficiency. This is permanent and self-limiting: it only shrinks batches 0 and
                // 1, so the worst-case cost on any library is two extra small job submissions, while
                // the fast-feedback benefit applies universally. (Motivation: 2026-07-18, a run went
                // 2+ hours with zero completed imports because the first full-size batch of multi-GB
                // scans hadn't finished packaging — no signal at all that import even functioned.)
                // Only items and bytes ramp; the version cap stays at its hard SPMI ceiling (a tiny
                // batch can't approach it anyway). A single file larger than a ramp byte cap still
                // forms its own batch via the Count > 0 guard below, so oversized files never stall.
                (int items, long bytes) CapsForBatch(int index) => index switch
                {
                    0 => (8,  512L * 1024 * 1024),          // ~1-2 scans, or 8 small files → import in minutes
                    1 => (32, 1L * 1024 * 1024 * 1024),     // medium step before full size
                    _ => (MaxItemsPerJob, MaxBytesPerJob),  // steady state
                };

                var batches      = new List<List<(CopyJob job, CopyResult result)>>();
                var currentBatch = new List<(CopyJob job, CopyResult result)>();
                int  currentVersions = 0;
                long currentBytes    = 0;
                foreach (var t in groupTasks)
                {
                    int  v = VersionsOf(t);
                    long b = BytesOf(t);
                    // Caps for the batch currently being filled (batches.Count is its index).
                    var (itemCap, byteCap) = CapsForBatch(batches.Count);
                    if (currentBatch.Count > 0 &&
                        (currentVersions + v > MaxVersionsPerJob || currentBatch.Count >= itemCap
                         || currentBytes + b > byteCap))
                    {
                        batches.Add(currentBatch);
                        currentBatch = [];
                        currentVersions = 0;
                        currentBytes    = 0;
                    }
                    currentBatch.Add(t);
                    currentVersions += v;
                    currentBytes    += b;
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
                // 4 balances the structural limits: SPMI imports run server-side (a held job costs
                // this app only a 2-30s polling loop), so extra slots exist to keep SharePoint's
                // queue fed with zero idle gap between batches — but SP processes only ~2 jobs per
                // site collection at once (queueing the rest), the prep producer is 3-wide so more
                // than ~4 imports can't be fed continuously anyway, and higher values (6+) have
                // shown SP-side "Operation canceled" soft-aborts plus a bigger blast radius when a
                // conflict abort forces a batch retry. Earlier 2-vs-4 wall-clock measurements were
                // confounded by other since-fixed issues; re-measure before changing.
                const int MaxConcurrentImports = 4;

                // isRetryPass omits the preflight counter/progress reporting — those already counted
                // this batch's files once during the original pass, and re-counting on a post-conflict
                // retry would overshoot the run's total and skew the progress bar. It also withholds
                // the prebuilt existing-file snapshot: the aborted first attempt imported files before
                // dying, and that snapshot predates them, so a retry reusing it re-inserts those files
                // and conflicts all over again (observed 2026-07-02 17:04 run — the retry's errors were
                // files attempt 1 had just imported). A null snapshot makes PrepareBatchAsync re-list
                // the target folders fresh (Graph + REST merged).
                Task<PreparedBatch> PrepTasks(
                    IList<(CopyJob job, CopyResult result)> tasks, string label, bool isRetryPass = false) => PrepareBatchAsync(
                    tasks, overwriteMode, maxVersions, maxParallel,
                    targetSiteUrl, webId, webRelUrl, siteId,
                    libraryServerRelUrl, listId, libraryTitle, cancellationToken,
                    copyCustomColumns, columnMappings, bulkFieldCache,
                    isRetryPass ? null : preflightCounter, preflightTotal, isRetryPass ? null : preflightProgress,
                    sharedFolderIdCache, isRetryPass ? null : sharedExistingByFolder,
                    batchLabel: label,
                    activityLog: activityLog,
                    onFilePacked: onFilePacked,
                    downloadController: downloadController,
                    largeFileGate: largeFileGate,
                    memoryBudget: memoryBudget,
                    folderUniqueIdCache: sharedFolderUniqueIdCache,
                    folderMetadata: folderMetadata,
                    metaCache: metaCache);

                Task<PreparedBatch> Prep(int idx) =>
                    PrepTasks(batches[idx], batches.Count > 1 ? $"{idx + 1}/{batches.Count}" : "");

                // Re-runs a batch once after SharePoint aborts the whole job over "already exists"
                // conflicts. SPMI cancels the ENTIRE job once ~100 items conflict — but it KEEPS the
                // files it already imported, and every other valid file is discarded as collateral
                // (observed 2026-07-02: 246-file batch, 100 reported conflicts, 0-of-246 credited).
                // Deletes the specific conflicting targets (with verification + re-delete for orphaned
                // rows), resets the batch back to Copying, and re-submits once; the retry's fresh
                // pre-flight listing also catches whatever attempt 1 imported before dying. A second
                // fatal abort is left as a genuine failure rather than retried again.
                async Task RetryBatchAfterConflictAsync(
                    PreparedBatch prepared, List<(CopyJob job, CopyResult result)> conflicts)
                {
                    var rpfx = string.IsNullOrEmpty(prepared.BatchLabel) ? "" : $"[{prepared.BatchLabel}] ";

                    // Skip mode: "already exists" is exactly what Skip means — the conflicting files
                    // are Skipped (not Failed, and nothing is deleted), and the collateral valid files
                    // SPMI discarded with the abort get one retry. The retry's fresh pre-flight
                    // re-skips whatever exists (including files attempt 1 imported before dying).
                    if (overwriteMode == OverwriteMode.Skip)
                    {
                        activityLog?.Report(
                            $"↻ {rpfx}SharePoint aborted the batch after {conflicts.Count} 'already exists' conflict{(conflicts.Count == 1 ? "" : "s")} — marking those Skipped and retrying the rest once...");
                        var conflictSet = new HashSet<CopyResult>(conflicts.Select(c => c.result));
                        foreach (var (_, result) in prepared.FileTasks)
                        {
                            if (conflictSet.Contains(result))
                            {
                                result.Status       = CopyStatus.Skipped;
                                result.ErrorMessage = "Already exists";
                            }
                            else if (result.Status == CopyStatus.Failed)
                            {
                                result.Status       = CopyStatus.Copying;
                                result.ErrorMessage = null;
                            }
                        }
                        if (!prepared.FileTasks.Any(t => t.result.Status == CopyStatus.Copying)) return;
                        var retriedSkip = await PrepTasks(prepared.FileTasks, prepared.BatchLabel, isRetryPass: true);
                        await SubmitAndPollBatchAsync(retriedSkip, webId, targetSiteUrl, activityLog, cancellationToken);
                        return;
                    }

                    activityLog?.Report(
                        $"↻ {rpfx}SharePoint aborted the batch after {conflicts.Count} name conflict{(conflicts.Count == 1 ? "" : "s")} — clearing and retrying the batch once...");

                    var stillBlocked = new HashSet<CopyResult>();
                    var blockedLock  = new object();
                    int clearedCount = 0, blockedCount = 0;
                    await Parallel.ForEachAsync(conflicts,
                        new ParallelOptions { MaxDegreeOfParallelism = 8, CancellationToken = cancellationToken },
                        async (conflict, ct) =>
                    {
                        var (job, result) = conflict;
                        var subPath = job.TargetSubFolderPath ?? string.Empty;
                        var fileServerRelUrl = string.IsNullOrEmpty(subPath)
                            ? $"{libraryServerRelUrl}/{job.SourceName}"
                            : $"{libraryServerRelUrl}/{subPath}/{job.SourceName}";

                        string? failReason = null;
                        string? deleteTrace = null;
                        bool ok = await spService.PermanentlyDeleteFileAsync(
                            targetSiteUrl, fileServerRelUrl, r => failReason = r, t => deleteTrace = t);
                        if (ok)
                        {
                            // Confirm it's ACTUALLY gone before trusting the recycle/purge HTTP status —
                            // and if the URL still resolves, DELETE AGAIN. Observed 2026-07-02 17:09:
                            // recycle+purge reported success on 100 conflicts yet all 100 still resolved
                            // ("cleared 0/100"), consistent with a second row surviving at the same URL
                            // (aborted SPMI imports leave orphaned rows alongside the list item). A second
                            // delete pass finds no list item, so recycleObject fails over to deleteObject,
                            // which removes rows at the AllDocs level — the one path the first pass never
                            // took while the recycle succeeded.
                            for (int v = 0; ; v++)
                            {
                                if (await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) == null) break;
                                if (v >= 2) { ok = false; failReason ??= "still resolves after delete+verify"; break; }
                                await spService.PermanentlyDeleteFileAsync(
                                    targetSiteUrl, fileServerRelUrl, r => failReason = r, t => deleteTrace = t);
                                await Task.Delay(TimeSpan.FromSeconds(3 * (v + 1)), ct);
                            }
                        }

                        if (ok) { Interlocked.Increment(ref clearedCount); return; }
                        int nowBlocked;
                        lock (blockedLock)
                        {
                            nowBlocked = ++blockedCount;
                            stillBlocked.Add(result);
                        }
                        if (nowBlocked <= 5)
                            activityLog?.Report($"  ⚠ {rpfx}delete failed [{failReason ?? "unknown"}] (trace: {deleteTrace ?? "n/a"}): {fileServerRelUrl}");
                        else if (nowBlocked == 6)
                            activityLog?.Report($"  ⚠ {rpfx}…more still-blocked conflicts (suppressing)");
                    });
                    activityLog?.Report($"  {rpfx}cleared {clearedCount}/{conflicts.Count} conflicting file(s)" +
                        (blockedCount > 0 ? $", {blockedCount} still blocked" : ""));

                    foreach (var (_, result) in prepared.FileTasks)
                    {
                        if (stillBlocked.Contains(result))
                        {
                            result.Status       = CopyStatus.Failed;
                            result.ErrorMessage = "Could not remove the existing file before retry — re-run to try again.";
                            continue;
                        }
                        // Only files the abort failed go back to Copying. Skipped results (IfNewer's
                        // "Up to date" decisions from the first pass) must survive the retry — resetting
                        // them forced the fresh pre-flight to re-evaluate (and, pre-tolerance-fix,
                        // sub-second skew could convert a former Skip into a full delete+re-copy).
                        if (result.Status == CopyStatus.Failed)
                        {
                            result.Status       = CopyStatus.Copying;
                            result.ErrorMessage = null;
                        }
                    }

                    if (!prepared.FileTasks.Any(t => t.result.Status == CopyStatus.Copying))
                    {
                        activityLog?.Report($"✗ {rpfx}Retry skipped — every conflicting file is still blocked.");
                        return;
                    }

                    var retried = await PrepTasks(prepared.FileTasks, prepared.BatchLabel, isRetryPass: true);
                    await SubmitAndPollBatchAsync(retried, webId, targetSiteUrl, activityLog, cancellationToken);
                }

                // Connect throttle → download-slot step-down now that downloads are about to begin.
                // Analysis throttles (above) are isolated; only download-phase throttles shrink the
                // gate. Unsubscribed per-iteration (finally below): subscribing here but removing
                // only once in the method's outer finally leaked one handler per additional target
                // library, and left group 1's download handler attached during group 2's analysis.
                spService.Throttled += onMigrationThrottle;
                try
                {

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
                    {
                        var conflicts = await SubmitAndPollBatchAsync(prepared, webId, targetSiteUrl, activityLog, cancellationToken);
                        if (conflicts.Count > 0)
                            await RetryBatchAfterConflictAsync(prepared, conflicts);
                    }
                }, cancellationToken)).ToArray();

                await Task.WhenAll(importWorkers);
                await prepProducer;

                }
                finally
                {
                    spService.Throttled -= onMigrationThrottle;
                }

                // Post-import correction (see CorrectFolderMetadataAsync/CorrectProgIdAsync above):
                // SPMI's <Folder> element doesn't honor Author/ModifiedBy or (likely) ProgId, so
                // both are patched via REST after the imports complete.
                await CorrectFolderMetadataAsync();
                await CorrectProgIdAsync();
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                // Cancelled, not Failed: this item was still Copying (never actually attempted)
                // when the run stopped — see CopyStatus.Cancelled.
                result.Status       = CopyStatus.Cancelled;
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
            if (onThrottleLog != null) spService.Throttled -= onThrottleLog;
            heartbeatCts.Cancel();
            try { await heartbeatTask; } catch (OperationCanceledException) { }
        }
    }

    private static async Task RunHeartbeatAsync(IProgress<string> activityLog, CancellationToken ct)
    {
        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(30));
        while (await timer.WaitForNextTickAsync(ct))
            activityLog.Report(ProcessDiagnostics.Snapshot());
    }

    // Forces a blocking gen-2 collection with LOH compaction, at most once per minute across all
    // concurrent batches. Called at batch boundaries (a natural lull, and right after a batch's
    // buffers die). The blocking pause runs on a worker thread and costs well under a second even
    // at multi-GB heaps — trivial next to batch import times, and far cheaper than the paging and
    // GC thrash a 15 GB fragmented floor causes on a 32 GB machine.
    private static long _lastLohCompactTicks;
    private static void CompactLargeObjectHeap()
    {
        long now  = DateTime.UtcNow.Ticks;
        long last = Interlocked.Read(ref _lastLohCompactTicks);
        if (now - last < TimeSpan.FromSeconds(60).Ticks) return;
        if (Interlocked.CompareExchange(ref _lastLohCompactTicks, now, last) != last) return;

        System.Runtime.GCSettings.LargeObjectHeapCompactionMode =
            System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
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

    // Merges Graph's folder listing with the SP REST view of the same folder. SPMI's duplicate
    // check runs against SharePoint's internal file table, which can hold rows Graph doesn't
    // return (files landed by a previous SPMI job that fatal-aborted mid-import). A pre-flight
    // that trusts Graph alone reports those folders as empty, skips the deletes, and every such
    // file then bounces off the import with "already exists" (observed 2026-07-02, 100+ per
    // batch). REST-only names carry an empty ItemId — nothing downstream reads ItemId; the
    // Overwrite/IfNewer branches delete by server-relative URL and Skip matches by name.
    private async Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>
        FetchTargetFolderSnapshotAsync(string driveId, string folderId, string siteUrl, string folderServerRelUrl)
    {
        var graphTask = spService.FetchFolderItemsAsync(driveId, folderId);
        var restTask  = spService.FetchFolderFileNamesRestAsync(siteUrl, folderServerRelUrl);
        await Task.WhenAll(graphTask, restTask);
        var merged = graphTask.Result;
        foreach (var (name, modified) in restTask.Result)
            if (!merged.ContainsKey(name))
                merged[name] = (string.Empty, modified);
        return merged;
    }

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
        SemaphoreSlim? largeFileGate = null,
        TransferMemoryBudget? memoryBudget = null,
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
                        // Graph + SP REST merged (see FetchTargetFolderSnapshotAsync). Also the path
                        // taken by a post-abort retry (prebuilt snapshot deliberately withheld): the
                        // aborted job imported files before dying, so the retry MUST re-list or it
                        // re-inserts those files and conflicts all over again.
                        var folderRelUrl = string.IsNullOrEmpty(kvp.Key)
                            ? libraryServerRelUrl
                            : $"{libraryServerRelUrl}/{kvp.Key}";
                        existingByFolder[kvp.Key] = await FetchTargetFolderSnapshotAsync(
                            fileTasks[0].job.TargetDriveId, kvp.Value, targetSiteUrl, folderRelUrl);
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

                    // A task can arrive here already Failed (a post-abort retry keeps still-blocked
                    // conflicts Failed while re-running the rest of its batch). Deleting that task's
                    // target here would destroy the existing file with no re-import to replace it —
                    // only files actually being copied may have their targets cleared.
                    if (result.Status != CopyStatus.Copying)
                    {
                        if (preflightCounter != null)
                        {
                            int doneNow = Interlocked.Increment(ref preflightCounter[0]);
                            preflightProgress?.Report((doneNow, preflightTotal));
                        }
                        return;
                    }

                    var subPath          = job.TargetSubFolderPath ?? string.Empty;
                    var fileServerRelUrl = string.IsNullOrEmpty(subPath)
                        ? $"{libraryServerRelUrl}/{job.SourceName}"
                        : $"{libraryServerRelUrl}/{subPath}/{job.SourceName}";

                    // Marks the item Failed with a clear, actionable message instead of letting it fall
                    // through to SPMI, which would otherwise reject it with a confusing "already exists"
                    // error (blank modified-by/date — SPMI's own message template, not ours) once the
                    // stale target survives into the import.
                    void failDeleteConflict() {
                        result.Status       = CopyStatus.Failed;
                        result.ErrorMessage = "Could not remove the existing file before overwrite (after retries) — skipped to avoid a duplicate-version import error. Re-run to retry.";
                    }

                    if (overwriteMode == OverwriteMode.Overwrite)
                    {
                        bool deleteOk = true;
                        if (existingByFolder[subPath].ContainsKey(job.SourceName))
                        {
                            // Real file: delete first so SPMI does a fresh INSERT.
                            // SPMI UPDATE (with existing GUID) appends imported versions to the
                            // existing version history instead of replacing it, causing duplication.
                            deleteOk = await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                        else if (existingByFolder[subPath].Count > 0 &&
                                 await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) != null)
                        {
                            // Zombie blob (AllDocs row without SPListItem) — delete via deleteObject,
                            // which bypasses the SPListItem requirement that recycleObject needs.
                            // Guard: if the folder is empty, no prior SPMI import ran here so no zombies exist.
                            deleteOk = await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                        if (!deleteOk) failDeleteConflict();
                        else existingFileIds[fileServerRelUrl] = null;
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
                            if (srcModifiedDate is { } srcModified && existing.Modified is { } targetModified &&
                                TimestampComparer.IsUpToDate(srcModified, targetModified))
                            {
                                result.Status       = CopyStatus.Skipped;
                                result.ErrorMessage = CopyResult.UpToDate;
                            }
                            else
                            {
                                // Source is newer — delete so SPMI does a fresh INSERT
                                // (same reasoning as overwrite mode).
                                if (!await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl))
                                    failDeleteConflict();
                            }
                        }
                        else if (existingByFolder[subPath].Count > 0 &&
                                 await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) != null)
                        {
                            // Zombie SPFile blob — purge so SPMI can import cleanly.
                            if (!await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl))
                                failDeleteConflict();
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
                            if (!await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl))
                                failDeleteConflict();
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
            // NetworkTimeout extends the per-request timeout from the 60-second Azure default so
            // large files don't cancel mid-upload on slow connections. The retry settings make
            // Azure's OWN pipeline ride out transient mid-transfer connection drops (the "Error
            // while copying content to a stream" failures — a dropped TCP write, the upload-side
            // twin of the download connection resets) with exponential backoff before it gives up
            // and surfaces to our outer retry loop. Our outer loop then adds several more attempts;
            // together with the reduced upload concurrency below this is what keeps the upload
            // failure rate near zero on a flaky connection (2026-07-18: ~40% of files were failing
            // here because too many multi-GB uploads ran at once and saturated the upload path).
            var blobOptions = new BlobClientOptions();
            blobOptions.Retry.NetworkTimeout = TimeSpan.FromMinutes(30);
            blobOptions.Retry.MaxRetries     = 6;
            blobOptions.Retry.Mode           = Azure.Core.RetryMode.Exponential;
            blobOptions.Retry.Delay          = TimeSpan.FromSeconds(2);
            blobOptions.Retry.MaxDelay       = TimeSpan.FromSeconds(30);
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

                            // This path buffers each version fully in memory (MemoryStream is capped at
                            // int.MaxValue bytes) — a version above 2 GiB cannot be packaged. Fail fast
                            // with an honest message instead of downloading it 3 times and surfacing
                            // "Stream was too long" as a connection problem.
                            // knownFileBytes falls back to the scan-captured size for cache-miss files —
                            // with the old 0 fallback a throttled run's misses bypassed BOTH this guard
                            // and the large-file gate below, so 16 multi-GB files buffered concurrently
                            // (2026-07-18: 44 GB heap, connection-reset storms, mass failures).
                            long knownFileBytes = cachedMeta?.Size ?? job.SourceSize ?? 0;
                            long largestVersionBytes = Math.Max(knownFileBytes,
                                cachedVersions?.Max(v => v.Size ?? 0) ?? 0);
                            if (largestVersionBytes > int.MaxValue)
                            {
                                result.Status       = CopyStatus.Failed;
                                result.ErrorMessage = $"File version is {largestVersionBytes / (1024.0 * 1024 * 1024):F1} GB — larger than the 2 GB this mode can buffer. Copy this file with Enhanced REST mode.";
                                return;
                            }

                            // Oversized files (multi-GB genomics/microscopy files etc.) get an extra gate on
                            // top of the normal download slot: each one holds its full content in memory
                            // twice (raw + AES-encrypted) until upload completes, so only a couple may be
                            // in flight at once regardless of maxParallel — acquired BEFORE the general
                            // slot so a large file waiting on this gate doesn't tie up a slot small files
                            // could otherwise use.
                            // Gate on the file's TOTAL bytes across all versions, not just the current
                            // version: all versions download concurrently and are held (raw + encrypted)
                            // until upload — a small current version with gigabytes of history is exactly
                            // the OOM class this gate exists for.
                            long totalVersionBytes = 0;
                            if (cachedVersions != null)
                                foreach (var v in cachedVersions)
                                    totalVersionBytes += v.Size ?? knownFileBytes;
                            // A file whose size is genuinely unknown (cache miss AND built outside the
                            // scan walk) is gated as large: the safe assumption costs at worst a little
                            // queueing; the unsafe one is the OOM class this gate exists for.
                            bool sizeUnknown = cachedMeta?.Size == null && job.SourceSize == null;
                            bool isLargeFile = sizeUnknown ||
                                Math.Max(knownFileBytes, totalVersionBytes) >= LargeFileThresholdBytes;
                            if (isLargeFile && largeFileGate != null)
                                await largeFileGate.WaitAsync(ct);

                            // Reserve the file's slice of the global byte budget BEFORE taking a
                            // download slot — waiting on memory while holding a slot would idle the
                            // slot for every other file. Charge ≈ 2× the payload (raw + encrypted
                            // copies coexist from encrypt until upload completes); unknown sizes
                            // charge the large-file threshold as a conservative nominal.
                            long budgetCharge = 0;
                            if (memoryBudget != null)
                            {
                                long payloadBytes = Math.Max(knownFileBytes, totalVersionBytes);
                                if (payloadBytes <= 0) payloadBytes = LargeFileThresholdBytes;
                                budgetCharge = memoryBudget.ClampCharge(2 * payloadBytes);
                                await memoryBudget.WaitAsync(budgetCharge, ct);
                            }

                            // Acquire a global download slot — limits total concurrent Graph
                            // content downloads across all SPMI batches to maxParallel.
                            if (downloadController != null)
                                await downloadController.WaitAsync(ct);
                            bool handedToConsumer = false;
                            try
                            {
                                if (verbosePerFile || isLargeFile) activityLog?.Report($"{pfx}↓ {job.SourceName}");
                                var data = await DownloadFileDataAsync(
                                    job, result, maxVersions, versionParallelism, existingFileId, ct,
                                    cachedMeta, cachedVersions,
                                    copyCustomColumns, columnMappings, bulkFieldCache,
                                    pfxLabel: pfx, activityLog: activityLog);
                                // Transfer the large slot AND the byte reservation to the consumer
                                // (see HoldsLargeSlot/HeldBudgetBytes): both are released only after
                                // this file's encrypted blobs finish uploading, so they bound the
                                // file's WHOLE in-memory lifetime, not just the download.
                                if (isLargeFile && largeFileGate != null)
                                    data = data with { HoldsLargeSlot = true };
                                if (budgetCharge > 0)
                                    data = data with { HeldBudgetBytes = budgetCharge };
                                await pipe.Writer.WriteAsync(data, ct);
                                handedToConsumer = true;
                            }
                            catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                            catch (Exception ex)
                            {
                                result.Status       = CopyStatus.Failed;
                                result.ErrorMessage = $"Download failed: {ex.Message}";
                            }
                            finally
                            {
                                downloadController?.Release();
                                if (!handedToConsumer)
                                {
                                    if (isLargeFile) largeFileGate?.Release();
                                    if (budgetCharge > 0) memoryBudget?.Release(budgetCharge);
                                }
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
            //
            // Capped at 4 (was min(16, maxParallel)): this is a PER-BATCH gate, and up to
            // MaxConcurrentPrep batches package at once, so 16 meant up to ~48 simultaneous multi-GB
            // uploads to Azure — which saturated the upload path and caused ~40% of files to fail
            // with "Error while copying content to a stream" after exhausting all retries (2026-07-18).
            // 4 per batch → ~12 concurrent uploads across batches: still plenty of parallelism, far
            // below the saturation point. Uploads are network-bound, so fewer-but-reliable beats
            // more-but-failing (a failed file is pure rework). Small files are unaffected — they
            // upload in one fast PUT regardless.
            int uploadConcurrency = Math.Max(2, Math.Min(4, maxParallel));
            using var uploadGate  = new SemaphoreSlim(uploadConcurrency);
            var uploadTasks       = new List<Task>();

            // GUID → CopyResult, built as files are added, so per-item SP import errors can be
            // attributed back to the exact file. Written only in this sequential consumer loop.
            var listItemMap = new Dictionary<string, CopyResult>(StringComparer.OrdinalIgnoreCase);

            await foreach (var data in pipe.Reader.ReadAllAsync(cancellationToken))
            {
                int filesBefore = builder.Files.Count;
                // This iteration owns the file's large slot and byte reservation (if any) until
                // ownership passes to its upload task; every failure path before that hand-off
                // must release them here.
                bool ownsGateAndBudget = data.HoldsLargeSlot || data.HeldBudgetBytes > 0;
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
                    // Register EVERY GUID the manifest mints for this file — SP import errors reference
                    // different ids depending on error type ([File] errors carry the FileId, list-item
                    // errors the ListItemId, stream errors the StreamId). Keying only ListItemId left
                    // [File]-level errors unattributed (observed 2026-07-01: 3 MD5 errors reported, 1
                    // attributed — 2 failed files shown as Success in the grid).
                    listItemMap[entry.ListItemId] = dataResult;
                    listItemMap[entry.FileId]     = dataResult;
                    foreach (var v in entry.Versions)
                    {
                        listItemMap[v.FileId]   = dataResult;
                        listItemMap[v.StreamId] = dataResult;
                    }
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
                                // SPMI requires an MD5 for every content blob ("Missing MD5 property in
                                // Manifest and Azure Blob" otherwise). Small blobs get one implicitly:
                                // a single-request upload has Azure compute/store Content-MD5 itself. Blobs
                                // above the SDK's single-shot threshold (~256 MB) upload as blocks, which
                                // get NO automatic blob-level MD5 — so every large file failed import
                                // (observed 2026-07-01: all >2 GB files in a batch rejected). Compute it
                                // ourselves over exactly the uploaded bytes (ciphertext after the 16-byte
                                // IV prefix — SP validates the stored blob, which is encrypted).
                                var md5 = System.Security.Cryptography.MD5.HashData(content.AsSpan(16));
                                var opts  = new BlobUploadOptions
                                {
                                    Metadata    = new Dictionary<string, string> { ["IV"] = ivB64 },
                                    HttpHeaders = new BlobHttpHeaders { ContentHash = md5 },
                                };
                                var blob = dataClient.GetBlobClient(version.StreamId);

                                // Azure Storage's own internal retry (ClientOptions.Retry) already exhausts
                                // itself on sustained network blips — when it does it throws an
                                // AggregateException ("Retry failed after 6 tries. ... (Error while copying
                                // content to a stream.)"). The old catch listed specific types (IOException,
                                // HttpRequestException, RequestFailedException, OperationCanceledException)
                                // and AggregateException is NONE of them, so that exhausted-retry failure
                                // slipped straight past this loop to the outer handler with ZERO extra
                                // retries — the root cause of the ~40% upload failure rate on 2026-07-18.
                                // Retry ANY failure except genuine user cancellation, with a longer,
                                // more patient backoff. A fresh MemoryStream is needed each attempt since
                                // UploadAsync consumes it.
                                const int UploadMaxAttempts = 5;
                                for (int attempt = 0; ; attempt++)
                                {
                                    try
                                    {
                                        using var ms = new MemoryStream(content, 16, content.Length - 16);
                                        await blob.UploadAsync(ms, opts, cancellationToken);
                                        await blob.CreateSnapshotAsync(cancellationToken: cancellationToken);
                                        break;
                                    }
                                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
                                    // Anything else — including the AggregateException from Azure's own
                                    // exhausted retry, or an OperationCanceledException from the SDK's
                                    // 30-minute NetworkTimeout (ruled out as user cancellation above) —
                                    // is a transient upload failure; retry it rather than failing the file
                                    // on one bad window.
                                    catch (Exception) when (attempt < UploadMaxAttempts - 1)
                                    {
                                        int waitsecs = Math.Min(60, (attempt + 1) * 10);
                                        activityLog?.Report($"⚠ {pfx}{fileName} — upload interrupted, retrying in {waitsecs}s ({attempt + 1}/{UploadMaxAttempts - 1})");
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
                        // Let a real user cancellation propagate unchanged; anything else — including an
                        // OperationCanceledException from an exhausted-retry client timeout above — fails
                        // just this file so the rest of the batch can still complete.
                        catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
                        {
                            entry.Failed = true; // exclude from manifest — its blobs weren't all uploaded
                            dataResult.Status       = CopyStatus.Failed;
                            dataResult.ErrorMessage = $"Upload failed ({fileName}): {ex.Message}";
                            // Never silent: this used to set ErrorMessage only, so a failed upload was
                            // invisible in the activity log and showed up solely as a bare "K failed"
                            // count with no reason (2026-07-18 — made this class of failure very hard to
                            // diagnose). Surface it like the download/import failure paths do.
                            activityLog?.Report($"✗ {pfx}Upload failed after retries: {fileName} — {ex.Message}");
                        }
                        finally
                        {
                            uploadGate.Release();
                            // End of this file's in-memory life (uploaded or failed, encrypted
                            // buffers nulled/dropped) — return its large slot and byte reservation.
                            if (data.HoldsLargeSlot) largeFileGate?.Release();
                            if (data.HeldBudgetBytes > 0) memoryBudget?.Release(data.HeldBudgetBytes);
                        }
                    }, cancellationToken));
                    ownsGateAndBudget = false; // upload task now owns both (released in its finally)
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
                    // Slot/reservation never reached an upload task (encrypt failed, or the upload
                    // hand-off threw) — release here or the gates starve for the rest of the run.
                    if (ownsGateAndBudget)
                    {
                        if (data.HoldsLargeSlot) largeFileGate?.Release();
                        if (data.HeldBudgetBytes > 0) memoryBudget?.Release(data.HeldBudgetBytes);
                    }
                }
            }

            await producerTask;
            await Task.WhenAll(uploadTasks); // all blobs uploaded before the manifest is built

            // This batch's multi-hundred-MB buffers just became garbage — compact the Large Object
            // Heap so the memory actually returns to the OS. Without this the LOH fragments and the
            // heap floor ratchets upward across the run (observed 2026-07-18: 0.5 GB → 15 GB floor
            // over ~90 minutes; the resulting GC pauses surfaced as connection-reset storms).
            CompactLargeObjectHeap();

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
                // Cancelled, not Failed: this item was still Copying (never actually attempted)
                // when the run stopped — see CopyStatus.Cancelled.
                result.Status       = CopyStatus.Cancelled;
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

    // Phase 2 of a batch: submit the import job and poll to completion. Runs at most
    // MaxConcurrentImports wide (measured sweet spot ~2; SharePoint soft-cancels imports when too
    // many run concurrently). No-op when the prepared batch has nothing to import (all files
    // skipped, or prep failed).
    //
    // Returns the (job, result) pairs that were attributed a genuine "already exists" conflict when
    // the job hit a JobFatalError — empty otherwise. The caller uses this to retry the whole batch
    // once after clearing those specific conflicts, instead of accepting SPMI's blanket abort as final
    // for every other (perfectly valid) file in the batch.
    private async Task<List<(CopyJob job, CopyResult result)>> SubmitAndPollBatchAsync(
        PreparedBatch batch, string webId, string targetSiteUrl,
        IProgress<string>? activityLog, CancellationToken cancellationToken)
    {
        if (batch.CopyingCount == 0) return new List<(CopyJob job, CopyResult result)>();

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
            bool   sawJobEnd      = false;
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
                        sawJobEnd = true;
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
                        // Attribute to the exact file and mark it Failed so it shows in the Failed filter
                        // and can be re-copied. The GUID isn't always in Message (the MD5 errors carry
                        // none there) — scan the WHOLE event for any GUID we minted for this batch; only
                        // manifest-minted ids can match the map, so job/correlation GUIDs are inert.
                        foreach (System.Text.RegularExpressions.Match gm in GuidRegex.Matches(evt.ToString()))
                        {
                            if (batch.ListItemMap.TryGetValue(gm.Value, out var failedRes)
                                && failedRes.Status == CopyStatus.Copying)
                            {
                                failedRes.Status       = CopyStatus.Failed;
                                failedRes.ErrorMessage = msg;
                                break;
                            }
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

            // Never fabricate Success: if polling ended without a JobEnd (cancellation, hung poll
            // endpoint, job evicted server-side) the import outcome is unknown for every file
            // still in flight — surface that honestly instead of promoting them below.
            cancellationToken.ThrowIfCancellationRequested();
            if (!sawJobEnd && !fatal)
                throw new Exception("Import polling ended before SharePoint reported completion — outcome unknown; re-run in Copy-If-Newer mode to reconcile");

            // Fetch SP's report BEFORE final marking: the .err file names every failed object with its
            // GUID, while the live queue events sometimes omit it (observed with the MD5 errors) — any
            // failure left unattributed at this point would be marked Success below.
            if (fatal || liveErrorCount > 0 || totalErrorsReported > 0)
                await TryLogMigrationReportAsync(metadataClient, jobId, batch.EncryptionKey, activityLog, pfx,
                    attributeMap: batch.ListItemMap);

            // If SP reported more errors than we could attribute to a specific file (no minted GUID
            // in the live event or any .err line), never promote the remainder to Success blindly —
            // confirm each still-in-flight file actually landed on the target. A lookup failure here
            // marks the file Failed (re-runnable), which is the safe direction; fabricated Success
            // is not.
            int attributedFailed = fileTasks.Count(t => t.result.Status == CopyStatus.Failed);
            int reportedErrors   = Math.Max(liveErrorCount, totalErrorsReported);
            if (!fatal && reportedErrors > attributedFailed)
            {
                var unconfirmed = fileTasks.Where(t => t.result.Status == CopyStatus.Copying).ToList();
                if (unconfirmed.Count > 0)
                {
                    activityLog?.Report($"⚠ {pfx}{reportedErrors - attributedFailed} import error(s) could not be attributed — confirming {unconfirmed.Count:N0} file(s) on target...");
                    await Parallel.ForEachAsync(unconfirmed,
                        new ParallelOptions { MaxDegreeOfParallelism = 8, CancellationToken = cancellationToken },
                        async (t, ct) =>
                        {
                            var (job, result) = t;
                            var subPath = job.TargetSubFolderPath ?? string.Empty;
                            var lib     = job.TargetLibraryServerRelativeUrl;
                            var fileUrl = string.IsNullOrEmpty(subPath)
                                ? $"{lib}/{job.SourceName}"
                                : $"{lib}/{subPath}/{job.SourceName}";
                            if (await spService.GetFileUniqueIdAsync(targetSiteUrl, fileUrl) == null)
                            {
                                result.Status       = CopyStatus.Failed;
                                result.ErrorMessage = "SharePoint reported import errors and this file was not found on the target afterward — re-run to retry";
                            }
                        });
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
            }
            else if (failedCount > 0 || liveErrorCount > 0 || totalErrorsReported > 0)
            {
                // Reconcile counts: if SP/JobError reported more errors than we could attribute to a
                // specific GUID, surface the discrepancy so a silent shortfall is never hidden.
                int reported = Math.Max(failedCount, Math.Max(liveErrorCount, totalErrorsReported));
                var extra = reported > failedCount ? $" ({reported} errors reported, {failedCount} attributed)" : "";
                activityLog?.Report($"⚠ {pfx}Import finished with errors — {importedCount} of {batch.CopyingCount} imported, {failedCount} failed{extra}");
            }
            else
            {
                activityLog?.Report($"{pfx}✓ Import complete — {importedCount} file{(importedCount == 1 ? "" : "s")} imported");
            }

            // Only a fatal abort warrants a batch-wide retry — the plain "some items failed" path
            // already reports honest per-file results and doesn't need a whole-batch redo.
            return fatal
                ? fileTasks.Where(t => t.result.Status == CopyStatus.Failed &&
                    t.result.ErrorMessage != null &&
                    t.result.ErrorMessage.Contains("already exists", StringComparison.OrdinalIgnoreCase)).ToList()
                : new List<(CopyJob job, CopyResult result)>();
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                // Cancelled, not Failed: this item was still Copying (never actually attempted)
                // when the run stopped — see CopyStatus.Cancelled.
                result.Status       = CopyStatus.Cancelled;
                result.ErrorMessage = "Cancelled";
            }
            return new List<(CopyJob job, CopyResult result)>();
        }
        catch (Exception ex)
        {
            // Unlike every other failure path in this method, a throw here (e.g. job submission
            // itself failing/timing out before a job ID is even obtained) used to mark files Failed
            // with no activity-log line at all — a batch could die silently mid-run with nothing to
            // show why. Always surface it, matching the "✗ ... FAILED" convention used elsewhere.
            activityLog?.Report($"✗ {pfx}Import FAILED — {ex.Message}");
            foreach (var (_, result) in fileTasks.Where(t => t.result.Status == CopyStatus.Copying))
            {
                result.Status       = CopyStatus.Failed;
                result.ErrorMessage = ex.Message;
            }
            return new List<(CopyJob job, CopyResult result)>();
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
        Dictionary<string, string>? CustomFields)
    {
        // True when this file still owns a largeFileGate slot. The producer transfers ownership
        // here instead of releasing after download: the file's bytes stay in memory well past the
        // download (raw stream in the pipe, then the AES-encrypted copy queued for upload), so
        // releasing at download-end let more large files start while earlier ones' encrypted
        // buffers were still piled up awaiting upload — the gate bounded only the first third of
        // each file's memory lifetime. Whoever ends the file's in-memory life (upload task's
        // finally, or the consumer's failure path) releases the slot.
        public bool HoldsLargeSlot { get; init; }

        // Bytes this file has reserved from the global TransferMemoryBudget (0 = none). Same
        // ownership rules and release sites as HoldsLargeSlot.
        public long HeldBudgetBytes { get; init; }
    }

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
                // Pre-size the buffer when the version's byte size is known: letting MemoryStream
                // grow by doubling allocates ~2× the file size in cumulative LOH churn per version,
                // which at multi-GB scale is real GC pressure on top of an already-bounded budget.
                long expectedBytes = version.Size ?? metadata.Size ?? 0;
                var ms = expectedBytes > 0 && expectedBytes <= int.MaxValue
                    ? new MemoryStream((int)expectedBytes)
                    : new MemoryStream();
                for (int attempt = 0; ; attempt++)
                {
                    // RESUME, don't restart: a reset partway through a multi-GB scan used to discard
                    // everything received so far and redownload the entire file from byte 0 on every
                    // retry — on a run where resets hit almost every large file 1-3 times, that meant
                    // downloading several GB two or three times over for a single file (observed
                    // 2026-07-18: ~1 file packaged per 1.8 minutes system-wide, ETA 160+ hours). ms's
                    // current Length IS how many bytes survived the previous attempt, so ask the
                    // server to continue from exactly there via Range instead of clearing it.
                    long resumeFrom = ms.Length;
                    try
                    {
                        var content = await spService.DownloadContentRangeAsync(
                            job.SourceDriveId, job.SourceItemId, isLast ? null : version.Id, resumeFrom, ct);

                        // Some intermediary in the redirect chain can ignore the Range header and
                        // return the WHOLE file from byte 0 (200, not 206) — appending that to bytes
                        // already buffered would silently corrupt the content, so only append when
                        // the server actually confirmed it's continuing from our offset; otherwise
                        // discard the partial buffer and restart clean, same as before this change.
                        bool resumed = resumeFrom > 0 && content.IsPartial && content.StartOffset == resumeFrom;
                        if (resumeFrom > 0 && !resumed) ms.SetLength(0);

                        await using (content.Content)
                        {
                            ms.Position = ms.Length; // append point (0 on a fresh/reset stream)
                            await content.Content.CopyToAsync(ms, ct);
                        }
                        ms.Position = 0;
                        slots[idx] = ms;
                        break;
                    }
                    catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                    // A MemoryStream past int.MaxValue is a hard size limit, not a transient network
                    // failure — retrying re-downloads the whole multi-GB file for nothing.
                    catch (System.IO.IOException ex) when (ex.Message.Contains("Stream was too long", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new NotSupportedException(
                            $"{job.SourceName} has a version larger than the 2 GB this mode can buffer — copy it with Enhanced REST mode.", ex);
                    }
                    // HTTP/2 stream resets surface as HttpRequestException, not IOException — must be
                    // caught alongside it or a single mid-stream RST_STREAM fails the file outright. A
                    // stalled connection surfaces as OperationCanceledException from the client's own
                    // 30-minute request timeout (ruled out as real cancellation above) and deserves the
                    // same retry, not silent treatment as a cancel that then stalls the whole batch.
                    catch (Exception ex) when (attempt < 3 && (ex is System.IO.IOException || ex is System.Net.Http.HttpRequestException || ex is OperationCanceledException))
                    {
                        int waitsecs = (attempt + 1) * 5;
                        var resumeNote = ms.Length > 0 ? $", resuming from {ms.Length / (1024.0 * 1024):F0} MB" : "";
                        activityLog?.Report($"⚠ {pfxLabel}{job.SourceName} — connection {(ex is OperationCanceledException ? "timed out" : "reset")}, retrying in {waitsecs}s ({attempt + 1}/3{resumeNote})");
                        await Task.Delay(TimeSpan.FromSeconds(waitsecs), ct);
                    }
                }
            });

        // Custom field values are NOT packaged: the manifest's standalone SPListItem objects (the
        // only element SPMI accepts <Fields> on) were removed because they caused "Missing file
        // info" import failures (see MigrationPackageBuilder.EmitStandaloneListItems). Fetching
        // per-file field data here was pure waste — one sequential Graph round-trip per file,
        // during the throttle-sensitive download phase, feeding data the manifest never emits.
        // CopyService surfaces a warning when custom columns are requested in this mode.
        Dictionary<string, string>? customFields = null;

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

    private static async Task TryLogMigrationReportAsync(
        Azure.Storage.Blobs.BlobContainerClient metadataClient, string jobId, byte[] key,
        IProgress<string>? activityLog = null, string? pfx = null,
        IReadOnlyDictionary<string, CopyResult>? attributeMap = null)
    {
        foreach (var suffix in new[] { ".err", ".log" })
        {
            // SP splits large reports into numbered segments (Import-{jobId}-1.err, -2.err, …);
            // reading only segment 1 left errors in later segments unattributed. Walk segments
            // until the first missing one.
            for (int segment = 1; segment <= 20; segment++)
            {
            var name = $"Import-{jobId}-{segment}{suffix}";
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
                    // Blob not written by SP — stop walking segments for this suffix.
                    System.Diagnostics.Debug.WriteLine($"[SP-{suffix[1..]}] {name} not present (SP did not write it)");
                    break;
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

                if (suffix == ".err")
                {
                    var lines = reportText.Split('\n')
                        .Select(l => l.Trim()).Where(l => l.Length > 0).ToList();

                    // Attribute each error line to its file: every line names the failed object's GUID,
                    // which the batch minted (FileId/ListItemId/StreamId are all registered). This is the
                    // authoritative failure list — live queue events can omit the GUID entirely.
                    if (attributeMap != null)
                    {
                        foreach (var line in lines.Where(l => l.Contains("[Error]", StringComparison.OrdinalIgnoreCase)))
                        {
                            foreach (System.Text.RegularExpressions.Match gm in GuidRegex.Matches(line))
                            {
                                if (attributeMap.TryGetValue(gm.Value, out var res)
                                    && res.Status == CopyStatus.Copying)
                                {
                                    res.Status       = CopyStatus.Failed;
                                    res.ErrorMessage = line;
                                    break;
                                }
                            }
                        }
                    }

                    // Surface the first few lines of SP's .err report into the activity feed so the
                    // actual failure reason is visible to the user immediately (not just in VS Debug
                    // Output). First segment only — later segments repeat the same error shapes.
                    if (activityLog != null && segment == 1)
                    {
                        foreach (var line in lines.Take(5))
                            activityLog.Report($"   {pfx}SP: {line}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SP-{suffix[1..]}] Cannot read {name}: {ex.Message}");
                break; // don't probe further segments after an unexpected read failure
            }
            } // end segment loop
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
