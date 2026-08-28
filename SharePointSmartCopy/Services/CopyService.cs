using System.Collections.ObjectModel;
using System.IO;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class CopyService(SharePointService spService, MigrationJobService migrationJobService)
{
    // Fired when adaptive throttling changes the effective parallelism during a copy run.
    public event Action<int>? ParallelismChanged;

    // Per-target-folder snapshot of existing files (name -> id/modified), used by the Enhanced REST
    // Skip/IfNewer decision (see CopySingleFileAsync) instead of one GetFileInfoAsync Graph call PER
    // FILE. On an all-skip Copy-If-Newer re-run of 100k files that was 100k round trips for a
    // decision one per-folder listing already answers for every file in it — ~5,000 folders means
    // ~20x fewer calls. Reset per run (see ExecuteCoreAsync) since target state can change between
    // runs; the Lazy dedupes concurrent first-access races to the same folder into one Graph call,
    // matching the pattern _folderSegmentTasks already uses.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string,
        Lazy<Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>>> _folderSnapshotCache = new();

    // Files at or above this size are spilled to a temp file rather than buffered fully in memory.
    // Enhanced REST previously buffered EVERY file into a MemoryStream with no size gate at all —
    // unbounded per file, multiplied by maxParallel concurrent copies (the same OOM incident class
    // a large-file gate + memory budget were built for on the Migration API side, just never
    // ported to this engine). MemoryStream also caps at int.MaxValue, so a file over ~2 GB threw
    // "Stream was too long" here — even though MigrationJobService's own >2 GB error message tells
    // the user to retry with Enhanced REST. That advice did not actually work until this fix. An
    // unknown size is treated as large, the same defensive default the Migration API gate uses.
    private const long EnhancedRestLargeFileThresholdBytes = 100L * 1024 * 1024; // 100 MB

    // Bounds how many large files are buffered (spilled to disk) at once, independent of
    // maxParallel — mirrors MigrationJobService's largeFileGate/MaxConcurrentLargeFiles. Small
    // files never touch this gate. A plain field (not scoped per-run) is fine: the count always
    // returns to its max once every holder releases in its `finally`, so nothing needs resetting
    // between runs.
    private readonly SemaphoreSlim _largeFileGate = new(2);

    private static bool IsLargeForBuffering(long? knownSize) =>
        !knownSize.HasValue || knownSize.Value >= EnhancedRestLargeFileThresholdBytes;

    private static Stream CreateTransferBuffer(bool spillToDisk) =>
        spillToDisk
            // DeleteOnClose: the temp file is removed as soon as the caller's `using` disposes the
            // stream, including on an exception — no separate cleanup path needed.
            ? new FileStream(Path.GetTempFileName(), FileMode.Create, FileAccess.ReadWrite, FileShare.None,
                bufferSize: 81920, FileOptions.DeleteOnClose)
            : new MemoryStream();

    private Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>> GetOrBuildFolderSnapshotAsync(
        string driveId, string parentItemId)
    {
        var cacheKey = $"{driveId}|{parentItemId}";
        var d = driveId; var p = parentItemId;
        var lazy = _folderSnapshotCache.GetOrAdd(cacheKey,
            _ => new Lazy<Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>>(
                () => spService.FetchFolderItemsAsync(d, p),
                System.Threading.LazyThreadSafetyMode.ExecutionAndPublication));
        return AwaitOrEvict(cacheKey, lazy);
    }

    private async Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>> AwaitOrEvict(
        string cacheKey, Lazy<Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>> lazy)
    {
        try
        {
            return await lazy.Value;
        }
        catch
        {
            // Never cache a faulted resolution — a single transient failure would otherwise make
            // every remaining file in this folder, for the rest of the run, fall back as if the
            // folder were empty (Skip mode would then re-upload everything; IfNewer would too).
            _folderSnapshotCache.TryRemove(new KeyValuePair<string,
                Lazy<Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>>>(cacheKey, lazy));
            throw;
        }
    }

    public async Task ExecuteAsync(
        IList<CopyJob> jobs,
        ObservableCollection<CopyResult> results,
        OverwriteMode overwriteMode,
        bool copyVersions,
        int maxParallel,
        int maxVersions,
        CopyMode copyMode,
        CancellationToken cancellationToken,
        IProgress<bool>? onMetadataDone = null,
        bool copyCustomColumns = false,
        List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null,
        bool copyPages = false,
        bool remapPageWebPartUrls = true,
        bool preserveMetadata = true,
        bool copyPermissions = false,
        PermissionCopyService? permissionService = null,
        Dictionary<string, bool>? permissionFlags = null,
        IProgress<(int, int)>? preflightProgress = null,
        IProgress<string>? activityLog = null,
        IProgress<long>? onFilePacked = null,
        IProgress<(int done, int total)>? onFolderProgress = null,
        bool reapplyFolderMetadata = true,
        // Source Modified/Created/Editor/Author, bulk-read alongside bulkFieldCache (same
        // "{listId}:{itemId}" keys) — lets the Migration API custom-fields restamp reuse that
        // already-paid-for scan instead of a per-file GetFileMetadataAsync Graph call. Only
        // consulted in Migration API mode; falls back to a live per-file fetch on a cache miss.
        Dictionary<string, FileMetadata>? sourceMetaCache = null)
    {
        // In SPMI mode the controller semaphore is never used as a download gate
        // (MigrationJobService has its own download controller). Suppress cosmetic step-downs.
        // Migration API engages whenever the mode is selected — independent of the Copy Versions
        // toggle. With versions off we copy current-only via the fast batched path (see maxVersions
        // translation below) rather than silently falling back to slow per-file REST.
        bool isMigrationMode = copyMode == CopyMode.MigrationApi;

        using var controller = new AdaptiveParallelismController(maxParallel);
        controller.LimitChanged += n => ParallelismChanged?.Invoke(n);
        if (activityLog != null && !isMigrationMode)
        {
            int lastLimit = maxParallel;
            controller.LimitChanged += n =>
            {
                bool down = n < lastLimit;
                lastLimit = n;
                activityLog.Report(down
                    ? $"↓ Parallelism: {n}/{maxParallel} (throttled)"
                    : $"⬆ Parallelism: {n}/{maxParallel} (recovering)");
            };
        }
        void onThrottled(TimeSpan delay, int attempt, int max, string? reason)
        {
            if (!isMigrationMode) controller.StepDown(delay);
        }
        // Both handlers MUST come off in a finally: spService outlives this run, so any handler
        // left behind (the old code unsubscribed mid-body, skipped on exception, and never
        // unsubscribed the logging one at all) kept firing on later runs — duplicate throttle log
        // lines and StepDown calls into this run's disposed controller.
        Action<TimeSpan, int, int, string?>? onThrottleLog = null;
        spService.Throttled += onThrottled;
        if (activityLog != null)
        {
            var throttleLogLock  = new object();
            var lastThrottleLog  = DateTimeOffset.MinValue;
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

        try
        {
            await ExecuteCoreAsync();
        }
        catch
        {
            // The core threw (cancel during the scan, migration fatal, …) before it could hand
            // off to — or complete — the metadata phase. Without this, IsUpdatingMetadata stayed
            // true forever: wizard wedged on "updating metadata" and sleep prevention held until
            // app exit. Reporting false is safe: the fire-and-forget metadata pass only starts as
            // the core's final statement, so a throw here means it never started.
            onMetadataDone?.Report(false);
            throw;
        }
        finally
        {
            spService.Throttled -= onThrottled;
            if (onThrottleLog != null) spService.Throttled -= onThrottleLog;
        }
        return;

        async Task ExecuteCoreAsync()
        {
        // Target folder item IDs cached in a previous run go stale if folders were deleted or
        // renamed between runs — and a faulted entry must never poison a fresh run.
        spService.ResetFolderSegmentCache();
        spService.ResetColumnCache();
        _folderSnapshotCache.Clear();

        var allTasks  = new List<(CopyJob job, CopyResult result)>();

        // Empty folders (SourceFileEntry.IsEmptyFolder) are created directly in the scan loop below
        // and never become file jobs in allTasks — so the dirty-path set built from allTasks (further
        // down) never included them, and the separate folder-metadata pass then filtered them out as
        // "not touched this run", leaving them permanently stamped with today's date and the
        // migrating account regardless of preserved source metadata. Newly-created ones are recorded
        // here so their paths can be folded into that same dirty set.
        var newlyCreatedEmptyFolderPaths = new List<string>();

        // Ordinary (non-special, non-root) folder IDENTITIES discovered by the scan for free — see
        // SourceFileEntry.IsFolder — captured here per top-level job (keyed by reference; CopyJob has
        // no custom Equals, so this is exact identity) instead of being rediscovered later by a
        // separate re-walk (Enhanced REST's old EnumerateFoldersAsync pass) or an ancestor "hopsUp"
        // derivation from a sample descendant FILE (SPMI's old scheme). Both engines still do their
        // own per-folder metadata+color fetch using these identities directly — what this eliminates
        // is the extra round trip(s) each engine previously needed just to FIND the right folder item
        // id in the first place (a full redundant tree walk for Enhanced REST; 1-3 parentReference
        // hops plus a ProgId probe per folder for SPMI).
        var scannedFoldersByJob = new Dictionary<CopyJob,
            List<(string driveId, string itemId, string relativePath)>>();

        // Pre-seeded rows (e.g. from a "select individual files" UI flow that creates placeholder
        // rows before the copy starts) are looked up by source path once per top-level job below.
        // Built once, O(n), instead of FindResult's old per-job linear scan of the whole (and
        // growing) results collection — O(n²) on a selection of many individual files.
        var resultsBySourcePath = new Dictionary<string, CopyResult>(StringComparer.Ordinal);
        foreach (var r in results)
            resultsBySourcePath.TryAdd(r.SourcePath, r);

        // Buffer new result rows and flush them to the bound collection in chunks. Adding tens of
        // thousands of rows one at a time via a *synchronous* Dispatcher.Invoke saturates the UI
        // thread and back-pressures enumeration — the progress display appears to "freeze" around
        // ~47k files on huge copies. Chunked async adds collapse 120k UI round-trips into a few
        // hundred and keep both the file listing and the window responsive.
        var pendingResults = new List<CopyResult>(256);
        async Task FlushPendingResultsAsync()
        {
            if (pendingResults.Count == 0) return;
            var chunk = pendingResults.ToArray();
            pendingResults.Clear();
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
            {
                foreach (var r in chunk) results.Add(r);
            }).Task;
        }

        // Expansion-scoped adaptive gate for the source walk, same pattern as the pre-flight gate in
        // MigrationJobService: without it the walk either serialized (the old one-call-at-a-time
        // recursive enumeration — ~30 silent minutes on a 3,000-folder library) or would burst at a
        // fixed width straight back into a depleted throttle budget.
        const int ScanMaxParallelism = 8;
        using var scanController = new AdaptiveParallelismController(ScanMaxParallelism);

        // Diagnostic-only counters (added to investigate escalating scan-phase throttle waits —
        // see activity-20260803-093215.log): total time spent waiting out throttles and how many
        // throttle events fired during this scan, so the "throttle overhead" of a run is visible
        // in its own activity log instead of requiring after-the-fact log-timestamp arithmetic.
        var scanThrottleStatsLock = new object();
        int      scanThrottleEvents    = 0;
        TimeSpan scanThrottleWaitTotal = TimeSpan.Zero;
        void onScanThrottle(TimeSpan delay, int _, int __, string? ___)
        {
            scanController.StepDown(delay);
            lock (scanThrottleStatsLock)
            {
                scanThrottleEvents++;
                scanThrottleWaitTotal += delay;
            }
        }

        // Surfaces the scan's own concurrency ceiling alongside the generic "Graph throttled" log
        // line (which fires from a shared handler with no visibility into which phase/controller is
        // active) — without this, a log review can see THAT the tenant throttled but not whether the
        // scan's own step-down/restore logic is oscillating back into it. Same pattern already used
        // for the main controller (above) and for Analysis/Downloads/Uploads in MigrationJobService.
        if (activityLog != null)
        {
            int lastScanLimit = ScanMaxParallelism;
            scanController.LimitChanged += n =>
            {
                bool down = n < lastScanLimit;
                lastScanLimit = n;
                activityLog.Report(down
                    ? $"↓ Scan concurrency: {n}/{ScanMaxParallelism} (throttled)"
                    : $"⬆ Scan concurrency: {n}/{ScanMaxParallelism} (recovering)");
            };
        }

        bool anyFolderJobs = jobs.Any(j => j.IsFolder);
        var scanStartTime = DateTimeOffset.UtcNow;
        if (anyFolderJobs)
            activityLog?.Report("Scanning source for files to copy...");
        int scannedFiles = 0;
        var lastScanReport = DateTimeOffset.UtcNow;

        // Graph's native /copy action has no overwrite concept — a same-named item at the target
        // just fails the copy outright (nameAlreadyExists). Overwrite clears it first; Skip/IfNewer
        // both leave an existing target alone (a folder-level "newer than" comparison isn't
        // meaningful the way a per-file one is, and this only affects the rare special-folder case).
        // Returns true if the caller should proceed with the native copy, false to skip it.
        async Task<bool> PrepareNativeCopyTargetAsync(string driveId, string parentId, string name)
        {
            if (overwriteMode == OverwriteMode.Overwrite)
            {
                await spService.DeleteChildIfExistsAsync(driveId, parentId, name);
                return true;
            }
            return !await spService.ChildExistsAsync(driveId, parentId, name);
        }

        spService.Throttled += onScanThrottle;
        try
        {
            foreach (var job in jobs)
            {
                if (!job.IsFolder)
                {
                    // Every current caller pre-seeds `results` with a placeholder row for each
                    // directly-selected file before calling ExecuteAsync, so the TryGetValue branch
                    // is what actually runs today. But if a future caller doesn't, CreateResult(job)
                    // was previously added only to allTasks — never to pendingResults/results — so
                    // that file would copy successfully yet never appear in the grid or the saved
                    // report. Flushing it here closes that gap defensively.
                    bool hadExisting = resultsBySourcePath.TryGetValue(job.SourceDisplayPath, out var existingResult);
                    var result = hadExisting ? existingResult! : CreateResult(job);
                    if (!hadExisting)
                    {
                        pendingResults.Add(result);
                        if (pendingResults.Count >= 200) await FlushPendingResultsAsync();
                    }
                    allTasks.Add((job, result));
                }
                else if (await spService.IsRootFolderSpecialContainerAsync(job.SourceDriveId, job.SourceItemId))
                {
                    // The job's OWN root is the special folder (e.g. a notebook selected directly
                    // as the copy source, not discovered as a descendant during a walk) — the
                    // per-child check inside EnumerateFilesForCopyAsync never sees this case since
                    // the walk starts AT this item rather than encountering it as someone's child.
                    // Uses the same cheap package-facet/content-type check as that per-child path —
                    // not GetFolderProgIdAsync's REST probe, which used to run unconditionally on
                    // every directly-selected folder root (including plain, non-special ones) purely
                    // to learn "no" (observed 2026-07-21: this extra read against the SOURCE root was
                    // the suspected cause of the source folder showing an unexpected Modified By
                    // stamp — a plain-folder root has no reason to be probed for a ProgID at all).
                    var folderResult = new CopyResult
                    {
                        FileName   = job.SourceName,
                        SourcePath = job.SourceDisplayPath,
                        TargetPath = job.TargetDisplayPath,
                        Status     = CopyStatus.Copying
                    };
                    pendingResults.Add(folderResult);
                    activityLog?.Report($"Copying special folder '{job.SourceName}' natively (preserves notebook/package association)...");
                    try
                    {
                        var parentId = await spService.GetOrCreateFolderPathAsync(
                            job.TargetDriveId, job.TargetParentItemId, job.TargetSubFolderPath);
                        if (!await PrepareNativeCopyTargetAsync(job.TargetDriveId, parentId, job.SourceName))
                        {
                            folderResult.Status = CopyStatus.Skipped;
                            activityLog?.Report($"⏭ Skipped '{job.SourceName}' — already exists at target");
                        }
                        else
                        {
                            var copyError = await spService.CopyFolderNativeAsync(
                                job.SourceDriveId, job.SourceItemId, job.TargetDriveId, parentId, job.SourceName, cancellationToken);
                            folderResult.Status       = copyError == null ? CopyStatus.Success : CopyStatus.Failed;
                            folderResult.ErrorMessage = copyError;
                            activityLog?.Report(copyError == null
                                ? $"✓ Native copy of '{job.SourceName}' complete"
                                : $"⚠ Native copy of '{job.SourceName}' failed: {copyError}");
                        }
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        folderResult.Status       = CopyStatus.Failed;
                        folderResult.ErrorMessage = ex.Message;
                        activityLog?.Report($"⚠ Native copy of '{job.SourceName}' failed: {ex.Message}");
                    }
                    await FlushPendingResultsAsync();
                }
                else
                {
                    // Seeded unconditionally (not lazily on the first IsFolder hit below) so a job
                    // whose walk finds files but ZERO subfolders — e.g. a single leaf folder with no
                    // nested directories — still gets an entry here. Without this, such a job was
                    // entirely absent from scannedFoldersByJob, so the ownKey loop further down (which
                    // iterates `foreach (var (folderJob, folders) in scannedFoldersByJob)`) never ran
                    // for it either, leaving the job's OWN top-level folder with no identity at all —
                    // not just a missing ancestor, but the job's actual root (confirmed 2026-08-04:
                    // a "java" folder holding 11 flat files, no subfolders, never appeared in
                    // scannedFolderIdentities even though it was the copy's own target).
                    var jobFolders = new List<(string driveId, string itemId, string relativePath)>();
                    scannedFoldersByJob[job] = jobFolders;

                    // Every file in the same source folder computes the IDENTICAL TargetSubFolderPath
                    // (see ComputeTargetSubFolder) — without this cache, a 250k-file/5,000-folder job
                    // allocates 250k separate-but-equal-content strings for a value that only has 5,000
                    // distinct contents. Keyed by the file's directory portion of relativePath, scoped
                    // to this top-level job (ComputeTargetSubFolder's other inputs are fixed per job).
                    var targetSubFolderCache = new Dictionary<string, string>(StringComparer.Ordinal);

                    await foreach (var entry in spService.EnumerateFilesForCopyAsync(
                        job.SourceDriveId, job.SourceItemId, "", scanController, cancellationToken))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var (driveId, itemId, name, relativePath) = (entry.DriveId, entry.ItemId, entry.Name, entry.RelativePath);

                        // Identity-only entry for an ordinary folder (see SourceFileEntry.IsFolder) —
                        // never creates a grid row or copy job here; just record its (driveId, itemId)
                        // for the folder-metadata passes further down (both engines), which still do
                        // their own per-folder metadata+color fetch — this only saves them from having
                        // to FIND the folder's item id themselves. Excluded from scannedFiles: counting
                        // it there inflated the "file(s) found" tally with folders that never become a
                        // grid row, so it drifted far ahead of the progress bar's real total (CopyResults.Count).
                        if (entry.IsFolder)
                        {
                            jobFolders.Add((driveId, itemId, relativePath));
                            continue;
                        }

                        scannedFiles++;
                        if (DateTimeOffset.UtcNow - lastScanReport >= TimeSpan.FromSeconds(3))
                        {
                            lastScanReport = DateTimeOffset.UtcNow;
                            activityLog?.Report($"Scanning source: {scannedFiles:N0} file(s) found so far...");
                        }

                        // Special folder (e.g. a OneNote notebook — see SourceFileEntry.IsSpecialFolder):
                        // copy it as a single native Graph operation right here in the scan loop rather
                        // than expanding it into per-file CopyJobs — SPMI/Enhanced REST would silently
                        // lose the property that makes it a notebook (see CopyFolderNativeAsync).
                        if (entry.IsSpecialFolder)
                        {
                            var parentSubFolder = ComputeTargetSubFolder(
                                relativePath, job.SourceName, job.IsLibrary, job.TargetSubFolderPath);
                            var folderResult = new CopyResult
                            {
                                FileName   = name,
                                SourcePath = $"{job.SourceDisplayPath}/{relativePath}",
                                TargetPath = $"{job.TargetDisplayPath}/{relativePath}",
                                Status     = CopyStatus.Copying
                            };
                            pendingResults.Add(folderResult);
                            activityLog?.Report($"Copying special folder '{name}' natively (preserves notebook/package association)...");
                            try
                            {
                                var parentId = await spService.GetOrCreateFolderPathAsync(
                                    job.TargetDriveId, job.TargetParentItemId, parentSubFolder);
                                if (!await PrepareNativeCopyTargetAsync(job.TargetDriveId, parentId, name))
                                {
                                    folderResult.Status = CopyStatus.Skipped;
                                    activityLog?.Report($"⏭ Skipped '{name}' — already exists at target");
                                }
                                else
                                {
                                    var copyError = await spService.CopyFolderNativeAsync(
                                        driveId, itemId, job.TargetDriveId, parentId, name, cancellationToken);
                                    folderResult.Status       = copyError == null ? CopyStatus.Success : CopyStatus.Failed;
                                    folderResult.ErrorMessage = copyError;
                                    activityLog?.Report(copyError == null
                                        ? $"✓ Native copy of '{name}' complete"
                                        : $"⚠ Native copy of '{name}' failed: {copyError}");
                                }
                            }
                            catch (Exception ex) when (ex is not OperationCanceledException)
                            {
                                folderResult.Status       = CopyStatus.Failed;
                                folderResult.ErrorMessage = ex.Message;
                                activityLog?.Report($"⚠ Native copy of '{name}' failed: {ex.Message}");
                            }
                            if (pendingResults.Count >= 200) await FlushPendingResultsAsync();
                            continue;
                        }

                        // A folder with no files anywhere in its subtree (see SourceFileEntry.IsEmptyFolder) —
                        // nothing else would ever create it at the target, since every other entry only
                        // provisions its ancestor chain as a side effect of copying an actual file.
                        if (entry.IsEmptyFolder)
                        {
                            var displayName = string.IsNullOrEmpty(relativePath) ? job.SourceName : name;
                            var emptyFolderTarget = ComputeTargetFolderPath(
                                relativePath, job.SourceName, job.IsLibrary, job.TargetSubFolderPath);
                            var folderResult = new CopyResult
                            {
                                FileName   = displayName,
                                SourcePath = string.IsNullOrEmpty(relativePath)
                                    ? job.SourceDisplayPath : $"{job.SourceDisplayPath}/{relativePath}",
                                TargetPath = string.IsNullOrEmpty(relativePath)
                                    ? job.TargetDisplayPath : $"{job.TargetDisplayPath}/{relativePath}",
                                Status     = CopyStatus.Copying
                            };
                            pendingResults.Add(folderResult);
                            try
                            {
                                // Unconditionally calling GetOrCreateFolderPathAsync (a get-OR-create) used to
                                // report Success on every run regardless of whether the folder already existed
                                // — on an otherwise fully up-to-date re-run, every empty folder in the source
                                // still showed as a fresh "Success", with no way to tell it apart from a real
                                // creation. Check existence first (the same read-only path-walk verification
                                // already uses) so an already-there folder reports Skipped, matching how every
                                // other already-exists case in this file is reported.
                                var existingId = await spService.ResolveItemIdByPathAsync(
                                    job.TargetDriveId, job.TargetParentItemId, emptyFolderTarget);
                                if (existingId != null)
                                {
                                    folderResult.Status = CopyStatus.Skipped;
                                    activityLog?.Report($"⏭ Skipped '{displayName}' — already exists at target");
                                }
                                else
                                {
                                    await spService.GetOrCreateFolderPathAsync(
                                        job.TargetDriveId, job.TargetParentItemId, emptyFolderTarget);
                                    folderResult.Status = CopyStatus.Success;
                                    newlyCreatedEmptyFolderPaths.Add(emptyFolderTarget);
                                }
                            }
                            catch (Exception ex) when (ex is not OperationCanceledException)
                            {
                                folderResult.Status       = CopyStatus.Failed;
                                folderResult.ErrorMessage = ex.Message;
                                activityLog?.Report($"⚠ Creating empty folder '{displayName}' failed: {ex.Message}");
                            }
                            if (pendingResults.Count >= 200) await FlushPendingResultsAsync();
                            continue;
                        }

                        var fileDirKey = System.IO.Path.GetDirectoryName(relativePath)?.Replace('\\', '/') ?? string.Empty;
                        if (!targetSubFolderCache.TryGetValue(fileDirKey, out var targetSubFolder))
                        {
                            targetSubFolder = ComputeTargetSubFolder(
                                relativePath, job.SourceName, job.IsLibrary, job.TargetSubFolderPath);
                            targetSubFolderCache[fileDirKey] = targetSubFolder;
                        }
                        var fileJob = new CopyJob
                        {
                            SourceDriveId                  = driveId,
                            SourceItemId                   = itemId,
                            SourceName                     = name,
                            SourceModified                 = entry.Modified,
                            SourceSize                     = entry.Size,
                            SourceSiteUrl                  = job.SourceSiteUrl,
                            SourceDisplayPath              = $"{job.SourceDisplayPath}/{relativePath}",
                            TargetDriveId                  = job.TargetDriveId,
                            TargetParentItemId             = job.TargetParentItemId,
                            TargetSiteId                   = job.TargetSiteId,
                            TargetSiteUrl                  = job.TargetSiteUrl,
                            TargetSubFolderPath            = targetSubFolder,
                            TargetLibraryServerRelativeUrl = job.TargetLibraryServerRelativeUrl,
                            TargetDisplayPath              = $"{job.TargetDisplayPath}/{relativePath}",
                            IsPage                         = copyPages,
                            IsFolder                       = false
                        };

                        var result = new CopyResult
                        {
                            FileName   = name,
                            SourcePath = fileJob.SourceDisplayPath,
                            TargetPath = fileJob.TargetDisplayPath
                        };

                        pendingResults.Add(result);
                        allTasks.Add((fileJob, result));
                        if (pendingResults.Count >= 200) await FlushPendingResultsAsync();
                    }
                }
            }
        }
        finally
        {
            spService.Throttled -= onScanThrottle;
        }
        if (anyFolderJobs)
        {
            var scanElapsed = DateTimeOffset.UtcNow - scanStartTime;
            int      throttleEvents;
            TimeSpan throttleWaitTotal;
            lock (scanThrottleStatsLock) { throttleEvents = scanThrottleEvents; throttleWaitTotal = scanThrottleWaitTotal; }
            // Overhead ratio = fraction of the scan's wall-clock time spent waiting out throttles —
            // a coarse but immediate signal of whether it's worth tuning scan concurrency/backoff at
            // all, without needing an external log-timestamp analysis after the fact.
            var overheadRatio = scanElapsed > TimeSpan.Zero
                ? throttleWaitTotal.TotalSeconds / scanElapsed.TotalSeconds : 0;
            activityLog?.Report($"Source scan complete: {scannedFiles:N0} file(s) found in {scanElapsed.TotalSeconds:0}s"
                + (throttleEvents > 0
                    ? $" ({throttleEvents} throttle event(s), {throttleWaitTotal.TotalSeconds:0}s waited, {overheadRatio:P0} throttle overhead)"
                    : ""));
        }
        await FlushPendingResultsAsync();

        // Flattens scannedFoldersByJob into the SAME key shape SPMI's folder-metadata correction
        // already keys by (TargetSubFolderPath.Trim('/') — see MigrationJobService's directFolderGroups):
        // ComputeTargetFolderPath treats a folder's own relativePath exactly like a hypothetical file
        // living directly inside it, which is the same computation a file job's TargetSubFolderPath
        // already uses (ComputeTargetSubFolder) — so keys from both sources line up. Passed to
        // MigrationJobService so it can skip its ancestor "hopsUp" derivation entirely and go straight
        // to the folder's real item id.
        var scannedFolderIdentities = new Dictionary<string, (string driveId, string itemId)>(StringComparer.OrdinalIgnoreCase);
        foreach (var (folderJob, folders) in scannedFoldersByJob)
        {
            // WalkFilesForCopyAsync only emits an IsFolder entry for folders it discovers as
            // CHILDREN of something else — the folder it was actually asked to walk (the job's own
            // top-level folder) never appears in `folders` at all. Without this, that one folder is
            // absent from scannedFolderIdentities, so MigrationJobService can never fetch its real
            // date/author and it falls back to the manifest's 2000-01-01 placeholder (renders as
            // "Dec 31, 1999" locally) — even though every descendant folder gets its real metadata.
            // Not needed for a library job: the library root is a separate concern, keyed by ""
            // and fetched directly in MigrationJobService.
            if (!folderJob.IsLibrary)
            {
                var ownKey = ComputeTargetFolderPath(
                    "", folderJob.SourceName, folderJob.IsLibrary, folderJob.TargetSubFolderPath).Trim('/');
                if (ownKey.Length > 0)
                    scannedFolderIdentities[ownKey] = (folderJob.SourceDriveId, folderJob.SourceItemId);
            }
            foreach (var (driveId, itemId, relativePath) in folders)
            {
                var key = ComputeTargetFolderPath(
                    relativePath, folderJob.SourceName, folderJob.IsLibrary, folderJob.TargetSubFolderPath).Trim('/');
                if (key.Length > 0) scannedFolderIdentities[key] = (driveId, itemId);
            }
        }

        if (copyMode == CopyMode.MigrationApi)
        {
            // Mode A: batch all files into migration jobs. When Copy Versions is off, callers pass
            // maxVersions = 0; the Migration path reads 0 as "all versions", so translate it to 1
            // (current version only) to honor the toggle. With versions on, maxVersions is already
            // the intended cap (and 0 there legitimately means "all versions").
            int migrationMaxVersions = copyVersions ? maxVersions : 1;
            // The SPMI manifest no longer carries per-item <Fields> (standalone SPListItem objects
            // caused "Missing file info" import failures — see MigrationPackageBuilder), so custom
            // column values can't be applied as part of the import itself. Instead they're applied
            // in a REST post-import pass below, the same mechanism Enhanced REST mode uses per file.
            await migrationJobService.ExecuteAsync(allTasks, overwriteMode, migrationMaxVersions, maxParallel, cancellationToken,
                copyCustomColumns, columnMappings, bulkFieldCache, preflightProgress, activityLog, onFilePacked,
                reapplyFolderMetadata: preserveMetadata && reapplyFolderMetadata,
                scannedFolderIdentities: scannedFolderIdentities);

            // Permissions: run after the migration job completes so the target items exist.
            // We can't use Graph item IDs here (the migration API doesn't surface them), so we
            // resolve the target via its server-relative URL using the SP REST file endpoint.
            // Skip the whole pass when the bulk flags say no item has unique permissions — the
            // per-file GetSharePointIdsAsync resolution below is one Graph round-trip per file,
            // which on a 100k-file run is pure waste when 0 items need it. When some do, resolve
            // in parallel (bounded) instead of strictly sequentially.
            if (copyPermissions && permissionService != null && permissionFlags != null &&
                permissionFlags.Values.Any(v => v))
            {
                var permCandidates = allTasks.Where(t =>
                    t.result.Status == CopyStatus.Success ||
                    // Files skipped as "Up to date" by Copy-if-newer still get their
                    // permissions refreshed — only permission changes may have occurred.
                    (t.result.Status == CopyStatus.Skipped && t.result.ErrorMessage == CopyResult.UpToDate)).ToList();

                await Parallel.ForEachAsync(permCandidates,
                    new ParallelOptions { MaxDegreeOfParallelism = 8, CancellationToken = cancellationToken },
                    async (t, ct) =>
                {
                    var (job, result) = t;
                    try
                    {
                        var srcIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
                        if (!srcIds.HasValue) return;
                        var flagKey = SharePointService.PermissionFlagKey(srcIds.Value.listId, srcIds.Value.listItemId);
                        if (!permissionFlags.TryGetValue(flagKey, out var hu) || !hu) return;

                        var sub = job.TargetSubFolderPath?.Trim('/');
                        var tgtRelUrl = string.IsNullOrEmpty(sub)
                            ? $"{job.TargetLibraryServerRelativeUrl}/{job.SourceName}"
                            : $"{job.TargetLibraryServerRelativeUrl.TrimEnd('/')}/{sub}/{job.SourceName}";

                        // *ByServerRelativePath(decodedurl=…): the *Url variant cannot resolve
                        // paths containing '#'/'%'/'+' even when percent-encoded (see
                        // ServerRelativePathArg), so those files silently lost their permissions.
                        var perm = await permissionService.CopyObjectPermissionsAsync(
                            job.SourceSiteUrl, job.TargetSiteUrl,
                            $"web/lists('{srcIds.Value.listId}')/items({srcIds.Value.listItemId})",
                            $"web/GetFileByServerRelativePath(decodedurl='{Uri.EscapeDataString(tgtRelUrl.Replace("'", "''"))}')/ListItemAllFields",
                            hasUniquePermissions: true,
                            job.SourceName, ct);

                        if (perm.HasActivity)
                            AddPermissionRow(results, perm, result);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch { /* non-fatal */ }
                });
            }

            // Custom columns: same constraint and pattern as the permissions pass above — the
            // migration API doesn't surface Graph item IDs for imported files, so each target item
            // is resolved via its REST server-relative path (ApplyFileCustomFieldsByPathAsync)
            // instead of the Graph-based ApplyFileCustomFieldsAsync that Enhanced REST mode uses.
            if (copyCustomColumns && bulkFieldCache != null && columnMappings != null && bulkFieldCache.Count > 0)
            {
                var fieldCandidates = allTasks.Where(t =>
                    t.result.Status == CopyStatus.Success ||
                    (t.result.Status == CopyStatus.Skipped && t.result.ErrorMessage == CopyResult.UpToDate)).ToList();

                if (fieldCandidates.Count > 0)
                {
                    activityLog?.Report($"Applying custom column values to {fieldCandidates.Count:N0} file(s)...");
                    // Target listId is the same for every file in a given target library — resolve
                    // it once per distinct (site, library) pair rather than once per file.
                    var targetListIdCache = new System.Collections.Concurrent.ConcurrentDictionary<string, Task<string>>();

                    await Parallel.ForEachAsync(fieldCandidates,
                        new ParallelOptions { MaxDegreeOfParallelism = 8, CancellationToken = cancellationToken },
                        async (t, ct) =>
                    {
                        var (job, result) = t;
                        try
                        {
                            var srcIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
                            if (!srcIds.HasValue) return;
                            var cacheKey = $"{srcIds.Value.listId}:{srcIds.Value.listItemId}";
                            if (!bulkFieldCache.TryGetValue(cacheKey, out var customFields) || customFields.Count == 0)
                                return;

                            var sub = job.TargetSubFolderPath?.Trim('/');
                            var tgtRelUrl = string.IsNullOrEmpty(sub)
                                ? $"{job.TargetLibraryServerRelativeUrl}/{job.SourceName}"
                                : $"{job.TargetLibraryServerRelativeUrl.TrimEnd('/')}/{sub}/{job.SourceName}";

                            var listIdKey = $"{job.TargetSiteUrl}|{job.TargetLibraryServerRelativeUrl}";
                            var targetListId = await targetListIdCache.GetOrAdd(listIdKey,
                                _ => spService.GetListIdByServerRelativeUrlAsync(job.TargetSiteUrl, job.TargetLibraryServerRelativeUrl));

                            // ValidateUpdateListItem bumps Modified/Editor to the migrating account and
                            // now — the Migration API import already stamped the correct Modified/Editor
                            // via SPMI's <File> element, so this write must re-stamp them from the
                            // SOURCE file's own metadata. Folded into the SAME call as the custom-field
                            // write (via `restamp`) rather than a separate call afterward — one call
                            // instead of two, and one list-item-ID resolution instead of two.
                            //
                            // Prefer the bulk-read cache (same scan that fetched customFields, same
                            // cacheKey) over a live Graph call — falls back to GetFileMetadataAsync
                            // only on a cache miss (e.g. sourceMetaCache wasn't built, or this item
                            // was added to the source library after the bulk scan ran).
                            FileMetadata? srcMeta = null;
                            if (preserveMetadata)
                            {
                                srcMeta = sourceMetaCache != null && sourceMetaCache.TryGetValue(cacheKey, out var cachedMeta)
                                    ? cachedMeta
                                    : await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
                            }
                            var cfErr = await spService.ApplyFileCustomFieldsByPathAsync(
                                job.TargetSiteUrl, targetListId, tgtRelUrl, customFields, columnMappings, srcMeta, ct);
                            result.CustomFieldStatus  = cfErr != null ? CopyStatus.Failed : CopyStatus.Success;
                            result.CustomFieldDetails = cfErr;
                            if (cfErr != null) result.ErrorMessage ??= cfErr;
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            result.CustomFieldStatus  = CopyStatus.Failed;
                            result.CustomFieldDetails = ex.Message;
                            result.ErrorMessage ??= $"Custom fields: {ex.Message}";
                        }
                    });
                }
            }
        }
        else
        {
            // Mode B: enhanced REST, parallel per-file. Parallel.ForEachAsync rather than
            // allTasks.Select(...) + Task.WhenAll: the Select form launches EVERY task immediately —
            // on a 250k-file run that's 250k async state machines plus a 250k-node SemaphoreSlim wait
            // queue plus a 250k-element array for WhenAll, well over 100 MB of pure scheduling
            // overhead before any file transfers, even though CopySingleFileAsync's own
            // AdaptiveParallelismController gate still limits how many run at once. ForEachAsync
            // only materializes MaxDegreeOfParallelism bodies at a time; the controller still governs
            // live (adaptive, shrink-on-throttle) width exactly as before — per-item exceptions are
            // already fully contained inside CopySingleFileAsync, so ForEachAsync's fail-fast
            // semantics don't change behavior.
            await Parallel.ForEachAsync(allTasks,
                new ParallelOptions { MaxDegreeOfParallelism = maxParallel, CancellationToken = cancellationToken },
                async (t, ct) => await CopySingleFileAsync(t.job, t.result, overwriteMode, copyVersions, maxVersions, controller, ct,
                    copyCustomColumns, columnMappings, bulkFieldCache, copyPages, remapPageWebPartUrls, preserveMetadata,
                    copyPermissions, permissionService, permissionFlags, permissionResults: results));
        }

        // One combined summary regardless of copy mode — both modes set CustomFieldStatus per row
        // above (Migration API via the post-import REST pass, Enhanced REST via CopySingleFileAsync).
        if (copyCustomColumns)
        {
            int customFieldFailures = allTasks.Count(t => t.result.CustomFieldStatus == CopyStatus.Failed);
            if (customFieldFailures > 0)
                activityLog?.Report($"⚠ Custom column values could not be fully applied for {customFieldFailures:N0} file(s) — see the Custom Fields column for details");
        }

        // SPMI already stamps folder timestamps via the manifest's TimeLastModified / TimeCreated /
        // Author / ModifiedBy attributes on <SPFolder> elements during import — no metadata
        // post-processing needed. Folder-level unique PERMISSIONS, however, are not representable
        // in the manifest, so with permissions enabled run the folder pass in permission-only mode
        // (applyMetadata: false) — previously folders with broken inheritance silently kept target
        // defaults in this mode.
        if (isMigrationMode)
        {
            var spmiFolderJobs = jobs.Where(j => j.IsFolder).ToList();
            if (copyPermissions && permissionService != null && spmiFolderJobs.Count > 0)
                _ = ApplyAllFolderMetadataAsync(spmiFolderJobs, maxParallel, onMetadataDone, cancellationToken,
                    copyPermissions, permissionService, results, onFolderProgress,
                    dirtyFolderPaths: null, applyMetadata: false, activityLog: activityLog,
                    scannedFoldersByJob: scannedFoldersByJob);
            else
                onMetadataDone?.Report(true);
            return;
        }

        // For REST mode: only update folders that received at least one successful copy.
        // Build an ancestor-inclusive set from successful file job paths so we skip
        // every clean branch (e.g. unchanged folders when running "If Newer").
        static IEnumerable<string> AncestorInclusivePaths(string path)
        {
            var parts = path.Split('/');
            return Enumerable.Range(1, parts.Length).Select(i => string.Join("/", parts.Take(i)));
        }
        var dirtyFolderPaths = allTasks
            .Where(t => t.result.Status == CopyStatus.Success
                     && !string.IsNullOrEmpty(t.job.TargetSubFolderPath))
            .SelectMany(t => AncestorInclusivePaths(t.job.TargetSubFolderPath!))
            // Newly-created empty folders never appear in allTasks (see
            // newlyCreatedEmptyFolderPaths' declaration) — folded in here so they aren't filtered
            // out of the folder-metadata pass as "not touched this run".
            .Concat(newlyCreatedEmptyFolderPaths.Where(p => !string.IsNullOrEmpty(p)).SelectMany(AncestorInclusivePaths))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var folderJobs = jobs.Where(j => j.IsFolder).ToList();
        bool anyFileCopied = results.Any(r => r.Status == CopyStatus.Success);
        // Permissions must not be gated on metadata options: with "Preserve metadata" off, or an
        // If-Newer re-run where every file was up to date but folder permissions changed at
        // source, folder permissions were silently skipped (files, by contrast, refresh their
        // permissions even when skipped as up to date). With permissions on, walk ALL folders
        // (dirty tracking only knows about file copies, not permission changes).
        bool wantsFolderPermissions = copyPermissions && permissionService != null;
        // reapplyFolderMetadata previously did nothing in this (Enhanced REST) branch — it was
        // consulted only on the Migration API path (see the call into migrationJobService above).
        // Its own tooltip promises "repairs folder dates/authors/color for every folder in the
        // selection on each run — needed to fix folder metadata on an already-copied target". With
        // it on, run the pass regardless of anyFileCopied and walk the WHOLE tree (dirtyFolderPaths
        // null), matching what SPMI already does; with it off, keep the existing
        // only-touch-what-changed behavior.
        bool wantsMetadataPass = preserveMetadata && (anyFileCopied || reapplyFolderMetadata);
        if (folderJobs.Count > 0 && (wantsMetadataPass || wantsFolderPermissions))
            _ = ApplyAllFolderMetadataAsync(folderJobs, maxParallel, onMetadataDone, cancellationToken,
                copyPermissions, permissionService, results, onFolderProgress,
                dirtyFolderPaths: (wantsFolderPermissions || reapplyFolderMetadata || dirtyFolderPaths.Count == 0) ? null : dirtyFolderPaths,
                applyMetadata: wantsMetadataPass, activityLog: activityLog,
                scannedFoldersByJob: scannedFoldersByJob);
        else
            onMetadataDone?.Report(true);
        } // end ExecuteCoreAsync
    }

    private async Task ApplyAllFolderMetadataAsync(
        IEnumerable<CopyJob> folderJobs, int maxParallel,
        IProgress<bool>? onDone, CancellationToken ct,
        bool copyPermissions = false,
        PermissionCopyService? permissionService = null,
        ObservableCollection<CopyResult>? permissionResults = null,
        IProgress<(int done, int total)>? folderProgress = null,
        HashSet<string>? dirtyFolderPaths = null,
        bool applyMetadata = true,
        IProgress<string>? activityLog = null,
        Dictionary<CopyJob, List<(string driveId, string itemId, string relativePath)>>? scannedFoldersByJob = null)
    {
        bool completed = true;
        // This pass was the one phase in the app with NO throttle protection at all: launched
        // fire-and-forget (`_ =`) after ExecuteCoreAsync already returned and its own Throttled
        // subscription was torn down, so its inner Parallel.ForEachAsync ran at raw maxParallel with
        // no adaptive gate — full width straight into a tenant the copy phase had just finished
        // depleting. Same soft-start + throttle-window-inheriting gate every other analysis phase
        // uses (see CreateThrottleAwareGate).
        using var gate = spService.CreateThrottleAwareGate(maxParallel, Math.Min(Math.Max(1, maxParallel), 2));
        void onThrottle(TimeSpan delay, int _, int __, string? ___) => gate.StepDown(delay);
        spService.Throttled += onThrottle;
        try
        {
            int[] done  = { 0 };
            int[] total = { 0 };
            foreach (var job in folderJobs)
            {
                List<(string driveId, string itemId, string relativePath)>? jobFolders = null;
                scannedFoldersByJob?.TryGetValue(job, out jobFolders);
                await ApplyFolderMetadataRecursiveAsync(job, maxParallel, gate, ct,
                    copyPermissions, permissionService, permissionResults,
                    done, total, folderProgress, dirtyFolderPaths, applyMetadata,
                    jobFolders);
            }
        }
        catch (OperationCanceledException) { completed = false; }
        catch (Exception ex)
        {
            // Previously a bare `catch { }` left `completed = true` — an exception on folder #1 of
            // thousands (throttle retries exhausted, transient error) abandoned every remaining
            // folder in the sequential `foreach` above, yet the wizard reported "Folder metadata
            // updated" with no error visible anywhere, since this pass took no activityLog either.
            completed = false;
            activityLog?.Report($"⚠ Folder metadata pass stopped early: {ex.Message}");
        }
        finally { spService.Throttled -= onThrottle; }
        onDone?.Report(completed);
    }

    private async Task CopySingleFileAsync(
        CopyJob job,
        CopyResult result,
        OverwriteMode overwriteMode,
        bool copyVersions,
        int maxVersions,
        AdaptiveParallelismController controller,
        CancellationToken ct,
        bool copyCustomColumns = false,
        List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null,
        bool copyPages = false,
        bool remapPageWebPartUrls = true,
        bool preserveMetadata = true,
        bool copyPermissions = false,
        PermissionCopyService? permissionService = null,
        Dictionary<string, bool>? permissionFlags = null,
        ObservableCollection<CopyResult>? permissionResults = null)
    {
        bool semaphoreAcquired = false;
        try { await controller.WaitAsync(ct); semaphoreAcquired = true; }
        catch (OperationCanceledException)
        {
            // Cancelled, not Skipped: this item never started (or didn't finish) — Skipped
            // otherwise means "compared and found already up to date." See CopyStatus.Cancelled.
            result.Status       = CopyStatus.Cancelled;
            result.ErrorMessage = "Cancelled";
            return;
        }
        // Set once the file itself has copied/skipped successfully — used below to tell "cancelled
        // before the file ever landed" from "file landed fine, cancellation only interrupted the
        // best-effort permission refresh that follows". Without this distinction, a fully-copied
        // file whose permission step was cancelled got its Success/Skipped status silently
        // overwritten with Cancelled ("never actually attempted" per CopyResult's own status docs).
        bool copyCompleted = false;
        try
        {
            result.Status = CopyStatus.Copying;

            var targetParentId = await ResolveTargetParentAsync(job, ct);

            // Set when IfNewer finds the target already current: the file copy is skipped
            // but permissions (below) still refresh when enabled.
            string? upToDateItemId = null;

            if (overwriteMode != OverwriteMode.Overwrite)
            {
                // One paginated listing per TARGET FOLDER (cached across every file in it — see
                // GetOrBuildFolderSnapshotAsync) instead of a GetFileInfoAsync Graph call per FILE.
                var folderSnapshot = await GetOrBuildFolderSnapshotAsync(job.TargetDriveId, targetParentId);
                if (folderSnapshot.TryGetValue(job.SourceName, out var existing))
                {
                    if (overwriteMode == OverwriteMode.Skip)
                    {
                        result.Status = CopyStatus.Skipped;
                        return;
                    }
                    // IfNewer: copy only when the source changed since the target was written.
                    // job.SourceModified was already captured by the scan (SourceFileEntry.Modified,
                    // see the walk below) — fall back to a per-file Graph read only on the rare miss
                    // (an individually-selected file, which never goes through the walk, or a file
                    // whose metadata fetch failed during the scan).
                    var srcModified = job.SourceModified
                        ?? (await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId)).ModifiedDateTime;
                    if (srcModified is { } sm && existing.Modified is { } tgtModified &&
                        TimestampComparer.IsUpToDate(sm, tgtModified))
                        upToDateItemId = existing.ItemId;
                }
            }

            string? targetGraphItemId;
            if (upToDateItemId != null)
            {
                targetGraphItemId   = upToDateItemId;
                result.Status       = CopyStatus.Skipped;
                result.ErrorMessage = CopyResult.UpToDate;
            }
            else
            {
                // Whether uploads should replace an existing target file. With Skip we returned
                // above if the file existed; with IfNewer we only reach here when replacing.
                bool overwrite = overwriteMode != OverwriteMode.Skip;

                // When overwriting with version history: delete the file first so the imported
                // versions replace the history rather than being appended to it.
                if (overwrite && copyVersions)
                    await spService.DeleteFileIfExistsAsync(job.TargetDriveId, targetParentId, job.SourceName);

                if (copyVersions)
                    targetGraphItemId = await CopyWithVersionsEnhancedRestAsync(job, result, targetParentId, overwrite, maxVersions, ct,
                        copyCustomColumns, columnMappings, bulkFieldCache, preserveMetadata);
                else
                    targetGraphItemId = await CopyCurrentVersionAsync(job, result, targetParentId, overwrite, ct,
                        copyCustomColumns, columnMappings, bulkFieldCache, copyPages, remapPageWebPartUrls, preserveMetadata);

                result.Status = CopyStatus.Success;
            }
            copyCompleted = true;

            // Per-file permission copy (skipped if not enabled or file has inherited permissions)
            if (copyPermissions && permissionService != null && !string.IsNullOrEmpty(targetGraphItemId))
            {
                try
                {
                    var srcIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
                    if (srcIds.HasValue)
                    {
                        var hasUnique = permissionFlags != null &&
                            permissionFlags.TryGetValue(
                                SharePointService.PermissionFlagKey(srcIds.Value.listId, srcIds.Value.listItemId),
                                out var hu) && hu;
                        if (hasUnique)
                        {
                            var tgtIds = await spService.GetSharePointIdsAsync(job.TargetDriveId, targetGraphItemId);
                            if (tgtIds.HasValue)
                            {
                                var perm = await permissionService.CopyObjectPermissionsAsync(
                                    job.SourceSiteUrl, job.TargetSiteUrl,
                                    $"web/lists('{srcIds.Value.listId}')/items({srcIds.Value.listItemId})",
                                    $"web/lists('{tgtIds.Value.listId}')/items({tgtIds.Value.listItemId})",
                                    hasUniquePermissions: true,
                                    job.SourceName, ct);
                                if ((perm.HasActivity) && permissionResults != null)
                                    AddPermissionRow(permissionResults, perm, result);
                            }
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch { /* non-fatal — permissions best-effort */ }
            }
        }
        catch (OperationCanceledException)
        {
            if (!copyCompleted)
            {
                // Cancelled, not Skipped: this item never started (or didn't finish) — Skipped
                // otherwise means "compared and found already up to date." See CopyStatus.Cancelled.
                result.Status       = CopyStatus.Cancelled;
                result.ErrorMessage = "Cancelled";
            }
            // else: the file itself already copied/skipped successfully (result.Status is already
            // Success or Skipped) — cancellation only interrupted the best-effort permission
            // refresh afterward. Leave the real outcome in place rather than reporting a fully
            // present file as "never actually attempted".
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError oe)
        {
            var detail = oe.Error?.Message ?? oe.Message;
            System.Diagnostics.Debug.WriteLine($"[CopySingle] ODataError HTTP {oe.ResponseStatusCode}: code={oe.Error?.Code}, message={detail}");
            result.Status       = CopyStatus.Failed;
            result.ErrorMessage = $"SharePoint error ({oe.ResponseStatusCode}): {detail}";
        }
        catch (Exception ex)
        {
            result.Status       = CopyStatus.Failed;
            result.ErrorMessage = ex.Message;
        }
        finally
        {
            if (semaphoreAcquired) controller.Release();
        }
    }

    private async Task<string> ResolveTargetParentAsync(CopyJob job, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(job.TargetParentItemId))
            throw new Exception("No target parent folder specified.");

        if (string.IsNullOrEmpty(job.TargetSubFolderPath))
            return job.TargetParentItemId;

        return await spService.GetOrCreateFolderPathAsync(
            job.TargetDriveId, job.TargetParentItemId, job.TargetSubFolderPath);
    }

    private async Task<string?> CopyCurrentVersionAsync(
        CopyJob job, CopyResult result, string targetParentId, bool overwrite, CancellationToken ct,
        bool copyCustomColumns = false, List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null,
        bool copyPages = false, bool remapPageWebPartUrls = true, bool preserveMetadata = true)
    {
        ct.ThrowIfCancellationRequested();
        System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] START: {job.SourceName} isPage={job.IsPage}");

        var metadata = preserveMetadata
            ? await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId)
            : new FileMetadata();
        System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] metadata fetched");

        string targetItemId;
        int    targetSitePagesId = 0;

        if (job.IsPage)
        {
            if (string.IsNullOrEmpty(job.TargetLibraryServerRelativeUrl))
                throw new Exception("Cannot create page: target library server-relative URL is not set.");
            var targetFolderRelUrl = string.IsNullOrEmpty(job.TargetSubFolderPath)
                ? job.TargetLibraryServerRelativeUrl
                : $"{job.TargetLibraryServerRelativeUrl.TrimEnd('/')}/{job.TargetSubFolderPath}";

            // Pre-fetch source canvas BEFORE creating the stub.
            // Any file operation between CreatePageStub and SavePage (e.g. PatchFileSystemDate
            // via Graph) ends the SitePages editing session, causing SavePage to return 409.
            // By fetching first we can call SavePage the instant the stub exists.
            PageMetadata? pageMeta = null;
            string? metaErr = null;
            if (copyPages && !string.IsNullOrEmpty(job.SourceSiteUrl))
            {
                var sourceLibRel = await spService.GetLibraryServerRelativeUrlAsync(job.SourceDriveId);
                var pageRel = $"{sourceLibRel}/{job.SourceName}";
                System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] pre-fetching source canvas…");
                (pageMeta, metaErr) = await spService.GetPageMetadataAsync(job.SourceSiteUrl, pageRel);
                System.Diagnostics.Debug.WriteLine(
                    $"[CopyCurrentVersion] GetPageMetadata: {(pageMeta == null ? $"null — {metaErr}" : $"CanvasContent1={pageMeta.CanvasContent1?.Length ?? 0} chars")}");
            }

            System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] page: creating stub in {targetFolderRelUrl}…");
            (targetItemId, targetSitePagesId) = await spService.CreatePageStubAsync(
                job.TargetSiteUrl, targetFolderRelUrl,
                job.TargetDriveId, targetParentId, job.SourceName, overwrite);
            System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] stub created: graphItemId={targetItemId} sitePagesId={targetSitePagesId}");

            // SavePage + Publish immediately (do not allow any other file operation in between)
            if (pageMeta != null)
            {
                var effectiveSrc = remapPageWebPartUrls ? job.SourceSiteUrl : job.TargetSiteUrl;
                var saveErr = await spService.SavePageContentAsync(
                    job.TargetSiteUrl, targetSitePagesId, pageMeta, effectiveSrc);
                if (saveErr != null)
                {
                    // Fail loudly: a stub whose content never saved is a blank page, and marking
                    // it Success hid exactly that. Re-running with overwrite recreates it.
                    System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] SavePage FAILED: {saveErr}");
                    throw new Exception($"Page created but its content could not be saved: {saveErr}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] SavePage OK");
                }

                var pubErr = await spService.PublishPageAsync(job.TargetSiteUrl, targetSitePagesId);
                if (pubErr != null)
                    System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] Publish warning: {pubErr}");
                else
                    System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] page published OK");
            }
            else if (copyPages)
            {
                // Fail loudly, same as the SavePage branch above: with copyPages ON, a null
                // pageMeta here means the source canvas genuinely could not be fetched (not the
                // "content copy disabled" case, which never reaches this branch) — the stub that
                // CreatePageStubAsync already created is a blank page, and letting the caller mark
                // this Success hid exactly that.
                throw new Exception($"Page created but its source content could not be read: {metaErr ?? "unknown error"}");
            }
            else
            {
                // copyPages is deliberately off — pageMeta was never fetched, so this stub is an
                // intentionally shallow copy, not a failure.
                result.ErrorMessage = "Page copied without content (Copy Pages option is off)";
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] downloading…");
            using var stream = await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId);
            bool isLargeFile = IsLargeForBuffering(job.SourceSize);
            if (isLargeFile) await _largeFileGate.WaitAsync(ct);
            try
            {
                using var ms = CreateTransferBuffer(isLargeFile);
                await stream.CopyToAsync(ms, ct);
                ms.Position = 0;
                System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] downloaded {ms.Length} bytes, uploading…");
                targetItemId = await spService.UploadFileAsync(job.TargetDriveId, targetParentId, job.SourceName, ms, overwrite);
                System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] upload complete, targetItemId={targetItemId}");
            }
            finally { if (isLargeFile) _largeFileGate.Release(); }
        }

        result.VersionsCopied = 1;
        result.VersionsTotal  = 1;
        if (!string.IsNullOrEmpty(targetItemId))
        {
            // For pages these run AFTER Publish so the editing session is already closed —
            // no 409 conflicts from Graph PATCH competing with the SitePages session.

            // Custom columns FIRST: ValidateUpdateListItem bumps Modified/Editor, so the
            // metadata stamp below must come last for preserved dates to survive.
            if (copyCustomColumns && bulkFieldCache != null && columnMappings != null)
            {
                System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] applying custom columns…");
                var spIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
                if (spIds.HasValue && bulkFieldCache.TryGetValue($"{spIds.Value.listId}:{spIds.Value.listItemId}", out var customFields))
                {
                    var cfErr = await spService.ApplyFileCustomFieldsAsync(
                        job.TargetDriveId, targetItemId, customFields, columnMappings, ct);
                    result.CustomFieldStatus  = cfErr != null ? CopyStatus.Failed : CopyStatus.Success;
                    result.CustomFieldDetails = cfErr;
                    if (cfErr != null) result.ErrorMessage ??= cfErr;
                }
            }

            if (preserveMetadata)
            {
                // PatchFileSystemDate FIRST: it creates a phantom version attributed to the
                // migrating user (see the version-replay path's design note), so the Editor/dates
                // stamp must come after it — the old Apply→Patch order left the newest version
                // attributed to the copying account.
                if (metadata.ModifiedDateTime.HasValue)
                {
                    var fsErr = await spService.PatchFileSystemDateAsync(
                        job.TargetDriveId, targetItemId,
                        metadata.ModifiedDateTime.Value, metadata.CreatedDateTime);
                    if (fsErr != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] PatchFileSystemDate warning: {fsErr}");
                        result.ErrorMessage ??= fsErr;
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] applying file metadata…");
                var err = await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetItemId, job.TargetSiteId, metadata);
                if (err != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] ApplyFileMetadata warning: {err}");
                    result.ErrorMessage ??= err;
                }
            }
        }
        System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] DONE: {job.SourceName}");
        return targetItemId;
    }

    // Mode B: enhanced REST version copy.
    // For each version (oldest-first):
    //   Upload → record upload-version ID U
    //   PatchFileSystemDate → creates phantom P with correct date
    //   ValidateUpdateListItem on P → sets per-version Editor/Author (NEW vs v1)
    //   DeleteItemVersion(U) → removes upload-time version
    // Result: versions 2,4,6,… (2× count) with correct dates AND correct per-version editors.
    private async Task<string?> CopyWithVersionsEnhancedRestAsync(
        CopyJob job, CopyResult result, string targetParentId, bool overwrite, int maxVersions,
        CancellationToken ct,
        bool copyCustomColumns = false, List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null,
        bool preserveMetadata = true)
    {
        var metadata    = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
        var allVersions = await spService.GetVersionsAsync(job.SourceDriveId, job.SourceItemId);
        var versions    = maxVersions > 0 && allVersions.Count > maxVersions
            ? allVersions.TakeLast(maxVersions).ToList()
            : allVersions;
        result.VersionsTotal      = versions.Count;
        result.VersionsBytesTotal = versions.Sum(v => v.Size ?? metadata.Size ?? job.SourceSize ?? 0);

        string targetItemId = string.Empty;

        foreach (var version in versions)
        {
            ct.ThrowIfCancellationRequested();
            if (version.Id == null) continue;

            bool isLast = version == versions[^1];

            using var stream = isLast
                ? await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId)
                : await spService.DownloadVersionAsync(job.SourceDriveId, job.SourceItemId, version.Id);

            // version.Size (this specific version's byte count) is the accurate figure; fall back to
            // the file's current size if a version didn't report one, same fallback order
            // MigrationJobService uses for the same decision.
            bool isLargeFile = IsLargeForBuffering(version.Size ?? job.SourceSize);
            if (isLargeFile) await _largeFileGate.WaitAsync(ct);
            try
            {
                using var ms = CreateTransferBuffer(isLargeFile);
                await stream.CopyToAsync(ms, ct);
                ms.Position = 0;

                // Always overwrite during version replay: after version 1 uploads, the file exists,
                // so the final (current) version's upload must replace it too — with overwrite=false
                // (Skip mode) the ≥4MB upload-session path 409'd and left an OLD version as the
                // target's current content. Skip semantics are enforced by the existence check in
                // CopySingleFileAsync before replay ever starts.
                targetItemId = await spService.UploadFileAsync(
                    job.TargetDriveId, targetParentId, job.SourceName, ms, overwrite: true);
            }
            finally { if (isLargeFile) _largeFileGate.Release(); }

            if (preserveMetadata)
            {
                // Record the upload version before PatchFileSystemDate creates the phantom
                var uploadVersionId = await spService.GetCurrentVersionIdAsync(job.TargetDriveId, targetItemId);

                // PatchFileSystemDate: sets date visible in version history, creates phantom P
                var versionDate = version.LastModifiedDateTime ?? DateTimeOffset.UtcNow;
                var fsErr = await spService.PatchFileSystemDateAsync(
                    job.TargetDriveId, targetItemId, versionDate,
                    isLast ? metadata.CreatedDateTime : null);
                if (fsErr != null) result.ErrorMessage ??= fsErr;

                // ValidateUpdateListItem on phantom P: set per-version editor
                var versionEditorEmail = SharePointService.GetIdentityEmail(version.LastModifiedBy?.User)
                                         ?? metadata.ModifiedByEmail;
                var perVersionMeta = new FileMetadata
                {
                    CreatedDateTime  = isLast ? metadata.CreatedDateTime : null,
                    CreatedByEmail   = isLast ? metadata.CreatedByEmail : null,
                    ModifiedDateTime = versionDate,
                    ModifiedByEmail  = versionEditorEmail,
                };
                var metaErr = await spService.ApplyFileMetadataAsync(
                    job.TargetDriveId, targetItemId, job.TargetSiteId, perVersionMeta);
                if (metaErr != null) result.ErrorMessage ??= metaErr;

                // Delete the upload version U; keep phantom P with correct date + editor
                if (uploadVersionId != null)
                {
                    var delErr = await spService.DeleteItemVersionAsync(
                        job.TargetDriveId, targetItemId, uploadVersionId);
                    if (delErr != null) result.ErrorMessage ??= delErr;
                }
                else
                {
                    // GetCurrentVersionIdAsync failed (transient Graph error) — previously silent,
                    // since the `if` above simply skipped the delete with no error recorded. The
                    // upload-time version now stays behind as a permanent duplicate entry in the
                    // target's version history alongside the correctly-dated phantom.
                    result.ErrorMessage ??= "Could not identify the temporary upload version to remove — an extra version may remain in history";
                }
            }

            result.VersionsCopied++;
        }

        // Apply custom column values once (on the final version target item)
        if (copyCustomColumns && bulkFieldCache != null && columnMappings != null &&
            !string.IsNullOrEmpty(targetItemId))
        {
            var spIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
            if (spIds.HasValue && bulkFieldCache.TryGetValue($"{spIds.Value.listId}:{spIds.Value.listItemId}", out var customFields) &&
                customFields.Count > 0)
            {
                var cfErr = await spService.ApplyFileCustomFieldsAsync(
                    job.TargetDriveId, targetItemId, customFields, columnMappings);
                result.CustomFieldStatus  = cfErr != null ? CopyStatus.Failed : CopyStatus.Success;
                result.CustomFieldDetails = cfErr;
                if (cfErr != null) result.ErrorMessage ??= cfErr;

                // ValidateUpdateListItem bumps Modified/Editor — re-stamp the final
                // version's metadata so the preserved dates survive the field write.
                if (preserveMetadata && versions.Count > 0)
                {
                    var lastVersion = versions[^1];
                    var finalMeta = new FileMetadata
                    {
                        CreatedDateTime  = metadata.CreatedDateTime,
                        CreatedByEmail   = metadata.CreatedByEmail,
                        ModifiedDateTime = lastVersion.LastModifiedDateTime ?? metadata.ModifiedDateTime,
                        ModifiedByEmail  = SharePointService.GetIdentityEmail(lastVersion.LastModifiedBy?.User)
                                           ?? metadata.ModifiedByEmail,
                    };
                    var restampErr = await spService.ApplyFileMetadataAsync(
                        job.TargetDriveId, targetItemId, job.TargetSiteId, finalMeta);
                    if (restampErr != null) result.ErrorMessage ??= restampErr;
                }
            }
        }
        return targetItemId;
    }

    private async Task ApplyFolderMetadataRecursiveAsync(
        CopyJob job, int maxParallel, AdaptiveParallelismController gate, CancellationToken ct,
        bool copyPermissions = false,
        PermissionCopyService? permissionService = null,
        ObservableCollection<CopyResult>? permissionResults = null,
        int[]? folderDone = null, int[]? folderTotal = null,
        IProgress<(int done, int total)>? folderProgress = null,
        HashSet<string>? dirtyFolderPaths = null,
        bool applyMetadata = true,
        List<(string driveId, string itemId, string relativePath)>? scannedFolders = null)
    {
        var prefix = string.IsNullOrEmpty(job.TargetSubFolderPath) ? "" : job.TargetSubFolderPath + "/";

        // Folder color is written via the path-based foldercoloring endpoint, so this pass needs the
        // target library's server-relative URL to build each folder's path. Resolved once per job
        // (one Graph call) and only when metadata is actually being applied; a failure here just
        // disables color for this job rather than failing the pass.
        string? targetLibRelUrl = null;
        if (applyMetadata)
        {
            try { targetLibRelUrl = await spService.GetLibraryServerRelativeUrlAsync(job.TargetDriveId); }
            catch { /* color skipped for this job */ }
        }

        // Applies folder color, ordered BEFORE the date/author correction: the coloring endpoint
        // writes the folder's list item, so it stamps Editor/Modified — doing it first means
        // ApplyFileMetadataAsync overwrites those side effects. Non-fatal by design.
        async Task StampColorIfAnyAsync(FileMetadata meta, string targetRelativePath)
        {
            if (string.IsNullOrEmpty(meta.ColorHex) || targetLibRelUrl == null) return;
            try
            {
                await spService.StampFolderColorAsync(
                    job.TargetSiteUrl, $"{targetLibRelUrl}/{targetRelativePath}", meta.ColorHex!, ct);
            }
            catch (OperationCanceledException) { throw; }
            catch { /* color is cosmetic — never fail a folder over it */ }
        }

        // With dirty tracking, only update the root folder if a file was copied into it or a descendant.
        bool hasRoot = !job.IsLibrary && job.SourceItemId != "root"
                    && (dirtyFolderPaths == null || dirtyFolderPaths.Contains(prefix + job.SourceName));
        if (hasRoot)
        {
            if (folderTotal != null) Interlocked.Increment(ref folderTotal[0]);
            folderProgress?.Report((folderDone?[0] ?? 0, folderTotal?[0] ?? 0));

            var rootTargetId = await spService.GetOrCreateFolderPathAsync(
                job.TargetDriveId, job.TargetParentItemId, prefix + job.SourceName);
            if (applyMetadata)
            {
                // Folder, not a file — opt into the folder-color read (see GetFileMetadataAsync).
                var rootMeta = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId, includeFolderColor: true);
                await StampColorIfAnyAsync(rootMeta, prefix + job.SourceName);
                await spService.ApplyFileMetadataAsync(job.TargetDriveId, rootTargetId, job.TargetSiteId, rootMeta);
                if (rootMeta.ModifiedDateTime.HasValue)
                    await spService.PatchFileSystemDateAsync(job.TargetDriveId, rootTargetId,
                        rootMeta.ModifiedDateTime.Value, rootMeta.CreatedDateTime);
            }

            if (copyPermissions && permissionService != null && permissionResults != null)
            {
                try
                {
                    var srcIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
                    var tgtIds = await spService.GetSharePointIdsAsync(job.TargetDriveId, rootTargetId);
                    if (srcIds.HasValue && tgtIds.HasValue)
                    {
                        var srcApiPath = $"web/lists('{srcIds.Value.listId}')/items({srcIds.Value.listItemId})";
                        var hasUnique  = await spService.GetHasUniqueRoleAssignmentsAsync(job.SourceSiteUrl, srcApiPath, ct);
                        var perm = await permissionService.CopyObjectPermissionsAsync(
                            job.SourceSiteUrl, job.TargetSiteUrl,
                            srcApiPath,
                            $"web/lists('{tgtIds.Value.listId}')/items({tgtIds.Value.listItemId})",
                            hasUniquePermissions: hasUnique,
                            job.SourceName, ct);
                        if (perm.HasActivity)
                            AddPermissionRow(permissionResults, perm, null); // folder result — no matching file row
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch { /* non-fatal */ }
            }

            if (folderDone != null) Interlocked.Increment(ref folderDone[0]);
            folderProgress?.Report((folderDone?[0] ?? 0, folderTotal?[0] ?? 0));
        }

        // Sourced from the main scan (see SourceFileEntry.IsFolder / scannedFoldersByJob), not a
        // separate EnumerateFoldersAsync re-walk of the whole tree — that re-walk was redundant (the
        // scan already visited every folder), serial, and had no throttle protection at all. Falls
        // back to the old re-walk only if the scan didn't capture anything for this job (e.g. a
        // caller that bypassed the normal scan path), so a genuinely empty/unexpected gap doesn't
        // silently skip every subfolder's metadata.
        List<(string driveId, string itemId, string relativePath)> subFolders;
        if (scannedFolders != null)
        {
            subFolders = scannedFolders;
        }
        else
        {
            subFolders = [];
            await foreach (var item in spService.EnumerateFoldersAsync(job.SourceDriveId, job.SourceItemId, ct: ct))
                subFolders.Add(item);
        }

        // With dirty tracking, skip subfolders that received no successful copies.
        // EnumerateFoldersAsync returns all descendants, so we filter the flat list here.
        if (dirtyFolderPaths != null)
            subFolders = subFolders
                .Where(sf =>
                {
                    var tp = job.IsLibrary ? prefix + sf.relativePath
                                           : prefix + $"{job.SourceName}/{sf.relativePath}";
                    return dirtyFolderPaths.Contains(tp);
                })
                .ToList();

        if (folderTotal != null) Interlocked.Add(ref folderTotal[0], subFolders.Count);
        folderProgress?.Report((folderDone?[0] ?? 0, folderTotal?[0] ?? 0));

        await Parallel.ForEachAsync(subFolders,
            new ParallelOptions { MaxDegreeOfParallelism = maxParallel, CancellationToken = ct },
            async (item, innerCt) =>
            {
                // The gate (not MaxDegreeOfParallelism) governs live concurrency: it soft-starts and
                // shrinks on throttle, so effective width = min(maxParallel, gate limit) — see
                // ApplyAllFolderMetadataAsync, which owns the gate and its Throttled subscription.
                await gate.WaitAsync(innerCt);
                try
                {
                var (driveId, itemId, relativePath) = item;
                var targetPath = job.IsLibrary ? prefix + relativePath : prefix + $"{job.SourceName}/{relativePath}";
                var targetFolderId = await spService.GetOrCreateFolderPathAsync(
                    job.TargetDriveId, job.TargetParentItemId, targetPath);
                if (applyMetadata)
                {
                    // Folder, not a file — opt into the folder-color read (see GetFileMetadataAsync).
                    var meta = await spService.GetFileMetadataAsync(driveId, itemId, includeFolderColor: true);
                    await StampColorIfAnyAsync(meta, targetPath);
                    await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetFolderId, job.TargetSiteId, meta);
                    if (meta.ModifiedDateTime.HasValue)
                        await spService.PatchFileSystemDateAsync(job.TargetDriveId, targetFolderId,
                            meta.ModifiedDateTime.Value, meta.CreatedDateTime);
                }

                if (copyPermissions && permissionService != null && permissionResults != null)
                {
                    try
                    {
                        var srcIds = await spService.GetSharePointIdsAsync(driveId, itemId);
                        var tgtIds = await spService.GetSharePointIdsAsync(job.TargetDriveId, targetFolderId);
                        if (srcIds.HasValue && tgtIds.HasValue)
                        {
                            var srcApiPath = $"web/lists('{srcIds.Value.listId}')/items({srcIds.Value.listItemId})";
                            var hasUnique  = await spService.GetHasUniqueRoleAssignmentsAsync(job.SourceSiteUrl, srcApiPath, innerCt);
                            var perm = await permissionService.CopyObjectPermissionsAsync(
                                job.SourceSiteUrl, job.TargetSiteUrl,
                                srcApiPath,
                                $"web/lists('{tgtIds.Value.listId}')/items({tgtIds.Value.listItemId})",
                                hasUniquePermissions: hasUnique,
                                System.IO.Path.GetFileName(relativePath), innerCt);
                            if (perm.HasActivity)
                                AddPermissionRow(permissionResults, perm, null); // folder result — no matching file row
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch { /* non-fatal */ }
                }

                if (folderDone != null) Interlocked.Increment(ref folderDone[0]);
                folderProgress?.Report((folderDone?[0] ?? 0, folderTotal?[0] ?? 0));
                }
                finally { gate.Release(); }
            });
    }

    // Computes the TargetSubFolderPath for a file expanded from a folder job.
    // For library jobs the file's directory becomes the subfolder directly.
    // For folder jobs the source folder name is prepended to form the relative path.
    internal static string ComputeTargetSubFolder(
        string relativePath, string jobSourceName, bool isLibrary, string jobTargetSubFolderPath)
    {
        var fileDir     = System.IO.Path.GetDirectoryName(relativePath)?.Replace('\\', '/') ?? string.Empty;
        var relToParent = isLibrary
            ? fileDir
            : (string.IsNullOrEmpty(fileDir) ? jobSourceName : $"{jobSourceName}/{fileDir}");
        return string.IsNullOrEmpty(jobTargetSubFolderPath)
            ? relToParent
            : string.IsNullOrEmpty(relToParent)
                ? jobTargetSubFolderPath
                : $"{jobTargetSubFolderPath}/{relToParent}";
    }

    // Same shape as ComputeTargetSubFolder, but for an IsEmptyFolder entry: relativePath is already
    // the folder's OWN path (there's no filename to strip a directory from), so it's used as-is
    // rather than via GetDirectoryName.
    internal static string ComputeTargetFolderPath(
        string relativePath, string jobSourceName, bool isLibrary, string jobTargetSubFolderPath)
    {
        var relToParent = isLibrary
            ? relativePath
            : (string.IsNullOrEmpty(relativePath) ? jobSourceName : $"{jobSourceName}/{relativePath}");
        return string.IsNullOrEmpty(jobTargetSubFolderPath)
            ? relToParent
            : string.IsNullOrEmpty(relToParent)
                ? jobTargetSubFolderPath
                : $"{jobTargetSubFolderPath}/{relToParent}";
    }

    private static CopyResult CreateResult(CopyJob job) => new()
    {
        FileName   = job.SourceName,
        SourcePath = job.SourceDisplayPath,
        TargetPath = job.TargetDisplayPath
    };

    // Stamps the permission outcome onto the caller's own row when it has one (every per-FILE call
    // site already holds its CopyResult in scope — no need to search for it). Pass null for
    // folder-level results, which have no matching file row; this silently no-ops for those, same as
    // before, but without an O(n) (or O(n)+O(n) fallback) scan of the results collection per file —
    // on a large single-file selection with unique permissions per file, that scan was O(n²) overall.
    private static void AddPermissionRow(ObservableCollection<CopyResult> results, PermissionCopyResult perm, CopyResult? row)
    {
        string detail;
        CopyStatus status;
        if (perm.Error != null)
        {
            detail = perm.Error;
            status = CopyStatus.Failed;
        }
        else
        {
            detail = perm.Applied == 1 ? "1 role assignment applied" : $"{perm.Applied} role assignments applied";
            if (perm.SkippedPrincipals.Count > 0)
                detail += $"; skipped {perm.SkippedPrincipals.Count} unresolvable: {string.Join(", ", perm.SkippedPrincipals)}";
            if (perm.FailedRoles is { Count: > 0 })
                detail += $"; {perm.FailedRoles.Count} failed: {string.Join(", ", perm.FailedRoles.Take(3))}";
            status = CopyStatus.Success;
        }

        if (row == null) return;

        // No explicit dispatch needed: CopyResult's property setters already marshal to the UI
        // thread via Dispatcher.BeginInvoke (see CopyResult.OnPropertyChanged). A synchronous
        // Dispatcher.Invoke here was both redundant and a real hazard — up to maxParallel copy
        // threads plus the 8-wide folder/SPMI permission passes could all block on the dispatcher
        // queue at once, the same pattern that caused the UCEERR_RENDERTHREADFAILURE crash this
        // codebase already fixed once for CopyResult itself.
        row.PermissionStatus  = status;
        row.PermissionDetails = detail;
    }
}
