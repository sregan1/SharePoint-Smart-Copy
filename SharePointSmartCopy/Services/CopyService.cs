using System.Collections.ObjectModel;
using System.IO;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class CopyService(SharePointService spService, MigrationJobService migrationJobService)
{
    // Fired when adaptive throttling changes the effective parallelism during a copy run.
    public event Action<int>? ParallelismChanged;

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
        IProgress<int>? onFilePacked = null,
        IProgress<(int done, int total)>? onFolderProgress = null)
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

        var allTasks  = new List<(CopyJob job, CopyResult result)>();

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
        void onScanThrottle(TimeSpan delay, int _, int __, string? ___) => scanController.StepDown(delay);

        bool anyFolderJobs = jobs.Any(j => j.IsFolder);
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
                    var result = FindResult(results, job.SourceDisplayPath) ?? CreateResult(job);
                    allTasks.Add((job, result));
                }
                else if (!string.IsNullOrEmpty(await spService.GetFolderProgIdAsync(job.SourceDriveId, job.SourceItemId)))
                {
                    // The job's OWN root is the special folder (e.g. a notebook selected directly
                    // as the copy source, not discovered as a descendant during a walk) — the
                    // per-child check inside EnumerateFilesForCopyAsync never sees this case since
                    // the walk starts AT this item rather than encountering it as someone's child.
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
                    await foreach (var entry in spService.EnumerateFilesForCopyAsync(
                        job.SourceDriveId, job.SourceItemId, "", scanController, cancellationToken))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var (driveId, itemId, name, relativePath) = (entry.DriveId, entry.ItemId, entry.Name, entry.RelativePath);
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
                                await spService.GetOrCreateFolderPathAsync(
                                    job.TargetDriveId, job.TargetParentItemId, emptyFolderTarget);
                                folderResult.Status = CopyStatus.Success;
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

                        var targetSubFolder = ComputeTargetSubFolder(
                            relativePath, job.SourceName, job.IsLibrary, job.TargetSubFolderPath);
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
            activityLog?.Report($"Source scan complete: {scannedFiles:N0} file(s) found");
        await FlushPendingResultsAsync();

        if (copyMode == CopyMode.MigrationApi)
        {
            // Mode A: batch all files into migration jobs. When Copy Versions is off, callers pass
            // maxVersions = 0; the Migration path reads 0 as "all versions", so translate it to 1
            // (current version only) to honor the toggle. With versions on, maxVersions is already
            // the intended cap (and 0 there legitimately means "all versions").
            int migrationMaxVersions = copyVersions ? maxVersions : 1;
            // The SPMI manifest no longer carries per-item <Fields> (standalone SPListItem objects
            // caused "Missing file info" import failures — see MigrationPackageBuilder), so custom
            // column values cannot be applied by this mode. Say so instead of silently ignoring the
            // option.
            if (copyCustomColumns)
                activityLog?.Report("⚠ Custom column values are not applied in Migration API mode — after this run, re-run in Enhanced REST mode with Copy-If-Newer to stamp custom columns");
            await migrationJobService.ExecuteAsync(allTasks, overwriteMode, migrationMaxVersions, maxParallel, cancellationToken,
                copyCustomColumns, columnMappings, bulkFieldCache, preflightProgress, activityLog, onFilePacked);

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
                            AddPermissionRow(results, perm, result.TargetPath);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch { /* non-fatal */ }
                });
            }
        }
        else
        {
            // Mode B: enhanced REST, parallel per-file
            var parallelTasks = allTasks.Select(t =>
                CopySingleFileAsync(t.job, t.result, overwriteMode, copyVersions, maxVersions, controller, cancellationToken,
                    copyCustomColumns, columnMappings, bulkFieldCache, copyPages, remapPageWebPartUrls, preserveMetadata,
                    copyPermissions, permissionService, permissionFlags, permissionResults: results));
            await Task.WhenAll(parallelTasks);
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
                    dirtyFolderPaths: null, applyMetadata: false);
            else
                onMetadataDone?.Report(true);
            return;
        }

        // For REST mode: only update folders that received at least one successful copy.
        // Build an ancestor-inclusive set from successful file job paths so we skip
        // every clean branch (e.g. unchanged folders when running "If Newer").
        var dirtyFolderPaths = allTasks
            .Where(t => t.result.Status == CopyStatus.Success
                     && !string.IsNullOrEmpty(t.job.TargetSubFolderPath))
            .SelectMany(t =>
            {
                var parts = t.job.TargetSubFolderPath!.Split('/');
                return Enumerable.Range(1, parts.Length)
                                 .Select(i => string.Join("/", parts.Take(i)));
            })
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var folderJobs = jobs.Where(j => j.IsFolder).ToList();
        bool anyFileCopied = results.Any(r => r.Status == CopyStatus.Success);
        // Permissions must not be gated on metadata options: with "Preserve metadata" off, or an
        // If-Newer re-run where every file was up to date but folder permissions changed at
        // source, folder permissions were silently skipped (files, by contrast, refresh their
        // permissions even when skipped as up to date). With permissions on, walk ALL folders
        // (dirty tracking only knows about file copies, not permission changes).
        bool wantsFolderPermissions = copyPermissions && permissionService != null;
        if (folderJobs.Count > 0 && ((preserveMetadata && anyFileCopied) || wantsFolderPermissions))
            _ = ApplyAllFolderMetadataAsync(folderJobs, maxParallel, onMetadataDone, cancellationToken,
                copyPermissions, permissionService, results, onFolderProgress,
                dirtyFolderPaths: (wantsFolderPermissions || dirtyFolderPaths.Count == 0) ? null : dirtyFolderPaths,
                applyMetadata: preserveMetadata && anyFileCopied);
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
        bool applyMetadata = true)
    {
        bool completed = true;
        try
        {
            int[] done  = { 0 };
            int[] total = { 0 };
            foreach (var job in folderJobs)
                await ApplyFolderMetadataRecursiveAsync(job, maxParallel, ct,
                    copyPermissions, permissionService, permissionResults,
                    done, total, folderProgress, dirtyFolderPaths, applyMetadata);
        }
        catch (OperationCanceledException) { completed = false; }
        catch { }
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
        try
        {
            result.Status = CopyStatus.Copying;

            var targetParentId = await ResolveTargetParentAsync(job, ct);

            // Set when IfNewer finds the target already current: the file copy is skipped
            // but permissions (below) still refresh when enabled.
            string? upToDateItemId = null;

            if (overwriteMode != OverwriteMode.Overwrite)
            {
                var existing = await spService.GetFileInfoAsync(job.TargetDriveId, targetParentId, job.SourceName);
                if (existing != null)
                {
                    if (overwriteMode == OverwriteMode.Skip)
                    {
                        result.Status = CopyStatus.Skipped;
                        return;
                    }
                    // IfNewer: copy only when the source changed since the target was written.
                    var srcMeta = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
                    if (srcMeta.ModifiedDateTime is { } srcModified && existing.Value.Modified is { } tgtModified &&
                        TimestampComparer.IsUpToDate(srcModified, tgtModified))
                        upToDateItemId = existing.Value.ItemId;
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
                                    AddPermissionRow(permissionResults, perm, result.TargetPath);
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
            // Cancelled, not Skipped: this item never started (or didn't finish) — Skipped
            // otherwise means "compared and found already up to date." See CopyStatus.Cancelled.
            result.Status       = CopyStatus.Cancelled;
            result.ErrorMessage = "Cancelled";
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
            else
            {
                result.ErrorMessage = metaErr ?? "Source page metadata unavailable";
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] downloading…");
            using var stream = await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId);
            using var ms     = new MemoryStream();
            await stream.CopyToAsync(ms, ct);
            ms.Position = 0;
            System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] downloaded {ms.Length} bytes, uploading…");
            targetItemId = await spService.UploadFileAsync(job.TargetDriveId, targetParentId, job.SourceName, ms, overwrite);
            System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] upload complete, targetItemId={targetItemId}");
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
        result.VersionsTotal = versions.Count;

        string targetItemId = string.Empty;

        foreach (var version in versions)
        {
            ct.ThrowIfCancellationRequested();
            if (version.Id == null) continue;

            bool isLast = version == versions[^1];

            using var stream = isLast
                ? await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId)
                : await spService.DownloadVersionAsync(job.SourceDriveId, job.SourceItemId, version.Id);
            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms, ct);
            ms.Position = 0;

            // Always overwrite during version replay: after version 1 uploads, the file exists,
            // so the final (current) version's upload must replace it too — with overwrite=false
            // (Skip mode) the ≥4MB upload-session path 409'd and left an OLD version as the
            // target's current content. Skip semantics are enforced by the existence check in
            // CopySingleFileAsync before replay ever starts.
            targetItemId = await spService.UploadFileAsync(
                job.TargetDriveId, targetParentId, job.SourceName, ms, overwrite: true);

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
        CopyJob job, int maxParallel, CancellationToken ct,
        bool copyPermissions = false,
        PermissionCopyService? permissionService = null,
        ObservableCollection<CopyResult>? permissionResults = null,
        int[]? folderDone = null, int[]? folderTotal = null,
        IProgress<(int done, int total)>? folderProgress = null,
        HashSet<string>? dirtyFolderPaths = null,
        bool applyMetadata = true)
    {
        var prefix = string.IsNullOrEmpty(job.TargetSubFolderPath) ? "" : job.TargetSubFolderPath + "/";

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
                var rootMeta = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
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
                            AddPermissionRow(permissionResults, perm, $"{job.TargetSiteUrl.TrimEnd('/')}/{job.SourceName}");
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch { /* non-fatal */ }
            }

            if (folderDone != null) Interlocked.Increment(ref folderDone[0]);
            folderProgress?.Report((folderDone?[0] ?? 0, folderTotal?[0] ?? 0));
        }

        var subFolders = new List<(string driveId, string itemId, string relativePath)>();
        await foreach (var item in spService.EnumerateFoldersAsync(job.SourceDriveId, job.SourceItemId))
            subFolders.Add(item);

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
                var (driveId, itemId, relativePath) = item;
                var targetPath = job.IsLibrary ? prefix + relativePath : prefix + $"{job.SourceName}/{relativePath}";
                var targetFolderId = await spService.GetOrCreateFolderPathAsync(
                    job.TargetDriveId, job.TargetParentItemId, targetPath);
                if (applyMetadata)
                {
                    var meta = await spService.GetFileMetadataAsync(driveId, itemId);
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
                                AddPermissionRow(permissionResults, perm, $"{job.TargetSiteUrl.TrimEnd('/')}/{relativePath}");
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch { /* non-fatal */ }
                }

                if (folderDone != null) Interlocked.Increment(ref folderDone[0]);
                folderProgress?.Report((folderDone?[0] ?? 0, folderTotal?[0] ?? 0));
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

    private static CopyResult? FindResult(IEnumerable<CopyResult> results, string sourcePath)
        => results.FirstOrDefault(r => r.SourcePath == sourcePath);

    private static CopyResult CreateResult(CopyJob job) => new()
    {
        FileName   = job.SourceName,
        SourcePath = job.SourceDisplayPath,
        TargetPath = job.TargetDisplayPath
    };

    // Stamps the permission outcome onto the existing file row. Silently no-ops for
    // folder-level results where no matching row exists.
    private static void AddPermissionRow(ObservableCollection<CopyResult> results, PermissionCopyResult perm, string targetPath)
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

        // Match by full target path first — the same file name routinely exists in multiple
        // folders, and a name-only match stamped the outcome on the first row with that name.
        var row = results.FirstOrDefault(r =>
                !r.IsPermissionResult &&
                string.Equals(r.TargetPath, targetPath, StringComparison.OrdinalIgnoreCase))
            ?? results.FirstOrDefault(r =>
                !r.IsPermissionResult &&
                string.Equals(r.FileName, perm.ItemName, StringComparison.OrdinalIgnoreCase));

        if (row == null) return;

        System.Windows.Application.Current.Dispatcher.Invoke(() =>
        {
            row.PermissionStatus  = status;
            row.PermissionDetails = detail;
        });
    }
}
