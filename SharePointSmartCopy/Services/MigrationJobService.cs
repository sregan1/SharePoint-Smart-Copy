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
        // Soft-start at 6 slots (safely below the Graph bandwidth throttle threshold for
        // typical SharePoint file sizes) and ramp up to maxParallel via the restore heartbeat.
        int migrationSoftStart = Math.Min(maxParallel, 6);
        using var downloadController = new AdaptiveParallelismController(maxParallel, migrationSoftStart);
        void onMigrationThrottle(TimeSpan delay, int __, int ___, string? ____) =>
            downloadController.StepDown(delay);
        spService.Throttled += onMigrationThrottle;
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
                if (subFolderPaths.Count > 0)
                {
                    activityLog?.Report($"Provisioning {subFolderPaths.Count} target subfolder{(subFolderPaths.Count == 1 ? "" : "s")}...");
                    var cacheLock = new object();
                    // Parallel creation at MaxDegree=4 — the segment cache (SharePointService._folderSegmentCache)
                    // makes concurrent creates safe: races on shared path prefixes resolve via 409 conflict
                    // recovery, and the cache ensures each unique segment is created at most once.
                    await Parallel.ForEachAsync(subFolderPaths,
                        new ParallelOptions { MaxDegreeOfParallelism = 4, CancellationToken = cancellationToken },
                        async (folderPath, ct) =>
                        {
                            var id = await spService.GetOrCreateFolderPathAsync(
                                firstJob.TargetDriveId, firstJob.TargetParentItemId, folderPath);
                            lock (cacheLock) sharedFolderIdCache[folderPath] = id;
                        });
                }

                // Build the target-folder file listing once and share across all SPMI batches.
                // Without sharing, N parallel batches each independently fetch all M folders = N×M calls.
                // With sharing: M calls total (parallel at 8), then all batches read the snapshot.
                activityLog?.Report($"Scanning {sharedFolderIdCache.Count} target folder{(sharedFolderIdCache.Count == 1 ? "" : "s")} for existing files...");
                var sharedExistingByFolder = new System.Collections.Concurrent.ConcurrentDictionary<
                    string, Dictionary<string, (string ItemId, DateTimeOffset? Modified)>>(
                    StringComparer.OrdinalIgnoreCase);
                await Parallel.ForEachAsync(sharedFolderIdCache,
                    new ParallelOptions { MaxDegreeOfParallelism = 8, CancellationToken = cancellationToken },
                    async (kvp, ct) =>
                    {
                        sharedExistingByFolder[kvp.Key] = await spService.FetchFolderItemsAsync(
                            firstJob.TargetDriveId, kvp.Value);
                    });

                // Split into SEQUENTIAL jobs of at most MaxItemsPerJob files. SharePoint's migration
                // performance guidance is to keep each package under 250 items / 250 MB; a single job
                // over that limit imports the SPFile blobs but fails every SPListItem with "Missing file
                // info for list item", hitting the 100-failure threshold and canceling the whole job.
                // (Verified empirically: a 32-item copy imports cleanly; 512 — and even 150-item batches —
                // fail this way.) The binding limit is the 250 MB byte cap, not item count: version
                // history multiplies bytes, so a count-based cap must stay well below 250 items to keep
                // each package under ~250 MB for typical file sizes. 30 (at/below the proven-good 32)
                // keeps packages around tens of MB. Jobs run ONE AT A TIME — concurrent jobs make SP
                // soft-cancel the extras with the same error, so sequential is required.
                // NOTE: this is count-based; a library of large files could still exceed 250 MB at 30
                // items. If that occurs, switch to a cumulative-byte cap (version sizes are available via
                // SharePointService.BatchFetchMetadataAndVersionsAsync).
                const int MaxItemsPerJob = 30;

                var batches = groupTasks
                    .Select((t, i) => (t, i))
                    .GroupBy(x => x.i / MaxItemsPerJob)
                    .Select(g => g.Select(x => x.t).ToList())
                    .ToList();

                for (int bi = 0; bi < batches.Count; bi++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await RunSingleJobAsync(
                        batches[bi], overwriteMode, maxVersions, maxParallel,
                        targetSiteUrl, webId, webRelUrl, siteId,
                        libraryServerRelUrl, listId, libraryTitle, cancellationToken,
                        copyCustomColumns, columnMappings, bulkFieldCache,
                        preflightCounter, preflightTotal, preflightProgress,
                        sharedFolderIdCache, sharedExistingByFolder,
                        batchLabel: batches.Count > 1 ? $"{bi + 1}/{batches.Count}" : "",
                        activityLog: activityLog,
                        onFilePacked: onFilePacked,
                        downloadController: downloadController);
                }
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

    private async Task RunSingleJobAsync(
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
        AdaptiveParallelismController? downloadController = null)
    {
        var pfx = string.IsNullOrEmpty(batchLabel) ? "" : $"[{batchLabel}] ";
        try
        {
            // Step 1: provision SP-provided encrypted containers (one set per job)
            activityLog?.Report($"{pfx}Provisioning Azure migration containers...");
            var (dataUri, metadataUri, encryptionKey) =
                await spService.ProvisionMigrationContainersAsync(targetSiteUrl);

            // Step 2: for overwrite mode, delete any existing files so SPMI does a fresh INSERT
            // (SPMI UPDATE appends versions instead of replacing them, causing duplication).
            // For non-overwrite mode, mark files that already exist (Graph) as Skipped, and
            // purge any zombies (AllDocs entry without SPListItem) so SPMI won't reject them.
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
                            var srcMeta = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
                            if (srcMeta.ModifiedDateTime is { } srcModified && srcModified <= existing.Modified)
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
                return;
            }

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

            var producerTask = Task.Run(async () =>
            {
                try
                {
                    // Batch-fetch metadata + versions 10 files at a time (20 Graph sub-requests per
                    // $batch call). Double-buffer: fire the next chunk's metadata fetch in the
                    // background while the current chunk's files are downloading, so the $batch
                    // round-trip (~200–600 ms) is hidden inside download time rather than serial.
                    var copyingTasks = fileTasks.Where(t => t.result.Status == CopyStatus.Copying).ToList();
                    var chunks = copyingTasks.Chunk(10).ToArray();

                    if (chunks.Length > 0)
                    {
                        // Prime the first chunk's metadata fetch before the loop begins.
                        var prefetchTask = spService.BatchFetchMetadataAndVersionsAsync(
                            chunks[0].Select(t => (t.job.SourceDriveId, t.job.SourceItemId)).ToList(),
                            cancellationToken);

                        for (int chunkIdx = 0; chunkIdx < chunks.Length; chunkIdx++)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            // Await metadata for this chunk (usually already done; waits for remainder if not).
                            var batchCache = await prefetchTask;

                            // Fire the next chunk's fetch in the background BEFORE downloads start,
                            // so its $batch round-trip overlaps with this chunk's download time.
                            if (chunkIdx + 1 < chunks.Length)
                                prefetchTask = spService.BatchFetchMetadataAndVersionsAsync(
                                    chunks[chunkIdx + 1].Select(t => (t.job.SourceDriveId, t.job.SourceItemId)).ToList(),
                                    cancellationToken);

                            // Download all files in this chunk in parallel.
                            // batchCache is read-only here; existingFileIds is ConcurrentDictionary — both are safe.
                            await Parallel.ForEachAsync(chunks[chunkIdx],
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

                                    // Acquire a global download slot — limits total concurrent Graph
                                    // content downloads across all SPMI batches to maxParallel.
                                    if (downloadController != null)
                                        await downloadController.WaitAsync(ct);
                                    try
                                    {
                                        batchCache.TryGetValue(job.SourceItemId, out var prefetched);
                                        if (verbosePerFile) activityLog?.Report($"{pfx}↓ {job.SourceName}");
                                        var data = await DownloadFileDataAsync(
                                            job, result, maxVersions, versionParallelism, existingFileId, ct,
                                            prefetched.Metadata, prefetched.Versions,
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
                    }
                }
                finally { pipe.Writer.Complete(); }
            }, cancellationToken);

            int packedInBatch = 0;
            await foreach (var data in pipe.Reader.ReadAllAsync(cancellationToken))
            {
                int filesBefore = builder.Files.Count;
                if (verbosePerFile) activityLog?.Report($"{pfx}↑ {data.Job.SourceName} ({data.Versions.Count} version{(data.Versions.Count == 1 ? "" : "s")})");
                try
                {
                    var versionStreams = data.Slots
                        .Select((ms, i) => (version: data.Versions[i], content: (Stream?)ms))
                        .Where(t => t.content != null)
                        .Select(t => (t.version, content: t.content!))
                        .ToList();

                    await builder.AddFileAsync(data.Job.SourceName, data.FolderRelPath, data.Metadata,
                        versionStreams, data.ExistingFileId, data.CustomFields);

                    // Versions are now encrypted. Free the original downloaded buffers immediately
                    // so only the encrypted bytes remain in memory during uploads — not both.
                    foreach (var ms in data.Slots) ms?.Dispose();
                    for (int si = 0; si < data.Slots.Length; si++) data.Slots[si] = null;

                    // Upload this file's version blobs now, then release the encrypted bytes.
                    var fileEntry = builder.Files[^1];
                    await Parallel.ForEachAsync(
                        fileEntry.Versions,
                        new ParallelOptions { MaxDegreeOfParallelism = Math.Max(2, maxParallel / 2), CancellationToken = cancellationToken },
                        async (version, ct) =>
                        {
                            var content = version.EncryptedContent
                                ?? throw new InvalidOperationException("Encrypted content already released.");
                            var ivB64 = Convert.ToBase64String(content[..16]);
                            var opts  = new BlobUploadOptions { Metadata = new Dictionary<string, string> { ["IV"] = ivB64 } };
                            using var ms = new MemoryStream(content, 16, content.Length - 16);
                            var blob = dataClient.GetBlobClient(version.StreamId);
                            await blob.UploadAsync(ms, opts, ct);
                            await blob.CreateSnapshotAsync(cancellationToken: ct);
                            version.EncryptedContent = null;
                        });

                    data.Result.VersionsCopied = data.Versions.Count;
                    onFilePacked?.Report(1);
                    int n = ++packedInBatch;
                    if (!verbosePerFile && (n == 1 || n % milestoneStep == 0 || n == copyingCount))
                        activityLog?.Report($"{pfx}{n:N0} / {copyingCount:N0} files packaged");
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    // Drop the partially-added entry so the manifest never references
                    // blobs that were not uploaded.
                    if (builder.Files.Count > filesBefore)
                        builder.RemoveLastFile();
                    data.Result.Status       = CopyStatus.Failed;
                    data.Result.ErrorMessage = $"Package build failed: {ex.Message}";
                }
                finally
                {
                    foreach (var ms in data.Slots)
                        ms?.Dispose();
                }
            }

            await producerTask;

            // Step 5: upload manifest XML blobs to the metadata container.
            // Fetch the root folder GUID so the manifest can include an explicit SPFolder entry.
            // This is required for newly created empty libraries — without it SPMI cannot resolve
            // the parent folder for files and fails with "Missing file info for list item".
            var rootFolderGuid = await spService.GetLibraryRootFolderUniqueIdAsync(
                targetSiteUrl, libraryServerRelUrl);

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
                var libRel    = libraryServerRelUrl.TrimEnd('/');
                var guidLock  = new object();
                await Parallel.ForEachAsync(allFolderPaths,
                    new ParallelOptions { MaxDegreeOfParallelism = 8, CancellationToken = cancellationToken },
                    async (path, ct) =>
                    {
                        var guid = await spService.GetLibraryRootFolderUniqueIdAsync(targetSiteUrl, $"{libRel}/{path}");
                        if (!string.IsNullOrEmpty(guid))
                            lock (guidLock) folderGuids[path] = guid;
                    });
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
                spmiOverwrite, rootFolderGuid, folderGuids);

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

            // Step 6: submit the migration job
            var jobId = await spService.CreateMigrationJobEncryptedAsync(
                targetSiteUrl, webId, dataUri, metadataUri, encryptionKey);
            System.Diagnostics.Debug.WriteLine($"[Migration] submitted job: {jobId}");
            activityLog?.Report($"{pfx}Submitted to SharePoint — waiting for import...");

            // Step 7: poll until JobEnd
            string? jobError = null;
            bool seenJobStart = false;
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
                            activityLog?.Report($"{pfx}SP importing: {proc.GetInt32():N0} / {copyingCount:N0} files");
                    }
                    if (name == "JobEnd")
                    {
                        if (evt.TryGetProperty("TotalErrors", out var te) &&
                            te.ValueKind == JsonValueKind.Number && te.GetInt32() > 0)
                            jobError = $"Migration job completed with {te.GetInt32()} error(s).";
                        break;
                    }
                    if (name == "JobFatalError")
                    {
                        var msg = evt.TryGetProperty("Message", out var m) ? m.GetString() : "Unknown error";
                        activityLog?.Report($"⚠ {pfx}Fatal error: {msg}");
                        // Log the full SP event — may contain an error code or richer reason
                        // beyond Message (e.g. "Operation canceled" can mean concurrent-job limit,
                        // SAS expiry, bad manifest, etc.)
                        activityLog?.Report($"  SP event: {evt}");
                        jobError = $"Migration job fatal error: {msg}";
                    }
                    else if (name == "JobError")
                    {
                        var msg = evt.TryGetProperty("Message", out var m) ? m.GetString() : "Unknown error";
                        System.Diagnostics.Debug.WriteLine($"[Migration] non-fatal JobError: {msg}");
                    }
                }
            }

            var importedCount = fileTasks.Count(t => t.result.Status == CopyStatus.Copying);
            if (jobError == null)
                activityLog?.Report($"{pfx}✓ Import complete — {importedCount} file{(importedCount == 1 ? "" : "s")} imported");
            else
                activityLog?.Report($"⚠ {pfx}Import finished with errors: {jobError}");

            _ = TryLogMigrationReportAsync(metadataClient, jobId, encryptionKey);

            // Step 8: mark results
            foreach (var (_, result) in fileTasks)
            {
                if (result.Status == CopyStatus.Copying)
                {
                    result.Status = jobError == null ? CopyStatus.Success : CopyStatus.Failed;
                    if (jobError != null) result.ErrorMessage ??= jobError;
                }
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

        if (prefetchedMetadata != null && prefetchedVersions != null)
        {
            metadata    = prefetchedMetadata;
            allVersions = prefetchedVersions;
        }
        else
        {
            // Batch fetch missed this item — fall back to individual Graph calls.
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
                if (version.Id == null) return;
                bool isLast = idx == versions.Count - 1;
                var ms = new MemoryStream();
                for (int attempt = 0; ; attempt++)
                {
                    ms.SetLength(0);
                    try
                    {
                        var networkStream = isLast
                            ? await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId)
                            : await spService.DownloadVersionAsync(job.SourceDriveId, job.SourceItemId, version.Id);
                        await using (networkStream)
                            await networkStream.CopyToAsync(ms, ct);
                        ms.Position = 0;
                        slots[idx] = ms;
                        break;
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (System.IO.IOException) when (attempt < 3)
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

    private static async Task TryLogMigrationReportAsync(
        Azure.Storage.Blobs.BlobContainerClient metadataClient, string jobId, byte[] key)
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

                System.Diagnostics.Debug.WriteLine($"[SP-{suffix[1..]}] {System.Text.Encoding.UTF8.GetString(plain)}");
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
