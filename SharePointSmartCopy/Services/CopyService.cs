using System.Collections.ObjectModel;
using System.IO;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class CopyService(SharePointService spService, MigrationJobService migrationJobService)
{
    public async Task ExecuteAsync(
        IList<CopyJob> jobs,
        ObservableCollection<CopyResult> results,
        bool overwrite,
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
        bool preserveMetadata = true)
    {
        var semaphore = new SemaphoreSlim(maxParallel, maxParallel);
        var allTasks  = new List<(CopyJob job, CopyResult result)>();

        foreach (var job in jobs)
        {
            if (!job.IsFolder)
            {
                var result = FindResult(results, job.SourceDisplayPath) ?? CreateResult(job);
                allTasks.Add((job, result));
            }
            else
            {
                await foreach (var (driveId, itemId, name, relativePath)
                    in spService.EnumerateFilesAsync(job.SourceDriveId, job.SourceItemId))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var targetSubFolder = ComputeTargetSubFolder(
                        relativePath, job.SourceName, job.IsLibrary, job.TargetSubFolderPath);
                    var fileJob = new CopyJob
                    {
                        SourceDriveId                  = driveId,
                        SourceItemId                   = itemId,
                        SourceName                     = name,
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

                    System.Windows.Application.Current.Dispatcher.Invoke(() => results.Add(result));
                    allTasks.Add((fileJob, result));
                }
            }
        }

        if (copyMode == CopyMode.MigrationApi && copyVersions)
        {
            // Mode A: batch all files into a single migration job
            await migrationJobService.ExecuteAsync(allTasks, overwrite, maxVersions, maxParallel, cancellationToken,
                copyCustomColumns, columnMappings, bulkFieldCache);
        }
        else
        {
            // Mode B: enhanced REST, parallel per-file
            var parallelTasks = allTasks.Select(t =>
                CopySingleFileAsync(t.job, t.result, overwrite, copyVersions, maxVersions, semaphore, cancellationToken,
                    copyCustomColumns, columnMappings, bulkFieldCache, copyPages, remapPageWebPartUrls, preserveMetadata));
            await Task.WhenAll(parallelTasks);
        }

        // Apply folder metadata in the background (skipped when preserveMetadata is off)
        var folderJobs = jobs.Where(j => j.IsFolder).ToList();
        if (folderJobs.Count > 0 && preserveMetadata)
            _ = ApplyAllFolderMetadataAsync(folderJobs, maxParallel, onMetadataDone, cancellationToken);
        else
            onMetadataDone?.Report(true);
    }

    private async Task ApplyAllFolderMetadataAsync(
        IEnumerable<CopyJob> folderJobs, int maxParallel,
        IProgress<bool>? onDone, CancellationToken ct)
    {
        try
        {
            foreach (var job in folderJobs)
                await ApplyFolderMetadataRecursiveAsync(job, maxParallel, ct);
        }
        catch { }
        finally
        {
            onDone?.Report(true);
        }
    }

    private async Task CopySingleFileAsync(
        CopyJob job,
        CopyResult result,
        bool overwrite,
        bool copyVersions,
        int maxVersions,
        SemaphoreSlim semaphore,
        CancellationToken ct,
        bool copyCustomColumns = false,
        List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null,
        bool copyPages = false,
        bool remapPageWebPartUrls = true,
        bool preserveMetadata = true)
    {
        await semaphore.WaitAsync(ct);
        try
        {
            result.Status = CopyStatus.Copying;

            var targetParentId = await ResolveTargetParentAsync(job, ct);

            if (!overwrite && await spService.FileExistsAsync(job.TargetDriveId, targetParentId, job.SourceName))
            {
                result.Status = CopyStatus.Skipped;
                return;
            }

            // When overwriting with version history: delete the file first so the imported
            // versions replace the history rather than being appended to it.
            if (overwrite && copyVersions)
                await spService.DeleteFileIfExistsAsync(job.TargetDriveId, targetParentId, job.SourceName);

            if (copyVersions)
                await CopyWithVersionsEnhancedRestAsync(job, result, targetParentId, overwrite, maxVersions, ct,
                    copyCustomColumns, columnMappings, bulkFieldCache, preserveMetadata);
            else
                await CopyCurrentVersionAsync(job, result, targetParentId, overwrite, ct,
                    copyCustomColumns, columnMappings, bulkFieldCache, copyPages, remapPageWebPartUrls, preserveMetadata);

            result.Status = CopyStatus.Success;
        }
        catch (OperationCanceledException)
        {
            result.Status       = CopyStatus.Failed;
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
            semaphore.Release();
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

    private async Task CopyCurrentVersionAsync(
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
                    System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] SavePage warning: {saveErr}");
                    result.ErrorMessage = saveErr;
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
            if (preserveMetadata)
            {
                System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] applying file metadata…");
                var err = await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetItemId, job.TargetSiteId, metadata);
                if (err != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] ApplyFileMetadata warning: {err}");
                    result.ErrorMessage ??= err;
                }
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
            }

            if (copyCustomColumns && bulkFieldCache != null && columnMappings != null)
            {
                System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] applying custom columns…");
                var spIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
                if (spIds.HasValue && bulkFieldCache.TryGetValue(spIds.Value.listItemId, out var customFields))
                {
                    var cfErr = await spService.ApplyFileCustomFieldsAsync(
                        job.TargetDriveId, targetItemId, customFields, columnMappings);
                    if (cfErr != null) result.ErrorMessage ??= cfErr;
                }
            }
        }
        System.Diagnostics.Debug.WriteLine($"[CopyCurrentVersion] DONE: {job.SourceName}");
    }

    // Mode B: enhanced REST version copy.
    // For each version (oldest-first):
    //   Upload → record upload-version ID U
    //   PatchFileSystemDate → creates phantom P with correct date
    //   ValidateUpdateListItem on P → sets per-version Editor/Author (NEW vs v1)
    //   DeleteItemVersion(U) → removes upload-time version
    // Result: versions 2,4,6,… (2× count) with correct dates AND correct per-version editors.
    private async Task CopyWithVersionsEnhancedRestAsync(
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

            targetItemId = await spService.UploadFileAsync(
                job.TargetDriveId, targetParentId, job.SourceName, ms,
                overwrite: isLast ? overwrite : true);

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
            if (spIds.HasValue && bulkFieldCache.TryGetValue(spIds.Value.listItemId, out var customFields))
            {
                var cfErr = await spService.ApplyFileCustomFieldsAsync(
                    job.TargetDriveId, targetItemId, customFields, columnMappings);
                if (cfErr != null) result.ErrorMessage ??= cfErr;
            }
        }
    }

    private async Task ApplyFolderMetadataRecursiveAsync(CopyJob job, int maxParallel, CancellationToken ct)
    {
        var prefix = string.IsNullOrEmpty(job.TargetSubFolderPath) ? "" : job.TargetSubFolderPath + "/";

        if (!job.IsLibrary && job.SourceItemId != "root")
        {
            var rootTargetId = await spService.GetOrCreateFolderPathAsync(
                job.TargetDriveId, job.TargetParentItemId, prefix + job.SourceName);
            var rootMeta = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
            await spService.ApplyFileMetadataAsync(job.TargetDriveId, rootTargetId, job.TargetSiteId, rootMeta);
            if (rootMeta.ModifiedDateTime.HasValue)
                await spService.PatchFileSystemDateAsync(job.TargetDriveId, rootTargetId,
                    rootMeta.ModifiedDateTime.Value, rootMeta.CreatedDateTime);
        }

        var subFolders = new List<(string driveId, string itemId, string relativePath)>();
        await foreach (var item in spService.EnumerateFoldersAsync(job.SourceDriveId, job.SourceItemId))
            subFolders.Add(item);

        await Parallel.ForEachAsync(subFolders,
            new ParallelOptions { MaxDegreeOfParallelism = maxParallel, CancellationToken = ct },
            async (item, innerCt) =>
            {
                var (driveId, itemId, relativePath) = item;
                var targetPath = job.IsLibrary ? prefix + relativePath : prefix + $"{job.SourceName}/{relativePath}";
                var targetFolderId = await spService.GetOrCreateFolderPathAsync(
                    job.TargetDriveId, job.TargetParentItemId, targetPath);
                var meta = await spService.GetFileMetadataAsync(driveId, itemId);
                await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetFolderId, job.TargetSiteId, meta);
                if (meta.ModifiedDateTime.HasValue)
                    await spService.PatchFileSystemDateAsync(job.TargetDriveId, targetFolderId,
                        meta.ModifiedDateTime.Value, meta.CreatedDateTime);
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

    private static CopyResult? FindResult(IEnumerable<CopyResult> results, string sourcePath)
        => results.FirstOrDefault(r => r.SourcePath == sourcePath);

    private static CopyResult CreateResult(CopyJob job) => new()
    {
        FileName   = job.SourceName,
        SourcePath = job.SourceDisplayPath,
        TargetPath = job.TargetDisplayPath
    };
}
