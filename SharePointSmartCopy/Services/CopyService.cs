using System.Collections.ObjectModel;
using System.IO;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class CopyService(SharePointService spService)
{
    // Entry point: resolves jobs (expanding folder jobs) and runs copies in parallel.
    public async Task ExecuteAsync(
        IList<CopyJob> jobs,
        ObservableCollection<CopyResult> results,
        bool overwrite,
        bool copyVersions,
        int maxParallel,
        int maxVersions,
        bool preserveVersionMetadata,
        CancellationToken cancellationToken)
    {
        var semaphore = new SemaphoreSlim(maxParallel, maxParallel);

        // Expand folder jobs into individual file tasks
        var allTasks = new List<(CopyJob job, CopyResult result)>();

        foreach (var job in jobs)
        {
            if (!job.IsFolder)
            {
                var result = FindResult(results, job.SourceDisplayPath) ?? CreateResult(job);
                allTasks.Add((job, result));
            }
            else
            {
                // Enumerate files under the folder; add results dynamically
                await foreach (var (driveId, itemId, name, relativePath)
                    in spService.EnumerateFilesAsync(job.SourceDriveId, job.SourceItemId))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var fileDir = System.IO.Path.GetDirectoryName(relativePath)?.Replace('\\', '/') ?? string.Empty;
                    // Recreate the selected folder at the target, then sub-folders beneath it
                    var targetSubFolder = string.IsNullOrEmpty(fileDir)
                        ? job.SourceName
                        : $"{job.SourceName}/{fileDir}";
                    var fileJob = new CopyJob
                    {
                        SourceDriveId       = driveId,
                        SourceItemId        = itemId,
                        SourceName          = name,
                        SourceDisplayPath   = $"{job.SourceDisplayPath}/{relativePath}",
                        TargetDriveId       = job.TargetDriveId,
                        TargetParentItemId  = job.TargetParentItemId,
                        TargetSiteId        = job.TargetSiteId,
                        TargetSubFolderPath = targetSubFolder,
                        TargetDisplayPath   = $"{targetSubFolder}/{name}",
                        IsFolder            = false
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

        // Run all file copies in parallel
        var parallelTasks = allTasks.Select(t =>
            CopySingleFileAsync(t.job, t.result, overwrite, copyVersions, maxVersions, preserveVersionMetadata, semaphore, cancellationToken));

        await Task.WhenAll(parallelTasks);

        // Apply metadata to all created folders (done after files so folders exist)
        foreach (var job in jobs.Where(j => j.IsFolder))
            await ApplyFolderMetadataRecursiveAsync(job, cancellationToken);
    }

    private async Task CopySingleFileAsync(
        CopyJob job,
        CopyResult result,
        bool overwrite,
        bool copyVersions,
        int maxVersions,
        bool preserveVersionMetadata,
        SemaphoreSlim semaphore,
        CancellationToken ct)
    {
        await semaphore.WaitAsync(ct);
        try
        {
            result.Status = CopyStatus.Copying;

            // Compute target parent folder (handle sub-folder paths for folder jobs)
            var targetParentId = await ResolveTargetParentAsync(job, ct);

            if (!overwrite && await spService.FileExistsAsync(job.TargetDriveId, targetParentId, job.SourceName))
            {
                result.Status = CopyStatus.Skipped;
                return;
            }

            if (copyVersions)
                await CopyWithVersionsAsync(job, result, targetParentId, overwrite, maxVersions, preserveVersionMetadata, ct);
            else
                await CopyCurrentVersionAsync(job, result, targetParentId, overwrite, ct);

            result.Status = CopyStatus.Success;
        }
        catch (OperationCanceledException)
        {
            result.Status  = CopyStatus.Failed;
            result.ErrorMessage = "Cancelled";
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
        CopyJob job, CopyResult result, string targetParentId, bool overwrite, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        var metadata = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
        using var stream = await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId);
        using var ms     = new MemoryStream();
        await stream.CopyToAsync(ms, ct);
        ms.Position = 0;
        var targetItemId = await spService.UploadFileAsync(job.TargetDriveId, targetParentId, job.SourceName, ms, overwrite);
        result.VersionsCopied = 1;
        result.VersionsTotal  = 1;
        if (!string.IsNullOrEmpty(targetItemId))
        {
            var err = await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetItemId, job.TargetSiteId, metadata);
            if (err != null) result.ErrorMessage = err;

            if (metadata.ModifiedDateTime.HasValue)
            {
                var fsErr = await spService.PatchFileSystemDateAsync(
                    job.TargetDriveId, targetItemId,
                    metadata.ModifiedDateTime.Value, metadata.CreatedDateTime);
                if (fsErr != null) result.ErrorMessage ??= fsErr;
            }
        }
    }

    private async Task CopyWithVersionsAsync(
        CopyJob job, CopyResult result, string targetParentId, bool overwrite, int maxVersions,
        bool preserveVersionMetadata, CancellationToken ct)
    {
        var metadata = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
        var allVersions = await spService.GetVersionsAsync(job.SourceDriveId, job.SourceItemId);
        // Take the N most-recent versions when a limit is set (list is sorted oldest-first)
        var versions = maxVersions > 0 && allVersions.Count > maxVersions
            ? allVersions.TakeLast(maxVersions).ToList()
            : allVersions;
        result.VersionsTotal = versions.Count;

        string targetItemId = string.Empty;

        foreach (var version in versions)
        {
            ct.ThrowIfCancellationRequested();
            if (version.Id == null) continue;

            bool isLast = version == versions[^1];

            // The current (last) version cannot be downloaded via the versions endpoint
            using var stream = isLast
                ? await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId)
                : await spService.DownloadVersionAsync(job.SourceDriveId, job.SourceItemId, version.Id);
            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms, ct);
            ms.Position = 0;

            targetItemId = await spService.UploadFileAsync(
                job.TargetDriveId, targetParentId, job.SourceName, ms,
                overwrite: isLast ? overwrite : true);

            if (preserveVersionMetadata)
            {
                // PRESERVE METADATA strategy: PATCH fileSystemInfo so the version's date matches
                // the source, then delete the upload-time version. fileSystemInfo is the field
                // shown in SharePoint version history. Side-effect: each version consumes two
                // slots (upload + phantom), so target version numbers are non-sequential
                // (e.g. 2, 4, 6 for 3 source versions). Tradeoff selected by the user.
                if (isLast)
                {
                    // Apply full metadata (Modified, Created, Author, Editor) to the listItem
                    // before the phantom+delete cycle so the final visible version carries it all.
                    var err = await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetItemId, job.TargetSiteId, metadata);
                    if (err != null) result.ErrorMessage ??= err;
                }

                var uploadVersionId = await spService.GetCurrentVersionIdAsync(job.TargetDriveId, targetItemId);

                var fsErr = await spService.PatchFileSystemDateAsync(
                    job.TargetDriveId, targetItemId,
                    version.LastModifiedDateTime ?? DateTimeOffset.UtcNow,
                    isLast ? metadata.CreatedDateTime : null);
                if (fsErr != null) result.ErrorMessage ??= fsErr;

                if (uploadVersionId != null)
                {
                    var delErr = await spService.DeleteItemVersionAsync(job.TargetDriveId, targetItemId, uploadVersionId);
                    if (delErr != null) result.ErrorMessage ??= delErr;
                }
            }
            else if (isLast)
            {
                // SEQUENTIAL strategy: for the latest version, sync all metadata.
                // VULI sets listItem fields (Modified, Created, Editor, Author).
                // PatchFileSystemDate also sets fileSystemInfo.lastModifiedDateTime, which is the
                // field the SharePoint library "Modified" column actually displays. Without this
                // second call the library shows the upload timestamp, not the source date.
                // The PATCH creates one phantom version (target ends up with N+1 versions for N
                // source versions), but all slots remain sequential: 1, 2, …, N, N+1.
                var err = await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetItemId, job.TargetSiteId, metadata);
                if (err != null) result.ErrorMessage ??= err;

                var fsErr = await spService.PatchFileSystemDateAsync(
                    job.TargetDriveId, targetItemId,
                    version.LastModifiedDateTime ?? DateTimeOffset.UtcNow,
                    metadata.CreatedDateTime);
                if (fsErr != null) result.ErrorMessage ??= fsErr;
            }

            result.VersionsCopied++;
        }
    }

    private async Task ApplyFolderMetadataRecursiveAsync(CopyJob job, CancellationToken ct)
    {
        // Apply to the selected folder itself (skip library roots — "root" has no meaningful metadata)
        if (job.SourceItemId != "root")
        {
            var rootTargetId = await spService.GetOrCreateFolderPathAsync(
                job.TargetDriveId, job.TargetParentItemId, job.SourceName);
            var rootMeta = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
            await spService.ApplyFileMetadataAsync(job.TargetDriveId, rootTargetId, job.TargetSiteId, rootMeta);
            // Folders have no version history, so patching fileSystemInfo has no phantom-version side effect.
            // This sets the dates the library view actually displays (fileSystemInfo, not listItem fields).
            if (rootMeta.ModifiedDateTime.HasValue)
                await spService.PatchFileSystemDateAsync(job.TargetDriveId, rootTargetId,
                    rootMeta.ModifiedDateTime.Value, rootMeta.CreatedDateTime);
        }

        // Apply to every sub-folder recursively
        await foreach (var (driveId, itemId, relativePath) in spService.EnumerateFoldersAsync(job.SourceDriveId, job.SourceItemId))
        {
            ct.ThrowIfCancellationRequested();
            var targetPath = $"{job.SourceName}/{relativePath}";
            var targetFolderId = await spService.GetOrCreateFolderPathAsync(
                job.TargetDriveId, job.TargetParentItemId, targetPath);
            var meta = await spService.GetFileMetadataAsync(driveId, itemId);
            await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetFolderId, job.TargetSiteId, meta);
            if (meta.ModifiedDateTime.HasValue)
                await spService.PatchFileSystemDateAsync(job.TargetDriveId, targetFolderId,
                    meta.ModifiedDateTime.Value, meta.CreatedDateTime);
        }
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
