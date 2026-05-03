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
            CopySingleFileAsync(t.job, t.result, overwrite, copyVersions, maxVersions, semaphore, cancellationToken));

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
        SemaphoreSlim semaphore,
        CancellationToken ct)
    {
        await semaphore.WaitAsync(ct);
        try
        {
            result.Status = CopyStatus.Copying;

            // Compute target parent folder (handle sub-folder paths for folder jobs)
            var targetParentId = await ResolveTargetParentAsync(job, ct);

            if (copyVersions)
                await CopyWithVersionsAsync(job, result, targetParentId, overwrite, maxVersions, ct);
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
            await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetItemId, job.TargetSiteId, metadata);
    }

    private async Task CopyWithVersionsAsync(
        CopyJob job, CopyResult result, string targetParentId, bool overwrite, int maxVersions, CancellationToken ct)
    {
        var metadata = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
        var allVersions = await spService.GetVersionsAsync(job.SourceDriveId, job.SourceItemId);
        // Take the N most-recent versions when a limit is set (list is sorted oldest-first)
        var versions = maxVersions > 0 && allVersions.Count > maxVersions
            ? allVersions.TakeLast(maxVersions).ToList()
            : allVersions;
        result.VersionsTotal = versions.Count;

        string targetItemId = string.Empty;

        // Upload oldest first so version history builds up naturally.
        // After each upload, patch the version's timestamp before the next upload snapshots it.
        foreach (var version in versions)
        {
            ct.ThrowIfCancellationRequested();
            if (version.Id == null) continue;

            bool isLast = version == versions[^1];

            // The current (last) version cannot be downloaded via the versions endpoint;
            // use the regular content endpoint for it.
            using var stream = isLast
                ? await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId)
                : await spService.DownloadVersionAsync(job.SourceDriveId, job.SourceItemId, version.Id);
            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms, ct);
            ms.Position = 0;

            targetItemId = await spService.UploadFileAsync(job.TargetDriveId, targetParentId, job.SourceName, ms,
                overwrite: isLast ? overwrite : true);

            // Patch this version's timestamp immediately so it's captured in the version snapshot
            // before the next upload overwrites the item. For the last version, full metadata
            // (including Author/Editor) is applied below.
            if (!string.IsNullOrEmpty(targetItemId) && !isLast)
                await spService.PatchTimestampsAsync(
                    job.TargetDriveId, targetItemId,
                    created: metadata.CreatedDateTime,
                    modified: version.LastModifiedDateTime);

            result.VersionsCopied++;
        }

        if (!string.IsNullOrEmpty(targetItemId))
            await spService.ApplyFileMetadataAsync(job.TargetDriveId, targetItemId, job.TargetSiteId, metadata);
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
