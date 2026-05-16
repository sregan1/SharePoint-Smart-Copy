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
        CancellationToken cancellationToken)
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
                    var fileDir = System.IO.Path.GetDirectoryName(relativePath)?.Replace('\\', '/') ?? string.Empty;
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
                        TargetSiteUrl       = job.TargetSiteUrl,
                        TargetSubFolderPath = targetSubFolder,
                        TargetDisplayPath   = $"{job.TargetDisplayPath}/{relativePath}",
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

        if (copyMode == CopyMode.MigrationApi && copyVersions)
        {
            // Mode A: batch all files into a single migration job
            await migrationJobService.ExecuteAsync(allTasks, overwrite, maxVersions, maxParallel, cancellationToken);
        }
        else
        {
            // Mode B: enhanced REST, parallel per-file
            var parallelTasks = allTasks.Select(t =>
                CopySingleFileAsync(t.job, t.result, overwrite, copyVersions, maxVersions, semaphore, cancellationToken));
            await Task.WhenAll(parallelTasks);
        }

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

            var targetParentId = await ResolveTargetParentAsync(job, ct);

            if (!overwrite && await spService.FileExistsAsync(job.TargetDriveId, targetParentId, job.SourceName))
            {
                result.Status = CopyStatus.Skipped;
                return;
            }

            if (copyVersions)
                await CopyWithVersionsEnhancedRestAsync(job, result, targetParentId, overwrite, maxVersions, ct);
            else
                await CopyCurrentVersionAsync(job, result, targetParentId, overwrite, ct);

            result.Status = CopyStatus.Success;
        }
        catch (OperationCanceledException)
        {
            result.Status       = CopyStatus.Failed;
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

    // Mode B: enhanced REST version copy.
    // For each version (oldest-first):
    //   Upload → record upload-version ID U
    //   PatchFileSystemDate → creates phantom P with correct date
    //   ValidateUpdateListItem on P → sets per-version Editor/Author (NEW vs v1)
    //   DeleteItemVersion(U) → removes upload-time version
    // Result: versions 2,4,6,… (2× count) with correct dates AND correct per-version editors.
    private async Task CopyWithVersionsEnhancedRestAsync(
        CopyJob job, CopyResult result, string targetParentId, bool overwrite, int maxVersions,
        CancellationToken ct)
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

            // Record the upload version before PatchFileSystemDate creates the phantom
            var uploadVersionId = await spService.GetCurrentVersionIdAsync(job.TargetDriveId, targetItemId);

            // PatchFileSystemDate: sets date visible in version history, creates phantom P
            var versionDate = version.LastModifiedDateTime ?? DateTimeOffset.UtcNow;
            var fsErr = await spService.PatchFileSystemDateAsync(
                job.TargetDriveId, targetItemId, versionDate,
                isLast ? metadata.CreatedDateTime : null);
            if (fsErr != null) result.ErrorMessage ??= fsErr;

            // ValidateUpdateListItem on phantom P: set per-version editor (improvement over v1)
            var versionEditorEmail = SharePointService.GetIdentityEmail(version.LastModifiedBy?.User)
                                     ?? metadata.ModifiedByEmail;
            var perVersionMeta = new FileMetadata(
                isLast ? metadata.CreatedDateTime : null,
                isLast ? metadata.CreatedByEmail : null,
                versionDate,
                versionEditorEmail);
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

            result.VersionsCopied++;
        }
    }

    private async Task ApplyFolderMetadataRecursiveAsync(CopyJob job, CancellationToken ct)
    {
        if (job.SourceItemId != "root")
        {
            var rootTargetId = await spService.GetOrCreateFolderPathAsync(
                job.TargetDriveId, job.TargetParentItemId, job.SourceName);
            var rootMeta = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
            await spService.ApplyFileMetadataAsync(job.TargetDriveId, rootTargetId, job.TargetSiteId, rootMeta);
            if (rootMeta.ModifiedDateTime.HasValue)
                await spService.PatchFileSystemDateAsync(job.TargetDriveId, rootTargetId,
                    rootMeta.ModifiedDateTime.Value, rootMeta.CreatedDateTime);
        }

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
