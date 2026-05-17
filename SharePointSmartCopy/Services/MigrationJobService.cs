using System.IO;
using System.Text.Json;
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
        bool overwrite,
        int maxVersions,
        int maxParallel,
        CancellationToken cancellationToken)
    {
        if (fileTasks.Count == 0) return;

        var targetSiteUrl = fileTasks[0].job.TargetSiteUrl;
        if (string.IsNullOrEmpty(targetSiteUrl))
            throw new InvalidOperationException("TargetSiteUrl must be set on CopyJob for Migration API mode.");

        foreach (var (_, result) in fileTasks)
            result.Status = CopyStatus.Copying;

        try
        {
            // Fetch shared pre-flight info once (shared across all parallel jobs)
            var (webId, webRelUrl) = await spService.GetWebInfoAsync(targetSiteUrl);
            var rawSiteId          = await spService.GetSiteIdAsync(targetSiteUrl);
            var siteId             = rawSiteId.Contains(',') ? rawSiteId.Split(',')[1] : rawSiteId;

            var firstJob            = fileTasks[0].job;
            var libraryServerRelUrl = firstJob.TargetLibraryServerRelativeUrl;
            if (string.IsNullOrEmpty(libraryServerRelUrl))
                libraryServerRelUrl = await spService.GetLibraryServerRelativeUrlAsync(firstJob.TargetDriveId);
            var listId       = await spService.GetListIdByServerRelativeUrlAsync(targetSiteUrl, libraryServerRelUrl);
            var libraryTitle = libraryServerRelUrl.Split('/').Last();

            // Pre-create subfolders before parallel split — concurrent creates for the same path conflict
            var subFolderPaths = fileTasks
                .Select(t => t.job.TargetSubFolderPath)
                .Where(p => !string.IsNullOrEmpty(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            foreach (var folderPath in subFolderPaths)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await spService.GetOrCreateFolderPathAsync(
                    firstJob.TargetDriveId, firstJob.TargetParentItemId, folderPath);
            }

            // Cap SPMI job count at 5 — beyond that SP throttling negates the gain
            var jobCount = Math.Min(Math.Max(maxParallel, 1), 5);

            if (jobCount <= 1)
            {
                await RunSingleJobAsync(
                    fileTasks, overwrite, maxVersions, maxParallel,
                    targetSiteUrl, webId, webRelUrl, siteId,
                    libraryServerRelUrl, listId, libraryTitle, cancellationToken);
            }
            else
            {
                var batches = Enumerable.Range(0, jobCount)
                    .Select(_ => new List<(CopyJob, CopyResult)>())
                    .ToArray();
                for (int i = 0; i < fileTasks.Count; i++)
                    batches[i % jobCount].Add(fileTasks[i]);

                await Task.WhenAll(batches
                    .Where(b => b.Count > 0)
                    .Select(batch => RunSingleJobAsync(
                        batch, overwrite, maxVersions, maxParallel,
                        targetSiteUrl, webId, webRelUrl, siteId,
                        libraryServerRelUrl, listId, libraryTitle, cancellationToken)));
            }
        }
        catch (OperationCanceledException)
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

    private async Task RunSingleJobAsync(
        IList<(CopyJob job, CopyResult result)> fileTasks,
        bool overwrite,
        int maxVersions,
        int maxParallel,
        string targetSiteUrl,
        string webId,
        string webRelUrl,
        string siteId,
        string libraryServerRelUrl,
        string listId,
        string libraryTitle,
        CancellationToken cancellationToken)
    {
        try
        {
            // Step 1: provision SP-provided encrypted containers (one set per job)
            var (dataUri, metadataUri, encryptionKey) =
                await spService.ProvisionMigrationContainersAsync(targetSiteUrl);

            // Step 2: skip files that already exist at the target (overwrite=false)
            if (!overwrite)
            {
                var folderIdCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var (job, result) in fileTasks)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var subPath = job.TargetSubFolderPath ?? string.Empty;
                    if (!folderIdCache.TryGetValue(subPath, out var parentId))
                    {
                        parentId = string.IsNullOrEmpty(subPath)
                            ? job.TargetParentItemId
                            : await spService.GetOrCreateFolderPathAsync(job.TargetDriveId, job.TargetParentItemId, subPath);
                        folderIdCache[subPath] = parentId;
                    }
                    if (await spService.FileExistsAsync(job.TargetDriveId, parentId, job.SourceName))
                        result.Status = CopyStatus.Skipped;
                }

                if (fileTasks.All(t => t.result.Status == CopyStatus.Skipped))
                    return;
            }

            // Step 3: build the package — download all versions, encrypt blobs, build manifests
            var builder = new MigrationPackageBuilder(encryptionKey);

            foreach (var (job, result) in fileTasks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (result.Status != CopyStatus.Copying) continue;
                try
                {
                    await AddFileToPackageAsync(builder, job, result, maxVersions, libraryServerRelUrl, cancellationToken);
                }
                catch (Exception ex)
                {
                    result.Status       = CopyStatus.Failed;
                    result.ErrorMessage = $"Package build failed: {ex.Message}";
                }
            }

            System.Diagnostics.Debug.WriteLine($"[Migration] encryptionKey length={encryptionKey.Length}");
            System.Diagnostics.Debug.WriteLine($"[Migration] dataUri prefix={dataUri[..Math.Min(dataUri.Length,80)]}");
            System.Diagnostics.Debug.WriteLine($"[Migration] metaUri prefix={metadataUri[..Math.Min(metadataUri.Length,80)]}");

            // Step 4: upload content blobs to the data container — parallel
            var dataClient = new BlobContainerClient(new Uri(dataUri));
            var allBlobs   = builder.Files.SelectMany(f => f.Versions).ToList();
            await Parallel.ForEachAsync(
                allBlobs,
                new ParallelOptions { MaxDegreeOfParallelism = maxParallel, CancellationToken = cancellationToken },
                async (version, ct) =>
                {
                    var ivB64 = Convert.ToBase64String(version.EncryptedContent[..16]);
                    var opts  = new BlobUploadOptions { Metadata = new Dictionary<string, string> { ["IV"] = ivB64 } };
                    using var ms = new MemoryStream(version.EncryptedContent, 16, version.EncryptedContent.Length - 16);
                    var blob = dataClient.GetBlobClient(version.StreamId);
                    await blob.UploadAsync(ms, opts, ct);
                    await blob.CreateSnapshotAsync(cancellationToken: ct);
                });

            // Step 5: upload manifest XML blobs to the metadata container
            var metadataClient = new BlobContainerClient(new Uri(metadataUri));
            var manifests = builder.BuildManifestXml(
                siteId, webId, listId,
                targetSiteUrl, webRelUrl, libraryTitle, libraryServerRelUrl);

            foreach (var (blobName, data) in manifests)
            {
                cancellationToken.ThrowIfCancellationRequested();
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

            // Step 7: poll until JobEnd
            string? jobError = null;
            await foreach (var evt in spService.PollMigrationJobAsync(targetSiteUrl, jobId, cancellationToken))
            {
                if (evt.TryGetProperty("Event", out var evtName))
                {
                    var name = evtName.GetString();
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
                        jobError = $"Migration job fatal error: {msg}";
                    }
                    else if (name == "JobError")
                    {
                        var msg = evt.TryGetProperty("Message", out var m) ? m.GetString() : "Unknown error";
                        System.Diagnostics.Debug.WriteLine($"[Migration] non-fatal JobError: {msg}");
                    }
                }
            }

            await TryLogMigrationReportAsync(metadataClient, jobId, encryptionKey);

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
        catch (OperationCanceledException)
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

    private async Task AddFileToPackageAsync(
        MigrationPackageBuilder builder,
        CopyJob job,
        CopyResult result,
        int maxVersions,
        string libraryServerRelUrl,
        CancellationToken ct)
    {
        var metadata    = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
        var allVersions = await spService.GetVersionsAsync(job.SourceDriveId, job.SourceItemId);
        var versions    = maxVersions > 0 && allVersions.Count > maxVersions
            ? allVersions.TakeLast(maxVersions).ToList()
            : allVersions;

        result.VersionsTotal = versions.Count;

        var versionStreams = new List<(Microsoft.Graph.Models.DriveItemVersion version, Stream content)>();
        foreach (var version in versions)
        {
            ct.ThrowIfCancellationRequested();
            if (version.Id == null) continue;

            bool isLast = version == versions[^1];
            var stream = isLast
                ? await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId)
                : await spService.DownloadVersionAsync(job.SourceDriveId, job.SourceItemId, version.Id);
            versionStreams.Add((version, stream));
        }

        var folderRelPath = string.IsNullOrEmpty(job.TargetSubFolderPath)
            ? ""
            : job.TargetSubFolderPath.TrimStart('/');

        await builder.AddFileAsync(job.SourceName, folderRelPath, metadata, versionStreams);

        foreach (var (_, s) in versionStreams)
            s.Dispose();

        result.VersionsCopied = versions.Count;
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
                catch { /* metadata read may not be permitted by SAS */ }

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
