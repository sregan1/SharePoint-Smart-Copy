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
        OverwriteMode overwriteMode,
        int maxVersions,
        int maxParallel,
        CancellationToken cancellationToken,
        bool copyCustomColumns = false,
        List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null)
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

                // Pre-create subfolders before parallel split — concurrent creates for the same path conflict
                var subFolderPaths = groupTasks
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

                if (jobCount <= 1 || groupTasks.Count <= 1)
                {
                    await RunSingleJobAsync(
                        groupTasks, overwriteMode, maxVersions, maxParallel,
                        targetSiteUrl, webId, webRelUrl, siteId,
                        libraryServerRelUrl, listId, libraryTitle, cancellationToken,
                        copyCustomColumns, columnMappings, bulkFieldCache);
                }
                else
                {
                    var batches = Enumerable.Range(0, jobCount)
                        .Select(_ => new List<(CopyJob, CopyResult)>())
                        .ToArray();
                    for (int i = 0; i < groupTasks.Count; i++)
                        batches[i % jobCount].Add(groupTasks[i]);

                    await Task.WhenAll(batches
                        .Where(b => b.Count > 0)
                        .Select(batch => RunSingleJobAsync(
                            batch, overwriteMode, maxVersions, maxParallel,
                            targetSiteUrl, webId, webRelUrl, siteId,
                            libraryServerRelUrl, listId, libraryTitle, cancellationToken,
                            copyCustomColumns, columnMappings, bulkFieldCache)));
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
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null)
    {
        try
        {
            // Step 1: provision SP-provided encrypted containers (one set per job)
            var (dataUri, metadataUri, encryptionKey) =
                await spService.ProvisionMigrationContainersAsync(targetSiteUrl);

            // Step 2: for overwrite mode, delete any existing files so SPMI does a fresh INSERT
            // (SPMI UPDATE appends versions instead of replacing them, causing duplication).
            // For non-overwrite mode, mark files that already exist (Graph) as Skipped, and
            // purge any zombies (AllDocs entry without SPListItem) so SPMI won't reject them.
            //
            // Step 2a (serial): resolve all unique subfolder IDs — concurrent GetOrCreate calls
            // on the same path can conflict, so this must remain sequential.
            var folderIdCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var subPath in fileTasks
                .Select(t => t.job.TargetSubFolderPath ?? string.Empty)
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                cancellationToken.ThrowIfCancellationRequested();
                folderIdCache[subPath] = string.IsNullOrEmpty(subPath)
                    ? fileTasks[0].job.TargetParentItemId
                    : await spService.GetOrCreateFolderPathAsync(
                        fileTasks[0].job.TargetDriveId, fileTasks[0].job.TargetParentItemId, subPath);
            }

            // Step 2b (parallel): check existence and delete files concurrently.
            var existingFileIds = new System.Collections.Concurrent.ConcurrentDictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            await Parallel.ForEachAsync(fileTasks,
                new ParallelOptions { MaxDegreeOfParallelism = maxParallel, CancellationToken = cancellationToken },
                async (task, ct) =>
                {
                    var (job, result) = task;
                    var subPath          = job.TargetSubFolderPath ?? string.Empty;
                    var parentId         = folderIdCache[subPath];
                    var fileServerRelUrl = string.IsNullOrEmpty(subPath)
                        ? $"{libraryServerRelUrl}/{job.SourceName}"
                        : $"{libraryServerRelUrl}/{subPath}/{job.SourceName}";

                    if (overwriteMode == OverwriteMode.Overwrite)
                    {
                        if (await spService.FileExistsAsync(job.TargetDriveId, parentId, job.SourceName))
                        {
                            // Real file: delete first so SPMI does a fresh INSERT.
                            // SPMI UPDATE (with existing GUID) appends imported versions to the
                            // existing version history instead of replacing it, causing duplication.
                            await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                        else if (await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) != null)
                        {
                            // Zombie blob (AllDocs row without SPListItem) — delete via deleteObject,
                            // which bypasses the SPListItem requirement that recycleObject needs.
                            await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                        existingFileIds[fileServerRelUrl] = null;
                    }
                    else if (overwriteMode == OverwriteMode.IfNewer)
                    {
                        var existing = await spService.GetFileInfoAsync(job.TargetDriveId, parentId, job.SourceName);
                        if (existing != null)
                        {
                            var srcMeta = await spService.GetFileMetadataAsync(job.SourceDriveId, job.SourceItemId);
                            if (srcMeta.ModifiedDateTime is { } srcModified && srcModified <= existing.Value.Modified)
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
                        else if (await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) != null)
                        {
                            // Zombie SPFile blob — purge so SPMI can import cleanly.
                            await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                    }
                    else // Skip
                    {
                        if (await spService.FileExistsAsync(job.TargetDriveId, parentId, job.SourceName))
                            result.Status = CopyStatus.Skipped;
                        else if (await spService.GetFileUniqueIdAsync(targetSiteUrl, fileServerRelUrl) != null)
                        {
                            // Zombie SPFile blob (AllDocs row exists but Graph returns 404).
                            // Purge it so SPMI can import cleanly.
                            await spService.PermanentlyDeleteFileAsync(targetSiteUrl, fileServerRelUrl);
                        }
                    }
                });

            if (fileTasks.All(t => t.result.Status == CopyStatus.Skipped))
                return;

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

            foreach (var (job, result) in fileTasks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (result.Status != CopyStatus.Copying) continue;
                int filesBefore = builder.Files.Count;
                try
                {
                    var subPath          = job.TargetSubFolderPath ?? string.Empty;
                    var fileServerRelUrl = string.IsNullOrEmpty(subPath)
                        ? $"{libraryServerRelUrl}/{job.SourceName}"
                        : $"{libraryServerRelUrl}/{subPath}/{job.SourceName}";
                    existingFileIds.TryGetValue(fileServerRelUrl, out var existingFileId);
                    await AddFileToPackageAsync(builder, job, result, maxVersions, libraryServerRelUrl,
                        existingFileId, cancellationToken,
                        copyCustomColumns, columnMappings, bulkFieldCache);

                    // Upload this file's version blobs now, then release the bytes.
                    var fileEntry = builder.Files[^1];
                    await Parallel.ForEachAsync(
                        fileEntry.Versions,
                        new ParallelOptions { MaxDegreeOfParallelism = maxParallel, CancellationToken = cancellationToken },
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
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    // Drop the partially-added entry so the manifest never references
                    // blobs that were not uploaded.
                    if (builder.Files.Count > filesBefore)
                        builder.RemoveLastFile();
                    result.Status       = CopyStatus.Failed;
                    result.ErrorMessage = $"Package build failed: {ex.Message}";
                }
            }

            // Step 5: upload manifest XML blobs to the metadata container.
            // Fetch the root folder GUID so the manifest can include an explicit SPFolder entry.
            // This is required for newly created empty libraries — without it SPMI cannot resolve
            // the parent folder for files and fails with "Missing file info for list item".
            var rootFolderGuid = await spService.GetLibraryRootFolderUniqueIdAsync(
                targetSiteUrl, libraryServerRelUrl);

            // IfNewer behaves like overwrite from SPMI's perspective: stale targets were
            // already deleted in step 2b, so surviving imports must be allowed to replace.
            var spmiOverwrite = overwriteMode != OverwriteMode.Skip;
            System.Diagnostics.Debug.WriteLine(
                $"[Migration] manifest params: siteId={siteId} webId={webId} listId={listId}" +
                $" webRelUrl={webRelUrl} libraryTitle={libraryTitle} libraryServerRelUrl={libraryServerRelUrl}" +
                $" rootFolderGuid={rootFolderGuid ?? "(null)"} overwrite={spmiOverwrite}");

            var metadataClient = new BlobContainerClient(new Uri(metadataUri));
            var manifests = builder.BuildManifestXml(
                siteId, webId, listId,
                targetSiteUrl, webRelUrl, libraryTitle, libraryServerRelUrl,
                spmiOverwrite, rootFolderGuid);

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

    private async Task AddFileToPackageAsync(
        MigrationPackageBuilder builder,
        CopyJob job,
        CopyResult result,
        int maxVersions,
        string libraryServerRelUrl,
        string? existingFileId,
        CancellationToken ct,
        bool copyCustomColumns = false,
        List<ColumnMapping>? columnMappings = null,
        Dictionary<string, Dictionary<string, object?>>? bulkFieldCache = null)
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

        // Look up custom field values from bulk cache (keyed by "{listId}:{listItemId}")
        Dictionary<string, string>? customFieldsForManifest = null;
        if (copyCustomColumns && bulkFieldCache != null && columnMappings != null)
        {
            var spIds = await spService.GetSharePointIdsAsync(job.SourceDriveId, job.SourceItemId);
            if (spIds.HasValue && bulkFieldCache.TryGetValue($"{spIds.Value.listId}:{spIds.Value.listItemId}", out var rawFields))
            {
                var mappingLookup = ColumnMapping.BuildTargetNameMap(columnMappings);
                customFieldsForManifest = new Dictionary<string, string>();
                foreach (var (srcName, value) in rawFields)
                {
                    if (value == null) continue;
                    string targetName;
                    if (mappingLookup.TryGetValue(srcName, out var mapped))
                    {
                        if (mapped == null) continue; // explicitly skipped
                        targetName = mapped;
                    }
                    else
                    {
                        targetName = srcName;
                    }
                    // SPMI lookup-value encoding: "-1;#value" asks SP to resolve at import.
                    // Person fields resolve claims logins; taxonomy fields resolve Label|guid
                    // (term GUIDs are valid as-is within the same tenant's term store).
                    // Lookup fields resolve by display value — SPMI matches the value against
                    // the lookup target list that already exists at the destination.
                    customFieldsForManifest[targetName] = value switch
                    {
                        PersonFieldValue p   => string.Join(";#", p.Logins.Select(l => $"-1;#{l}")),
                        TaxonomyFieldValue t => string.Join(";#", t.Terms.Select(x => $"-1;#{x.Label}|{x.TermGuid}")),
                        LookupFieldValue l   => string.Join(";#", l.Entries.Select(e => $"-1;#{e.DisplayValue}")),
                        _ => value.ToString() ?? "",
                    };
                }
            }
        }

        await builder.AddFileAsync(job.SourceName, folderRelPath, metadata, versionStreams, existingFileId,
            customFieldsForManifest);

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
