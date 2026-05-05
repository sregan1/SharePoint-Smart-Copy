using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class SharePointService
{
    private readonly AuthService _authService;
    private GraphServiceClient? _graphClient;
    private readonly HttpClient _httpClient = new();
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int> _userIdCache = new();
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (string siteUrl, string listId, string listItemId)> _spIdsCache = new();

    public SharePointService(AuthService authService)
    {
        _authService = authService;
    }

    public void Initialize()
    {
        var provider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(_authService));
        _graphClient = new GraphServiceClient(provider);
    }

    private GraphServiceClient Graph => _graphClient
        ?? throw new InvalidOperationException("Not initialized. Please sign in first.");

    // ── Site ──────────────────────────────────────────────────────────────────

    public async Task<string> GetSiteIdAsync(string siteUrl)
    {
        var uri = new Uri(siteUrl.TrimEnd('/'));
        var hostname = uri.Host;
        var path = uri.AbsolutePath.TrimStart('/');
        var key = string.IsNullOrEmpty(path) ? hostname : $"{hostname}:/{path}";

        var site = await Graph.Sites[key].GetAsync();
        return site?.Id ?? throw new Exception($"Could not find site at {siteUrl}");
    }

    // ── Libraries ─────────────────────────────────────────────────────────────

    public async Task<List<SharePointNode>> GetLibrariesAsync(string siteId)
    {
        var drives = await Graph.Sites[siteId].Drives.GetAsync();
        var result = new List<SharePointNode>();

        foreach (var d in drives?.Value ?? [])
        {
            if (d.Id == null || d.Name == null) continue;
            var node = new SharePointNode
            {
                Id       = "root",
                DriveId  = d.Id,
                SiteId   = siteId,
                Name     = d.Name,
                Type     = NodeType.Library,
                HasChildren = true
            };
            node.Children.Add(Placeholder());
            result.Add(node);
        }
        return result;
    }

    // ── Children ──────────────────────────────────────────────────────────────

    public async Task<List<SharePointNode>> GetChildrenAsync(
        string driveId, string itemId, string siteId, bool foldersOnly = false)
    {
        DriveItemCollectionResponse? page;

        var resolvedId = itemId == "root" ? "root" : itemId;
        page = await Graph.Drives[driveId].Items[resolvedId].Children
            .GetAsync(cfg => cfg.QueryParameters.Top = 1000);

        var items = new List<DriveItem>();
        while (page != null)
        {
            items.AddRange(page.Value ?? []);
            if (page.OdataNextLink == null) break;
            page = await Graph.Drives[driveId].Items[resolvedId].Children
                .WithUrl(page.OdataNextLink).GetAsync();
        }

        var result = new List<SharePointNode>();
        foreach (var item in items)
        {
            if (item.Id == null || item.Name == null) continue;
            if (foldersOnly && item.Folder == null) continue;

            bool isFolder = item.Folder != null;
            var node = new SharePointNode
            {
                Id          = item.Id,
                DriveId     = driveId,
                SiteId      = siteId,
                Name        = item.Name,
                Type        = isFolder ? NodeType.Folder : NodeType.File,
                Size        = item.Size,
                WebUrl      = item.WebUrl,
                HasChildren = isFolder && (item.Folder?.ChildCount ?? 0) > 0
            };

            if (node.HasChildren)
                node.Children.Add(Placeholder());

            result.Add(node);
        }
        return result;
    }

    // ── Enumerate all files under a folder (for copy) ─────────────────────────

    public async IAsyncEnumerable<(string driveId, string itemId, string name, string relativePath)>
        EnumerateFilesAsync(string driveId, string rootItemId, string basePath = "")
    {
        var children = await GetChildrenAsync(driveId, rootItemId, string.Empty);
        foreach (var child in children)
        {
            var childPath = string.IsNullOrEmpty(basePath) ? child.Name : $"{basePath}/{child.Name}";
            if (child.Type == NodeType.File)
            {
                yield return (driveId, child.Id, child.Name, childPath);
            }
            else
            {
                await foreach (var item in EnumerateFilesAsync(driveId, child.Id, childPath))
                    yield return item;
            }
        }
    }

    // ── Enumerate all sub-folders under a folder (for metadata) ──────────────

    public async IAsyncEnumerable<(string driveId, string itemId, string relativePath)>
        EnumerateFoldersAsync(string driveId, string rootItemId, string basePath = "")
    {
        var children = await GetChildrenAsync(driveId, rootItemId, string.Empty);
        foreach (var child in children.Where(c => c.Type == NodeType.Folder))
        {
            var childPath = string.IsNullOrEmpty(basePath) ? child.Name : $"{basePath}/{child.Name}";
            yield return (driveId, child.Id, childPath);
            await foreach (var item in EnumerateFoldersAsync(driveId, child.Id, childPath))
                yield return item;
        }
    }

    // ── File content ──────────────────────────────────────────────────────────

    public async Task<Stream> DownloadFileAsync(string driveId, string itemId)
    {
        var stream = await Graph.Drives[driveId].Items[itemId].Content.GetAsync();
        return stream ?? throw new Exception("Empty response downloading file.");
    }

    public async Task<Stream> DownloadVersionAsync(string driveId, string itemId, string versionId)
    {
        var stream = await Graph.Drives[driveId].Items[itemId].Versions[versionId].Content.GetAsync();
        return stream ?? throw new Exception("Empty response downloading version content.");
    }

    public async Task<List<DriveItemVersion>> GetVersionsAsync(string driveId, string itemId)
    {
        var all  = new List<DriveItemVersion>();
        var page = await Graph.Drives[driveId].Items[itemId].Versions.GetAsync();
        while (page != null)
        {
            all.AddRange(page.Value ?? []);
            if (page.OdataNextLink == null) break;
            page = await Graph.Drives[driveId].Items[itemId].Versions
                .WithUrl(page.OdataNextLink).GetAsync();
        }
        return all.OrderBy(v => v.LastModifiedDateTime).ToList();
    }

    // ── Metadata ──────────────────────────────────────────────────────────────

    public async Task<FileMetadata> GetFileMetadataAsync(string driveId, string itemId)
    {
        var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
            cfg.QueryParameters.Select = ["createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy"]);

        return new FileMetadata(
            item?.CreatedDateTime,
            GetIdentityEmail(item?.CreatedBy?.User),
            item?.LastModifiedDateTime,
            GetIdentityEmail(item?.LastModifiedBy?.User));
    }

    internal static string? GetIdentityEmail(Microsoft.Graph.Models.Identity? identity)
    {
        if (identity?.AdditionalData == null) return null;
        if (identity.AdditionalData.TryGetValue("email", out var email) && email != null)
            return email.ToString();
        // Some Graph responses use userPrincipalName instead of email
        if (identity.AdditionalData.TryGetValue("userPrincipalName", out var upn) && upn != null)
            return upn.ToString();
        return null;
    }

    private async Task<int?> LookupSharePointUserIdAsync(string siteId, string email)
    {
        var key = $"{siteId}|{email}";
        if (_userIdCache.TryGetValue(key, out var cached)) return cached;
        try
        {
            var response = await Graph.Sites[siteId].Lists["User Information List"].Items
                .GetAsync(cfg =>
                {
                    cfg.QueryParameters.Filter = $"fields/EMail eq '{email}'";
                    cfg.QueryParameters.Expand = ["fields"];
                    cfg.QueryParameters.Top    = 1;
                });

            var fields = response?.Value?.FirstOrDefault()?.Fields?.AdditionalData;
            if (fields != null && fields.TryGetValue("ID", out var idObj) && idObj != null)
            {
                var id = Convert.ToInt32(idObj);
                _userIdCache[key] = id;
                return id;
            }
        }
        catch { /* user not found or list not queryable — skip */ }
        return null;
    }

    // Returns null on success, or an error string describing what failed.
    public async Task<string?> PatchTimestampsAsync(
        string driveId, string itemId, DateTimeOffset? created, DateTimeOffset? modified,
        string? createdByEmail = null, string? modifiedByEmail = null)
    {
        // ValidateUpdateListItem with bNewDocumentUpdate=true sets all four metadata fields
        // without triggering a new version. No fileSystemInfo fallback here — on intermediate
        // versions that would create phantom versions in version history.
        return await PatchTimestampsViaRestAsync(driveId, itemId, modified, created, createdByEmail, modifiedByEmail);
    }

    private async Task<(string siteUrl, string listId, string listItemId)?> GetSharePointIdsAsync(
        string driveId, string itemId)
    {
        var key = $"{driveId}|{itemId}";
        if (_spIdsCache.TryGetValue(key, out var cached)) return cached;

        // SharePoint may take a moment to propagate sharepointIds for a newly uploaded item.
        // Retry with backoff before giving up.
        for (int attempt = 0; attempt < 3; attempt++)
        {
            if (attempt > 0) await Task.Delay(attempt * 1500);
            try
            {
                var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
                    cfg.QueryParameters.Select = ["sharepointIds"]);

                var ids = item?.SharepointIds;
                if (ids?.SiteUrl == null || ids.ListId == null || ids.ListItemId == null)
                {
                    System.Diagnostics.Debug.WriteLine($"GetSharePointIds attempt {attempt + 1}/3: sharepointIds null/incomplete for item {itemId} (SiteUrl={ids?.SiteUrl}, ListId={ids?.ListId}, ListItemId={ids?.ListItemId})");
                    continue;
                }

                var result = (ids.SiteUrl, ids.ListId, ids.ListItemId);
                _spIdsCache[key] = result;
                return result;
            }
            catch (Exception ex)
            {
                var odataEx = ex as Microsoft.Graph.Models.ODataErrors.ODataError;
                var detail  = odataEx?.Error?.Message ?? ex.Message;
                System.Diagnostics.Debug.WriteLine($"GetSharePointIds attempt {attempt + 1}/3 failed (HTTP {odataEx?.ResponseStatusCode}): {detail}");
            }
        }
        return null;
    }

    // Returns null on success, or an error string describing what failed.
    private async Task<string?> PatchTimestampsViaRestAsync(
        string driveId, string itemId, DateTimeOffset? modified, DateTimeOffset? created,
        string? createdByEmail = null, string? modifiedByEmail = null)
    {
        var ids = await GetSharePointIdsAsync(driveId, itemId);
        if (ids == null) return "SP IDs unavailable — item not found or sharepointIds not propagated";

        var (siteUrl, listId, listItemId) = ids.Value;
        System.Diagnostics.Debug.WriteLine($"PatchTimestampsViaRest: siteUrl={siteUrl} listId={listId} listItemId={listItemId}");

        var formValues = new List<object>();
        // ValidateUpdateListItem parses dates using the site's regional settings (not ISO 8601).
        // US English SharePoint sites expect M/d/yyyy H:mm:ss (24-hour, InvariantCulture).
        if (modified.HasValue)
            formValues.Add(new { FieldName = "Modified", FieldValue = modified.Value.ToUniversalTime().ToString("M/d/yyyy H:mm:ss", System.Globalization.CultureInfo.InvariantCulture) });
        if (created.HasValue)
            formValues.Add(new { FieldName = "Created",  FieldValue = created.Value.ToUniversalTime().ToString("M/d/yyyy H:mm:ss", System.Globalization.CultureInfo.InvariantCulture) });
        // Author/Editor via claims identity — same call, no new version side-effect
        if (!string.IsNullOrEmpty(modifiedByEmail))
            formValues.Add(new { FieldName = "Editor", FieldValue = $"i:0#.f|membership|{modifiedByEmail}" });
        if (!string.IsNullOrEmpty(createdByEmail))
            formValues.Add(new { FieldName = "Author", FieldValue = $"i:0#.f|membership|{createdByEmail}" });

        if (formValues.Count == 0) return null;

        try
        {
            // SharePoint REST API requires a SharePoint-scoped token, not the Graph token.
            var token = await _authService.GetSharePointTokenAsync(siteUrl);
            var url   = $"{siteUrl}/_api/web/lists('{listId}')/items({listItemId})/ValidateUpdateListItem()";
            using var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new System.Net.Http.StringContent(
                    System.Text.Json.JsonSerializer.Serialize(new { formValues, bNewDocumentUpdate = true }),
                    System.Text.Encoding.UTF8, "application/json")
            };
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            req.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            var response = await _httpClient.SendAsync(req);
            var body = await response.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine($"ValidateUpdateListItem {(int)response.StatusCode}: {body}");

            if (!response.IsSuccessStatusCode)
            {
                var preview = body.Length > 200 ? body[..200] : body;
                return $"ValidateUpdateListItem HTTP {(int)response.StatusCode}: {preview}";
            }

            // HTTP 200 OK — but ValidateUpdateListItem returns per-field error codes in the body.
            // A 200 response does NOT mean every field was updated; check each field's HasException.
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("value", out var arr))
                {
                    var fieldErrors = new List<string>();
                    foreach (var field in arr.EnumerateArray())
                    {
                        bool hasEx = field.TryGetProperty("HasException", out var he) && he.GetBoolean();
                        int  ec    = field.TryGetProperty("ErrorCode",    out var ecEl) ? ecEl.GetInt32() : 0;
                        if (hasEx || ec != 0)
                        {
                            var fn = field.TryGetProperty("FieldName",     out var fnEl) ? fnEl.GetString() ?? "" : "";
                            var em = field.TryGetProperty("ErrorMessage",  out var emEl) ? emEl.GetString() ?? "" : "";
                            fieldErrors.Add($"{fn}: {em} (code {ec})");
                        }
                    }
                    if (fieldErrors.Count > 0)
                        return $"Field errors: {string.Join("; ", fieldErrors)}";
                }
            }
            catch { /* JSON parse failure — treat body as success */ }

            return null; // all fields updated successfully
        }
        catch (Exception ex)
        {
            var msg = $"ValidateUpdateListItem exception: {ex.Message}";
            System.Diagnostics.Debug.WriteLine(msg);
            return msg;
        }
    }

    // Returns null on success, or an error string describing what failed.
    public async Task<string?> ApplyFileMetadataAsync(
        string driveId, string itemId, string siteId, FileMetadata metadata)
    {
        System.Diagnostics.Debug.WriteLine($"ApplyFileMetadata: itemId={itemId} createdBy={metadata.CreatedByEmail} modifiedBy={metadata.ModifiedByEmail} created={metadata.CreatedDateTime} modified={metadata.ModifiedDateTime}");
        // ValidateUpdateListItem sets all four fields (Modified, Created, Editor, Author)
        // atomically without creating a new version. Author/Editor are set via claims identity
        // (i:0#.f|membership|email), confirmed working from debug output.
        // The previous listItem/fields PATCH and fileSystemInfo fallback are removed — both
        // could trigger a version bump in versioned libraries, causing phantom versions.
        return await PatchTimestampsViaRestAsync(driveId, itemId,
            metadata.ModifiedDateTime, metadata.CreatedDateTime,
            metadata.CreatedByEmail, metadata.ModifiedByEmail);
    }

    // Returns the drive item version ID of the current (newest) version.
    // Used to record which version to delete after a fileSystemInfo PATCH creates a phantom.
    public async Task<string?> GetCurrentVersionIdAsync(string driveId, string itemId)
    {
        try
        {
            var page = await Graph.Drives[driveId].Items[itemId].Versions.GetAsync();
            // Graph returns versions newest-first; first entry is the current version.
            return page?.Value?.FirstOrDefault()?.Id;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"GetCurrentVersionId exception: {ex.Message}");
            return null;
        }
    }

    // Patches driveItem.fileSystemInfo.lastModifiedDateTime, which creates a new version in a
    // versioned library and is the field SharePoint actually displays in version history.
    // (The listItem.Modified field — set by ValidateUpdateListItem — is NOT shown in version history.)
    public async Task<string?> PatchFileSystemDateAsync(string driveId, string itemId,
        DateTimeOffset lastModifiedDateTime, DateTimeOffset? createdDateTime = null)
    {
        try
        {
            await Graph.Drives[driveId].Items[itemId].PatchAsync(new DriveItem
            {
                FileSystemInfo = new Microsoft.Graph.Models.FileSystemInfo
                {
                    LastModifiedDateTime = lastModifiedDateTime,
                    CreatedDateTime      = createdDateTime
                }
            });
            System.Diagnostics.Debug.WriteLine($"PatchFileSystemDate: modified={lastModifiedDateTime:o} created={createdDateTime:o} on {itemId}");
            return null;
        }
        catch (Exception ex)
        {
            var msg = $"PatchFileSystemDate exception: {ex.Message}";
            System.Diagnostics.Debug.WriteLine(msg);
            return msg;
        }
    }

    // Deletes a specific historical drive item version.
    public async Task<string?> DeleteItemVersionAsync(string driveId, string itemId, string versionId)
    {
        try
        {
            await Graph.Drives[driveId].Items[itemId].Versions[versionId].DeleteAsync();
            System.Diagnostics.Debug.WriteLine($"DeleteItemVersion: {versionId} on {itemId}");
            return null;
        }
        catch (Exception ex)
        {
            var msg = $"DeleteItemVersion exception: {ex.Message}";
            System.Diagnostics.Debug.WriteLine(msg);
            return msg;
        }
    }

    // ── Upload ────────────────────────────────────────────────────────────────

    public async Task<bool> FileExistsAsync(string driveId, string parentItemId, string fileName)
    {
        try
        {
            await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName)).GetAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<string> UploadFileAsync(
        string targetDriveId,
        string targetParentItemId,
        string fileName,
        Stream content,
        bool overwrite,
        IProgress<int>? progress = null)
    {
        content.Position = 0;
        long size = content.Length;

        var conflictBehavior = overwrite ? "replace" : "fail";

        if (size < 4 * 1024 * 1024)
        {
            var item = await Graph.Drives[targetDriveId].Items[targetParentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName)).Content.PutAsync(content);
            progress?.Report(100);
            return item?.Id ?? string.Empty;
        }

        // Large-file upload session.
        // NOTE: SharePoint supports fileSystemInfo in the session body only for NEW file uploads,
        // not for replacement uploads (conflictBehavior:"replace" on an existing file). Attempting
        // to set it on a replacement session returns 200 OK with an upload URL that responds 404
        // ("session not found") on the first chunk PUT. Dates must be fixed post-upload via
        // PatchFileSystemDateAsync instead.
        var sessionBody = new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody
        {
            Item = new DriveItemUploadableProperties
            {
                AdditionalData = new Dictionary<string, object>
                {
                    { "@microsoft.graph.conflictBehavior", conflictBehavior }
                }
            }
        };

        var session = await Graph.Drives[targetDriveId].Items[targetParentItemId]
            .ItemWithPath(Uri.EscapeDataString(fileName)).CreateUploadSession.PostAsync(sessionBody);

        var uploadTask = new Microsoft.Graph.LargeFileUploadTask<DriveItem>(
            session!, content, 320 * 1024, Graph.RequestAdapter);

        var uploadResult = await uploadTask.UploadAsync(new Progress<long>(uploaded =>
        {
            if (size > 0) progress?.Report((int)(uploaded * 100 / size));
        }));

        if (!uploadResult.UploadSucceeded)
            throw new Exception("Large file upload failed.");

        progress?.Report(100);
        return uploadResult.ItemResponse?.Id ?? string.Empty;
    }

    // ── Folder creation ───────────────────────────────────────────────────────

    public async Task<string> GetOrCreateFolderPathAsync(string driveId, string parentItemId, string relativePath)
    {
        var current = parentItemId;
        foreach (var part in relativePath.Split('/').Where(p => !string.IsNullOrEmpty(p)))
        {
            current = await GetOrCreateFolderAsync(driveId, current, part);
        }
        return current;
    }

    private async Task<string> GetOrCreateFolderAsync(string driveId, string parentItemId, string folderName)
    {
        try
        {
            var existing = await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(folderName)).GetAsync();
            return existing?.Id ?? throw new Exception("Null ID from existing folder.");
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            try
            {
                var created = await Graph.Drives[driveId].Items[parentItemId].Children.PostAsync(new DriveItem
                {
                    Name   = folderName,
                    Folder = new Folder()
                });
                return created?.Id ?? throw new Exception("Null ID from created folder.");
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError conflict) when (conflict.ResponseStatusCode == 409)
            {
                // Another parallel copy task created the folder first — fetch it
                var existing = await Graph.Drives[driveId].Items[parentItemId]
                    .ItemWithPath(Uri.EscapeDataString(folderName)).GetAsync();
                return existing?.Id ?? throw new Exception("Null ID after concurrent folder creation.");
            }
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static SharePointNode Placeholder() =>
        new() { Name = "__placeholder__" };

    private sealed class TokenProvider : IAccessTokenProvider
    {
        private readonly AuthService _auth;
        public TokenProvider(AuthService auth) => _auth = auth;

        public AllowedHostsValidator AllowedHostsValidator { get; } =
            new AllowedHostsValidator(["graph.microsoft.com"]);

        public async Task<string> GetAuthorizationTokenAsync(
            Uri uri,
            Dictionary<string, object>? additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
            => await _auth.GetAccessTokenAsync();
    }
}
