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
        var versions = await Graph.Drives[driveId].Items[itemId].Versions.GetAsync();
        return (versions?.Value ?? [])
            .OrderBy(v => v.LastModifiedDateTime)
            .ToList();
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

    private static string? GetIdentityEmail(Microsoft.Graph.Models.Identity? identity)
    {
        if (identity == null) return null;
        if (identity.AdditionalData?.TryGetValue("email", out var val) == true)
            return val?.ToString();
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

    public async Task PatchTimestampsAsync(
        string driveId, string itemId, DateTimeOffset? created, DateTimeOffset? modified)
    {
        try
        {
            await Graph.Drives[driveId].Items[itemId].PatchAsync(new DriveItem
            {
                FileSystemInfo = new Microsoft.Graph.Models.FileSystemInfo
                {
                    CreatedDateTime      = created,
                    LastModifiedDateTime = modified
                }
            });
        }
        catch { /* best-effort */ }
    }

    public async Task ApplyFileMetadataAsync(
        string driveId, string itemId, string siteId, FileMetadata metadata)
    {
        // Patch timestamps via fileSystemInfo — works with Files.ReadWrite.All delegated permission
        try
        {
            if (metadata.CreatedDateTime.HasValue || metadata.ModifiedDateTime.HasValue)
            {
                await Graph.Drives[driveId].Items[itemId].PatchAsync(new DriveItem
                {
                    FileSystemInfo = new Microsoft.Graph.Models.FileSystemInfo
                    {
                        CreatedDateTime      = metadata.CreatedDateTime,
                        LastModifiedDateTime = metadata.ModifiedDateTime
                    }
                });
            }
        }
        catch { /* best-effort */ }

        // Patch Author/Editor via listItem/fields — requires user to exist in the target site
        try
        {
            var userFields = new Dictionary<string, object>();

            if (!string.IsNullOrEmpty(metadata.CreatedByEmail))
            {
                var id = await LookupSharePointUserIdAsync(siteId, metadata.CreatedByEmail);
                if (id.HasValue) userFields["AuthorLookupId"] = id.Value.ToString();
            }
            if (!string.IsNullOrEmpty(metadata.ModifiedByEmail))
            {
                var id = await LookupSharePointUserIdAsync(siteId, metadata.ModifiedByEmail);
                if (id.HasValue) userFields["EditorLookupId"] = id.Value.ToString();
            }

            if (userFields.Count > 0)
            {
                var token = await _authService.GetAccessTokenAsync();
                var url = $"https://graph.microsoft.com/v1.0/drives/{driveId}/items/{itemId}/listItem/fields";
                using var req = new HttpRequestMessage(HttpMethod.Patch, url)
                {
                    Content = new System.Net.Http.StringContent(
                        System.Text.Json.JsonSerializer.Serialize(userFields),
                        System.Text.Encoding.UTF8, "application/json")
                };
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                req.Headers.Accept.ParseAdd("application/json");
                await _httpClient.SendAsync(req);
            }
        }
        catch { /* best-effort */ }
    }

    // ── Upload ────────────────────────────────────────────────────────────────

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

        // Large-file upload session
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
