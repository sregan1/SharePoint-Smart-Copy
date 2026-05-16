using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
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

    public async Task<List<SharePointNode>> GetLibrariesAsync(string siteId, string siteUrl)
    {
        // No $select — the Kiota SDK does not reliably populate webUrl when explicitly selected
        // on the drives collection; the default response includes it.
        var drives = await Graph.Sites[siteId].Drives.GetAsync();
        var result = new List<SharePointNode>();

        foreach (var d in drives?.Value ?? [])
        {
            if (d.Id == null || d.Name == null) continue;

            string? serverRelUrl = null;
            if (d.WebUrl != null)
            {
                var uri = new Uri(d.WebUrl);
                var path = Uri.UnescapeDataString(uri.AbsolutePath.TrimEnd('/'));
                // Exclude view URLs like /Forms/AllItems.aspx — these aren't the library root
                if (!path.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
                    serverRelUrl = path;
            }

            var node = new SharePointNode
            {
                Id                 = "root",
                DriveId            = d.Id,
                SiteId             = siteId,
                SiteUrl            = siteUrl,
                Name               = d.Name,
                Type               = NodeType.Library,
                HasChildren        = true,
                ServerRelativePath = serverRelUrl
            };
            node.Children.Add(Placeholder());
            result.Add(node);
        }
        return result;
    }

    // ── Children ──────────────────────────────────────────────────────────────

    public async Task<List<SharePointNode>> GetChildrenAsync(
        string driveId, string itemId, string siteId, string siteUrl = "", bool foldersOnly = false)
    {
        var resolvedId = itemId == "root" ? "root" : itemId;
        var page = await Graph.Drives[driveId].Items[resolvedId].Children
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
                SiteUrl     = siteUrl,
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

    // ── Enumerate all sub-folders under a folder ──────────────────────────────

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
        // Sort oldest-first so we replay versions in chronological order
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

    public static string? GetIdentityEmail(Microsoft.Graph.Models.Identity? identity)
    {
        if (identity?.AdditionalData == null) return null;
        if (identity.AdditionalData.TryGetValue("email", out var email) && email != null)
            return email.ToString();
        if (identity.AdditionalData.TryGetValue("userPrincipalName", out var upn) && upn != null)
            return upn.ToString();
        return null;
    }

    private async Task<(string siteUrl, string listId, string listItemId)?> GetSharePointIdsAsync(
        string driveId, string itemId)
    {
        var key = $"{driveId}|{itemId}";
        if (_spIdsCache.TryGetValue(key, out var cached)) return cached;

        for (int attempt = 0; attempt < 3; attempt++)
        {
            if (attempt > 0) await Task.Delay(attempt * 1500);
            try
            {
                var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
                    cfg.QueryParameters.Select = ["sharepointIds"]);

                var ids = item?.SharepointIds;
                if (ids?.SiteUrl == null || ids.ListId == null || ids.ListItemId == null)
                    continue;

                var result = (ids.SiteUrl, ids.ListId, ids.ListItemId);
                _spIdsCache[key] = result;
                return result;
            }
            catch { /* retry */ }
        }
        return null;
    }

    private async Task<string?> PatchTimestampsViaRestAsync(
        string driveId, string itemId, DateTimeOffset? modified, DateTimeOffset? created,
        string? createdByEmail = null, string? modifiedByEmail = null)
    {
        var ids = await GetSharePointIdsAsync(driveId, itemId);
        if (ids == null) return "SP IDs unavailable — item not found or sharepointIds not propagated";

        var (siteUrl, listId, listItemId) = ids.Value;

        var formValues = new List<object>();
        if (modified.HasValue)
            formValues.Add(new { FieldName = "Modified", FieldValue = modified.Value.ToUniversalTime().ToString("M/d/yyyy H:mm:ss", System.Globalization.CultureInfo.InvariantCulture) });
        if (created.HasValue)
            formValues.Add(new { FieldName = "Created",  FieldValue = created.Value.ToUniversalTime().ToString("M/d/yyyy H:mm:ss", System.Globalization.CultureInfo.InvariantCulture) });
        if (!string.IsNullOrEmpty(modifiedByEmail))
            formValues.Add(new { FieldName = "Editor", FieldValue = $"i:0#.f|membership|{modifiedByEmail}" });
        if (!string.IsNullOrEmpty(createdByEmail))
            formValues.Add(new { FieldName = "Author", FieldValue = $"i:0#.f|membership|{createdByEmail}" });

        if (formValues.Count == 0) return null;

        try
        {
            var token = await _authService.GetSharePointTokenAsync(siteUrl);
            var url   = $"{siteUrl}/_api/web/lists('{listId}')/items({listItemId})/ValidateUpdateListItem()";
            using var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new System.Net.Http.StringContent(
                    JsonSerializer.Serialize(new { formValues, bNewDocumentUpdate = true }),
                    System.Text.Encoding.UTF8, "application/json")
            };
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            req.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            var response = await _httpClient.SendAsync(req);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var preview = body.Length > 200 ? body[..200] : body;
                return $"ValidateUpdateListItem HTTP {(int)response.StatusCode}: {preview}";
            }

            try
            {
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("value", out var arr))
                {
                    var fieldErrors = new List<string>();
                    foreach (var field in arr.EnumerateArray())
                    {
                        bool hasEx = field.TryGetProperty("HasException", out var he) && he.GetBoolean();
                        int  ec    = field.TryGetProperty("ErrorCode",    out var ecEl) ? ecEl.GetInt32() : 0;
                        if (hasEx || ec != 0)
                        {
                            var fn = field.TryGetProperty("FieldName",    out var fnEl) ? fnEl.GetString() ?? "" : "";
                            var em = field.TryGetProperty("ErrorMessage", out var emEl) ? emEl.GetString() ?? "" : "";
                            fieldErrors.Add($"{fn}: {em} (code {ec})");
                        }
                    }
                    if (fieldErrors.Count > 0)
                        return $"Field errors: {string.Join("; ", fieldErrors)}";
                }
            }
            catch { }

            return null;
        }
        catch (Exception ex)
        {
            return $"ValidateUpdateListItem exception: {ex.Message}";
        }
    }

    public async Task<string?> ApplyFileMetadataAsync(
        string driveId, string itemId, string siteId, FileMetadata metadata)
    {
        return await PatchTimestampsViaRestAsync(driveId, itemId,
            metadata.ModifiedDateTime, metadata.CreatedDateTime,
            metadata.CreatedByEmail, metadata.ModifiedByEmail);
    }

    public async Task<string?> GetCurrentVersionIdAsync(string driveId, string itemId)
    {
        try
        {
            var page = await Graph.Drives[driveId].Items[itemId].Versions.GetAsync();
            return page?.Value?.FirstOrDefault()?.Id;
        }
        catch { return null; }
    }

    // Patches fileSystemInfo.lastModifiedDateTime — the field shown in SharePoint version history.
    // Always creates a phantom version in a versioned library.
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
            return null;
        }
        catch (Exception ex)
        {
            return $"PatchFileSystemDate exception: {ex.Message}";
        }
    }

    public async Task<string?> DeleteItemVersionAsync(string driveId, string itemId, string versionId)
    {
        try
        {
            await Graph.Drives[driveId].Items[itemId].Versions[versionId].DeleteAsync();
            return null;
        }
        catch (Exception ex)
        {
            return $"DeleteItemVersion exception: {ex.Message}";
        }
    }

    // ── Migration API helpers ─────────────────────────────────────────────────

    // Calls /_api/site/ProvisionMigrationContainers on the target site.
    // Returns (dataContainerUri, metadataContainerUri, encryptionKey).
    public async Task<(string dataUri, string metadataUri, byte[] encryptionKey)>
        ProvisionMigrationContainersAsync(string siteUrl)
    {
        var token = await _authService.GetSharePointTokenAsync(siteUrl);
        var url = $"{siteUrl.TrimEnd('/')}/_api/site/ProvisionMigrationContainers";
        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.ParseAdd("application/json;odata=nometadata");
        req.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");

        var response = await _httpClient.SendAsync(req);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"ProvisionMigrationContainers HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;
        var dataUri     = root.GetProperty("DataContainerUri").GetString()!;
        var metadataUri = root.GetProperty("MetadataContainerUri").GetString()!;
        var keyBase64   = root.GetProperty("EncryptionKey").GetString()!;
        return (dataUri, metadataUri, Convert.FromBase64String(keyBase64));
    }

    // Returns the signed-in user's display name, email, and IsSiteAdmin flag for siteUrl.
    public async Task<(string title, string email, bool isSiteAdmin)> GetCurrentUserInfoAsync(string siteUrl)
    {
        var token = await _authService.GetSharePointTokenAsync(siteUrl, "AllSites.FullControl");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/currentuser?$select=Title,Email,LoginName,IsSiteAdmin";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.ParseAdd("application/json;odata=nometadata");

        var response = await _httpClient.SendAsync(req);
        if (!response.IsSuccessStatusCode)
            return ("(unknown)", "(HTTP error)", false);

        var body = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;
        var title   = root.TryGetProperty("Title",      out var t) ? t.GetString() ?? "" : "";
        var email   = root.TryGetProperty("Email",      out var e) ? e.GetString() ?? "" : "";
        var login   = root.TryGetProperty("LoginName",  out var l) ? l.GetString() ?? "" : "";
        var isAdmin = root.TryGetProperty("IsSiteAdmin", out var a) && a.GetBoolean();
        return (title, string.IsNullOrEmpty(email) ? login : email, isAdmin);
    }

    // Returns (webId, serverRelativeUrl) for the root web of the given site URL.
    public async Task<(string webId, string serverRelativeUrl)> GetWebInfoAsync(string siteUrl)
    {
        var token = await _authService.GetSharePointTokenAsync(siteUrl);
        var url = $"{siteUrl.TrimEnd('/')}/_api/web?$select=Id,ServerRelativeUrl";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.ParseAdd("application/json;odata=nometadata");

        var response = await _httpClient.SendAsync(req);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"GetWebInfo HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        var webId  = doc.RootElement.GetProperty("Id").GetString()!;
        var relUrl = doc.RootElement.GetProperty("ServerRelativeUrl").GetString()!;
        return (webId, relUrl);
    }

    // Returns the GUID of a document library given its server-relative URL.
    public async Task<string> GetListIdByServerRelativeUrlAsync(string siteUrl, string serverRelativeUrl)
    {
        var token = await _authService.GetSharePointTokenAsync(siteUrl);
        var encodedUrl = Uri.EscapeDataString($"'{serverRelativeUrl}'");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetList({encodedUrl})?$select=Id";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.ParseAdd("application/json;odata=nometadata");

        var response = await _httpClient.SendAsync(req);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"GetListId HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        return doc.RootElement.GetProperty("Id").GetString()!;
    }

    // Gets the server-relative URL of a document library by fetching the drive's root item webUrl.
    // More reliable than deriving it from drive.WebUrl which the Kiota SDK may not populate.
    public async Task<string> GetLibraryServerRelativeUrlAsync(string driveId)
    {
        var root = await Graph.Drives[driveId].Root
            .GetAsync(cfg => cfg.QueryParameters.Select = ["webUrl"]);
        if (root?.WebUrl == null)
            throw new InvalidOperationException($"Cannot determine library path for drive {driveId}.");
        return Uri.UnescapeDataString(new Uri(root.WebUrl).AbsolutePath.TrimEnd('/'));
    }

    // Submits a migration job using SP-provided encrypted containers.
    // Returns the job ID GUID.
    //
    // Uses raw CSOM ProcessQuery XML rather than the CSOM client library because
    // the library's ExecutingWebRequest token injection is unreliable on .NET 8
    // (compiled for .NET 4.0; HttpWebRequest behaviour changed under the compat layer).
    public async Task<string> CreateMigrationJobEncryptedAsync(
        string siteUrl, string gWebId,
        string dataContainerUri, string metadataContainerUri,
        byte[] encryptionKey)
    {
        // AllSites.FullControl is required: Sites.ReadWrite.All caps effective privilege below
        // site-collection-admin level in SP's OAuth permission model, causing CreateMigrationJobEncrypted
        // to reject even explicit SCAs.  The Azure AD app must have AllSites.FullControl delegated
        // permission registered and admin-consented; the user will be prompted if not yet granted.
        var token    = await _authService.GetSharePointTokenAsync(siteUrl, "AllSites.FullControl");
        var keyB64   = Convert.ToBase64String(encryptionKey);

        // Build the ProcessQuery XML.  EncryptionOption TypeId is {85614ad4-7a40-49e0-b272-6d1807dbfcc6}.
        // AES256CBCKey is a byte[] serialised as Base64Binary.
        // SP reads the per-blob IV from the first 16 bytes of each encrypted blob,
        // so no IV is needed here — only the key.
        var ns = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/sharepoint/clientquery/2009");
        var requestXml = new System.Xml.Linq.XDocument(
            new System.Xml.Linq.XElement(ns + "Request",
                new System.Xml.Linq.XAttribute("SchemaVersion",   "15.0.0.0"),
                new System.Xml.Linq.XAttribute("LibraryVersion",  "16.0.0.0"),
                new System.Xml.Linq.XAttribute("ApplicationName", "SharePointSmartCopy"),
                new System.Xml.Linq.XElement(ns + "Actions",
                    new System.Xml.Linq.XElement(ns + "Method",
                        new System.Xml.Linq.XAttribute("Name",         "CreateMigrationJobEncrypted"),
                        new System.Xml.Linq.XAttribute("Id",           "1"),
                        new System.Xml.Linq.XAttribute("ObjectPathId", "2"),
                        new System.Xml.Linq.XElement(ns + "Parameters",
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("Type", "Guid"), gWebId),
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("Type", "String"), dataContainerUri),
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("Type", "String"), metadataContainerUri),
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("Type", "String"), ""),
                            new System.Xml.Linq.XElement(ns + "Parameter",
                                new System.Xml.Linq.XAttribute("TypeId", "{85614ad4-7a40-49e0-b272-6d1807dbfcc6}"),
                                new System.Xml.Linq.XElement(ns + "Property",
                                    new System.Xml.Linq.XAttribute("Name", "AES256CBCKey"),
                                    new System.Xml.Linq.XAttribute("Type", "Base64Binary"),
                                    keyB64))))),
                new System.Xml.Linq.XElement(ns + "ObjectPaths",
                    new System.Xml.Linq.XElement(ns + "Property",
                        new System.Xml.Linq.XAttribute("Id",       "2"),
                        new System.Xml.Linq.XAttribute("ParentId", "0"),
                        new System.Xml.Linq.XAttribute("Name",     "Site")),
                    new System.Xml.Linq.XElement(ns + "StaticProperty",
                        new System.Xml.Linq.XAttribute("Id",     "0"),
                        new System.Xml.Linq.XAttribute("TypeId", "{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}"),
                        new System.Xml.Linq.XAttribute("Name",   "Current")))));

        var xmlBody = requestXml.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        var url     = $"{siteUrl.TrimEnd('/')}/_vti_bin/client.svc/ProcessQuery";

        using var req = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(xmlBody, System.Text.Encoding.UTF8, "text/xml")
        };
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.ParseAdd("application/json;odata=nometadata");

        var response = await _httpClient.SendAsync(req);
        var body     = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"CreateMigrationJobEncrypted HTTP {(int)response.StatusCode}: {body}");

        // ProcessQuery returns a JSON array.  Index 0 is the header with ErrorInfo.
        // On success the job GUID appears as a bare string element later in the array.
        using var doc = JsonDocument.Parse(body);
        var arr = doc.RootElement;

        if (arr.GetArrayLength() > 0 &&
            arr[0].TryGetProperty("ErrorInfo", out var errInfo) &&
            errInfo.ValueKind != JsonValueKind.Null)
        {
            var msg = errInfo.TryGetProperty("ErrorMessage", out var m) ? m.GetString() : "Unknown CSOM error";
            var code = errInfo.TryGetProperty("ErrorCode", out var c) ? c.GetInt32() : 0;
            if (code == -2147024891) // E_ACCESSDENIED
            {
                var (uTitle, uEmail, uIsAdmin) = await GetCurrentUserInfoAsync(siteUrl);
                throw new UnauthorizedAccessException(
                    $"The Migration API requires explicit Site Collection Administrator membership on {siteUrl}.\n\n" +
                    $"SP sees you as: {uTitle} ({uEmail}), IsSiteAdmin={uIsAdmin}\n\n" +
                    "Note: Global Admin / SharePoint Admin does not automatically populate this list.\n" +
                    "Fix: go to the target site → Site Settings → Site Collection Administrators → add your account.",
                    new Exception(msg));
            }
            throw new Exception($"CreateMigrationJobEncrypted CSOM error {code}: {msg}");
        }

        for (int i = 1; i < arr.GetArrayLength(); i++)
        {
            var el = arr[i];
            if (el.ValueKind != JsonValueKind.String) continue;
            var s = el.GetString() ?? "";
            // CSOM returns Guids as "\/Guid(xxxxxxxx-...)\/"; also handle bare GUIDs.
            if (s.StartsWith("/Guid(", StringComparison.OrdinalIgnoreCase) && s.EndsWith(")/"))
                s = s[6..^2];
            if (Guid.TryParse(s, out var g))
                return g.ToString("D");
        }

        throw new Exception($"Could not find job ID GUID in ProcessQuery response: {body[..Math.Min(body.Length, 500)]}");
    }

    // Polls the migration job using the paging-based GetMigrationJobProgress endpoint.
    // Yields each event JSON string. Caller should stop when a JobEnd event is received.
    public async IAsyncEnumerable<JsonElement> PollMigrationJobAsync(
        string siteUrl, string jobId, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var token = await _authService.GetSharePointTokenAsync(siteUrl);
        int nextToken = 0;

        while (!cancellationToken.IsCancellationRequested)
        {
            await Task.Delay(TimeSpan.FromSeconds(15), cancellationToken);

            // Guid parameters in SP REST require the guid'...' syntax, not just '...'.
            // AllSites.FullControl is needed — the same reason CreateMigrationJobEncrypted needs it.
            var url = $"{siteUrl.TrimEnd('/')}/_api/site/GetMigrationJobProgress(jobId=guid'{jobId}',nextToken={nextToken})";
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer",
                await _authService.GetSharePointTokenAsync(siteUrl, "AllSites.FullControl", cancellationToken));
            req.Headers.Accept.ParseAdd("application/json;odata=nometadata");

            HttpResponseMessage response;
            try
            {
                response = await _httpClient.SendAsync(req, cancellationToken);
            }
            catch (TaskCanceledException) { yield break; }

            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync(cancellationToken);
                throw new Exception($"GetMigrationJobProgress HTTP {(int)response.StatusCode}: {err[..Math.Min(err.Length, 300)]}");
            }

            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            System.Diagnostics.Debug.WriteLine($"[Poll] raw response ({body.Length} bytes): {body[..Math.Min(body.Length, 3000)]}");
            JsonDocument doc;
            try { doc = JsonDocument.Parse(body); }
            catch { continue; }

            bool hitJobEnd = false;
            using (doc)
            {
                // SP GetMigrationJobProgress returns { "Logs": [...], "nextToken": N }
                // where each element of Logs is a JSON-encoded string (not an object).
                // Some older docs describe "value" or a bare array — handle all variants.
                JsonElement arr = default;
                if (doc.RootElement.ValueKind == JsonValueKind.Array)
                    arr = doc.RootElement;
                else if (doc.RootElement.TryGetProperty("Logs", out var logs))
                    arr = logs;
                else if (doc.RootElement.TryGetProperty("value", out var v))
                    arr = v;

                if (arr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var rawEvt in arr.EnumerateArray())
                    {
                        JsonElement evtObj;
                        if (rawEvt.ValueKind == JsonValueKind.String)
                        {
                            try
                            {
                                using var inner = JsonDocument.Parse(rawEvt.GetString()!);
                                evtObj = inner.RootElement.Clone();
                            }
                            catch { continue; }
                        }
                        else if (rawEvt.ValueKind == JsonValueKind.Object)
                            evtObj = rawEvt.Clone();
                        else
                            continue;

                        if (!evtObj.TryGetProperty("Event", out var evtName)) continue;
                        yield return evtObj;
                        if (evtName.GetString() == "JobEnd")
                        {
                            hitJobEnd = true;
                            break;
                        }
                    }
                }

                // nextToken may be a string or number, and may be null
                if (doc.RootElement.TryGetProperty("nextToken", out var nt))
                {
                    if (nt.ValueKind == JsonValueKind.Number) nextToken = nt.GetInt32();
                    else if (nt.ValueKind == JsonValueKind.String && int.TryParse(nt.GetString(), out var p)) nextToken = p;
                }
                else if (doc.RootElement.TryGetProperty("NextToken", out var nt2))
                {
                    if (nt2.ValueKind == JsonValueKind.Number) nextToken = nt2.GetInt32();
                    else if (nt2.ValueKind == JsonValueKind.String && int.TryParse(nt2.GetString(), out var p2)) nextToken = p2;
                }
            }

            if (hitJobEnd) yield break;
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
        catch { return false; }
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
            current = await GetOrCreateFolderAsync(driveId, current, part);
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
