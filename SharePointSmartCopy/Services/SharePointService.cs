using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Http.HttpClientLibrary;
using SharePointSmartCopy.Models;
using SpColumnDef  = SharePointSmartCopy.Models.ColumnDefinition;
using SpColumnMap  = SharePointSmartCopy.Models.ColumnMapping;

namespace SharePointSmartCopy.Services;

public class SharePointService
{
    private readonly AuthService _authService;
    private GraphServiceClient? _graphClient;
    // 30-minute timeout: site copies involve large file downloads and long-running REST calls.
    private readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromMinutes(30) };
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (string siteUrl, string listId, string listItemId)> _spIdsCache = new();

    public SharePointService(AuthService authService)
    {
        _authService = authService;
    }

    public Task<string> GetSharePointTokenAsync(string siteUrl, string scope = "Sites.ReadWrite.All")
        => _authService.GetSharePointTokenAsync(siteUrl, scope);

    // Returns the root item ID for a drive (needed to target uploads to a library root).
    public async Task<string?> GetLibraryRootItemIdAsync(string driveId)
    {
        try
        {
            var root = await Graph.Drives[driveId].Root.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id"]);
            return root?.Id;
        }
        catch { return null; }
    }

    // Returns the SharePoint folder UniqueId (GUID) for the library root folder.
    // Required for the SPMI manifest SPFolder entry — without it SPMI cannot resolve the
    // parent folder for files in newly created empty libraries ("Missing file info" error).
    public async Task<string?> GetLibraryRootFolderUniqueIdAsync(string siteUrl, string libraryServerRelativeUrl)
    {
        var encoded = Uri.EscapeDataString($"'{libraryServerRelativeUrl}'");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetFolderByServerRelativeUrl({encoded})?$select=UniqueId";
        try
        {
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);
            if (!response.IsSuccessStatusCode) return null;
            var body = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("UniqueId", out var uid) ? uid.GetString() : null;
        }
        catch { return null; }
    }

    public void Initialize()
    {
        var provider   = new BaseBearerTokenAuthenticationProvider(new TokenProvider(_authService));
        // KiotaClientFactory wires up all Graph middleware (retry, redirect, auth challenges, etc.)
        // then we extend the timeout so large file downloads don't cancel at the 100-second default.
        var httpClient = KiotaClientFactory.Create();
        httpClient.Timeout = TimeSpan.FromMinutes(30);
        var adapter    = new HttpClientRequestAdapter(provider, httpClient: httpClient);
        _graphClient   = new GraphServiceClient(adapter);
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
        // Sort oldest-first so we replay versions in chronological order.
        // Sort by the numeric version label ("2.0" → 2.0) rather than timestamp:
        // versions saved within the same second would otherwise keep Graph's
        // newest-first order and be replayed out of sequence.
        return all.OrderBy(v =>
        {
            var parts = (v.Id ?? "0").Split('.');
            double major = parts.Length > 0 && double.TryParse(parts[0], out var mj) ? mj : 0;
            double minor = parts.Length > 1 && double.TryParse(parts[1], out var mn) ? mn : 0;
            return major + minor / 1000.0;
        }).ToList();
    }

    // ── Metadata ──────────────────────────────────────────────────────────────

    public async Task<FileMetadata> GetFileMetadataAsync(string driveId, string itemId)
    {
        var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
            cfg.QueryParameters.Select = ["createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy"]);

        return new FileMetadata
        {
            CreatedDateTime  = item?.CreatedDateTime,
            CreatedByEmail   = GetIdentityEmail(item?.CreatedBy?.User),
            ModifiedDateTime = item?.LastModifiedDateTime,
            ModifiedByEmail  = GetIdentityEmail(item?.LastModifiedBy?.User),
        };
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

    public async Task<(string siteUrl, string listId, string listItemId)?> GetSharePointIdsAsync(
        string driveId, string itemId)
    {
        var key = $"{driveId}|{itemId}";
        if (_spIdsCache.TryGetValue(key, out var cached))
        {
            System.Diagnostics.Debug.WriteLine($"[GetSPIds] cache hit: itemId={itemId}");
            return cached;
        }

        for (int attempt = 0; attempt < 3; attempt++)
        {
            if (attempt > 0) await Task.Delay(attempt * 1500);
            try
            {
                System.Diagnostics.Debug.WriteLine($"[GetSPIds] attempt {attempt + 1}/3 for itemId={itemId}");
                var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
                    cfg.QueryParameters.Select = ["sharepointIds"]);

                var ids = item?.SharepointIds;
                if (ids?.SiteUrl == null || ids.ListId == null || ids.ListItemId == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[GetSPIds] attempt {attempt + 1}: sharepointIds not yet populated for itemId={itemId}, will retry");
                    continue;
                }

                System.Diagnostics.Debug.WriteLine($"[GetSPIds] success: listItemId={ids.ListItemId} siteUrl={ids.SiteUrl}");
                var result = (ids.SiteUrl, ids.ListId, ids.ListItemId);
                _spIdsCache[key] = result;
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetSPIds] attempt {attempt + 1} threw {ex.GetType().Name}: {ex.Message}");
            }
        }
        System.Diagnostics.Debug.WriteLine($"[GetSPIds] all 3 attempts failed for itemId={itemId}");
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
        System.Diagnostics.Debug.WriteLine($"[PatchFSDate] itemId={itemId} modified={lastModifiedDateTime:u}");
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
            System.Diagnostics.Debug.WriteLine($"[PatchFSDate] success for itemId={itemId}");
            return null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PatchFSDate] FAILED for itemId={itemId}: {ex.GetType().Name}: {ex.Message}");
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
        var url = $"{siteUrl.TrimEnd('/')}/_api/site/ProvisionMigrationContainers";
        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl);
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
        var url = $"{siteUrl.TrimEnd('/')}/_api/web?$select=Id,ServerRelativeUrl";
        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl);
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
        var encodedUrl = Uri.EscapeDataString($"'{serverRelativeUrl}'");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetList({encodedUrl})?$select=Id";
        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"GetListId HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        return doc.RootElement.GetProperty("Id").GetString()!;
    }

    // Gets the SharePoint list GUID for a document library via the Graph drive's associated list.
    // More reliable than GetListIdByServerRelativeUrlAsync when the server-relative URL is uncertain.
    public async Task<string?> GetListIdFromDriveAsync(string driveId)
    {
        try
        {
            var list = await Graph.Drives[driveId].List
                .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
            return list?.Id;
        }
        catch
        {
            return null;
        }
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
        var keyB64 = Convert.ToBase64String(encryptionKey);

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

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(xmlBody, System.Text.Encoding.UTF8, "text/xml")
            };
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl, "AllSites.FullControl");
        var body = await response.Content.ReadAsStringAsync();
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
            await Task.Delay(TimeSpan.FromSeconds(3), cancellationToken);

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
            System.Diagnostics.Debug.WriteLine($"[FileExists] EXISTS: {fileName}");
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[FileExists] NOT FOUND (expected 404): {fileName} — {ex.GetType().Name}: {ex.Message}");
            return false;
        }
    }

    public async Task DeleteFileIfExistsAsync(string driveId, string parentItemId, string fileName)
    {
        try
        {
            await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName)).DeleteAsync();
        }
        catch { }
    }

    // Returns the SharePoint UniqueId (AllDocs GUID) for a file by its server-relative URL.
    // Works via REST (not Graph) so it finds zombie files — SPFile blobs without a list item
    // that Graph returns 404 for. Returns null if the file doesn't exist.
    public async Task<string?> GetFileUniqueIdAsync(string siteUrl, string serverRelativeUrl)
    {
        var encodedPath = Uri.EscapeDataString($"'{serverRelativeUrl}'");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetFileByServerRelativeUrl({encodedPath})/UniqueId";
        try
        {
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!response.IsSuccessStatusCode) return null;
            var body = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.ValueKind == JsonValueKind.String
                ? doc.RootElement.GetString()
                : doc.RootElement.TryGetProperty("value", out var v) ? v.GetString() : null;
        }
        catch { return null; }
    }

    // Permanently deletes a file by recycling it then immediately purging the recycle bin entry.
    // Graph DeleteAsync only soft-deletes (recycle bin); soft-deleted list item records interfere
    // with SPMI imports at the same URL, producing "Missing file info for list item" errors.
    public async Task PermanentlyDeleteFileAsync(string siteUrl, string serverRelativeUrl)
    {
        // Step 1: recycle — returns the recycle bin item GUID
        var encodedPath = Uri.EscapeDataString($"'{serverRelativeUrl}'");
        var recycleUrl  = $"{siteUrl.TrimEnd('/')}/_api/web/GetFileByServerRelativeUrl({encodedPath})/recycleObject";

        string? recycleBinGuid = null;
        try
        {
            using var recycleResponse = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, recycleUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);

            System.Diagnostics.Debug.WriteLine($"[PermDelete] recycleObject status={recycleResponse.StatusCode} path={serverRelativeUrl}");

            if (!recycleResponse.IsSuccessStatusCode)
            {
                // recycleObject fails for zombie blobs (no SPListItem → can't create recycle entry).
                // Fall back to deleteObject which operates at the AllDocs level directly.
                await TryDeleteObjectAsync(siteUrl, encodedPath);
                return;
            }

            var body = await recycleResponse.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine($"[PermDelete] recycleObject body={body[..Math.Min(body.Length, 200)]}");
            using var doc = JsonDocument.Parse(body);
            recycleBinGuid = doc.RootElement.ValueKind == JsonValueKind.String
                ? doc.RootElement.GetString()
                : doc.RootElement.TryGetProperty("value", out var v) ? v.GetString() : null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PermDelete] recycleObject error: {ex.Message}");
            return;
        }

        System.Diagnostics.Debug.WriteLine($"[PermDelete] recycleBinGuid={recycleBinGuid}");
        if (string.IsNullOrEmpty(recycleBinGuid)) return;

        // Step 2: purge from recycle bin — removes all DB records so SPMI can import cleanly.
        // SP REST RecycleBin key uses single-quoted string (not guid'...' OData Guid literal).
        // Uses Sites.ReadWrite.All — user is SCA so this token is already cached and sufficient.
        try
        {
            var purgeUrl = $"{siteUrl.TrimEnd('/')}/_api/site/RecycleBin('{recycleBinGuid}')/DeleteObject";
            using var purgeResponse = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, purgeUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
            System.Diagnostics.Debug.WriteLine($"[PermDelete] purge status={purgeResponse.StatusCode}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PermDelete] purge error: {ex.Message}");
        }
    }

    private async Task TryDeleteObjectAsync(string siteUrl, string encodedPath)
    {
        var deleteUrl = $"{siteUrl.TrimEnd('/')}/_api/web/GetFileByServerRelativeUrl({encodedPath})/deleteObject";
        try
        {
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, deleteUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
            System.Diagnostics.Debug.WriteLine($"[PermDelete] deleteObject status={response.StatusCode}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PermDelete] deleteObject error: {ex.Message}");
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

    // Creates a modern Site Page via the SitePages REST API and returns the Graph item ID
    // and the SitePages integer ID. The SitePages ID is needed for SavePage / Publish calls.
    // AddTemplateFile(templateFileType=0) creates a classic page without CanvasContent1 support.
    public async Task<(string graphItemId, int sitePagesId)> CreatePageStubAsync(
        string siteUrl, string targetFolderRelUrl,
        string driveId, string parentItemId,
        string fileName, bool overwrite)
    {
        System.Diagnostics.Debug.WriteLine($"[CreatePageStub] {fileName} → {siteUrl}");

        // Delete existing page first — SitePages API returns 409 if the file already exists.
        if (overwrite)
        {
            try
            {
                var existing = await Graph.Drives[driveId].Items[parentItemId]
                    .ItemWithPath(Uri.EscapeDataString(fileName))
                    .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
                if (existing?.Id != null)
                {
                    await Graph.Drives[driveId].Items[existing.Id].DeleteAsync();
                    System.Diagnostics.Debug.WriteLine($"[CreatePageStub] deleted existing: {existing.Id}");
                }
            }
            catch { /* file doesn't exist — nothing to delete */ }
        }

        var title = System.IO.Path.GetFileNameWithoutExtension(fileName)
                        .Replace('-', ' ').Replace('_', ' ');

        var body = System.Text.Json.JsonSerializer.Serialize(new
        {
            __metadata     = new { type = "SP.Publishing.SitePage" },
            FileName       = fileName,
            Title          = title,
            PageLayoutType = "Article",
        });

        using var createResp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post,
                $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages");
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=verbose");
            r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
            r.Content.Headers.ContentType =
                System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            return r;
        }, siteUrl);

        if (!createResp.IsSuccessStatusCode)
        {
            var err = await createResp.Content.ReadAsStringAsync();
            throw new Exception(
                $"Page create HTTP {(int)createResp.StatusCode}: {err[..Math.Min(400, err.Length)]}");
        }

        // Parse the SitePages integer ID from the response — needed for SavePage and Publish.
        var respBody = await createResp.Content.ReadAsStringAsync();
        int sitePagesId = 0;
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(respBody);
            var root = doc.RootElement.TryGetProperty("d", out var d) ? d : doc.RootElement;
            if (root.TryGetProperty("Id", out var idProp))
                sitePagesId = idProp.GetInt32();
        }
        catch { }
        System.Diagnostics.Debug.WriteLine($"[CreatePageStub] sitePagesId={sitePagesId}");

        // Resolve Graph item ID with retry (Graph can lag 1–2 s after SitePages creation).
        for (int attempt = 0; attempt < 3; attempt++)
        {
            if (attempt > 0) await Task.Delay(1000);
            try
            {
                var item = await Graph.Drives[driveId].Items[parentItemId]
                    .ItemWithPath(Uri.EscapeDataString(fileName))
                    .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
                if (item?.Id != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[CreatePageStub] graphItemId={item.Id}");
                    return (item.Id, sitePagesId);
                }
            }
            catch { }
        }
        System.Diagnostics.Debug.WriteLine($"[CreatePageStub] could not resolve Graph item ID after 3 attempts");
        return (string.Empty, sitePagesId);
    }

    // Writes page content to a draft modern page via the SitePages SavePage API.
    // ValidateUpdateListItem writes the list item fields but the modern page renderer reads
    // content from the SitePages checkout state — SavePage is the correct write path.
    public async Task<string?> SavePageContentAsync(
        string siteUrl, int sitePagesId, PageMetadata pageMeta, string sourceSiteUrl)
    {
        if (sitePagesId == 0) return "SitePages ID unknown — cannot save content";

        string? SubstUrl(string? json) =>
            json == null ? null : SubstituteUrls(json, sourceSiteUrl, siteUrl);

        // Build a save body with only the fields that have values
        var fields = new Dictionary<string, object>
        {
            ["__metadata"] = (object)new { type = "SP.Publishing.SitePage" },
        };
        var canvas  = SubstUrl(pageMeta.CanvasContent1);
        var layout  = SubstUrl(pageMeta.LayoutWebpartsContent);
        var banner  = SubstUrl(pageMeta.BannerImageUrl);
        if (canvas  != null) fields["CanvasContent1"]        = canvas;
        if (layout  != null) fields["LayoutWebpartsContent"] = layout;
        if (banner  != null) fields["BannerImageUrl"]        = banner;
        if (pageMeta.Description    != null) fields["Description"]    = pageMeta.Description;
        if (pageMeta.PageLayoutType != null) fields["PageLayoutType"]  = pageMeta.PageLayoutType;
        fields["PromotedState"] = (object)pageMeta.PromotedState;

        var saveBody = System.Text.Json.JsonSerializer.Serialize(fields);
        var saveUrl  = $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages({sitePagesId})/SavePage";

        System.Diagnostics.Debug.WriteLine($"[SavePageContent] POST {saveUrl}");
        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, saveUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=verbose");
            r.Content = new StringContent(saveBody, System.Text.Encoding.UTF8, "application/json");
            r.Content.Headers.ContentType =
                System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            return r;
        }, siteUrl);

        System.Diagnostics.Debug.WriteLine($"[SavePageContent] status={resp.StatusCode}");
        if (!resp.IsSuccessStatusCode)
        {
            var err = await resp.Content.ReadAsStringAsync();
            return $"SavePage HTTP {(int)resp.StatusCode}: {err[..Math.Min(300, err.Length)]}";
        }
        return null;
    }

    // Publishes a draft modern page via the SitePages Publish endpoint.
    // Must be called after SavePageContentAsync so the published version contains the content.
    public async Task<string?> PublishPageAsync(string siteUrl, int sitePagesId)
    {
        if (sitePagesId == 0) return "SitePages ID unknown — cannot publish";
        var url = $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages({sitePagesId})/Publish";
        try
        {
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
            System.Diagnostics.Debug.WriteLine($"[PublishPage] status={response.StatusCode}");
            return response.IsSuccessStatusCode ? null : $"Publish HTTP {(int)response.StatusCode}";
        }
        catch (Exception ex)
        {
            return $"Publish: {ex.Message}";
        }
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
            System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] EXISTS: {folderName}");
            return existing?.Id ?? throw new Exception("Null ID from existing folder.");
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] not found, creating: {folderName}");
            try
            {
                var created = await Graph.Drives[driveId].Items[parentItemId].Children.PostAsync(new DriveItem
                {
                    Name   = folderName,
                    Folder = new Folder()
                });
                System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] created: {folderName}");
                return created?.Id ?? throw new Exception("Null ID from created folder.");
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError conflict) when (conflict.ResponseStatusCode == 409)
            {
                System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] 409 conflict (race), re-fetching: {folderName}");
                var existing = await Graph.Drives[driveId].Items[parentItemId]
                    .ItemWithPath(Uri.EscapeDataString(folderName)).GetAsync();
                return existing?.Id ?? throw new Exception("Null ID after concurrent folder creation.");
            }
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    // Sends a SharePoint REST request with resilience:
    //  - 401: retried once with a force-refreshed token.
    //  - 429/503 (throttling): retried up to 3 times, honoring the Retry-After header.
    // buildRequest receives the bearer token and must return a fresh HttpRequestMessage each call.
    internal async Task<HttpResponseMessage> SendSharePointRequestAsync(
        Func<string, HttpRequestMessage> buildRequest,
        string siteUrl,
        string spScope = "Sites.ReadWrite.All",
        CancellationToken cancellationToken = default)
    {
        const int maxThrottleRetries = 3;
        var token = await _authService.GetSharePointTokenAsync(siteUrl, spScope, cancellationToken);
        HttpResponseMessage response;
        bool refreshedToken = false;

        for (int attempt = 0; ; attempt++)
        {
            using var req = buildRequest(token);
            response = await _httpClient.SendAsync(req, cancellationToken);

            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized && !refreshedToken)
            {
                response.Dispose();
                token = await _authService.GetSharePointTokenAsync(siteUrl, spScope, cancellationToken, forceRefresh: true);
                refreshedToken = true;
                continue;
            }

            bool throttled = response.StatusCode is System.Net.HttpStatusCode.TooManyRequests
                                                 or System.Net.HttpStatusCode.ServiceUnavailable;
            if (throttled && attempt < maxThrottleRetries)
            {
                var delay = response.Headers.RetryAfter?.Delta
                    ?? TimeSpan.FromSeconds(Math.Pow(2, attempt + 1)); // 2s, 4s, 8s
                System.Diagnostics.Debug.WriteLine(
                    $"[SP-REST] {(int)response.StatusCode} throttled — retrying in {delay.TotalSeconds:N0}s (attempt {attempt + 1}/{maxThrottleRetries})");
                response.Dispose();
                await Task.Delay(delay, cancellationToken);
                continue;
            }

            return response;
        }
    }

    private static SharePointNode Placeholder() =>
        new() { Name = "__placeholder__" };

    // SP REST pagination link: "odata.nextLink" under odata=nometadata (classic), but
    // "@odata.nextLink" under OData v4 / minimalmetadata responses. Check both so
    // pagination never silently truncates at the first page.
    private static string? GetNextLink(JsonElement root) =>
        root.TryGetProperty("odata.nextLink",  out var nl1) ? nl1.GetString() :
        root.TryGetProperty("@odata.nextLink", out var nl2) ? nl2.GetString() : null;

    // ── Column / Custom Field Methods ─────────────────────────────────────────

    // Built-in field internal names that should never be treated as custom columns.
    private static readonly HashSet<string> _builtInFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "ID", "Title", "Author", "Editor", "Created", "Modified", "ContentType",
        "FileLeafRef", "FileDirRef", "FileRef", "_UIVersionString", "MetaInfo",
        "_ModerationStatus", "_ModerationComments", "Edit", "SelectTitle", "InstanceID",
        "Order", "GUID", "WorkflowVersion", "_HasCopyDestinations", "_CopySource",
        "owshiddenversion", "WorkflowInstanceID", "ParentLeafName", "UniqueId",
        "ProgId", "ScopeId", "HTML_x0020_File_x0020_Type", "_EditMenuTableStart",
        "_EditMenuTableEnd", "LinkFilenameNoMenu", "LinkFilename", "DocConcurrencyToken",
        "SelectFilename", "ItemChildCount", "FolderChildCount", "Restricted",
        "OriginatorId", "NoExecute", "ContentVersion", "_dlc_DocId", "_dlc_DocIdUrl",
        "_dlc_DocIdPersistId", "TemplateUrl", "xd_ProgID", "xd_Signature",
        "AppAuthor", "AppEditor", "_IsCurrentVersion", "SMTotalSize",
        "SMLastModifiedDate", "_Level", "_IsCurrentVersion", "CheckedOutUserId",
        "IsCheckedoutToLocal", "CheckoutUser", "SMTotalFileStreamSize",
        "ComplianceAssetId", "_ComplianceFlags", "_ComplianceTag",
        "_ComplianceTagWrittenTime", "_ComplianceTagUserId",
        "AccessPolicy", "BSN", "MicroBlogging",
    };

    private static readonly Dictionary<int, SupportedFieldType> _fieldTypeMap = new()
    {
        { 2,  SupportedFieldType.Text },
        { 3,  SupportedFieldType.Note },
        { 4,  SupportedFieldType.DateTime },
        { 6,  SupportedFieldType.Choice },
        { 8,  SupportedFieldType.Boolean },
        { 9,  SupportedFieldType.Number },
        { 15, SupportedFieldType.MultiChoice },
    };

    // Keyed by listId. Concurrent — read/written from parallel copy tasks.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, List<SpColumnDef>> _columnCache = new();

    // Returns the custom (non-built-in) columns for a library.
    public async Task<List<SpColumnDef>> GetLibraryColumnsAsync(string siteUrl, string listId, bool skipCache = false)
    {
        if (!skipCache && _columnCache.TryGetValue(listId, out var cached))
            return cached;

        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/fields" +
                  "?$filter=Hidden eq false and ReadOnlyField eq false and FromBaseType eq false" +
                  "&$select=InternalName,Title,FieldTypeKind,Choices";

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl);

        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"GetLibraryColumns HTTP {(int)response.StatusCode}: {body}");

        using var doc = JsonDocument.Parse(body);
        var result = new List<SpColumnDef>();

        if (doc.RootElement.TryGetProperty("value", out var values))
        {
            foreach (var field in values.EnumerateArray())
            {
                var internalName = field.GetProperty("InternalName").GetString() ?? "";
                if (_builtInFields.Contains(internalName)) continue;
                var typeKind = field.GetProperty("FieldTypeKind").GetInt32();
                if (!_fieldTypeMap.TryGetValue(typeKind, out var fieldType)) continue;

                string[]? choices = null;
                if (field.TryGetProperty("Choices", out var choicesProp) &&
                    choicesProp.ValueKind == JsonValueKind.Array)
                {
                    choices = choicesProp.EnumerateArray()
                        .Select(c => c.GetString() ?? "")
                        .Where(s => s.Length > 0)
                        .ToArray();
                }
                // Also handle OData verbose choice format: { "results": [...] }
                if (choices == null &&
                    field.TryGetProperty("Choices", out var choicesVerbose) &&
                    choicesVerbose.ValueKind == JsonValueKind.Object &&
                    choicesVerbose.TryGetProperty("results", out var results))
                {
                    choices = results.EnumerateArray()
                        .Select(c => c.GetString() ?? "")
                        .Where(s => s.Length > 0)
                        .ToArray();
                }

                result.Add(new SpColumnDef
                {
                    InternalName = internalName,
                    DisplayName  = field.GetProperty("Title").GetString() ?? internalName,
                    FieldType    = fieldType,
                    Choices      = choices,
                });
            }
        }

        _columnCache[listId] = result;
        return result;
    }

    // Bulk-reads custom field values for all items in a library in one paginated pass.
    // Returns a dictionary keyed by SP list item integer ID (as string, e.g. "42").
    // This key matches the listItemId returned by GetSharePointIdsAsync, enabling O(1) cache lookup per file.
    public async Task<Dictionary<string, Dictionary<string, object?>>> BulkReadCustomFieldsAsync(
        string siteUrl, string listId, IEnumerable<SpColumnDef> columns,
        IProgress<int>? progress = null, CancellationToken ct = default)
    {
        var cols = columns.ToList();
        if (cols.Count == 0) return [];

        // Chunk columns into groups of 20 to stay within URL length limits.
        var result = new Dictionary<string, Dictionary<string, object?>>();
        var chunks  = cols.Chunk(20).ToList();

        for (int chunk = 0; chunk < chunks.Count; chunk++)
        {
            var colNames = string.Join(",", new[] { "ID" }.Concat(chunks[chunk].Select(c => c.InternalName)));
            var nextUrl  = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items?$select={Uri.EscapeDataString(colNames)}&$top=1000";

            while (nextUrl != null)
            {
                ct.ThrowIfCancellationRequested();
                using var response = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, nextUrl);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl, cancellationToken: ct);

                var body = await response.Content.ReadAsStringAsync(ct);
                if (!response.IsSuccessStatusCode) break;

                using var doc = JsonDocument.Parse(body);
                var root = doc.RootElement;

                if (root.TryGetProperty("value", out var items))
                {
                    foreach (var item in items.EnumerateArray())
                    {
                        if (!item.TryGetProperty("ID", out var idProp)) continue;
                        var itemId = idProp.ValueKind == JsonValueKind.Number
                            ? idProp.GetInt32().ToString()
                            : idProp.GetString() ?? "";
                        if (string.IsNullOrEmpty(itemId)) continue;

                        if (!result.TryGetValue(itemId, out var fields))
                        {
                            fields = new Dictionary<string, object?>();
                            result[itemId] = fields;
                        }

                        foreach (var col in chunks[chunk])
                        {
                            if (!item.TryGetProperty(col.InternalName, out var valProp)) continue;
                            fields[col.InternalName] = ParseFieldValue(valProp, col.FieldType);
                        }
                    }
                }

                nextUrl = GetNextLink(root);
            }

            progress?.Report((chunk + 1) * 100 / chunks.Count);
        }

        return result;
    }

    private static object? ParseFieldValue(JsonElement el, SupportedFieldType type)
    {
        if (el.ValueKind == JsonValueKind.Null || el.ValueKind == JsonValueKind.Undefined)
            return null;

        return type switch
        {
            SupportedFieldType.MultiChoice when el.ValueKind == JsonValueKind.Array =>
                string.Join(";#", el.EnumerateArray().Select(v => v.GetString() ?? "")),
            SupportedFieldType.MultiChoice when el.ValueKind == JsonValueKind.Object &&
                el.TryGetProperty("results", out var r) =>
                string.Join(";#", r.EnumerateArray().Select(v => v.GetString() ?? "")),
            SupportedFieldType.Boolean =>
                el.ValueKind == JsonValueKind.True ? "1" : "0",
            SupportedFieldType.DateTime when el.ValueKind == JsonValueKind.String =>
                el.GetString(),
            _ => el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString()
        };
    }

    // Applies custom field values to a list item via ValidateUpdateListItem.
    // mappings translates source InternalName → target InternalName.
    public async Task<string?> ApplyFileCustomFieldsAsync(
        string driveId, string itemId,
        Dictionary<string, object?> fields,
        IEnumerable<SpColumnMap> mappings)
    {
        if (fields.Count == 0) return null;

        // Build mapping lookup: source internal name → target internal name
        var mappingLookup = mappings.ToDictionary(
            m => m.SourceColumn.InternalName,
            m => m.TargetColumn?.InternalName);

        var formValues = new List<object>();
        foreach (var (srcName, value) in fields)
        {
            if (value == null) continue;
            // Resolve target name via mapping, or use source name directly
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

            var formatted = FormatFieldValueForValidate(value);
            formValues.Add(new { FieldName = targetName, FieldValue = formatted });
        }

        if (formValues.Count == 0) return null;

        var ids = await GetSharePointIdsAsync(driveId, itemId)
            ?? throw new Exception($"Could not resolve SharePoint IDs for {driveId}/{itemId}");

        var url = $"{ids.siteUrl.TrimEnd('/')}/_api/web/lists('{ids.listId}')/items({ids.listItemId})/ValidateUpdateListItem()";
        var payload = JsonSerializer.Serialize(new
        {
            formValues,
            bNewDocumentUpdate = true,
        });

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, ids.siteUrl);

        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            return $"ApplyCustomFields HTTP {(int)response.StatusCode}";

        // Parse per-field errors
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var vals))
            {
                var errors = vals.EnumerateArray()
                    .Where(v => v.TryGetProperty("HasException", out var ex) && ex.GetBoolean())
                    .Select(v => v.TryGetProperty("FieldName", out var fn) ? fn.GetString() : "?")
                    .ToList();
                if (errors.Count > 0)
                    return $"Custom field errors: {string.Join(", ", errors)}";
            }
        }
        catch { /* ignore parse errors */ }

        return null;
    }

    private static string FormatFieldValueForValidate(object value)
    {
        var str = value.ToString() ?? "";
        // MultiChoice: ensure ";#val1;#val2;#" format
        if (str.Contains(";#") && !str.StartsWith(";#"))
            str = ";#" + str + ";#";
        return str;
    }

    // ── System libraries ──────────────────────────────────────────────────────

    // Returns a SharePointNode for a document library identified by its display title,
    // or null if the library does not exist on the site.
    // Useful for hidden/system libraries (Site Assets, Style Library) that the Graph
    // Drives API does not always enumerate.
    public async Task<SharePointNode?> GetLibraryNodeByTitleAsync(string siteId, string siteUrl, string title)
    {
        try
        {
            var encoded  = Uri.EscapeDataString(title.Replace("'", "''"));
            var endpoint = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{encoded}')?$select=Id";

            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, endpoint);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) return null;

            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("Id", out var idProp)) return null;
            var listGuid = idProp.GetString();
            if (string.IsNullOrEmpty(listGuid)) return null;

            var drive = await Graph.Sites[siteId].Lists[listGuid].Drive.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "webUrl"]);

            var driveId = drive?.Id;
            if (string.IsNullOrEmpty(driveId)) return null;

            string? serverRelUrl = null;
            if (drive?.WebUrl != null)
            {
                var uri  = new Uri(drive.WebUrl);
                var path = Uri.UnescapeDataString(uri.AbsolutePath.TrimEnd('/'));
                if (!path.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
                    serverRelUrl = path;
            }

            var node = new SharePointNode
            {
                Id                 = "root",
                DriveId            = driveId,
                SiteId             = siteId,
                SiteUrl            = siteUrl,
                Name               = title,
                Type               = NodeType.Library,
                HasChildren        = true,
                ServerRelativePath = serverRelUrl,
            };
            node.Children.Add(Placeholder());
            return node;
        }
        catch
        {
            return null;
        }
    }

    // ── Pages ─────────────────────────────────────────────────────────────────

    // Returns the Site Pages library node, or null if not found.
    // Site Pages uses BaseTemplate=119 and is NOT returned by the Drives API on most tenants.
    // Strategy: fast-path via Drives, then reliable fallback via SharePoint REST + Graph list drive.
    public async Task<SharePointNode?> GetSitePagesLibraryAsync(string siteId, string siteUrl)
    {
        // Fast path: some tenants expose Site Pages as a Graph drive
        var libraries = await GetLibrariesAsync(siteId, siteUrl);
        var byDrive = libraries.FirstOrDefault(l =>
            l.Name.Equals("Site Pages", StringComparison.OrdinalIgnoreCase));
        if (byDrive != null) return byDrive;

        // Reliable fallback: use SharePoint REST to get the list GUID, then Graph to get its drive.
        // This is more reliable than expanding 'drive' on the Lists collection response,
        // which may not populate in some Graph SDK versions.
        try
        {
            // Step 1: get the Site Pages list GUID via REST
            var listGuid = await GetSitePagesListIdViaRestAsync(siteUrl);
            if (string.IsNullOrEmpty(listGuid)) return null;

            // Step 2: get the drive for this specific list via Graph
            var drive = await Graph.Sites[siteId].Lists[listGuid].Drive.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "webUrl"]);

            var driveId = drive?.Id;
            if (string.IsNullOrEmpty(driveId)) return null;

            // Step 3: get the root item so we have a Graph item ID for tree expansion
            var root = await Graph.Drives[driveId].Root.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "webUrl"]);

            if (root?.Id == null) return null;

            string? serverRelUrl = null;
            if (root.WebUrl != null)
                serverRelUrl = Uri.UnescapeDataString(new Uri(root.WebUrl).AbsolutePath.TrimEnd('/'));

            var node = new SharePointNode
            {
                Id                 = root.Id,
                DriveId            = driveId,
                SiteId             = siteId,
                SiteUrl            = siteUrl,
                Name               = "Site Pages",
                Type               = NodeType.Library,
                HasChildren        = true,
                ServerRelativePath = serverRelUrl,
            };
            node.Children.Add(Placeholder());
            return node;
        }
        catch
        {
            return null;
        }
    }

    // Uses SharePoint REST to find the Site Pages list GUID — more reliable than Graph Lists API
    // because it matches by BaseTemplate=119 regardless of display name locale.
    private async Task<string?> GetSitePagesListIdViaRestAsync(string siteUrl)
    {
        // Try by BaseTemplate filter first (locale-independent)
        var filterUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists?$filter=BaseTemplate eq 119&$select=Id";
        try
        {
            using var filterResp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, filterUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (filterResp.IsSuccessStatusCode)
            {
                var body = await filterResp.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(body);
                var val = doc.RootElement.TryGetProperty("value", out var v) ? v : doc.RootElement;
                var first = val.EnumerateArray().FirstOrDefault();
                if (first.TryGetProperty("Id", out var id))
                    return id.GetString();
            }
        }
        catch { /* fall through to title-based lookup */ }

        // Fallback: look up by title "Site Pages" (English)
        var titleUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('Site Pages')?$select=Id";
        try
        {
            using var titleResp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, titleUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (titleResp.IsSuccessStatusCode)
            {
                var body = await titleResp.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("Id", out var id))
                    return id.GetString();
            }
        }
        catch { }

        return null;
    }

    // Reads page metadata from a SharePoint modern page using a two-step SitePages API lookup.
    // Step 1: find the page's integer ID by filename ($filter avoids any path encoding issues).
    // Step 2: fetch the full content by that ID (/_api/sitepages/pages/{id}).
    // Returns (metadata, null) on success, (null, errorMessage) on failure.
    public async Task<(PageMetadata? meta, string? error)> GetPageMetadataAsync(
        string siteUrl, string pageServerRelativeUrl)
    {
        var fileName    = System.IO.Path.GetFileName(pageServerRelativeUrl.Replace('\\', '/'));
        var escapedName = fileName.Replace("'", "''");

        // Step 1: find the SitePages integer ID by filename
        int pageId = 0;
        var listUrl = $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages" +
                      $"?$filter=FileName eq '{escapedName}'&$select=Id";
        System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] step1 GET {listUrl}");
        try
        {
            using var listResp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, listUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!listResp.IsSuccessStatusCode)
            {
                var err = await listResp.Content.ReadAsStringAsync();
                var msg = $"SitePages list HTTP {(int)listResp.StatusCode}: {err[..Math.Min(200, err.Length)]}";
                System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] {msg}");
                return (null, msg);
            }

            var body = await listResp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var arr))
            {
                foreach (var item in arr.EnumerateArray())
                {
                    if (item.TryGetProperty("Id", out var idProp))
                        pageId = idProp.GetInt32();
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] step1 exception: {ex.Message}");
            return (null, $"Page lookup exception: {ex.Message}");
        }

        if (pageId == 0)
        {
            System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] page '{fileName}' not found in SitePages list");
            return (null, $"Page '{fileName}' not found in source Site Pages library");
        }

        // Step 2: fetch full content by integer ID
        var pageUrl = $"{siteUrl.TrimEnd('/')}/_api/sitepages/pages({pageId})" +
                      "?$select=CanvasContent1,LayoutWebpartsContent,BannerImageUrl,Description,PageLayoutType,PromotedState";
        System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] step2 GET {pageUrl}");
        try
        {
            using var pageResp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, pageUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!pageResp.IsSuccessStatusCode)
            {
                var err = await pageResp.Content.ReadAsStringAsync();
                var msg = $"SitePages content HTTP {(int)pageResp.StatusCode}: {err[..Math.Min(200, err.Length)]}";
                System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] {msg}");
                return (null, msg);
            }

            var body = await pageResp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;

            var meta = new PageMetadata
            {
                CanvasContent1        = root.TryGetProperty("CanvasContent1",        out var c1) ? c1.GetString() : null,
                LayoutWebpartsContent = root.TryGetProperty("LayoutWebpartsContent", out var lw) ? lw.GetString() : null,
                BannerImageUrl        = root.TryGetProperty("BannerImageUrl",        out var bi) ? bi.GetString() : null,
                Description           = root.TryGetProperty("Description",           out var d)  ? d.GetString()  : null,
                PageLayoutType        = root.TryGetProperty("PageLayoutType",        out var pl) ? pl.GetString() : null,
                PromotedState         = root.TryGetProperty("PromotedState",         out var ps) && ps.TryGetInt32(out var psv) ? psv : 0,
            };
            System.Diagnostics.Debug.WriteLine(
                $"[GetPageMetadata] OK — CanvasContent1={meta.CanvasContent1?.Length ?? 0} chars");
            return (meta, null);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[GetPageMetadata] step2 exception: {ex.Message}");
            return (null, $"Page content fetch exception: {ex.Message}");
        }
    }

    // Applies page metadata to a copied page, substituting source URLs with target URLs.
    // Returns a warning string if web parts referencing source-specific list IDs are detected.
    public async Task<string?> ApplyPageMetadataAsync(
        string targetSiteUrl, string targetDriveId, string targetItemId,
        PageMetadata metadata, string sourceSiteUrl)
    {
        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] START targetItemId={targetItemId}");
        var formValues = new List<object>();

        string? SubstUrl(string? json) =>
            json == null ? null : SubstituteUrls(json, sourceSiteUrl, targetSiteUrl);

        var canvas  = SubstUrl(metadata.CanvasContent1);
        var layout  = SubstUrl(metadata.LayoutWebpartsContent);
        var banner  = SubstUrl(metadata.BannerImageUrl);

        if (canvas  != null) formValues.Add(new { FieldName = "CanvasContent1",        FieldValue = canvas });
        if (layout  != null) formValues.Add(new { FieldName = "LayoutWebpartsContent", FieldValue = layout });
        if (banner  != null) formValues.Add(new { FieldName = "BannerImageUrl",        FieldValue = banner });
        if (metadata.Description   != null) formValues.Add(new { FieldName = "Description",   FieldValue = metadata.Description });
        if (metadata.PageLayoutType != null) formValues.Add(new { FieldName = "PageLayoutType", FieldValue = metadata.PageLayoutType });
        formValues.Add(new { FieldName = "PromotedState", FieldValue = metadata.PromotedState.ToString() });

        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] {formValues.Count} fields to write, resolving SP IDs…");
        if (formValues.Count == 0) return null;

        var ids = await GetSharePointIdsAsync(targetDriveId, targetItemId)
            ?? throw new Exception($"Could not resolve SharePoint IDs for {targetDriveId}/{targetItemId}");
        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] SP IDs resolved, posting ValidateUpdateListItem…");

        var url = $"{ids.siteUrl.TrimEnd('/')}/_api/web/lists('{ids.listId}')/items({ids.listItemId})/ValidateUpdateListItem()";
        var payload = JsonSerializer.Serialize(new { formValues, bNewDocumentUpdate = true });

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, ids.siteUrl);

        if (!response.IsSuccessStatusCode)
        {
            var errBody = await response.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] ValidateUpdateListItem HTTP {(int)response.StatusCode}: {errBody[..Math.Min(errBody.Length, 300)]}");
            return $"ApplyPageMetadata HTTP {(int)response.StatusCode}";
        }

        // Detect web parts referencing source-specific list GUIDs
        bool hasListIdRefs = canvas != null &&
            System.Text.RegularExpressions.Regex.IsMatch(canvas,
                @"""[Ll]ist[Ii]d""\s*:\s*""[{]?[0-9a-fA-F\-]{36}[}]?""");

        var result = hasListIdRefs
            ? "Some web parts reference list IDs from the source site and may need manual review."
            : null;
        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] DONE, result={(result ?? "OK")}");
        return result;
    }

    internal static string SubstituteUrls(string json, string sourceUrl, string targetUrl)
    {
        var src     = sourceUrl.TrimEnd('/');
        var tgt     = targetUrl.TrimEnd('/');
        var srcEnc  = Uri.EscapeDataString(src);
        var tgtEnc  = Uri.EscapeDataString(tgt);
        var srcEnc2 = src.Replace(":", "%3A").Replace("/", "%2F");
        var tgtEnc2 = tgt.Replace(":", "%3A").Replace("/", "%2F");

        return json
            .Replace(srcEnc2, tgtEnc2, StringComparison.OrdinalIgnoreCase)
            .Replace(srcEnc,  tgtEnc,  StringComparison.OrdinalIgnoreCase)
            .Replace(src,     tgt,     StringComparison.OrdinalIgnoreCase);
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    public record NavigationNode(int Id, string Title, string Url, bool IsExternal, List<NavigationNode> Children);

    // Reads Quick Launch (quickLaunch=true) or Top Navigation Bar (quickLaunch=false) nodes.
    public async Task<List<NavigationNode>> GetNavigationNodesAsync(string siteUrl, bool quickLaunch)
    {
        var section  = quickLaunch ? "quicklaunch" : "topnavigationbar";
        var endpoint = $"{siteUrl.TrimEnd('/')}/_api/web/navigation/{section}?$expand=Children&$select=Id,Title,Url,IsExternal";

        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, endpoint);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("value", out var arr)
                ? ParseNavigationNodes(arr)
                : [];
        }
        catch { return []; }
    }

    // Copies Quick Launch and/or Top Navigation from source to target, remapping site URLs.
    public async Task CopyNavigationAsync(string sourceSiteUrl, string targetSiteUrl, bool quickLaunch)
    {
        var nodes   = await GetNavigationNodesAsync(sourceSiteUrl, quickLaunch);
        var section = quickLaunch ? "quicklaunch" : "topnavigationbar";

        // Clear existing nodes at target first
        var existing = await GetNavigationNodesAsync(targetSiteUrl, quickLaunch);
        foreach (var node in existing)
            await DeleteNavigationNodeAsync(targetSiteUrl, section, node.Id);

        // Recreate from source with URL substitution
        foreach (var node in nodes)
        {
            var mappedUrl = SubstituteUrls(node.Url, sourceSiteUrl, targetSiteUrl);
            var newId     = await CreateNavigationNodeAsync(targetSiteUrl, section, null, node.Title, mappedUrl, node.IsExternal);

            foreach (var child in node.Children)
            {
                var childUrl = SubstituteUrls(child.Url, sourceSiteUrl, targetSiteUrl);
                await CreateNavigationNodeAsync(targetSiteUrl, section, newId, child.Title, childUrl, child.IsExternal);
            }
        }
    }

    private static List<NavigationNode> ParseNavigationNodes(JsonElement array)
    {
        var nodes = new List<NavigationNode>();
        foreach (var el in array.EnumerateArray())
        {
            var id    = el.TryGetProperty("Id",         out var ip)  ? ip.GetInt32()    : 0;
            var title = el.TryGetProperty("Title",      out var tp)  ? tp.GetString()   ?? "" : "";
            var url   = el.TryGetProperty("Url",        out var up)  ? up.GetString()   ?? "" : "";
            var ext   = el.TryGetProperty("IsExternal", out var ep)  && ep.GetBoolean();

            var children = new List<NavigationNode>();
            if (el.TryGetProperty("Children", out var cp))
            {
                if (cp.ValueKind == JsonValueKind.Array)
                    children = ParseNavigationNodes(cp);
                else if (cp.TryGetProperty("value", out var cv))
                    children = ParseNavigationNodes(cv);
            }
            nodes.Add(new NavigationNode(id, title, url, ext, children));
        }
        return nodes;
    }

    private async Task DeleteNavigationNodeAsync(string siteUrl, string section, int nodeId)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/navigation/{section}/GetById({nodeId})";
        try
        {
            using var _ = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Delete, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Headers.Add("IF-MATCH", "*");
                return r;
            }, siteUrl);
        }
        catch { }
    }

    private async Task<int> CreateNavigationNodeAsync(
        string siteUrl, string section, int? parentId,
        string title, string url, bool isExternal)
    {
        var endpoint = parentId.HasValue
            ? $"{siteUrl.TrimEnd('/')}/_api/web/navigation/{section}/GetById({parentId.Value})/Children"
            : $"{siteUrl.TrimEnd('/')}/_api/web/navigation/{section}";

        var body = JsonSerializer.Serialize(new { Title = title, Url = url, IsExternal = isExternal });

        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, endpoint);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) return 0;
            var respBody = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(respBody);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetInt32() : 0;
        }
        catch { return 0; }
    }

    // ── Custom lists ──────────────────────────────────────────────────────────

    // System list templates to exclude from custom list enumeration.
    private static readonly HashSet<int> _systemListBaseTemplates = [101, 119, 171, 544, 890];

    // Returns non-hidden, non-catalog, non-library lists (Tasks, Calendars, Announcements, etc.).
    public async Task<List<(string Id, string Title, int BaseTemplate)>> GetCustomListsAsync(string siteUrl)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists" +
                  "?$filter=Hidden eq false and IsCatalog eq false" +
                  " and BaseTemplate ne 101 and BaseTemplate ne 119" +
                  "&$select=Id,Title,BaseTemplate&$orderby=Title";

        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);

            var result = new List<(string, string, int)>();
            if (!doc.RootElement.TryGetProperty("value", out var values)) return result;

            foreach (var el in values.EnumerateArray())
            {
                var id       = el.TryGetProperty("Id",           out var ip) ? ip.GetString()  ?? "" : "";
                var title    = el.TryGetProperty("Title",        out var tp) ? tp.GetString()  ?? "" : "";
                var baseTmpl = el.TryGetProperty("BaseTemplate", out var bp) ? bp.GetInt32()        : 0;
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(title)) continue;
                if (_systemListBaseTemplates.Contains(baseTmpl)) continue;
                result.Add((id, title, baseTmpl));
            }
            return result;
        }
        catch { return []; }
    }

    // Returns the list GUID for a list looked up by title, or null if not found.
    public async Task<string?> GetListIdByTitleAsync(string siteUrl, string title)
    {
        var encoded = Uri.EscapeDataString(title.Replace("'", "''"));
        var url     = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{encoded}')?$select=Id";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);
            if (!resp.IsSuccessStatusCode) return null;
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetString() : null;
        }
        catch { return null; }
    }

    // Reads all items from a list with the given custom field names plus Created/Modified.
    // Returns each item as a string-keyed dictionary. Handles pagination automatically.
    public async Task<List<Dictionary<string, object?>>> GetListItemsAsync(
        string siteUrl, string listId,
        IEnumerable<string> customFieldNames,
        CancellationToken ct = default)
    {
        var fields  = customFieldNames.ToList();
        var select  = "Id,Title,Created,Modified," + (fields.Count > 0 ? string.Join(",", fields) : "Title");
        var baseUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                      $"?$select={Uri.EscapeDataString(select)}&$top=5000";

        var result = new List<Dictionary<string, object?>>();
        string? next = baseUrl;

        while (next != null)
        {
            ct.ThrowIfCancellationRequested();
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, next);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) break;
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);

            if (doc.RootElement.TryGetProperty("value", out var values))
                foreach (var el in values.EnumerateArray())
                {
                    var item = new Dictionary<string, object?>();
                    foreach (var prop in el.EnumerateObject())
                        item[prop.Name] = ExtractJsonValue(prop.Value);
                    result.Add(item);
                }

            next = GetNextLink(doc.RootElement);
        }
        return result;
    }

    // Returns (Id, Title) for each item in a list — lightweight query for tree display only.
    public async Task<List<(string Id, string Title)>> GetListItemTitlesAsync(
        string siteUrl, string listId, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                  "?$select=Id,Title&$top=5000&$orderby=Id";

        var result = new List<(string, string)>();
        string? next = url;

        while (next != null)
        {
            ct.ThrowIfCancellationRequested();
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, next);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);

            if (!resp.IsSuccessStatusCode) break;
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);

            if (doc.RootElement.TryGetProperty("value", out var values))
                foreach (var el in values.EnumerateArray())
                {
                    var id    = el.TryGetProperty("Id",    out var ip) ? ip.GetInt32().ToString() : string.Empty;
                    var title = el.TryGetProperty("Title", out var tp) && tp.ValueKind == JsonValueKind.String
                                    ? tp.GetString() ?? string.Empty
                                    : string.Empty;
                    if (!string.IsNullOrEmpty(id))
                        result.Add((id, title));
                }

            next = GetNextLink(doc.RootElement);
        }
        return result;
    }

    // Creates a list item and optionally back-fills Created/Modified timestamps.
    // Returns the new item's integer ID as a string, or null if it could not be parsed.
    public async Task<string?> CreateListItemAsync(
        string siteUrl, string listId,
        Dictionary<string, object?> fields,
        string? createdDate, string? modifiedDate,
        CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();

        var createUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items";
        var body      = JsonSerializer.Serialize(fields.Where(kv => kv.Value != null)
                            .ToDictionary(kv => kv.Key, kv => kv.Value));

        string? newItemId = null;
        using (var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, createUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl))
        {
            if (!resp.IsSuccessStatusCode)
            {
                var errBody = await resp.Content.ReadAsStringAsync();
                throw new HttpRequestException($"Create item failed: {(int)resp.StatusCode} {resp.ReasonPhrase} — {errBody}");
            }
            var respBody = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(respBody);
            newItemId = doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetRawText() : null;
        }

        if (newItemId == null || (createdDate == null && modifiedDate == null)) return newItemId;

        var metaValues = new List<object>();
        if (createdDate  != null) metaValues.Add(new { FieldName = "Created",  FieldValue = createdDate });
        if (modifiedDate != null) metaValues.Add(new { FieldName = "Modified", FieldValue = modifiedDate });

        var metaPayload = JsonSerializer.Serialize(new { formValues = metaValues, bNewDocumentUpdate = true });
        var metaUrl     = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({newItemId})/ValidateUpdateListItem()";

        try
        {
            using var _ = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, metaUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent(metaPayload, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
        }
        catch { }
        return newItemId;
    }

    // Updates an existing list item's fields via MERGE/PATCH and optionally back-fills Created/Modified.
    public async Task UpdateListItemAsync(
        string siteUrl, string listId, string itemId,
        Dictionary<string, object?> fields,
        string? createdDate, string? modifiedDate,
        CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();

        var updateUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({itemId})";
        var body      = JsonSerializer.Serialize(fields.Where(kv => kv.Value != null)
                            .ToDictionary(kv => kv.Key, kv => kv.Value));

        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, updateUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Headers.Add("X-HTTP-Method", "MERGE");
            r.Headers.Add("If-Match", "*");
            r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl);

        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync();
            throw new HttpRequestException($"Update item failed: {(int)resp.StatusCode} {resp.ReasonPhrase} — {errBody}");
        }

        if (createdDate == null && modifiedDate == null) return;

        var metaValues = new List<object>();
        if (createdDate  != null) metaValues.Add(new { FieldName = "Created",  FieldValue = createdDate });
        if (modifiedDate != null) metaValues.Add(new { FieldName = "Modified", FieldValue = modifiedDate });

        var metaPayload = JsonSerializer.Serialize(new { formValues = metaValues, bNewDocumentUpdate = true });
        var metaUrl     = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({itemId})/ValidateUpdateListItem()";

        try
        {
            using var _ = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, metaUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent(metaPayload, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
        }
        catch { }
    }

    private static object? ExtractJsonValue(JsonElement el) => el.ValueKind switch
    {
        JsonValueKind.String  => el.GetString(),
        JsonValueKind.Number  => el.TryGetInt64(out var l) ? (object?)l : el.GetDouble(),
        JsonValueKind.True    => true,
        JsonValueKind.False   => false,
        JsonValueKind.Null    => null,
        _                     => null,
    };

    // Reads ID + HasUniqueRoleAssignments for all items in a library/list in one paginated request.
    // Returns a dictionary keyed by "{listId}:{itemId}" (see PermissionFlagKey), value =
    // hasUniquePermissions. The composite key lets flags from multiple source libraries
    // be merged into one dictionary without item-ID collisions.
    public async Task<Dictionary<string, bool>> BulkReadPermissionFlagsAsync(
        string siteUrl, string listId, CancellationToken ct = default)
    {
        var result  = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        string? next = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                       "?$select=Id,HasUniqueRoleAssignments&$top=5000";
        while (next != null)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var resp = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, next);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl, cancellationToken: ct);
                if (!resp.IsSuccessStatusCode) break;
                var body = await resp.Content.ReadAsStringAsync(ct);
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("value", out var vals))
                    foreach (var el in vals.EnumerateArray())
                    {
                        var id      = el.TryGetProperty("Id",                    out var ip) ? ip.GetInt32().ToString() : null;
                        var unique  = el.TryGetProperty("HasUniqueRoleAssignments", out var up) && up.GetBoolean();
                        if (id != null) result[PermissionFlagKey(listId, id)] = unique;
                    }
                next = GetNextLink(doc.RootElement);
            }
            catch { break; }
        }
        return result;
    }

    // Composite key for permission-flag dictionaries. GUID braces are stripped so keys
    // match whether the list ID came from SP REST or Graph sharepointIds.
    public static string PermissionFlagKey(string listId, string itemId) =>
        $"{listId.Trim('{', '}')}:{itemId}";

    // ── Permissions ───────────────────────────────────────────────────────────

    // Checks whether a single SP object has unique (broken) permissions.
    // apiPath examples: "web/lists('{guid}')", "web"
    public async Task<bool> GetHasUniqueRoleAssignmentsAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}?$select=HasUniqueRoleAssignments";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return false;
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("HasUniqueRoleAssignments", out var v) && v.GetBoolean();
        }
        catch { return false; }
    }

    public async Task<List<RoleAssignmentInfo>> GetRoleAssignmentsAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/roleassignments" +
                  "?$expand=Member,RoleDefinitionBindings&$select=Member/Id,Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("value", out var vals)) return [];

            var result = new List<RoleAssignmentInfo>();
            foreach (var el in vals.EnumerateArray())
            {
                if (!el.TryGetProperty("Member", out var member)) continue;
                var pid        = member.TryGetProperty("Id",            out var idEl)   ? idEl.GetInt32()          : 0;
                var loginName  = member.TryGetProperty("LoginName",     out var lnEl)   ? lnEl.GetString() ?? ""   : "";
                var title      = member.TryGetProperty("Title",         out var tEl)    ? tEl.GetString()  ?? ""   : "";
                var ptype      = member.TryGetProperty("PrincipalType", out var ptEl)   ? ptEl.GetInt32()          : 0;
                var roleNames  = new List<string>();
                if (el.TryGetProperty("RoleDefinitionBindings", out var rdbs))
                    foreach (var rdb in rdbs.EnumerateArray())
                        if (rdb.TryGetProperty("Name", out var nm))
                            roleNames.Add(nm.GetString() ?? "");
                result.Add(new RoleAssignmentInfo(pid, ptype, loginName, title, roleNames));
            }
            return result;
        }
        catch { return []; }
    }

    public async Task BreakPermissionInheritanceAsync(
        string siteUrl, string apiPath, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)";
        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(string.Empty, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct);
        if (!resp.IsSuccessStatusCode)
        {
            var body = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException($"BreakPermissionInheritance failed: {(int)resp.StatusCode} — {body}");
        }
    }

    public async Task AddRoleAssignmentAsync(
        string siteUrl, string apiPath, int principalId, int roleDefId, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/{apiPath}/roleassignments/addroleassignment(principalid={principalId},roleDefId={roleDefId})";
        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(string.Empty, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct);
        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException($"AddRoleAssignment failed: {(int)resp.StatusCode} — {errBody}");
        }
    }

    public async Task<int?> EnsureUserAsync(
        string siteUrl, string loginName, CancellationToken ct = default)
    {
        var url  = $"{siteUrl.TrimEnd('/')}/_api/web/ensureuser";
        var body = JsonSerializer.Serialize(new { logonName = loginName });
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return null;
            var respBody = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(respBody);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetInt32() : null;
        }
        catch { return null; }
    }

    public async Task<int?> GetSiteGroupIdAsync(
        string siteUrl, string groupTitle, CancellationToken ct = default)
    {
        var encoded = Uri.EscapeDataString(groupTitle.Replace("'", "''"));
        var url     = $"{siteUrl.TrimEnd('/')}/_api/web/sitegroups/getbyname('{encoded}')?$select=Id";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return null;
            var respBody = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(respBody);
            return doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetInt32() : null;
        }
        catch { return null; }
    }

    // Loads all role definitions for the site and returns a name→ID dictionary.
    public async Task<Dictionary<string, int>> GetAllRoleDefinitionsAsync(
        string siteUrl, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/roledefinitions?$select=Id,Name";
        try
        {
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode) return [];
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("value", out var vals)) return [];
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var el in vals.EnumerateArray())
            {
                var id   = el.TryGetProperty("Id",   out var idEl)   ? idEl.GetInt32()        : 0;
                var name = el.TryGetProperty("Name", out var nameEl) ? nameEl.GetString() ?? "" : "";
                if (id > 0 && !string.IsNullOrEmpty(name))
                    result[name] = id;
            }
            return result;
        }
        catch { return []; }
    }

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
