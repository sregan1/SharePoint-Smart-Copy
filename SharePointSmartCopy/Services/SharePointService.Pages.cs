using SharePointSmartCopy.Localization;
using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
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

// Partial class SharePointService: Pages. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
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
            _router.RegisterDriveForSite(driveId, siteUrl);

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
        // The OData quote-doubling alone isn't enough — the result still gets interpolated
        // directly into the query string. A page named "Q3 #Review.aspx" truncated the request
        // at the '#' (URL fragment), and '&'/'+' broke it similarly, misdiagnosing a real page as
        // "not found in source Site Pages library". Uri.EscapeDataString is the fix every other
        // REST filter/path builder in this file already applies (see ResolveLookupValueAsync).
        var escapedName = Uri.EscapeDataString(fileName.Replace("'", "''"));

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
                    {
                        pageId = idProp.GetInt32();
                        break;
                    }
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
            return (null, Loc.T("Svc_Sp_PageNotFoundIn", fileName));
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
            ?? throw new Exception(Loc.T("Svc_Sp_CouldNotResolveSharePoint", targetDriveId, targetItemId));
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
            ? Loc.T("Svc_Sp_SomeWebPartsReference")
            : null;
        System.Diagnostics.Debug.WriteLine($"[ApplyPageMetadata] DONE, result={(result ?? "OK")}");
        return result;
    }

    internal static string SubstituteUrls(string json, string sourceUrl, string targetUrl)
    {
        var src = sourceUrl.TrimEnd('/');
        var tgt = targetUrl.TrimEnd('/');

        // Boundary-checked replace: a raw substring replace corrupted sibling sites — with source
        // /sites/HR, a link to /sites/HRArchive became "<target>Archive". The character after the
        // match must not extend the same path segment.
        static string ReplaceBounded(string input, string find, string repl)
        {
            if (string.IsNullOrEmpty(find)) return input;
            var pattern = System.Text.RegularExpressions.Regex.Escape(find) + @"(?![A-Za-z0-9_\-])";
            return System.Text.RegularExpressions.Regex.Replace(
                input, pattern, repl.Replace("$", "$$"),
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        var result = json;
        result = ReplaceBounded(result, src.Replace(":", "%3A").Replace("/", "%2F"),
                                        tgt.Replace(":", "%3A").Replace("/", "%2F"));
        result = ReplaceBounded(result, Uri.EscapeDataString(src), Uri.EscapeDataString(tgt));
        result = ReplaceBounded(result, src, tgt);

        // Server-relative references too: modern-page CanvasContent1 stores image/web-part URLs
        // as "/sites/HR/SiteAssets/..." — the absolute-only replace left those pointing at the
        // source site even with URL remapping on.
        static string RelOf(string abs) =>
            Uri.TryCreate(abs, UriKind.Absolute, out var u) ? u.AbsolutePath.TrimEnd('/') : "";
        var srcRel = RelOf(src);
        var tgtRel = RelOf(tgt);
        if (srcRel.Length > 1 && !string.Equals(srcRel, tgtRel, StringComparison.OrdinalIgnoreCase))
        {
            result = ReplaceBounded(result, srcRel.Replace("/", "%2F"), tgtRel.Replace("/", "%2F"));
            result = ReplaceBounded(result, srcRel, tgtRel);
        }
        return result;
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    public record NavigationNode(int Id, string Title, string Url, bool IsExternal, List<NavigationNode> Children);

    // Reads Quick Launch (quickLaunch=true) or Top Navigation Bar (quickLaunch=false) nodes.
    // throwOnError: pass true when an empty result would be acted on destructively — the default
    // swallow-to-empty is only safe for display purposes.
    public async Task<List<NavigationNode>> GetNavigationNodesAsync(string siteUrl, bool quickLaunch, bool throwOnError = false)
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

            if (!resp.IsSuccessStatusCode)
            {
                if (throwOnError)
                    throw new HttpRequestException(Loc.T("Svc_Sp_NavigationReadFailedHTTP", (int)resp.StatusCode));
                return [];
            }
            var body = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("value", out var arr)
                ? ParseNavigationNodes(arr)
                : [];
        }
        catch when (!throwOnError) { return []; }
    }

    // Copies Quick Launch and/or Top Navigation from source to target, remapping site URLs.
    public async Task CopyNavigationAsync(string sourceSiteUrl, string targetSiteUrl, bool quickLaunch)
    {
        // Fail loudly if the SOURCE read fails: this method deletes the target's existing
        // navigation before recreating it, so a swallowed source failure (transient 503 after
        // retries) used to wipe the target nav, rebuild nothing, and report Success.
        var nodes   = await GetNavigationNodesAsync(sourceSiteUrl, quickLaunch, throwOnError: true);
        var section = quickLaunch ? "quicklaunch" : "topnavigationbar";

        // Clear existing nodes at target first
        var existing = await GetNavigationNodesAsync(targetSiteUrl, quickLaunch, throwOnError: true);
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

}
