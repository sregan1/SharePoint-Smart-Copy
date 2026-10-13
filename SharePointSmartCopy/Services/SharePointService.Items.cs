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

// Partial class SharePointService: Items. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
    // ── Upload ────────────────────────────────────────────────────────────────

    // Fetches all file children of a folder in one paginated pass.
    // Returns filename → (itemId, lastModifiedDateTime), case-insensitive.
    // Used by the migration pre-flight to avoid one Graph call per file.
    public async Task<Dictionary<string, (string ItemId, DateTimeOffset? Modified)>> FetchFolderItemsAsync(
        string driveId, string folderId)
    {
        // Retries the WHOLE paginated walk on a transient transport failure (see
        // IsTransientTransportError) rather than leaving a mid-pagination blip to propagate out of the
        // whole drive-group scan and fail every pending file. Re-walking from page 1 is cheap relative
        // to the cost of failing the entire run.
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                var result = new Dictionary<string, (string, DateTimeOffset?)>(StringComparer.OrdinalIgnoreCase);
                var page = await Graph.Drives[driveId].Items[folderId].Children
                    .GetAsync(cfg =>
                    {
                        cfg.QueryParameters.Top    = 1000;
                        cfg.QueryParameters.Select = ["id", "name", "lastModifiedDateTime", "file"];
                    });

                while (page != null)
                {
                    foreach (var item in page.Value ?? [])
                    {
                        if (item.File == null || item.Name == null || item.Id == null) continue;
                        result[item.Name] = (item.Id, item.LastModifiedDateTime);
                    }
                    if (page.OdataNextLink == null) break;
                    page = await Graph.Drives[driveId].Items[folderId].Children
                        .WithUrl(page.OdataNextLink).GetAsync();
                }
                return result;
            }
            catch (Exception ex) when (attempt < 4 && IsTransientTransportError(ex))
            {
                await Task.Delay(TimeSpan.FromSeconds((attempt + 1) * 3));
            }
        }
    }

    // Lists a folder's files via SP REST (GetFolderByServerRelativeUrl/Files) — the same AllDocs
    // store the Migration API validates imports against. Graph's Children listing does NOT return
    // rows left behind by a previous SPMI job that fatal-aborted mid-import, so an overwrite
    // pre-flight that trusts Graph alone reports "0 already exist" while SPMI then rejects those
    // same files with "already exists" (observed 2026-07-02: 100+ per batch). Returns name →
    // TimeLastModified; a missing folder (404) returns empty. Follows odata.nextLink paging.
    public async Task<Dictionary<string, DateTimeOffset?>> FetchFolderFileNamesRestAsync(
        string siteUrl, string folderServerRelativeUrl)
    {
        var result  = new Dictionary<string, DateTimeOffset?>(StringComparer.OrdinalIgnoreCase);
        var baseUrl = siteUrl.TrimEnd('/');
        // *Path (not *Url) API: folder names containing #, %, or + are unresolvable via
        // GetFolderByServerRelativeUrl, so its /Files listing 404s and the pre-flight wrongly
        // reports the folder empty. See ServerRelativePathArg.
        string? nextUrl = $"{baseUrl}/_api/web/GetFolderByServerRelativePath({ServerRelativePathArg(folderServerRelativeUrl)})/Files" +
                          "?$select=Name,TimeLastModified&$top=5000";

        for (int attempt = 0; nextUrl != null; )
        {
            try
            {
                var requestUrl = nextUrl;
                using var response = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl);

                if (!response.IsSuccessStatusCode)
                {
                    if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        // Folder doesn't exist yet (fresh target) — genuinely empty.
                        System.Diagnostics.Debug.WriteLine(
                            $"[RestFolderList] 404 (folder not found) for {folderServerRelativeUrl}");
                        return result;
                    }
                    // Any OTHER status used to be treated the same as 404 — "return what we have" —
                    // which silently truncated the snapshot on page 2+ of a large folder (12,000+
                    // files) under throttling. The caller (MigrationJobService's overwrite pre-flight)
                    // merges this result assuming it's complete: a truncated snapshot meant existing
                    // files past the failed page were never queued for deletion, and SPMI then
                    // rejected the whole re-import with "already exists" — precisely the regression
                    // this method exists to prevent. Throw instead so the failure is visible.
                    var errBody = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException(
                        Loc.T("Svc_Sp_ListingFolderFilesFailed", (int)response.StatusCode, folderServerRelativeUrl) + " " +
                        $"{errBody[..Math.Min(errBody.Length, 300)]}");
                }

                var body = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("value", out var items))
                {
                    foreach (var item in items.EnumerateArray())
                    {
                        if (!item.TryGetProperty("Name", out var n) || n.ValueKind != JsonValueKind.String) continue;
                        DateTimeOffset? modified = null;
                        if (item.TryGetProperty("TimeLastModified", out var tm) &&
                            tm.ValueKind == JsonValueKind.String &&
                            DateTimeOffset.TryParse(tm.GetString(), System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None, out var dt))
                            modified = dt;
                        result[n.GetString()!] = modified;
                    }
                }
                nextUrl = doc.RootElement.TryGetProperty("odata.nextLink", out var nl) &&
                          nl.ValueKind == JsonValueKind.String
                    ? nl.GetString()
                    : null;
                attempt = 0; // reset the retry budget per page
            }
            catch (Exception ex) when (attempt < 4 && IsTransientTransportError(ex))
            {
                attempt++;
                await Task.Delay(TimeSpan.FromSeconds(attempt * 3));
            }
        }
        return result;
    }

    // Cheap "is this folder completely empty?" check (files OR subfolders), via a single $top=1
    // request. Used to detect a fresh migration target so the per-folder existing-file scan can be
    // skipped entirely — on an empty target nothing can already exist.
    public async Task<bool> IsFolderEmptyAsync(string driveId, string folderId)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                var page = await Graph.Drives[driveId].Items[folderId].Children
                    .GetAsync(cfg =>
                    {
                        cfg.QueryParameters.Top    = 1;
                        cfg.QueryParameters.Select = ["id"];
                    });
                return (page?.Value?.Count ?? 0) == 0;
            }
            catch (Exception ex) when (attempt < 4 && IsTransientTransportError(ex))
            {
                await Task.Delay(TimeSpan.FromSeconds((attempt + 1) * 3));
            }
        }
    }

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

    // Returns the file's Graph item ID and last-modified timestamp, or null if it does
    // not exist. One call serves the existence check, the Copy-if-newer comparison, and
    // (via the ID) permission refresh on files that end up skipped as up to date.
    public async Task<(string ItemId, DateTimeOffset? Modified)?> GetFileInfoAsync(
        string driveId, string parentItemId, string fileName)
    {
        // Only a genuine 404 may map to null ("doesn't exist"): swallowing every failure here made
        // a transient error (post-retry 429, network blip) read as "not there", which in Skip mode
        // proceeded to upload — and the small-file PUT then silently clobbered the existing target.
        try
        {
            var item = await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName))
                .GetAsync(cfg => cfg.QueryParameters.Select = ["id", "lastModifiedDateTime"]);
            return item?.Id is { } id ? (id, item.LastModifiedDateTime) : null;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            return null;
        }
    }

    // Builds the OData "...ByServerRelativePath(decodedurl='…')" argument for a REST call.
    // The *Path (not *Url) API family is REQUIRED for any path containing #, %, or + — the
    // GetFileByServerRelativeUrl/GetFolderByServerRelativeUrl variants cannot resolve those even
    // when percent-encoded, so listings 404 (→ pre-flight misses the file) and deletes silently
    // fail (→ the file survives into the SPMI import as an "already exists" conflict). Observed
    // 2026-07-02 on folder and file names containing '+' and '#'.
    // decodedurl takes the LITERAL path: OData-escape embedded quotes (' → ''), then percent-encode
    // the whole thing for transport; SharePoint decodes it back to the literal server-relative path.
    private static string ServerRelativePathArg(string serverRelativeUrl) =>
        $"decodedurl='{Uri.EscapeDataString(serverRelativeUrl.Replace("'", "''"))}'";

    // Returns the SharePoint UniqueId (AllDocs GUID) for a file by its server-relative URL.
    // Works via REST (not Graph) so it finds zombie files — SPFile blobs without a list item
    // that Graph returns 404 for. Returns null if the file doesn't exist.
    public async Task<string?> GetFileUniqueIdAsync(string siteUrl, string serverRelativeUrl)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetFileByServerRelativePath({ServerRelativePathArg(serverRelativeUrl)})/UniqueId";
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
    //
    // Retries transient failures at each step and returns false (rather than silently swallowing
    // the error) if the file can't be confirmed gone. Previously a single dropped call here — under
    // sustained throttling on a large batch — left the old item in place while the caller proceeded
    // as if it had been deleted, so SPMI rejected the re-import with "already exists" even though
    // overwrite was requested (observed on a 5,000-file overwrite run, 2026-07-02). Callers must
    // check the result and fail the item cleanly rather than queue a doomed import.
    // onFail, when provided, receives a short human-readable reason (HTTP status + trimmed SharePoint
    // error body) whenever this returns false — so the caller can surface WHY a delete failed into the
    // activity log instead of only "still present after delete." SharePoint's REST error body usually
    // names the real cause (locked/checked-out/access-denied/not-found).
    // onFail receives a short reason when this returns false. onTrace (when provided) ALWAYS receives
    // a one-line trace of what each REST step actually returned — recycle status, bin id, purge status —
    // even on success. That trace is the only way to see these steps in a Release build (the old
    // Debug.WriteLine calls are compiled out unless a debugger is attached), and it's what finally
    // showed the delete "succeeding" while the file survived.
    public async Task<bool> PermanentlyDeleteFileAsync(
        string siteUrl, string serverRelativeUrl, Action<string>? onFail = null, Action<string>? onTrace = null)
    {
        const int maxAttempts = 4;
        var baseUrl = siteUrl.TrimEnd('/');
        string lastReason = "unknown";
        var trace = new System.Text.StringBuilder();
        void Trace(string s) { trace.Append(s).Append(' '); }
        bool Done(bool ok, string? reason = null)
        {
            onTrace?.Invoke(trace.ToString().TrimEnd());
            if (!ok) onFail?.Invoke(reason ?? lastReason);
            return ok;
        }

        var uniqueId = await GetFileUniqueIdAsync(siteUrl, serverRelativeUrl);
        if (string.IsNullOrEmpty(uniqueId)) { Trace("resolve=null(gone)"); return Done(true); }
        Trace($"id={uniqueId}");

        // Step 1: recycle by ID. GetFileById('{guid}') keeps the path (and its '#'/'+'/'%' hazards) out
        // of the URL entirely. A recycle 404 is NOT treated as "already gone" here — the caller's verify
        // loop is the real gone-check; a false 404 would otherwise mask a delete that never happened.
        var recycleUrl = $"{baseUrl}/_api/web/GetFileById('{uniqueId}')/recycleObject";
        string? recycleBinGuid = null;
        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
            bool lastAttempt = attempt == maxAttempts - 1;
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

                var body = await recycleResponse.Content.ReadAsStringAsync();
                Trace($"recycle={(int)recycleResponse.StatusCode}");

                if (!recycleResponse.IsSuccessStatusCode)
                {
                    lastReason = $"recycle HTTP {(int)recycleResponse.StatusCode}: {Trim(body)}";
                    // recycleObject fails for a file with no list item (zombie AllDocs row) — fall back to
                    // deleteObject by ID, which removes the row at the AllDocs level directly.
                    if (lastAttempt)
                    {
                        var (delOk, delReason) = await TryDeleteObjectByIdAsync(siteUrl, uniqueId, Trace);
                        return Done(delOk, delReason ?? lastReason);
                    }
                    await Task.Delay(TimeSpan.FromSeconds(3 * (attempt + 1)));
                    continue;
                }

                using var doc = JsonDocument.Parse(body);
                recycleBinGuid = doc.RootElement.ValueKind == JsonValueKind.String
                    ? doc.RootElement.GetString()
                    : doc.RootElement.TryGetProperty("value", out var v) ? v.GetString() : null;
                break;
            }
            catch (Exception ex)
            {
                lastReason = $"recycle exception: {ex.Message}";
                if (lastAttempt) return Done(false);
                await Task.Delay(TimeSpan.FromSeconds(3 * (attempt + 1)));
            }
        }

        Trace($"bin={recycleBinGuid ?? "null"}");
        if (string.IsNullOrEmpty(recycleBinGuid)) return Done(false, Loc.T("Svc_Sp_RecycleReturnedNoBin", lastReason));

        // Step 2: purge. recycleObject drops the item in the WEB (first-stage) recycle bin, so purge from
        // _api/web/RecycleBin; the previous code purged _api/site/RecycleBin (site collection / second
        // stage), where a web-bin id isn't found — that 404 was then wrongly treated as success, so the
        // file was only recycled, never purged (and the item lingered). Try web first, site as fallback.
        // NOTE: a purge failure still leaves the file recycled (out of the folder), which is enough for
        // SPMI to re-import — so we return true once recycle succeeded even if the purge is imperfect.
        foreach (var scope in new[] { "web", "site" })
        {
            var purgeUrl = $"{baseUrl}/_api/{scope}/RecycleBin('{recycleBinGuid}')/DeleteObject";
            try
            {
                using var purgeResponse = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Post, purgeUrl);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                    return r;
                }, siteUrl);
                Trace($"purge[{scope}]={(int)purgeResponse.StatusCode}");
                if (purgeResponse.IsSuccessStatusCode) return Done(true);
            }
            catch (Exception ex)
            {
                Trace($"purge[{scope}]=ex");
                lastReason = $"purge exception: {ex.Message}";
            }
        }

        // Recycle succeeded (file left the folder) but neither bin purge confirmed. Still good enough for
        // the re-import; report success but leave the trace so a lingering-file case is diagnosable.
        return Done(true);
    }

    // Trims a SharePoint REST error body to a short single line for logging.
    private static string Trim(string body) =>
        System.Text.RegularExpressions.Regex.Replace(body ?? "", @"\s+", " ").Trim() is { Length: > 0 } s
            ? s[..Math.Min(s.Length, 200)] : Loc.T("Svc_Sp_EmptyBody");

    private async Task<(bool ok, string? reason)> TryDeleteObjectByIdAsync(
        string siteUrl, string uniqueId, Action<string>? trace = null)
    {
        const int maxAttempts = 4;
        var deleteUrl = $"{siteUrl.TrimEnd('/')}/_api/web/GetFileById('{uniqueId}')/deleteObject";
        string? lastReason = null;
        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
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
                var body = await response.Content.ReadAsStringAsync();
                trace?.Invoke($"deleteObject={(int)response.StatusCode}");
                if (response.IsSuccessStatusCode || response.StatusCode == System.Net.HttpStatusCode.NotFound)
                    return (true, null);
                lastReason = $"deleteObject HTTP {(int)response.StatusCode}: {Trim(body)}";
            }
            catch (Exception ex)
            {
                lastReason = $"deleteObject exception: {ex.Message}";
            }
            if (attempt < maxAttempts - 1)
                await Task.Delay(TimeSpan.FromSeconds(3 * (attempt + 1)));
        }
        return (false, lastReason);
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
            var contentRb = Graph.Drives[targetDriveId].Items[targetParentItemId]
                .ItemWithPath(Uri.EscapeDataString(fileName)).Content;
            DriveItem? item;
            if (overwrite)
            {
                item = await contentRb.PutAsync(content);
            }
            else
            {
                // A plain PUT to /content always replaces — it ignores conflict semantics entirely,
                // so any path reaching here with overwrite=false (a race where the file appeared
                // after the existence check) silently clobbered the target. Ask Graph to fail on
                // conflict instead, matching the large-file upload-session behavior.
                var req = contentRb.ToPutRequestInformation(content);
                var url = req.URI.ToString();
                url += (url.Contains('?') ? "&" : "?") + "@microsoft.graph.conflictBehavior=fail";
                item = await contentRb.WithUrl(url).PutAsync(content);
            }
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

        // 10 MiB slices (a multiple of the required 320 KiB granularity): at the old 320 KiB a
        // 1 GB file took ~3,200 sequential round trips; this cuts them ~32×. Graph allows slices
        // up to 60 MiB.
        var uploadTask = new Microsoft.Graph.LargeFileUploadTask<DriveItem>(
            session!, content, 32 * 320 * 1024, Graph.RequestAdapter);

        var uploadResult = await uploadTask.UploadAsync(new Progress<long>(uploaded =>
        {
            if (size > 0) progress?.Report((int)(uploaded * 100 / size));
        }));

        if (!uploadResult.UploadSucceeded)
            throw new Exception(Loc.T("Svc_Sp_LargeFileUploadFailed"));

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
        if (sitePagesId == 0) return Loc.T("Svc_Sp_SitePagesIDUnknownCannot");

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
        if (sitePagesId == 0) return Loc.T("Svc_Sp_SitePagesIDUnknownCannot2");
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
        {
            // Capture loop variables for the lambda — the Lazy factory runs on first .Value access.
            var cacheKey = $"{driveId}|{current}|{part}";
            var d = driveId; var p = current; var n = part;
            // GetOrAdd may construct multiple Lazy instances under contention but only stores one.
            // Because Lazy.Value is called after GetOrAdd returns, the discarded instances never
            // invoke their factories — so only one Graph call is made per unique folder segment.
            var lazy = _folderSegmentTasks.GetOrAdd(cacheKey,
                _ => new Lazy<Task<string>>(
                    () => GetOrCreateFolderAsync(d, p, n),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication));
            try
            {
                current = await lazy.Value;
            }
            catch
            {
                // Never cache a faulted resolution: the service lives for the whole app session,
                // so a single transient failure here used to poison this folder for every later
                // file — and every later run.
                _folderSegmentTasks.TryRemove(new KeyValuePair<string, Lazy<Task<string>>>(cacheKey, lazy));
                throw;
            }
        }
        return current;
    }

    // Clears the folder-segment cache. Called at the start of each copy run: cached item IDs go
    // stale if a target folder was deleted/renamed between runs, and a session-lifetime cache
    // otherwise serves those dead IDs forever.
    public void ResetFolderSegmentCache() => _folderSegmentTasks.Clear();

    // Column definitions are cached per listId for the app's session lifetime with no reset
    // anywhere except one narrow list-copy path's explicit skipCache:true. If the user adds or
    // retypes a target column between two runs without restarting the app, the general file-copy
    // custom-column path (ApplyFileCustomFieldsAsync's internal GetLibraryColumnsAsync call) kept
    // resolving against the stale definition — e.g. a Lookup column whose LookupListId changed
    // would resolve against the old list. Call at the start of every run, alongside
    // ResetFolderSegmentCache().
    public void ResetColumnCache() => _columnCache.Clear();

    // True for exceptions that represent a dropped/failed connection rather than a real Graph response
    // (including a proper throttle response, which Kiota's own retry handler already resolves before
    // it would ever reach this code). VERIFIED (114k-file run, 2026-07-01): under sustained throttling
    // a single "An error occurred while sending the request" transport blip during pre-flight folder
    // provisioning/scanning went uncaught here, propagated out of the whole drive-group loop, and hit
    // MigrationJobService's top-level catch-all — which marked EVERY file still pending (i.e. nearly
    // the entire 114k-file job) as Failed. The download/upload paths elsewhere in this codebase already
    // retry on exactly these exception types; pre-flight Graph calls need the same treatment so one
    // blip doesn't take down an entire run.
    private static bool IsTransientTransportError(Exception ex) =>
        ex is System.Net.Http.HttpRequestException || ex is IOException;

    private async Task<string> GetOrCreateFolderAsync(string driveId, string parentItemId, string folderName)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                return await GetOrCreateFolderCoreAsync(driveId, parentItemId, folderName);
            }
            catch (Exception ex) when (attempt < 4 && IsTransientTransportError(ex))
            {
                await Task.Delay(TimeSpan.FromSeconds((attempt + 1) * 3));
            }
        }
    }

    private async Task<string> GetOrCreateFolderCoreAsync(string driveId, string parentItemId, string folderName)
    {
        try
        {
            var existing = await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(folderName)).GetAsync();
            System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] EXISTS: {folderName}");
            return existing?.Id ?? throw new Exception(Loc.T("Svc_Sp_NullIDFromExisting"));
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
                return created?.Id ?? throw new Exception(Loc.T("Svc_Sp_NullIDFromCreated"));
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError conflict) when (conflict.ResponseStatusCode == 409)
            {
                System.Diagnostics.Debug.WriteLine($"[GetOrCreateFolder] 409 conflict (race), re-fetching: {folderName}");
                var existing = await Graph.Drives[driveId].Items[parentItemId]
                    .ItemWithPath(Uri.EscapeDataString(folderName)).GetAsync();
                return existing?.Id ?? throw new Exception(Loc.T("Svc_Sp_NullIDAfterConcurrent"));
            }
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    // Raised when a request is throttled and a retry is scheduled — lets the UI show
    // "Throttled — retrying in Ns…" instead of appearing stalled.
    public event Action<TimeSpan, int, int, string?>? Throttled;

    // Sends a SharePoint REST request with resilience:
    //  - 401: retried once with a force-refreshed token.
    //  - 429/503 (throttling): retried up to 5 times, honoring the Retry-After header,
    //    with capped exponential backoff + jitter when the header is absent.
    // buildRequest receives the bearer token and must return a fresh HttpRequestMessage each call.
    internal async Task<HttpResponseMessage> SendSharePointRequestAsync(
        Func<string, HttpRequestMessage> buildRequest,
        string siteUrl,
        string spScope = "Sites.ReadWrite.All",
        CancellationToken cancellationToken = default)
    {
        const int maxThrottleRetries = 8;
        var siteAuth = _router.ForSiteUrl(siteUrl);
        var token = await siteAuth.GetSharePointTokenAsync(siteUrl, spScope, cancellationToken);
        HttpResponseMessage response;
        bool refreshedToken = false;

        for (int attempt = 0; ; attempt++)
        {
            using var req = buildRequest(token);
            response = await _httpClient.SendAsync(req, cancellationToken);

            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized && !refreshedToken)
            {
                response.Dispose();
                token = await siteAuth.GetSharePointTokenAsync(siteUrl, spScope, cancellationToken, forceRefresh: true);
                refreshedToken = true;
                continue;
            }

            bool throttled = response.StatusCode is System.Net.HttpStatusCode.TooManyRequests
                                                 or System.Net.HttpStatusCode.ServiceUnavailable;
            if (throttled && attempt < maxThrottleRetries)
            {
                // Retry-After can arrive as a delta or an absolute date.
                var delay = response.Headers.RetryAfter?.Delta
                    ?? (response.Headers.RetryAfter?.Date is { } when
                            ? when - DateTimeOffset.UtcNow
                            : TimeSpan.FromSeconds(Math.Pow(2, attempt + 1))
                              + TimeSpan.FromMilliseconds(Random.Shared.Next(0, 1000))); // jitter
                if (delay < TimeSpan.Zero) delay = TimeSpan.FromSeconds(1);
                if (delay > TimeSpan.FromSeconds(120)) delay = TimeSpan.FromSeconds(120);

                System.Diagnostics.Debug.WriteLine(
                    $"[SP-REST] {(int)response.StatusCode} throttled — retrying in {delay.TotalSeconds:N0}s (attempt {attempt + 1}/{maxThrottleRetries})");
                Throttled?.Invoke(delay, attempt + 1, maxThrottleRetries, null);
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

}
