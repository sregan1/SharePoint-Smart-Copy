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

// Partial class SharePointService: Metadata. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
    // ── Metadata ──────────────────────────────────────────────────────────────

    // includeFolderColor: opt IN to the extra listItem/fields expand that carries folder color.
    // Defaults to FALSE deliberately — this method is on the per-FILE hot path (1-2 calls per file
    // in Enhanced REST mode, so 100k+ per run), and folder color is meaningless for a file. Adding
    // the expand unconditionally forced a server-side listItem join plus a fields projection on
    // every file read, against this file's documented payload-minimization grain (see the $select
    // comment above IsSpecialContainer, and the note there about a removed per-folder probe having
    // been "the dominant hidden cost of the source scan"). Only folder callers pass true.
    public async Task<FileMetadata> GetFileMetadataAsync(
        string driveId, string itemId, bool includeFolderColor = false)
    {
        string[] select = ["createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy", "size"];
        DriveItem? item = null;
        string? colorTag = null, colorHex = null;

        if (includeFolderColor)
        {
            // Best-effort: the color field names are unverified (see FolderColorTagFieldName), and an
            // unknown name inside a nested fields($select=...) may well be a 400 rather than a
            // silently-omitted property. Never let that break the metadata read this method exists
            // for — on ANY failure, fall through to the plain request below.
            try
            {
                item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
                {
                    // `listItem` is repeated in $select alongside the $expand — the common working
                    // idiom, and harmless. Graph's fieldValueSet is an OData OPEN type, so an unknown
                    // color field name here is silently omitted rather than erroring (unlike a REST
                    // $select, which hard-fails — see GetFolderCurrentMetadataAsync). The try/catch
                    // remains for transport-level failures.
                    cfg.QueryParameters.Select = [.. select, "listItem"];
                    cfg.QueryParameters.Expand = [$"listItem($expand=fields($select={FolderColorTagODataName},{FolderColorHexODataName}))"];
                });
                (colorTag, colorHex) = GetFolderColorValues(item);
            }
            catch
            {
                item = null;
            }
        }

        item ??= await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
            cfg.QueryParameters.Select = select);

        return new FileMetadata
        {
            CreatedDateTime  = item?.CreatedDateTime,
            CreatedByEmail   = GetIdentityEmail(item?.CreatedBy?.User),
            ModifiedDateTime = item?.LastModifiedDateTime,
            ModifiedByEmail  = GetIdentityEmail(item?.LastModifiedBy?.User),
            Size             = item?.Size,
            ColorTag         = colorTag,
            ColorHex         = colorHex,
        };
    }

    // Fetches each SOURCE folder's created/modified metadata for the migration manifest (folders
    // otherwise get a hardcoded placeholder date, which surfaced as a wrong "1999"/"2000" timestamp
    // on the target). `folders` maps a caller key (the folder's relative path) to a sample FILE item
    // id and a hop count: hopsUp=0 means the file sits DIRECTLY in that folder (its parentReference
    // IS the folder); hopsUp=N means the caller is borrowing a file from N levels deeper and this
    // method must walk up N extra parentReference links to reach the actual target folder — needed
    // for a folder that contains only subfolders and no files of its own (see the 2026-07-07 fix in
    // MigrationJobService, where the shallowest such descendant file is chosen). Still entirely
    // ID-based, so folder/file names containing '#'/'%'/'+' are unaffected. The library root ("" key)
    // is read from the drive root. Parallel + the Graph client's own throttle retries; a folder that
    // can't be resolved is simply absent (the builder keeps the placeholder for it).
    public async Task<Dictionary<string, FileMetadata>> FetchFolderMetadataAsync(
        string rootDriveId,
        IReadOnlyList<(string folderKey, string driveId, string sampleFileItemId, int hopsUp)> folders,
        int maxConcurrency,
        IProgress<int>? progress = null,
        CancellationToken ct = default)
    {
        int completed = 0;
        var result = new System.Collections.Concurrent.ConcurrentDictionary<string, FileMetadata>();

        // This is the heaviest per-item analysis phase (~4-5 Graph/REST calls per folder). Previously
        // it ran at a fixed width on the Graph client's per-request retry alone, so under throttling it
        // kept every lane hammering. Route it through a soft-starting, throttle-window-inheriting gate
        // (see CreateThrottleAwareGate) so it eases in and steps down on 429s like the other phases.
        using var gate = CreateThrottleAwareGate(maxConcurrency, Math.Min(Math.Max(1, maxConcurrency), 2));
        void onThrottle(TimeSpan delay, int __, int ___, string? ____) => gate.StepDown(delay);
        Throttled += onThrottle;
        try
        {

        // Library root folder.
        try
        {
            string[] rootSelect = ["id", "createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy"];
            DriveItem? root = null;
            string? rootColorTag = null, rootColorHex = null;
            // Same best-effort contract as GetFileMetadataAsync: an unverified color field must never
            // cost us the root folder's dates/authorship.
            try
            {
                root = await Graph.Drives[rootDriveId].Root.GetAsync(cfg =>
                {
                    cfg.QueryParameters.Select = [.. rootSelect, "listItem"];
                    cfg.QueryParameters.Expand = [$"listItem($expand=fields($select={FolderColorTagODataName},{FolderColorHexODataName}))"];
                }, cancellationToken: ct);
                (rootColorTag, rootColorHex) = GetFolderColorValues(root);
            }
            catch { root = null; }
            root ??= await Graph.Drives[rootDriveId].Root.GetAsync(cfg =>
                cfg.QueryParameters.Select = rootSelect, cancellationToken: ct);
            if (root != null)
            {
                result[string.Empty] = new FileMetadata
                {
                    CreatedDateTime  = root.CreatedDateTime,
                    CreatedByEmail   = GetIdentityEmail(root.CreatedBy?.User),
                    ModifiedDateTime = root.LastModifiedDateTime,
                    ModifiedByEmail  = GetIdentityEmail(root.LastModifiedBy?.User),
                    ProgId           = await GetFolderProgIdAsync(rootDriveId, "root"),
                    ColorTag         = rootColorTag,
                    ColorHex         = rootColorHex,
                };
            }
        }
        catch { /* keep placeholder */ }

        await Parallel.ForEachAsync(folders,
            new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, maxConcurrency), CancellationToken = ct },
            async (f, c) =>
            {
                // The gate (not MaxDegreeOfParallelism) governs live concurrency: it soft-starts and
                // shrinks on throttle, so effective width = min(MDOP, gate limit).
                await gate.WaitAsync(c);
                try
                {
                    var file = await Graph.Drives[f.driveId].Items[f.sampleFileItemId].GetAsync(cfg =>
                        cfg.QueryParameters.Select = ["parentReference"], cancellationToken: c);
                    var currentId = file?.ParentReference?.Id;
                    for (int hop = 0; hop < f.hopsUp && !string.IsNullOrEmpty(currentId); hop++)
                    {
                        var folderItem = await Graph.Drives[f.driveId].Items[currentId].GetAsync(cfg =>
                            cfg.QueryParameters.Select = ["parentReference"], cancellationToken: c);
                        currentId = folderItem?.ParentReference?.Id;
                    }
                    if (string.IsNullOrEmpty(currentId)) return;
                    // currentId is a FOLDER here, so opt into the color read (see GetFileMetadataAsync).
                    var metadata = await GetFileMetadataAsync(f.driveId, currentId, includeFolderColor: true);
                    var progId   = await GetFolderProgIdAsync(f.driveId, currentId);
                    result[f.folderKey] = string.IsNullOrEmpty(progId) ? metadata : new FileMetadata
                    {
                        CreatedDateTime  = metadata.CreatedDateTime,
                        CreatedByEmail   = metadata.CreatedByEmail,
                        ModifiedDateTime = metadata.ModifiedDateTime,
                        ModifiedByEmail  = metadata.ModifiedByEmail,
                        Size             = metadata.Size,
                        ProgId           = progId,
                        ColorTag         = metadata.ColorTag,
                        ColorHex         = metadata.ColorHex,
                        CustomFields     = metadata.CustomFields,
                    };
                }
                catch { /* keep placeholder */ }
                finally { gate.Release(); progress?.Report(Interlocked.Increment(ref completed)); }
            });

        return new Dictionary<string, FileMetadata>(result);
        }
        finally { Throttled -= onThrottle; }
    }

    // Non-root counterpart to FetchFolderMetadataAsync's per-folder branch, for callers that already
    // know each folder's real (driveId, itemId) — from the source scan (see CopyService's
    // SourceFileEntry.IsFolder) — instead of only a sample descendant file. Skips BOTH things that
    // made the old scheme expensive and, in one case, wrong:
    //   • the ancestor "hopsUp" parentReference walk (1+ Graph calls per folder) needed to resolve a
    //     folder's identity from a borrowed descendant file — not needed when the identity is already
    //     known, and it's what let a target-side-only path segment shift the walk onto an unrelated
    //     folder (including the target library root itself in one observed case)
    //   • the GetFolderProgIdAsync probe (2 calls per folder) — always returns null here by
    //     construction: a folder reaching this method came from the main scan, which diverts special
    //     containers (the only things that ever have a ProgID) to native copy before they can produce
    //     one of these entries
    // What remains (metadata + color in one call via GetFileMetadataAsync) is unchanged from the old
    // per-folder cost — same gate/backoff pattern as FetchFolderMetadataAsync.
    public async Task<(Dictionary<string, FileMetadata> Metadata, List<(string FolderKey, string Error)> Failures)> FetchFolderMetadataByIdentityAsync(
        IReadOnlyList<(string folderKey, string driveId, string itemId)> folders,
        int maxConcurrency,
        IProgress<int>? progress = null,
        CancellationToken ct = default)
    {
        int completed = 0;
        var result = new System.Collections.Concurrent.ConcurrentDictionary<string, FileMetadata>();
        var failures = new System.Collections.Concurrent.ConcurrentBag<(string FolderKey, string Error)>();
        if (folders.Count == 0) return (new Dictionary<string, FileMetadata>(), []);

        using var gate = CreateThrottleAwareGate(maxConcurrency, Math.Min(Math.Max(1, maxConcurrency), 2));
        void onThrottle(TimeSpan delay, int __, int ___, string? ____) => gate.StepDown(delay);
        Throttled += onThrottle;
        try
        {
            await Parallel.ForEachAsync(folders,
                new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, maxConcurrency), CancellationToken = ct },
                async (f, c) =>
                {
                    await gate.WaitAsync(c);
                    try
                    {
                        result[f.folderKey] = await GetFileMetadataAsync(f.driveId, f.itemId, includeFolderColor: true);
                    }
                    // Previously a bare `catch { }` — swallowed the real reason (429 vs. 404 vs.
                    // auth vs. timeout) entirely, so "Graph lookup failed after retries" was a guess,
                    // not a diagnosis. Capturing the exception here is what let us tell throttling
                    // apart from a genuinely missing/inaccessible item (2026-08-04 investigation).
                    catch (Exception ex) { failures.Add((f.folderKey, $"{ex.GetType().Name}: {ex.Message}")); }
                    finally { gate.Release(); progress?.Report(Interlocked.Increment(ref completed)); }
                });
            return (new Dictionary<string, FileMetadata>(result), [.. failures]);
        }
        finally { Throttled -= onThrottle; }
    }

    // Reads a folder's ProgID (e.g. "OneNote.Notebook" for OneNote notebook containers) via
    // SharePoint REST — Graph has no equivalent property. ProgID is what tells SharePoint's UI a
    // folder is a special container rather than a plain folder (it's what makes a notebook folder
    // open in the OneNote web app instead of just browsing into it). This is a property of the
    // SP.Folder resource, reached here via the list item's /Folder navigation property rather than
    // GetFolderById(GUID) — an earlier version used Graph's sharepointIds.ListItemUniqueId as that
    // GUID and it silently 404'd (never confirmed to equal the folder's own UniqueId), which reads
    // as "no ProgID" identically to a real plain folder, so the bug was invisible without comparing
    // against a live tenant. Routing through GetSharePointIdsAsync (listId + list-item ID, retried,
    // already proven reliable elsewhere in this file) and then Folder?$select=ProgID avoids that
    // GUID-identity assumption entirely. It's silently dropped by the Migration API's <Folder>
    // import (which only carries TimeCreated/TimeLastModified — see PatchFolderMetadataAsync), so
    // it has to be read from the source and reapplied after import the same way Author/Editor
    // already are. Returns null (not empty string) for a plain folder with no ProgID, so callers
    // can tell "no value" apart from "fetch failed".
    internal async Task<string?> GetFolderProgIdAsync(string driveId, string itemId)
    {
        var ids = await GetSharePointIdsAsync(driveId, itemId);
        if (ids == null) return null;
        var (siteUrl, listId, listItemId) = ids.Value;
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({listItemId})/Folder?$select=ProgID";
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
            var value = doc.RootElement.TryGetProperty("ProgID", out var v) ? v.GetString() : null;
            return string.IsNullOrEmpty(value) ? null : value;
        }
        catch { return null; }
    }

    // Writes a folder's ProgID via CSOM's ProcessQuery endpoint. A plain REST MERGE on the
    // SP.Folder resource (GetFolderById(...) + X-HTTP-Method: MERGE) returns HTTP 200 but silently
    // drops the write — confirmed 2026-07-08 by reading the folder back afterward and finding
    // ProgID still null, the exact same "accepted but ignored" symptom already documented for
    // Author/Editor on list items (see PatchFolderViaCsomAsync above). CSOM's Folder.Update() after
    // a direct SetProperty on the folder object (not a list-item SetFieldValue call) is the
    // working path for this specific property.
    public async Task<string?> PatchFolderProgIdAsync(string siteUrl, string listId, string folderUniqueId, string progId)
    {
        string? digest = await GetFormDigestAsync(siteUrl);
        if (digest == null)
            return Loc.T("Svc_Sp_CouldNotObtainA");

        const string clientContextTypeId = "{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}";
        var objectPaths =
            $"<StaticProperty Id=\"1\" TypeId=\"{clientContextTypeId}\" Name=\"Current\" />" +
            "<Property Id=\"2\" ParentId=\"1\" Name=\"Web\" />" +
            $"<Method Id=\"3\" ParentId=\"2\" Name=\"GetFolderById\"><Parameters><Parameter Type=\"Guid\">{folderUniqueId}</Parameter></Parameters></Method>";
        var actions =
            $"<SetProperty Id=\"10\" ObjectPathId=\"3\" Name=\"ProgID\"><Parameter Type=\"String\">{System.Security.SecurityElement.Escape(progId)}</Parameter></SetProperty>" +
            "<Method Name=\"Update\" Id=\"11\" ObjectPathId=\"3\" />";

        var requestXml =
            "<Request AddExpandoFieldTypeSuffix=\"true\" SchemaVersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" " +
            "ApplicationName=\"SharePointSmartCopy\" xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\">" +
            $"<Actions>{actions}</Actions><ObjectPaths>{objectPaths}</ObjectPaths></Request>";

        try
        {
            var url = $"{siteUrl.TrimEnd('/')}/_vti_bin/client.svc/ProcessQuery";
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Add("X-RequestDigest", digest);
                r.Headers.Accept.ParseAdd("application/json");
                r.Content = new StringContent(requestXml, System.Text.Encoding.UTF8, "text/xml");
                return r;
            }, siteUrl);

            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
                return $"CSOM ProcessQuery HTTP {(int)response.StatusCode}: {body[..Math.Min(body.Length, 300)]}";

            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.ValueKind == JsonValueKind.Array && doc.RootElement.GetArrayLength() > 0)
            {
                var first = doc.RootElement[0];
                if (first.ValueKind == JsonValueKind.Object &&
                    first.TryGetProperty("ErrorInfo", out var errInfo) &&
                    errInfo.ValueKind == JsonValueKind.Object)
                {
                    var msg  = errInfo.TryGetProperty("ErrorMessage", out var m) ? m.GetString() : Loc.T("Svc_Sp_UnknownCSOMError");
                    var code = errInfo.TryGetProperty("ErrorCode", out var c) ? c.ToString() : "?";
                    return Loc.T("Svc_Sp_CSOMErrorCode", code, msg);
                }
            }
            return null;
        }
        catch (Exception ex)
        {
            return $"CSOM exception: {ex.Message}";
        }
    }

    // Checks whether a child with this exact name exists directly under the given parent — used to
    // decide whether a native folder copy (which has no overwrite concept of its own) should
    // proceed under Skip/IfNewer.
    public async Task<bool> ChildExistsAsync(string driveId, string parentItemId, string name)
    {
        // Only a genuine 404 may map to false ("doesn't exist") — same reasoning as GetFileInfoAsync:
        // swallowing every failure here made a transient error (post-retry 429, network blip) read as
        // "not there", which in Overwrite/Skip mode proceeded as if the target were free to use.
        try
        {
            var existing = await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(name))
                .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
            return existing?.Id != null;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            return false;
        }
    }

    // Deletes a child with this exact name directly under the given parent, if one exists — Graph's
    // native /copy action has no "overwrite" concept of its own (a same-named existing item just
    // 409s with nameAlreadyExists), so Overwrite mode for a native folder copy has to clear the old
    // item first. Mirrors the existing-item check CreatePageStubAsync already uses for pages.
    // Returns true if something was deleted.
    //
    // Only a genuine 404 on the existence check may be swallowed as "nothing to delete" — any other
    // failure (including one from the DeleteAsync call itself, e.g. a locked/checked-out OneNote
    // notebook or Document Set) is thrown, not reported as false. Swallowing it here previously meant
    // the caller proceeded as if the target were clear, CopyFolderNativeAsync then failed with
    // "nameAlreadyExists", and the real cause (the delete failing, often for an access/lock reason)
    // never surfaced — the user saw a confusing "already exists" instead.
    public async Task<bool> DeleteChildIfExistsAsync(string driveId, string parentItemId, string name)
    {
        string? existingId;
        try
        {
            var existing = await Graph.Drives[driveId].Items[parentItemId]
                .ItemWithPath(Uri.EscapeDataString(name))
                .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
            existingId = existing?.Id;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            return false;
        }
        if (existingId == null) return false;
        await Graph.Drives[driveId].Items[existingId].DeleteAsync();
        return true;
    }

    // Copies a special folder (see SourceFileEntry.IsSpecialFolder) as a single server-side Graph
    // operation instead of reconstructing it file-by-file. Confirmed 2026-07-08: SharePoint's own
    // "Copy to" feature preserves a OneNote notebook's special behavior across libraries — proof
    // that a native copy duplicates the real backing object rather than rebuilding a new one from a
    // manifest/re-upload, which is why this is the only path that can carry over properties like
    // ProgID that aren't writable through any public API (see PatchFolderProgIdAsync above). Graph's
    // /copy action is async: it 202s immediately with a Location header pointing at a monitor
    // resource, which is polled here until it reports completed/failed.
    public async Task<string?> CopyFolderNativeAsync(
        string sourceDriveId, string sourceItemId,
        string targetDriveId, string targetParentItemId, string name,
        CancellationToken ct = default)
    {
        try
        {
            var sourceAuth = _router.ForDrive(sourceDriveId);
            // The Graph /copy endpoint only works inside one tenant. Cross-tenant needs a
            // download + re-upload fallback (planned); until then fail clearly instead of with an
            // opaque Graph 4xx.
            if (!ReferenceEquals(sourceAuth, _router.ForDrive(targetDriveId)))
                return Loc.T("Svc_Sp_NativeCopyCrossTenant");

            var copyBody = JsonSerializer.Serialize(new
            {
                parentReference = new { driveId = targetDriveId, id = targetParentItemId },
                name
            });
            var req = new HttpRequestMessage(HttpMethod.Post, $"{GraphBaseUrl}/drives/{sourceDriveId}/items/{sourceItemId}/copy")
            {
                Content = new System.Net.Http.StringContent(copyBody, System.Text.Encoding.UTF8, "application/json")
            };
            var batch = new BatchRequestContentCollection(Graph);
            var stepId = batch.AddBatchRequestStep(req);
            BatchResponseContentCollection batchResponse;
            using (_router.Scope(sourceAuth))
                batchResponse = await Graph.Batch.PostAsync(batch, ct);
            using var http = await batchResponse.GetResponseByIdAsync(stepId);

            if (http.StatusCode != System.Net.HttpStatusCode.Accepted)
            {
                var body = await http.Content.ReadAsStringAsync(ct);
                return Loc.T("Svc_Sp_NativeCopyFailedTo", (int)http.StatusCode, body[..Math.Min(body.Length, 300)]);
            }

            var monitorUrl = http.Headers.Location?.ToString();
            if (string.IsNullOrEmpty(monitorUrl))
                return Loc.T("Svc_Sp_NativeCopyStartedBut");

            // The monitor URL isn't guaranteed to be on graph.microsoft.com — for a SharePoint
            // (as opposed to OneDrive personal) drive it points at a SharePoint-hosted tracking
            // endpoint instead, e.g. .../_api/v2.1/drives/.../operations/{id}?...&tempauth=v1.ey...
            // That tempauth query parameter IS a complete, self-signed bearer credential for this
            // one resource — attaching an ADDITIONAL Authorization header of our own (confirmed
            // 2026-07-09: even a correctly-audienced SharePoint token) 401s, presumably because the
            // endpoint expects exactly one credential and rejects the mismatch/redundancy. So: no
            // header at all when the URL already carries tempauth; only attach our own bearer token
            // for the unconfirmed-but-possible case of a bare graph.microsoft.com monitor URL with
            // no embedded credential of its own.
            var monitorHost   = new Uri(monitorUrl).Host;
            var hasEmbeddedAuth = monitorUrl.Contains("tempauth=", StringComparison.OrdinalIgnoreCase);
            string? token = hasEmbeddedAuth
                ? null
                : monitorHost.Equals("graph.microsoft.com", StringComparison.OrdinalIgnoreCase)
                    ? await sourceAuth.GetAccessTokenAsync(cancellationToken: ct)
                    : await sourceAuth.GetSharePointTokenAsync($"https://{monitorHost}", cancellationToken: ct);
            for (int attempt = 0; attempt < 120; attempt++) // up to ~10 minutes for a large notebook
            {
                await Task.Delay(TimeSpan.FromSeconds(attempt < 10 ? 2 : 5), ct);
                using var monReq = new HttpRequestMessage(HttpMethod.Get, monitorUrl);
                if (token != null) monReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                using var monResp = await _httpClient.SendAsync(monReq, ct);
                var monBody = await monResp.Content.ReadAsStringAsync(ct);
                if (!monResp.IsSuccessStatusCode)
                    return Loc.T("Svc_Sp_NativeCopyMonitorFailed", (int)monResp.StatusCode, monBody[..Math.Min(monBody.Length, 300)], monitorUrl);

                using var monDoc = JsonDocument.Parse(monBody);
                var status = monDoc.RootElement.TryGetProperty("status", out var s) ? s.GetString() : null;
                if (string.Equals(status, "completed", StringComparison.OrdinalIgnoreCase)) return null;
                if (string.Equals(status, "failed", StringComparison.OrdinalIgnoreCase))
                {
                    var err = monDoc.RootElement.TryGetProperty("error", out var e) ? e.ToString() : monBody;
                    return Loc.T("Svc_Sp_NativeCopyFailed", err[..Math.Min(err.Length, 300)]);
                }
                // otherwise "inProgress"/"notStarted" — keep polling
            }
            return Loc.T("Svc_Sp_NativeCopyTimedOut");
        }
        catch (Exception ex)
        {
            return Loc.T("Svc_Sp_NativeCopyException", ex.Message);
        }
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
        string driveId, string itemId, CancellationToken ct = default)
    {
        var key = $"{driveId}|{itemId}";
        if (_spIdsCache.TryGetValue(key, out var cached))
        {
            System.Diagnostics.Debug.WriteLine($"[GetSPIds] cache hit: itemId={itemId}");
            return cached;
        }

        for (int attempt = 0; attempt < 3; attempt++)
        {
            if (attempt > 0) await Task.Delay(attempt * 1500, ct);
            try
            {
                System.Diagnostics.Debug.WriteLine($"[GetSPIds] attempt {attempt + 1}/3 for itemId={itemId}");
                var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
                    cfg.QueryParameters.Select = ["sharepointIds"], cancellationToken: ct);

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
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw; // genuine cancellation — don't retry it 3 times before giving up
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetSPIds] attempt {attempt + 1} threw {ex.GetType().Name}: {ex.Message}");
            }
        }
        System.Diagnostics.Debug.WriteLine($"[GetSPIds] all 3 attempts failed for itemId={itemId}");
        return null;
    }

    // Corrects a folder's Created/Modified/Author/Editor via REST after a Migration API import.
    // SPMI's <Folder> manifest element supports TimeCreated/TimeLastModified (confirmed working)
    // but NOT Author/ModifiedBy — SharePoint silently attributes every imported folder to the
    // account running the import, even on a brand-new folder creation, even though the exact same
    // GetUserId/UserGroup.xml mechanism correctly sets Author/ModifiedBy on <File> elements. This
    // is the one thing SPMI can't do for folders, so it's patched separately via the same
    // ValidateUpdateListItem REST call already used for per-version file editor correction.
    //
    // Date and author are set in ONE call, not two: any write to a list item bumps its own
    // Modified timestamp to "now" by default unless that SAME write also explicitly re-specifies
    // Modified — a separate author-only follow-up call after SPMI already set the correct date
    // would silently clobber it back to "now" (observed 2026-07-08: a freshly-corrected folder
    // showed as modified "a minute ago" — exactly when this pass ran).
    //
    // Takes the folder's SharePoint UniqueId (GUID) and the target library's list ID —
    // GetFolderById resolves the folder's own list item ID from the GUID (still fully ID-based, so
    // folder/file names containing '#'/'%'/'+' are unaffected), and the actual field write goes
    // through the direct lists('{listId}')/items({itemId}) path, matching the file-patch method
    // exactly. NOTE: the real fix for the "invalid date range" misattribution bug turned out to be
    // the Person field VALUE FORMAT (see below), not this path — routing through the direct path
    // was tried first and kept since it now matches the proven file mechanism exactly, but it's
    // unconfirmed whether the nested GetFolderById(...)/ListItemAllFields/ValidateUpdateListItem()
    // path would have worked fine with the corrected Person format too.
    // Returns (Error, ColorWarning). ColorWarning is set when dates/authorship were applied fine but
    // the folder color could not be — kept separate so an unsupported or absent color field (the
    // names are unverified; Microsoft documents no supported write API for folder color at all) never
    // inflates the "could not correct metadata" diagnostics the caller reports.
    public async Task<(string? Error, string? ColorWarning)> PatchFolderMetadataAsync(
        string siteUrl, string listId, string folderUniqueId,
        DateTimeOffset? createdDateTime, DateTimeOffset? modifiedDateTime,
        string? createdByEmail, string? modifiedByEmail,
        string? colorTag = null, string? colorHex = null,
        string? folderServerRelativeUrl = null)
    {
        // Cross-tenant: stamp the target-tenant equivalent of the source author/editor. A no-op unless
        // identity mapping is configured (idempotent, so the later BuildDateAndPersonFormValuesAsync is safe).
        createdByEmail  = Identity.MapEmail(createdByEmail);
        modifiedByEmail = Identity.MapEmail(modifiedByEmail);

        if (createdDateTime == null && modifiedDateTime == null &&
            string.IsNullOrEmpty(createdByEmail) && string.IsNullOrEmpty(modifiedByEmail) &&
            string.IsNullOrEmpty(colorTag) && string.IsNullOrEmpty(colorHex))
            return (null, null);

        // FAST PATH (added 2026-07-10) — the whole-tree repair fix made this correction pass run
        // unconditionally on EVERY run, even an all-skip Copy-If-Newer re-run, because that's the
        // only way an already-completed migration's wrong folder metadata gets fixed without a full
        // re-transfer. But that means a 123,000-file library's routine "anything changed?" re-run
        // was paying full EnsureUser+digest+CSOM+read-back cost for every folder, forever, not just
        // the one repair run that actually needed it — observed as a genuinely slow steady-state
        // re-run after the one-time repair completed successfully. Fix: read the folder's CURRENT
        // metadata with a single cheap REST call first; if it already matches the source, skip
        // EnsureUser/digest/CSOM entirely for that folder. Only a folder that's actually wrong (new,
        // never corrected, or somehow reverted) pays the full cost.
        int folderItemId;
        // Whether the color write should even be attempted — unknown (assume yes, matching the
        // pre-existing author/date behavior in the fallback branch below) unless the fast-path read
        // below confirms the target already matches the source.
        bool needsColorWrite = !string.IsNullOrEmpty(colorTag) || !string.IsNullOrEmpty(colorHex);
        var current = await GetFolderCurrentMetadataAsync(siteUrl, folderUniqueId);
        if (current != null)
        {
            var (id, actualAuthorEmail, actualEditorEmail, actualCreated, actualModified, actualColorTag, actualColorHex) = current.Value;
            folderItemId = id;
            bool authorOk = string.IsNullOrEmpty(createdByEmail) || EmailsMatch(actualAuthorEmail, createdByEmail);
            bool editorOk = string.IsNullOrEmpty(modifiedByEmail) || EmailsMatch(actualEditorEmail, modifiedByEmail);
            bool createdOk = createdDateTime == null || DatesClose(actualCreated, createdDateTime.Value);
            bool modifiedOk = modifiedDateTime == null || DatesClose(actualModified, modifiedDateTime.Value);
            bool colorOk = (string.IsNullOrEmpty(colorTag) || string.Equals(actualColorTag, colorTag, StringComparison.OrdinalIgnoreCase))
                        && (string.IsNullOrEmpty(colorHex) || string.Equals(actualColorHex, colorHex, StringComparison.OrdinalIgnoreCase));
            needsColorWrite = !colorOk;
            if (authorOk && editorOk && createdOk && modifiedOk && colorOk)
                return (null, null); // already correct — nothing to write
        }
        else
        {
            // Fall back to the old ID-only resolution if the combined read failed for any reason
            // (e.g. a transient error) — don't let an optimization's failure block the real fix.
            var idUrl = $"{siteUrl.TrimEnd('/')}/_api/web/GetFolderById('{folderUniqueId}')/ListItemAllFields/Id";
            using var idResponse = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, idUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);
            var idBody = await idResponse.Content.ReadAsStringAsync();
            if (!idResponse.IsSuccessStatusCode)
                return (Loc.T("Svc_Sp_CouldNotResolveFolder", (int)idResponse.StatusCode, idBody[..Math.Min(idBody.Length, 200)]), null);
            try
            {
                using var idDoc = JsonDocument.Parse(idBody);
                folderItemId = idDoc.RootElement.ValueKind == JsonValueKind.Object && idDoc.RootElement.TryGetProperty("value", out var v)
                    ? v.GetInt32()
                    : idDoc.RootElement.GetInt32();
            }
            catch (Exception ex)
            {
                return (Loc.T("Svc_Sp_CouldNotParseFolder", idBody[..Math.Min(idBody.Length, 100)], ex.Message), null);
            }
        }

        // A claims login for a person who has never been referenced anywhere on the target site
        // doesn't resolve via ValidateUpdateListItem — the site has to know about the account
        // first (its hidden User Information List), the same requirement any principal-resolution
        // path in SharePoint has. EnsureUser registers it if missing.
        //
        // Unlike the original version of this method, an unresolved user is NOT sent to
        // ValidateUpdateListItem at all — submitting Editor/Author alongside Modified/Created in
        // one call and having the Person field fail is a known source of misleading errors:
        // SharePoint's per-field validation can misattribute a Person-field resolution failure to
        // an unrelated field in the SAME call (observed 2026-07-08: a folder's real editor doesn't
        // exist on the target — a plausible cross-tenant migration scenario — and the reported
        // error was "Modified: You must specify a valid date within the range of 1/1/1900 and
        // 12/31/8900", not anything mentioning the Editor field at all). Skipping unresolvable
        // fields up front means the date write still succeeds cleanly, and the reported error
        // names the actual person field, not a confusing date complaint.
        var skipped = new List<string>();
        string? resolvedCreatedBy = createdByEmail;
        string? resolvedModifiedBy = modifiedByEmail;
        int? createdById = null;
        int? modifiedById = null;
        if (!string.IsNullOrEmpty(createdByEmail))
        {
            createdById = await EnsureUserAsync(siteUrl, createdByEmail);
            if (createdById == null)
            {
                skipped.Add(Loc.T("Svc_Sp_AuthorNotAResolvable", createdByEmail));
                resolvedCreatedBy = null;
            }
        }
        if (!string.IsNullOrEmpty(modifiedByEmail))
        {
            if (string.Equals(modifiedByEmail, createdByEmail, StringComparison.OrdinalIgnoreCase))
            {
                resolvedModifiedBy = resolvedCreatedBy; // same account — reuse the Author check above
                modifiedById = createdById;
            }
            else
            {
                modifiedById = await EnsureUserAsync(siteUrl, modifiedByEmail);
                if (modifiedById == null)
                {
                    skipped.Add(Loc.T("Svc_Sp_EditorNotAResolvable", modifiedByEmail));
                    resolvedModifiedBy = null;
                }
            }
        }

        if (createdById == null && modifiedById == null && createdDateTime == null && modifiedDateTime == null && !needsColorWrite)
            return (skipped.Count > 0 ? Loc.T("Svc_Sp_NothingToWriteSkipped", string.Join(", ", skipped)) : null, null);

        // TENTH attempt 2026-07-09 — REST is now fully exhausted (Person-field format x2, REST path
        // x2, bNewDocumentUpdate x2, field combination x2 — nine variations, all producing the
        // identical "accepted but silently ignored" symptom for Author/Editor while Modified/Created
        // always worked). Research into how migration tools like ShareGate actually pull this off
        // turned up the real answer, from an archived Microsoft engineer note: setting Author/Editor
        // to an ARBITRARY (non-calling) user requires SharePoint's THIRD update mode,
        // `UpdateOverwriteVersion()` — distinct from both `Update()` (always stamps the calling
        // account as Editor, confirmed via research) and `SystemUpdate()`/`bNewDocumentUpdate`
        // (silently leaves Editor/Author unchanged, confirmed by our own read-back testing above).
        // REST's ValidateUpdateListItem has NO equivalent to UpdateOverwriteVersion — only CSOM
        // exposes it — which is exactly why every REST permutation failed identically no matter what
        // was changed in the request. Switched the actual field write to a CSOM ProcessQuery call.
        // Color goes FIRST, via the supported foldercoloring endpoint (see StampFolderColorAsync).
        // It is a list-item write, so it stamps Editor = calling account and Modified = now — doing it
        // before the correction below means those side effects get overwritten. The colour fields
        // themselves are read-only, so they cannot ride in the CSOM batch.
        string? colorWarning = null;
        if (needsColorWrite && !string.IsNullOrEmpty(colorHex) && !string.IsNullOrEmpty(folderServerRelativeUrl))
        {
            var stampError = await StampFolderColorAsync(siteUrl, folderServerRelativeUrl!, colorHex!);
            if (stampError != null) colorWarning = Loc.T("Svc_Sp_FolderColorNotApplied", stampError);
        }

        var csomError = await PatchFolderViaCsomAsync(
            siteUrl, listId, folderItemId, createdById, modifiedById, createdDateTime, modifiedDateTime,
            resolvedCreatedBy, resolvedModifiedBy);

        // The write may have succeeded, but still report a failure if a person field had to be
        // skipped up front — the folder is only partially corrected (date fixed, that one person
        // field still shows the importing account), and that's not "no error". Color is reported
        // separately so an unsupported/absent color field never counts as a metadata failure.
        if (csomError != null)
            return (csomError + (skipped.Count > 0 ? $" | skipped: {string.Join(", ", skipped)}" : ""), colorWarning);
        return (skipped.Count > 0 ? Loc.T("Svc_Sp_PartiallyCorrectedSkipped", string.Join(", ", skipped)) : null, colorWarning);
    }

    // Sets a folder's Author/Editor/Created/Modified via CSOM's ProcessQuery endpoint, using
    // UpdateOverwriteVersion() — the only SharePoint update mode that both (a) does not force
    // Editor to the calling account (unlike a regular Update()) and (b) actually applies an
    // explicitly-set Editor/Author value instead of silently ignoring it (unlike SystemUpdate(),
    // confirmed unable to set Editor per Microsoft Q&A 914732 and PnP-PowerShell#2016). REST's
    // ValidateUpdateListItem has no equivalent mode, which is why it could never do this.
    //
    // Deliberately does NOT carry folder color: the color fields are read-only, so color is applied
    // separately (and EARLIER) via StampFolderColorAsync. Keeping it out matters — ProcessQuery
    // aborts at the first failing action, so one rejected color field here would take the whole
    // date/authorship write down with it.
    private async Task<string?> PatchFolderViaCsomAsync(
        string siteUrl, string listId, int itemId, int? authorUserId, int? editorUserId,
        DateTimeOffset? created, DateTimeOffset? modified,
        string? expectedAuthorEmail, string? expectedEditorEmail)
    {
        if (authorUserId == null && editorUserId == null && created == null && modified == null)
            return null;

        string? digest = await GetFormDigestAsync(siteUrl);
        if (digest == null)
            return Loc.T("Svc_Sp_CouldNotObtainA");

        const string fieldUserValueTypeId = "{c956ab54-16bd-4c18-89d2-996f57282a6f}";
        const string clientContextTypeId = "{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}";

        var objectPaths = new StringBuilder();
        objectPaths.Append($"<StaticProperty Id=\"1\" TypeId=\"{clientContextTypeId}\" Name=\"Current\" />");
        objectPaths.Append("<Property Id=\"2\" ParentId=\"1\" Name=\"Web\" />");
        objectPaths.Append("<Property Id=\"3\" ParentId=\"2\" Name=\"Lists\" />");
        objectPaths.Append($"<Method Id=\"4\" ParentId=\"3\" Name=\"GetById\"><Parameters><Parameter Type=\"Guid\">{listId}</Parameter></Parameters></Method>");
        objectPaths.Append($"<Method Id=\"5\" ParentId=\"4\" Name=\"GetItemById\"><Parameters><Parameter Type=\"Number\">{itemId}</Parameter></Parameters></Method>");

        string BuildRequestXml()
        {
            var actions = new StringBuilder();
            int nextId = 10;
            void SetFieldValueLiteral(string fieldName, string xmlValue)
            {
                actions.Append($"<Method Name=\"SetFieldValue\" Id=\"{nextId++}\" ObjectPathId=\"5\"><Parameters>" +
                    $"<Parameter Type=\"String\">{System.Security.SecurityElement.Escape(fieldName)}</Parameter>{xmlValue}" +
                    "</Parameters></Method>");
            }
            // FieldUserValue is a CSOM "ValueObject", not a server-tracked object — it's passed INLINE
            // as a typed Parameter directly in the method call (a <Parameter TypeId="..."><Property .../>
            // </Parameter> block), never via a separate <Constructor> ObjectPath + SetProperty pair (that
            // pattern is for real objects with server-side identity, which a plain value type isn't).
            // CORRECTED 2026-07-09 after "Cannot find stub for type" — the GUID itself was right, the
            // Constructor-based usage of it was wrong.
            void SetPersonField(string fieldName, int userId)
            {
                SetFieldValueLiteral(fieldName,
                    $"<Parameter TypeId=\"{fieldUserValueTypeId}\"><Property Name=\"LookupId\" Type=\"Int32\">{userId}</Property></Parameter>");
            }

            if (authorUserId.HasValue) SetPersonField("Author", authorUserId.Value);
            if (editorUserId.HasValue) SetPersonField("Editor", editorUserId.Value);
            if (created.HasValue)
                SetFieldValueLiteral("Created", $"<Parameter Type=\"DateTime\">{created.Value.ToUniversalTime():yyyy-MM-ddTHH:mm:ssZ}</Parameter>");
            if (modified.HasValue)
                SetFieldValueLiteral("Modified", $"<Parameter Type=\"DateTime\">{modified.Value.ToUniversalTime():yyyy-MM-ddTHH:mm:ssZ}</Parameter>");
            actions.Append($"<Method Name=\"UpdateOverwriteVersion\" Id=\"{nextId++}\" ObjectPathId=\"5\" />");

            return "<Request AddExpandoFieldTypeSuffix=\"true\" SchemaVersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" " +
                "ApplicationName=\"SharePointSmartCopy\" xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\">" +
                $"<Actions>{actions}</Actions><ObjectPaths>{objectPaths}</ObjectPaths></Request>";
        }

        // Returns null on success, else an error string.
        async Task<string?> SendAsync(string requestXml)
        {
            try
            {
                var url = $"{siteUrl.TrimEnd('/')}/_vti_bin/client.svc/ProcessQuery";
                using var response = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Post, url);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Add("X-RequestDigest", digest);
                    r.Headers.Accept.ParseAdd("application/json");
                    r.Content = new StringContent(requestXml, System.Text.Encoding.UTF8, "text/xml");
                    return r;
                }, siteUrl);

                var body = await response.Content.ReadAsStringAsync();
                if (!response.IsSuccessStatusCode)
                    return $"CSOM ProcessQuery HTTP {(int)response.StatusCode}: {body[..Math.Min(body.Length, 300)]}";

                // ProcessQuery's response is a flat JSON array: the first element carries ErrorInfo
                // (null on success), interleaved afterward with [actionId, resultValue] pairs for any
                // action that returns data (none of ours do — SetFieldValue/UpdateOverwriteVersion are
                // void). ProcessQuery aborts at the first failing action and reports it there, so that
                // first element is sufficient.
                try
                {
                    using var doc = JsonDocument.Parse(body);
                    if (doc.RootElement.ValueKind == JsonValueKind.Array && doc.RootElement.GetArrayLength() > 0)
                    {
                        var first = doc.RootElement[0];
                        if (first.ValueKind == JsonValueKind.Object &&
                            first.TryGetProperty("ErrorInfo", out var errInfo) &&
                            errInfo.ValueKind == JsonValueKind.Object)
                        {
                            var msg = errInfo.TryGetProperty("ErrorMessage", out var m) ? m.GetString() : Loc.T("Svc_Sp_UnknownCSOMError");
                            var code = errInfo.TryGetProperty("ErrorCode", out var c) ? c.ToString() : "?";
                            return Loc.T("Svc_Sp_CSOMErrorCode", code, msg);
                        }
                    }
                }
                catch (Exception ex)
                {
                    return Loc.T("Svc_Sp_CouldNotParseCSOM", ex.Message, body[..Math.Min(body.Length, 300)]);
                }
                return null;
            }
            catch (Exception ex)
            {
                return $"CSOM exception: {ex.Message}";
            }
        }

        var error = await SendAsync(BuildRequestXml());
        if (error != null)
            return error;

        // As with the REST path, a clean response still isn't proof the Person fields actually
        // persisted — read the item back and compare.
        return (authorUserId.HasValue || editorUserId.HasValue)
            ? await VerifyFolderPersonFieldsAsync(siteUrl, listId, itemId, expectedAuthorEmail, expectedEditorEmail)
            : null;
    }

    // Applies a folder's color via SharePoint's dedicated folder-coloring endpoint — the SUPPORTED
    // write path, and the one SharePoint's own UI uses. The `_ColorHex`/`_ColorTag` list fields are
    // both ReadOnlyField="TRUE" on current SharePoint Online, so writing them through CSOM
    // SetFieldValue or ValidateUpdateListItem is rejected; cli-microsoft365 moved to this endpoint for
    // exactly that reason. `colorHex` is a palette INDEX as a string, "0".."15" (0/empty = yellow,
    // 1 = dark red, … 15 = light pink) — not an RGB value, despite the name. Out-of-range values
    // silently fall back to the default.
    //
    // MUST be called BEFORE the Created/Modified/Author/Editor correction for the same folder: this
    // is a write to the folder's list item, so it stamps Editor = calling account and Modified = now.
    // Running it first means the metadata correction that follows overwrites those side effects.
    // Ordering it the other way round is the bug that shipped in the first version of this feature.
    //
    // Takes the folder's SERVER-RELATIVE URL (e.g. /sites/Team/Shared Documents/Reports), not a GUID
    // or item id — DecodedUrl is a path-based parameter. Returns null on success, else an error string;
    // callers treat a failure as non-fatal (color is cosmetic, and this endpoint is undocumented).
    public async Task<string?> StampFolderColorAsync(
        string siteUrl, string serverRelativeFolderUrl, string colorHex, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(colorHex)) return null;

        try
        {
            // DecodedUrl is passed as an OData alias (@a1) rather than inline so that folder names
            // containing '#', '%', '+', '&' or a single quote can't break the request — the same
            // hazard the rest of this file avoids by staying ID-based. The alias VALUE is a quoted
            // OData string literal, so embedded single quotes are doubled, then URL-encoded.
            var literal = "'" + serverRelativeFolderUrl.Replace("'", "''") + "'";
            var url = $"{siteUrl.TrimEnd('/')}/_api/foldercoloring/stampcolor(DecodedUrl=@a1)" +
                      $"?@a1={Uri.EscapeDataString(literal)}";
            var payload = JsonSerializer.Serialize(new { coloringInformation = new { ColorHex = colorHex } });

            string? digest = await GetFormDigestAsync(siteUrl);

            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url)
                {
                    Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json")
                };
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                if (digest != null) r.Headers.Add("X-RequestDigest", digest);
                return r;
            }, siteUrl, cancellationToken: ct);

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync(ct);
                return $"stampcolor HTTP {(int)response.StatusCode}: {body[..Math.Min(body.Length, 300)]}";
            }
            return null;
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            return $"stampcolor exception: {ex.Message}";
        }
    }

    // Fetches an X-RequestDigest value via _api/contextinfo — required for CSOM ProcessQuery POSTs
    // even when authenticating with an OAuth bearer token (unlike REST endpoints, which don't need it).
    private async Task<string?> GetFormDigestAsync(string siteUrl)
    {
        if (_digestCache.TryGetValue(siteUrl, out var cached) &&
            DateTimeOffset.UtcNow - cached.fetchedUtc < DigestCacheLifetime)
            return cached.digest;

        try
        {
            var url = $"{siteUrl.TrimEnd('/')}/_api/contextinfo";
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent("", System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
            if (!response.IsSuccessStatusCode) return null;
            var body = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            string? digest = doc.RootElement.TryGetProperty("FormDigestValue", out var fd) ? fd.GetString()
                : doc.RootElement.TryGetProperty("GetContextWebInformation", out var gcwi) &&
                  gcwi.TryGetProperty("FormDigestValue", out var fd2) ? fd2.GetString() : null;
            if (digest != null) _digestCache[siteUrl] = (digest, DateTimeOffset.UtcNow);
            return digest;
        }
        catch { return null; }
    }

    // Single cheap read of a folder's CURRENT Id/Author/Editor/Created/Modified/color, used by
    // PatchFolderMetadataAsync's fast path to decide whether any write is needed at all. Returns
    // null on any failure so the caller can fall back to the old (slower but proven) resolution.
    // The color fields simply come back null if FolderColorTagFieldName/FolderColorHexFieldName
    // don't exist on this list — same "absent means no value" contract as everywhere else.
    private async Task<(int ItemId, string? AuthorEmail, string? EditorEmail, DateTimeOffset? Created, DateTimeOffset? Modified, string? ColorTag, string? ColorHex)?>
        GetFolderCurrentMetadataAsync(string siteUrl, string folderUniqueId)
    {
        try
        {
            // Attempt WITH the color fields, then WITHOUT if that's rejected. The color field names
            // are unverified (see FolderColorTagFieldName) and a bad name in a $select is a 400 for
            // the whole request — which would silently disable this entire fast path and put every
            // folder back on the full EnsureUser+digest+CSOM+read-back cost the 2026-07-10 fix was
            // added to avoid. The retry guarantees a color problem costs at most one extra GET on
            // the folders that need correcting, never the optimization itself.
            string BuildUrl(bool withColor) =>
                $"{siteUrl.TrimEnd('/')}/_api/web/GetFolderById('{folderUniqueId}')/ListItemAllFields" +
                "?$select=Id,Author/EMail,Editor/EMail,Created,Modified" +
                (withColor ? $",{FolderColorTagODataName},{FolderColorHexODataName}" : "") +
                "&$expand=Author,Editor";

            async Task<HttpResponseMessage> SendAsync(bool withColor) =>
                await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, BuildUrl(withColor));
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl);

            var response = await SendAsync(withColor: true);
            if (!response.IsSuccessStatusCode)
            {
                response.Dispose();
                response = await SendAsync(withColor: false);
            }
            using var _ = response;
            if (!response.IsSuccessStatusCode) return null;

            var body = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;
            if (!root.TryGetProperty("Id", out var idEl)) return null;

            string? authorEmail = root.TryGetProperty("Author", out var a) && a.ValueKind == JsonValueKind.Object &&
                a.TryGetProperty("EMail", out var ae) ? ae.GetString() : null;
            string? editorEmail = root.TryGetProperty("Editor", out var e) && e.ValueKind == JsonValueKind.Object &&
                e.TryGetProperty("EMail", out var ee) ? ee.GetString() : null;
            // Pinned to InvariantCulture: a non-Gregorian default calendar (e.g. ar-SA) could
            // otherwise misparse SharePoint's ISO 8601 string and silently read back as null.
            DateTimeOffset? created = root.TryGetProperty("Created", out var c) && c.ValueKind == JsonValueKind.String &&
                DateTimeOffset.TryParse(c.GetString(), System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var cv) ? cv : null;
            DateTimeOffset? modified = root.TryGetProperty("Modified", out var m) && m.ValueKind == JsonValueKind.String &&
                DateTimeOffset.TryParse(m.GetString(), System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var mv) ? mv : null;
            // `_ColorHex` is an INTEGER field, so accept a JSON number as well as a string — an
            // earlier string-only check read every colored folder back as null.
            static string? ReadColor(JsonElement obj, string name) =>
                obj.TryGetProperty(name, out var el)
                    ? el.ValueKind switch
                    {
                        JsonValueKind.String => el.GetString(),
                        JsonValueKind.Number => el.ToString(),
                        _ => null,
                    }
                    : null;
            string? colorTag = ReadColor(root, FolderColorTagODataName);
            string? colorHex = ReadColor(root, FolderColorHexODataName);

            return (idEl.GetInt32(), authorEmail, editorEmail, created, modified, colorTag, colorHex);
        }
        catch { return null; }
    }

    // Public wrapper around GetFolderCurrentMetadataAsync for callers outside this class that need a
    // pre-existing folder's OWN current dates/authorship — e.g. MigrationJobService stamping a
    // manifest SPFolder entry for a destination container folder that isn't part of the copied source
    // tree (no source-side FileMetadata exists for it). Returning its real current values, rather than
    // MigrationPackageBuilder's hardcoded placeholder date, makes that manifest entry a no-op restate
    // instead of overwriting the folder's real dates on import (see the 2026-08-02 fix).
    public async Task<FileMetadata?> GetFolderOwnMetadataAsync(string siteUrl, string folderUniqueId)
    {
        var m = await GetFolderCurrentMetadataAsync(siteUrl, folderUniqueId);
        if (m == null) return null;
        var (_, authorEmail, editorEmail, created, modified, colorTag, colorHex) = m.Value;
        return new FileMetadata
        {
            CreatedDateTime  = created,
            CreatedByEmail   = authorEmail,
            ModifiedDateTime = modified,
            ModifiedByEmail  = editorEmail,
            ColorTag         = colorTag,
            ColorHex         = colorHex,
        };
    }

    private static bool EmailsMatch(string? actual, string? expected) =>
        !string.IsNullOrEmpty(actual) && string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);

    // A couple of seconds' tolerance absorbs sub-second precision differences between how SPMI/CSOM
    // round-trip a DateTime and how Graph originally reported it — not a meaningful "is this stale"
    // threshold, just enough to avoid false mismatches on values that are actually identical.
    private static bool DatesClose(DateTimeOffset? actual, DateTimeOffset expected) =>
        actual.HasValue && Math.Abs((actual.Value - expected).TotalSeconds) < 2;

    // Reads Author/Editor straight back off the item after a ValidateUpdateListItem write and
    // compares against what was intended. ValidateUpdateListItem can report zero field errors while
    // still leaving a Person field unchanged (a documented SharePoint quirk with SystemUpdate-style
    // writes) — this is the only way to tell "actually persisted" from "silently ignored" without
    // depending on the (possibly cached) SharePoint UI.
    private async Task<string?> VerifyFolderPersonFieldsAsync(
        string siteUrl, string listId, int itemId, string? expectedCreatedBy, string? expectedModifiedBy)
    {
        if (string.IsNullOrEmpty(expectedCreatedBy) && string.IsNullOrEmpty(expectedModifiedBy))
            return null;

        try
        {
            var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({itemId})" +
                "?$select=Author/EMail,Author/Name,Editor/EMail,Editor/Name&$expand=Author,Editor";
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl);
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
                return null; // don't fail the whole correction over a diagnostic read-back

            using var doc = JsonDocument.Parse(body);
            string? actualAuthor = doc.RootElement.TryGetProperty("Author", out var a) && a.ValueKind == JsonValueKind.Object
                ? (a.TryGetProperty("EMail", out var ae) ? ae.GetString() : null) ?? (a.TryGetProperty("Name", out var an) ? an.GetString() : null)
                : null;
            string? actualEditor = doc.RootElement.TryGetProperty("Editor", out var e) && e.ValueKind == JsonValueKind.Object
                ? (e.TryGetProperty("EMail", out var ee) ? ee.GetString() : null) ?? (e.TryGetProperty("Name", out var en) ? en.GetString() : null)
                : null;

            var problems = new List<string>();
            if (!string.IsNullOrEmpty(expectedCreatedBy) &&
                !string.Equals(actualAuthor, expectedCreatedBy, StringComparison.OrdinalIgnoreCase) &&
                actualAuthor?.IndexOf(expectedCreatedBy, StringComparison.OrdinalIgnoreCase) < 0)
                problems.Add(Loc.T("Svc_Sp_AuthorStillExpected", actualAuthor, expectedCreatedBy));
            if (!string.IsNullOrEmpty(expectedModifiedBy) &&
                !string.Equals(actualEditor, expectedModifiedBy, StringComparison.OrdinalIgnoreCase) &&
                actualEditor?.IndexOf(expectedModifiedBy, StringComparison.OrdinalIgnoreCase) < 0)
                problems.Add(Loc.T("Svc_Sp_EditorStillExpected", actualEditor, expectedModifiedBy));

            return problems.Count > 0
                ? Loc.T("Svc_Sp_WriteReportedNoError", string.Join("; ", problems))
                : null;
        }
        catch
        {
            return null; // diagnostic-only — never fail the correction pass because the read-back broke
        }
    }

    private async Task<string?> PatchTimestampsViaRestAsync(
        string driveId, string itemId, DateTimeOffset? modified, DateTimeOffset? created,
        string? createdByEmail = null, string? modifiedByEmail = null)
    {
        var ids = await GetSharePointIdsAsync(driveId, itemId);
        if (ids == null) return Loc.T("Svc_Sp_SPIDsUnavailableItem");

        var (siteUrl, listId, listItemId) = ids.Value;
        return await PatchTimestampsCoreAsync(siteUrl, listId, listItemId, modified, created, createdByEmail, modifiedByEmail);
    }

    // Shared by PatchTimestampsCoreAsync (folders/files, standalone date+person correction) and
    // ApplyFileCustomFieldsCoreAsync's Migration API restamp (date+person folded into the same
    // ValidateUpdateListItem call as a custom-field write) — one place to build the
    // Modified/Created/Editor/Author formValues so both write paths use identical field-name/value
    // shapes and the same site-local date conversion.
    private async Task<List<(string FieldName, string FieldValue)>> BuildDateAndPersonFormValuesAsync(
        string siteUrl, DateTimeOffset? modified, DateTimeOffset? created,
        string? createdByEmail, string? modifiedByEmail)
    {
        var values = new List<(string, string)>();
        createdByEmail  = Identity.MapEmail(createdByEmail);
        modifiedByEmail = Identity.MapEmail(modifiedByEmail);
        if (modified.HasValue)
            values.Add(("Modified", await FormatAsSiteLocalDateStringAsync(siteUrl, modified.Value)));
        if (created.HasValue)
            values.Add(("Created", await FormatAsSiteLocalDateStringAsync(siteUrl, created.Value)));
        if (!string.IsNullOrEmpty(modifiedByEmail))
            values.Add(("Editor", $"i:0#.f|membership|{modifiedByEmail}"));
        if (!string.IsNullOrEmpty(createdByEmail))
            values.Add(("Author", $"i:0#.f|membership|{createdByEmail}"));
        return values;
    }

    private async Task<string?> PatchTimestampsCoreAsync(
        string siteUrl, string listId, string listItemId, DateTimeOffset? modified, DateTimeOffset? created,
        string? createdByEmail = null, string? modifiedByEmail = null)
    {
        // CORRECTED 2026-07-08 (was wrong all along, on both files and folders): ValidateUpdateListItem
        // does NOT reliably parse ISO 8601 strings with a "T"/"Z" designator ("yyyy-MM-ddTHH:mm:ssZ") —
        // this is a documented quirk (SharePoint/sp-dev-docs#4917, Microsoft Q&A 59836): the endpoint
        // parses formValues date strings the way a browser form field would, not as machine-readable
        // JSON dates, and a "T"/"Z"-bearing string can fail that parse and surface as a misleading
        // "you must specify a valid date within the range of 1/1/1900 and 12/31/8900" error — even
        // though the date itself is perfectly valid. This is what was actually behind the folder
        // correction pass failing on every folder regardless of Editor field format or REST path
        // (both red herrings chased first). The fix is a space-separated, locale-independent format
        // with no "T"/"Z": "yyyy-MM-dd HH:mm:ss".
        //
        // CORRECTED 2026-08-28: that plain string is NOT parsed as UTC — it's a second, separate
        // quirk from the one above. With no offset in the string, ValidateUpdateListItem interprets
        // it in the SITE's regional-settings time zone (the same "browser form field" behavior that
        // caused the original parse failure), not UTC. Submitting a UTC instant verbatim there gets
        // it shifted by the site's zone offset when SharePoint converts it to UTC for storage — e.g.
        // a UTC-7 site turned an actual 6:33am UTC modified time into 1:33pm.
        //
        // CORRECTED again 2026-08-28, same day: a first attempt tried to replicate the site's zone
        // locally by matching RegionalSettings/TimeZone's Description string against
        // TimeZoneInfo.GetSystemTimeZones() DisplayNames. That never actually matched anything —
        // SharePoint's classic time zone list uses "GMT"/"and" wording (e.g. "(GMT-08:00) Pacific
        // Time (US and Canada)") where .NET's DisplayName uses "UTC"/"&" — so the lookup silently
        // fell back to UTC every time, i.e. it changed nothing, matching the report that the time
        // was still off after that fix. Rather than hand-roll a second, still-guessable text-matching
        // heuristic, ask the SITE itself to do the conversion via its own
        // regionalsettings/timezone/UTCToLocalTime() REST function — this uses the exact same
        // zone/DST rules ValidateUpdateListItem will apply when it parses our submitted string back,
        // by construction, so there's nothing left to get wrong about which zone or rule set applies.
        var dateAndPersonValues = await BuildDateAndPersonFormValuesAsync(siteUrl, modified, created, createdByEmail, modifiedByEmail);
        var formValues = dateAndPersonValues.Select(v => (object)new { FieldName = v.FieldName, FieldValue = v.FieldValue }).ToList();
        // Folder color is deliberately NOT sent here. The color fields are ReadOnlyField="TRUE", and
        // ValidateUpdateListItem's commit is ALL-OR-NOTHING — the spec states that if any field
        // raises an exception the item is not committed at all, and an unknown field name faults the
        // whole call. Including color would therefore risk discarding the dates and authorship too.
        // Color is applied separately, and earlier, via StampFolderColorAsync.
        if (formValues.Count == 0) return null;

        return await PostValidateUpdateAsync(formValues);

        // Routed through SendSharePointRequestAsync like every other REST write in this file. It
        // previously called _httpClient.SendAsync directly — the only write here that did — which
        // meant this, the hottest write in the app (once per file, and once per version per file via
        // ApplyFileMetadataAsync), had NO 429/503 retry, NO 401 force-refresh, no cancellation, and
        // leaked its HttpResponseMessage. A throttle window or a token expiring mid-run therefore
        // silently left files with the migrating account as Author/Editor and today's dates.
        async Task<string?> PostValidateUpdateAsync(List<object> values)
        {
            try
            {
                var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({listItemId})/ValidateUpdateListItem()";
                var payload = JsonSerializer.Serialize(new { formValues = values, bNewDocumentUpdate = true });
                using var response = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Post, url)
                    {
                        Content = new System.Net.Http.StringContent(
                            payload, System.Text.Encoding.UTF8, "application/json")
                    };
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl);
                var body = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    var preview = body.Length > 200 ? body[..200] : body;
                    return $"ValidateUpdateListItem HTTP {(int)response.StatusCode}: {preview}";
                }

                // A parse failure here used to be swallowed and reported as SUCCESS, i.e. "metadata
                // applied" — the same class of silent-loss bug this file has hit repeatedly. Treat an
                // unreadable response as an error instead; the HTTP status alone is not proof, since
                // ValidateUpdateListItem reports per-field failures inside a 200.
                try
                {
                    using var doc = JsonDocument.Parse(body);
                    if (doc.RootElement.TryGetProperty("value", out var arr))
                    {
                        var fieldErrors = new List<string>();
                        foreach (var field in arr.EnumerateArray())
                        {
                            bool hasEx = field.TryGetProperty("HasException", out var he) &&
                                         he.ValueKind == JsonValueKind.True;
                            int  ec    = field.TryGetProperty("ErrorCode", out var ecEl) &&
                                         ecEl.ValueKind == JsonValueKind.Number ? ecEl.GetInt32() : 0;
                            if (hasEx || ec != 0)
                            {
                                var fn = field.TryGetProperty("FieldName",    out var fnEl) ? fnEl.GetString() ?? "" : "";
                                var em = field.TryGetProperty("ErrorMessage", out var emEl) ? emEl.GetString() ?? "" : "";
                                fieldErrors.Add($"{fn}: {em} (code {ec})");
                            }
                        }
                        if (fieldErrors.Count > 0)
                            return Loc.T("Svc_Sp_FieldErrors", string.Join("; ", fieldErrors));
                    }
                }
                catch (Exception ex)
                {
                    return Loc.T("Svc_Sp_CouldNotParseValidateUpdateListItem", ex.Message) + " " +
                           Loc.T("Svc_Sp_MetadataMayNotHave", body[..Math.Min(body.Length, 200)]);
                }

                return null;
            }
            catch (Exception ex)
            {
                return $"ValidateUpdateListItem exception: {ex.Message}";
            }
        }
    }

    // Returns the wall-clock string ValidateUpdateListItem must be given so that, once SharePoint
    // parses it using the SITE's regional-settings time zone (see PatchTimestampsCoreAsync), it lands
    // back on the exact UTC instant `utcInstant` represents. Delegates the conversion to the site's
    // own regionalsettings/timezone/UTCToLocalTime() REST function instead of guessing at the zone
    // locally — a prior attempt matched RegionalSettings/TimeZone's Description string against
    // TimeZoneInfo.GetSystemTimeZones() DisplayNames, but SharePoint's classic zone list wording
    // ("GMT", "and") never matches .NET's ("UTC", "&"), so that match always failed and silently fell
    // back to treating the instant as already-UTC — i.e. it changed nothing. Asking SharePoint itself
    // to convert removes any need to reproduce its zone/DST rules correctly. Falls back to formatting
    // `utcInstant` as-is (the pre-fix behavior) if the conversion call fails for any reason, so an
    // unreachable site can't turn into a hard failure of the whole metadata write.
    //
    // A FAILED conversion is never cached: it used to be (as the fallback string) for the rest of the
    // session, so one throttled call pinned that instant — typically the unchanged CreatedDateTime
    // reused across every version of a file — to a value off by the site's UTC offset. Now a failure
    // evicts the entry, is retried once, and only then falls back for that single call.
    private async Task<string> FormatAsSiteLocalDateStringAsync(string siteUrl, DateTimeOffset utcInstant)
    {
        var cacheKey = $"{siteUrl}|{utcInstant.UtcTicks}";
        var lazy = _siteLocalDateStringCache.GetOrAdd(cacheKey,
            _ => new Lazy<Task<string?>>(() => ConvertUtcToSiteLocalAsync(siteUrl, utcInstant)));
        var converted = await lazy.Value;
        if (converted != null) return converted;

        _siteLocalDateStringCache.TryRemove(cacheKey, out _);
        await Task.Delay(TimeSpan.FromSeconds(2));
        converted = await ConvertUtcToSiteLocalAsync(siteUrl, utcInstant);
        return converted
            ?? utcInstant.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
    }

    // A UTC ISO timestamp string → the site-local form ValidateUpdateListItem expects. An unparseable
    // value is passed through unchanged (SharePoint then reports it as a per-field error).
    private async Task<string> ToSiteLocalFormDateAsync(string siteUrl, string isoUtc)
        => DateTimeOffset.TryParse(isoUtc, System.Globalization.CultureInfo.InvariantCulture,
               System.Globalization.DateTimeStyles.AssumeUniversal, out var instant)
            ? await FormatAsSiteLocalDateStringAsync(siteUrl, instant)
            : isoUtc;

    // Null on any failure (the caller decides what to do); a string only for a real conversion.
    private async Task<string?> ConvertUtcToSiteLocalAsync(string siteUrl, DateTimeOffset utcInstant)
    {
        try
        {
            var utcParam = utcInstant.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", System.Globalization.CultureInfo.InvariantCulture);
            var url = $"{siteUrl.TrimEnd('/')}/_api/web/regionalsettings/timezone/utctolocaltime(datetime'{utcParam}')";
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
            string? raw = doc.RootElement.ValueKind == JsonValueKind.String
                ? doc.RootElement.GetString()
                : doc.RootElement.TryGetProperty("value", out var v) && v.ValueKind == JsonValueKind.String
                    ? v.GetString() : null;
            if (raw == null) return null;

            // The returned value is a naive local wall-clock string with no offset of its own —
            // parse its components as-is, don't let DateTime infer/attach any zone of its own.
            if (!DateTime.TryParse(raw, System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.NoCurrentDateDefault, out var local))
                return null;

            return local.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
        }
        catch { return null; }
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
            // Called once per replayed version during version-history copy, and each iteration has
            // added ~2 versions (upload + phantom) to this same item — an unbounded GetAsync() here
            // re-fetched the WHOLE, growing version collection just to read one Id, an O(N²)-bytes
            // cost across the replay. $top=1 + $select=id reduces it to the single field needed.
            var page = await Graph.Drives[driveId].Items[itemId].Versions.GetAsync(cfg =>
            {
                cfg.QueryParameters.Top    = 1;
                cfg.QueryParameters.Select = ["id"];
            });
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

}
