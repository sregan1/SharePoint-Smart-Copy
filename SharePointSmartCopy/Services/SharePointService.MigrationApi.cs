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

// Partial class SharePointService: MigrationApi. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
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
        var token = await _router.ForSiteUrl(siteUrl).GetSharePointTokenAsync(siteUrl, "AllSites.FullControl");
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
            throw new InvalidOperationException(Loc.T("Svc_Sp_CannotDetermineLibraryPath", driveId));
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
        byte[] encryptionKey, CancellationToken ct = default)
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
        }, siteUrl, "AllSites.FullControl", ct);
        var body = await response.Content.ReadAsStringAsync(ct);
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
            var msg = errInfo.TryGetProperty("ErrorMessage", out var m) ? m.GetString() : Loc.T("Svc_Sp_UnknownCSOMError2");
            var code = errInfo.TryGetProperty("ErrorCode", out var c) ? c.GetInt32() : 0;
            if (code == -2147024891) // E_ACCESSDENIED
            {
                var (uTitle, uEmail, uIsAdmin) = await GetCurrentUserInfoAsync(siteUrl);
                throw new UnauthorizedAccessException(
                    Loc.T("Svc_Sp_TheMigrationAPIRequires", siteUrl) +
                    Loc.T("Svc_Sp_SPSeesYouAs", uTitle, uEmail, uIsAdmin) +
                    Loc.T("Svc_Sp_NoteGlobalAdminSharePoint") +
                    Loc.T("Svc_Sp_FixGoToThe"),
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

        throw new Exception(Loc.T("Svc_Sp_CouldNotFindJob", body[..Math.Min(body.Length, 500)]));
    }

    // Polls the migration job using the paging-based GetMigrationJobProgress endpoint.
    // Yields each event JSON string. Caller should stop when a JobEnd event is received.
    public async IAsyncEnumerable<JsonElement> PollMigrationJobAsync(
        string siteUrl, string jobId, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        int nextToken  = 0;
        int pollCount  = 0;
        int consecutiveTimeouts = 0;
        var pollDelay  = TimeSpan.FromSeconds(2);
        // A job that never emits JobEnd (stuck in SharePoint's queue, or its log evicted) would otherwise be polled forever,
        // pinning one of the few import workers until the whole run stalls. Give up after a long stretch with no event.
        var lastEventUtc = DateTime.UtcNow;
        var stallLimit   = TimeSpan.FromHours(2);

        while (!cancellationToken.IsCancellationRequested)
        {
            if (DateTime.UtcNow - lastEventUtc > stallLimit)
                throw new TimeoutException(Loc.T("Svc_Sp_MigrationJobStalled", (int)stallLimit.TotalMinutes));
            // Poll immediately on the first pass (the previous fixed 3 s pre-delay added latency to
            // every batch); then back off for long-running jobs to reduce API noise:
            // 2 s for the first ~1 min → 10 s for the next ~6 min → 30 s thereafter.
            if (pollCount > 0)
                await Task.Delay(pollDelay, cancellationToken);
            pollCount++;
            if (pollCount == 40)  pollDelay = TimeSpan.FromSeconds(10);
            if (pollCount == 100) pollDelay = TimeSpan.FromSeconds(30);

            // Guid parameters in SP REST require the guid'...' syntax, not just '...'.
            // AllSites.FullControl is needed — the same reason CreateMigrationJobEncrypted needs it.
            var url = $"{siteUrl.TrimEnd('/')}/_api/site/GetMigrationJobProgress(jobId=guid'{jobId}',nextToken={nextToken})";
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer",
                await _router.ForSiteUrl(siteUrl).GetSharePointTokenAsync(siteUrl, "AllSites.FullControl", cancellationToken));
            req.Headers.Accept.ParseAdd("application/json;odata=nometadata");

            HttpResponseMessage response;
            try
            {
                response = await _httpClient.SendAsync(req, cancellationToken);
                consecutiveTimeouts = 0;
            }
            catch (TaskCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                // User cancellation — propagate so the caller marks the batch Cancelled
                // instead of the iterator ending "cleanly" (which used to read as success).
                throw new OperationCanceledException("Migration job polling cancelled", cancellationToken);
            }
            catch (TaskCanceledException)
            {
                // HttpClient timeout on a single poll GET — the import is still running
                // server-side, so retry; only give up after several hung polls in a row.
                if (++consecutiveTimeouts >= 3)
                    throw new Exception(Loc.T("Svc_Sp_MigrationJobProgressPolling"));
                continue;
            }

            bool hitJobEnd = false;
            using (response)
            {
            bool isThrottle = response.StatusCode is System.Net.HttpStatusCode.TooManyRequests
                                                  or System.Net.HttpStatusCode.ServiceUnavailable;
            bool isTransient = !isThrottle && (int)response.StatusCode >= 500;
            if (isThrottle || isTransient)
            {
                // Throttle: honour Retry-After; 5xx: SP sometimes returns 500 when internally
                // throttled — treat as retriable rather than fatal.
                var delay = isThrottle
                    ? (response.Headers.RetryAfter?.Delta
                       ?? (response.Headers.RetryAfter?.Date is { } when
                               ? when - DateTimeOffset.UtcNow
                               : TimeSpan.FromSeconds(30)))
                    : TimeSpan.FromSeconds(30);
                if (delay < TimeSpan.Zero) delay = TimeSpan.FromSeconds(1);
                if (delay > TimeSpan.FromSeconds(120)) delay = TimeSpan.FromSeconds(120);
                await Task.Delay(delay, cancellationToken);
                continue;
            }
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
                        lastEventUtc = DateTime.UtcNow;
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

            } // end using (response)

            if (hitJobEnd) yield break;
        }
    }

}
