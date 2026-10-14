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

public partial class SharePointService
{
    // OneNote notebooks (and similar container items) can come back from Graph with only a
    // "package" facet (packageType "oneNote") and no populated "folder" facet, even though the
    // item is a real folder with children. Treating those as files misclassifies them throughout
    // scan/copy/verify — the file-copy path ends up copying the underlying folder anyway (Graph
    // copies the real object), while our own scan and verification keep comparing it as a file.
    private static bool IsFolderLike(DriveItem item) => item.Folder != null || item.Package != null;

    // C2 (source-scan perf): opt-in, OFF by default. When true, EnumerateFilesForCopyAsync uses
    // EnumerateFilesForCopyViaDeltaAsync (one /delta walk instead of one Graph listing PER FOLDER)
    // instead of the proven WalkFilesForCopyAsync path. Unverified against a live tenant — flip this
    // only for controlled testing; do not surface it in the UI/settings until it has been. See
    // EnumerateFilesForCopyViaDeltaAsync's header comment for the full risk/tradeoff writeup.
    internal static bool UseDeltaScan = false;

    private readonly AuthService _authService;
    private GraphServiceClient? _graphClient;
    // 30-minute timeout: site copies involve large file downloads and long-running REST calls.
    // 2-min pooled connection lifetime (matching the Graph client): the default is infinite,
    // which pins DNS resolution for the whole multi-hour run.
    private readonly HttpClient _httpClient = new(new System.Net.Http.SocketsHttpHandler
    {
        PooledConnectionLifetime = TimeSpan.FromMinutes(2),
        AutomaticDecompression   = System.Net.DecompressionMethods.All,
    }) { Timeout = TimeSpan.FromMinutes(30) };
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (string siteUrl, string listId, string listItemId)> _spIdsCache = new();
    // Form digests are site-scoped and valid ~30 minutes; without this a folder-correction pass
    // paid one _api/contextinfo POST per folder needing a CSOM write (thousands per run) where ~1
    // per site would do. 20-minute TTL leaves margin below the ~30-minute server-side expiry.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (string digest, DateTimeOffset fetchedUtc)> _digestCache = new();
    private static readonly TimeSpan DigestCacheLifetime = TimeSpan.FromMinutes(20);
    // EnsureUser resolves a claims login to a user ID and registers it on the site if missing — the
    // set of distinct author/editor emails across a library is typically tens, not one per folder.
    // Keyed "{siteUrl}|{loginName}".
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int?> _ensureUserCache = new();
    // Deduplicates concurrent folder-creation calls for the same path segment.
    // Lazy<Task<string>> ensures only one Graph call is made per "{driveId}|{parentId}|{segment}"
    // key even when multiple parallel tasks race — all share the same Task and await its result.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, Lazy<Task<string>>> _folderSegmentTasks = new(StringComparer.OrdinalIgnoreCase);
    // ValidateUpdateListItem's plain "yyyy-MM-dd HH:mm:ss" date strings (no "T"/"Z", see
    // PatchTimestampsCoreAsync) are parsed by SharePoint as being in the SITE's regional-settings
    // time zone, not UTC — a second, separate quirk from the "T"/"Z" parse failure fixed 2026-07-08.
    // Submitting a UTC instant verbatim as that plain string gets it re-interpreted as site-local and
    // shifted by the zone offset (e.g. a UTC-7 site showed 1:33pm for an actual 6:33am UTC modified
    // time). Cached per (site, exact instant) — same site+timestamp recurs often enough (e.g. an
    // unchanged CreatedDateTime reused across every version of a file) to make the dedup worthwhile,
    // and it turns a repeat REST call into a completed-Task hit instead.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, Lazy<Task<string?>>> _siteLocalDateStringCache = new(StringComparer.Ordinal);
    // Limits concurrent Graph $batch calls across all parallel SPMI job pipelines.
    // Each batch call carries 20 sub-requests; 10 parallel jobs firing simultaneously
    // would produce 200 concurrent Graph ops, reliably hitting per-app throttle limits.
    // 3 permits → at most 60 concurrent Graph ops from batch calls at any instant.
    private readonly System.Threading.SemaphoreSlim _batchGate = new(3, 3);

    // Picks the tenant (AuthService) per request. With nothing registered it always returns
    // _authService, so single-tenant behavior is unchanged. See TenantRouter.
    private readonly TenantRouter _router;
    public TenantRouter Router => _router;

    // Cross-tenant identity translation applied wherever an identity is written to the target.
    // Inactive (returns inputs untouched) unless the view model installs a configured mapper.
    public IdentityMapper Identity { get; set; } = IdentityMapper.None;

    // People a cross-tenant run could not find in the target. Cleared at the start of each run.
    public UnresolvedIdentityLog Unresolved { get; } = new();

    // True while source and target are different tenants (hosts registered with the router). Gates
    // everything that only makes sense across tenants — user pre-checks and unresolved-user capture.
    internal bool IsCrossTenant => _router.IsMultiTenant;

    public SharePointService(AuthService authService, TenantRouter? router = null)
    {
        _authService = authService;
        _router = router ?? new TenantRouter(authService);
        // Decorate the raw HTTP client (content downloads + SharePoint REST) — see MigrationUserAgent.
        // Pipe-delimited value isn't a valid structured UA, so add it without validation.
        _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", MigrationUserAgent);
    }

    // SharePoint Online and Graph throttle UNDECORATED traffic more aggressively than traffic that
    // carries a stable, structured User-Agent. Microsoft's documented format for a line-of-business
    // (non-ISV) app is "NONISV|Company|AppName/Version". Undecorated before 2026-07-20 — both the
    // Graph (Kiota) client and this raw HttpClient sent only a default/empty UA, which lands us in the
    // harder-throttled tier. This is a general-purpose tool, so the app's own name fills the company
    // field too — nothing organization-specific is embedded.
    private const string UserAgentApp = "SharePointSmartCopy";
    private static readonly string MigrationUserAgent = BuildUserAgent();
    private static string BuildUserAgent()
    {
        var v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var version = v == null ? "1.0" : $"{v.Major}.{v.Minor}.{v.Build}";
        return $"NONISV|{UserAgentApp}|{UserAgentApp}/{version}";
    }

    // Forces the decorated User-Agent onto every Graph request, overriding Kiota's default
    // "kiota-dotnet/…" UA (undecorated as far as SharePoint's throttling tiers are concerned).
    private sealed class SpoUserAgentHandler(string userAgent) : System.Net.Http.DelegatingHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            request.Headers.Remove("User-Agent");
            request.Headers.TryAddWithoutValidation("User-Agent", userAgent);
            return base.SendAsync(request, ct);
        }
    }

    public Task<string> GetSharePointTokenAsync(string siteUrl, string scope = "Sites.ReadWrite.All")
        => _router.ForSiteUrl(siteUrl).GetSharePointTokenAsync(siteUrl, scope);

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
        // Use GetFolderByServerRelativePath(decodedurl=...) instead of ...ByServerRelativeUrl(...):
        // the Url variant mishandles special characters in folder names (e.g. '#', '%', '&') and
        // returns 404 for them, which then failed the whole batch's manifest with the
        // "Could not resolve target subfolder ID" guard. The Path/decodedurl variant resolves them
        // correctly. Single quotes are OData-escaped (doubled), then the literal is URL-encoded.
        var odataLiteral = libraryServerRelativeUrl.Replace("'", "''");
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetFolderByServerRelativePath(decodedurl='{Uri.EscapeDataString(odataLiteral)}')?$select=UniqueId";
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
        var provider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(_router));
        var handlers = KiotaClientFactory.CreateDefaultHandlers();
        // CreateDefaultHandlers() order: [0]=UriReplacementHandler [1]=RetryHandler [2]=RedirectHandler ...
        // Replace the RetryHandler (index 1) with one configured for large migrations.
        // Default MaxRetry=3; with 10 parallel SPMI jobs the Graph per-app rate limit is
        // hit frequently enough that 3 retries is insufficient — bump to 10 (library max).
        handlers[1] = new Microsoft.Kiota.Http.HttpClientLibrary.Middleware.RetryHandler(
            new Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options.RetryHandlerOption
            {
                MaxRetry = 10,
            });
        // Insert our throttle-notify handler after RetryHandler (now at index 2, after RedirectHandler)
        // so it sits inside Kiota's retry loop and sees raw 429/503 responses before they are absorbed.
        // This makes Graph throttles visible to the adaptive parallelism controller.
        handlers.Insert(3, new GraphThrottleNotifyHandler(
            (delay, attempt, max, reason) => Throttled?.Invoke(delay, attempt, max, reason)));
        // Decorate every Graph request with our structured User-Agent (see MigrationUserAgent),
        // appended last so it runs closest to the wire and overrides Kiota's default UA.
        handlers.Add(new SpoUserAgentHandler(MigrationUserAgent));
        // Configure the underlying socket handler: gzip so metadata/JSON responses transfer
        // compressed (binary downloads are already-compressed and unaffected), and recycle pooled
        // connections every few minutes for stability across multi-hour 20k+ file runs.
        // EnableMultipleHttp2Connections matters at our concurrency levels: SocketsHttpHandler
        // multiplexes ALL requests to a host over a SINGLE HTTP/2 connection by default, so up to
        // 16 concurrent Graph calls (the app's max parallel-copies setting) would otherwise share
        // one TCP connection's congestion window and frame-processing loop instead of spreading
        // across several — a client-side ceiling independent of (and easy to mistake for) Graph's
        // own throttling.
        var finalHandler = new System.Net.Http.SocketsHttpHandler
        {
            AutomaticDecompression         = System.Net.DecompressionMethods.All,
            PooledConnectionLifetime       = TimeSpan.FromMinutes(2),
            EnableMultipleHttp2Connections = true,
        };
        var httpClient = KiotaClientFactory.Create(handlers, finalHandler);
        httpClient.Timeout = TimeSpan.FromMinutes(30);
        var adapter  = new HttpClientRequestAdapter(provider, httpClient: httpClient);
        _graphClient = new GraphServiceClient(adapter);
        HookThrottleTracker();
    }

    private GraphServiceClient Graph => _graphClient
        ?? throw new InvalidOperationException(Loc.T("Svc_Sp_NotInitializedPleaseSign"));

    // ── Shared throttle awareness ───────────────────────────────────────────────
    // Every Throttled event (from any phase — scan, analysis, download, $batch) records how long the
    // tenant asked us to back off. A fresh adaptive gate created for a LATER phase then inherits that
    // window via CreateThrottleAwareGate, so back-to-back analysis phases don't each re-discover the
    // same throttle by bursting into it and earning another long Retry-After. State lives on the
    // SharePointService instance, so it also carries across runs in the same app session.
    private readonly object _throttleStateLock = new();
    private DateTimeOffset _throttleSensitiveUntilUtc = DateTimeOffset.MinValue;
    private bool _throttleTrackerHooked;

    private void HookThrottleTracker()
    {
        if (_throttleTrackerHooked) return;   // Initialize can run more than once; subscribe only once
        _throttleTrackerHooked = true;
        Throttled += (delay, _, _, _) =>
        {
            if (delay <= TimeSpan.Zero) return;
            lock (_throttleStateLock)
            {
                var until = DateTimeOffset.UtcNow + delay;
                if (until > _throttleSensitiveUntilUtc) _throttleSensitiveUntilUtc = until;
            }
        };
    }

    // Remaining time the tenant is expected to stay throttle-sensitive (TimeSpan.Zero once clear).
    public TimeSpan RecentThrottleBackoff()
    {
        lock (_throttleStateLock)
        {
            var remaining = _throttleSensitiveUntilUtc - DateTimeOffset.UtcNow;
            return remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero;
        }
    }

    // Creates an adaptive gate for a metadata/analysis phase that (a) SOFT-STARTS so it eases into the
    // tenant rate limit instead of bursting to full width on the first tick, and (b) inherits any
    // active throttle window from an earlier phase/run (via RecentThrottleBackoff) so it starts at its
    // floor and waits that window out rather than slamming straight back into the same 429.
    internal AdaptiveParallelismController CreateThrottleAwareGate(int max, int softStart)
    {
        var gate = new AdaptiveParallelismController(max, Math.Clamp(softStart, 1, Math.Max(1, max)));
        var backoff = RecentThrottleBackoff();
        if (backoff > TimeSpan.Zero) gate.StepDown(backoff);
        return gate;
    }

}
