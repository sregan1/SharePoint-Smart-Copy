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

// Partial class SharePointService: Scan. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
    // ── Site ──────────────────────────────────────────────────────────────────

    public async Task<string> GetSiteIdAsync(string siteUrl)
    {
        var uri = new Uri(siteUrl.TrimEnd('/'));
        var hostname = uri.Host;
        var path = uri.AbsolutePath.TrimStart('/');
        var key = string.IsNullOrEmpty(path) ? hostname : $"{hostname}:/{path}";

        var site = await Graph.Sites[key].GetAsync();
        return site?.Id ?? throw new Exception(Loc.T("Svc_Sp_CouldNotFindSite", siteUrl));
    }

    // ── Libraries ─────────────────────────────────────────────────────────────

    public async Task<List<SharePointNode>> GetLibrariesAsync(string siteId, string siteUrl)
    {
        // No $select — the Kiota SDK does not reliably populate webUrl when explicitly selected
        // on the drives collection; the default response includes it.
        // Paginated: reading only the first page silently missed libraries on sites with many
        // drives (Site scope then skipped them without a trace).
        var page      = await Graph.Sites[siteId].Drives.GetAsync();
        var allDrives = new List<Drive>();
        while (page != null)
        {
            allDrives.AddRange(page.Value ?? []);
            if (page.OdataNextLink == null) break;
            page = await Graph.Sites[siteId].Drives.WithUrl(page.OdataNextLink).GetAsync();
        }
        var result = new List<SharePointNode>();

        foreach (var d in allDrives)
        {
            if (d.Id == null || d.Name == null) continue;
            _router.RegisterDriveForSite(d.Id, siteUrl);

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
            if (foldersOnly && !IsFolderLike(item)) continue;

            bool isFolder = IsFolderLike(item);
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
                // Package-only folders (e.g. OneNote notebooks) have no Folder facet, so ChildCount
                // is unavailable — assume they may have children rather than hiding the expander.
                HasChildren = isFolder && (item.Folder != null ? item.Folder.ChildCount > 0 : true)
            };

            if (node.HasChildren)
                node.Children.Add(Placeholder());

            result.Add(node);
        }
        return result;
    }

    // ── Enumerate all files under a folder (for copy) ─────────────────────────

    // One discovered source file, with the modified date the Children listing already carries —
    // captured here so the If Newer overwrite decision never has to re-fetch it per file later.
    // IsSpecialFolder marks a different kind of entry entirely: a folder (e.g. a OneNote notebook)
    // that carries a ProgID SharePoint sets internally and doesn't expose as writable via any
    // public API (see PatchFolderProgIdAsync) — reconstructing it file-by-file through SPMI/Enhanced
    // REST would silently lose that association, so the walk stops descending into it here and
    // hands it back as a single unit for the caller to copy natively (Graph's server-side /copy
    // action, which duplicates the real object instead of rebuilding one). ItemId/Name/RelativePath
    // refer to the FOLDER itself in that case, not a file, and Modified is unused.
    // IsEmptyFolder: this entry is a folder with no files anywhere in its subtree (directly or via
    // descendants) — the walk below never visits such a folder as anyone's child copy target, so
    // without an explicit marker it silently never gets created at the destination. ItemId/Modified
    // are unused for this case; RelativePath is the folder's own path, not a file's.
    // Size: the file's byte size from the same Children listing (also already in the payload).
    // Migration API mode uses it as the fallback for its large-file memory gate and per-batch byte
    // budget when the upfront metadata cache missed the file under throttling — without it a
    // multi-GB file whose metadata fetch failed was treated as size 0, bypassed the gate, and
    // 16 of them downloading concurrently exhausted process memory (observed 2026-07-18).
    // IsFolder: a metadata-only entry for an ORDINARY (non-special) folder — emitted once per folder,
    // regardless of whether it turns out to have children or not, so downstream folder-metadata
    // passes (both engines) never need a second, separate per-folder fetch or walk. Never emitted for
    // the walk's own ROOT (job.SourceItemId) — that folder is never encountered as "someone's child"
    // by this walk, exactly like IsSpecialContainer's root exception (see
    // IsRootFolderSpecialContainerAsync); root metadata keeps using each engine's existing separate
    // root-only fetch. Consumers should NOT create a copy job or grid row from this entry — it exists
    // purely to carry Created/CreatedByEmail/ModifiedByEmail data folders don't otherwise get. A
    // folder that also turns out empty gets BOTH this entry AND its own IsEmptyFolder entry — they
    // are complementary (one is metadata, the other drives creation/row-reporting), not alternatives.
    public readonly record struct SourceFileEntry(
        string DriveId, string ItemId, string Name, string RelativePath, DateTimeOffset? Modified,
        long? Size = null, bool IsSpecialFolder = false, bool IsEmptyFolder = false,
        bool IsFolder = false, DateTimeOffset? Created = null,
        string? CreatedByEmail = null, string? ModifiedByEmail = null);

    // Concurrent replacement for the old sequential recursive walk (same channel + sibling-fan-out
    // pattern as EnumerateFilesWithMetadataAsync below). The sequential version issued ONE Graph
    // call at a time — on a 3,000+-folder library that alone took ~30 minutes before the copy could
    // even start, with nothing logged. The controller bounds concurrent listings and backs off on
    // throttle; $select trims 100k+ full DriveItem payloads down to the five fields the copy needs.
    internal async IAsyncEnumerable<SourceFileEntry> EnumerateFilesForCopyAsync(
        string driveId, string rootItemId, string basePath,
        AdaptiveParallelismController controller,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
    {
        // See UseDeltaScan's declaration — opt-in, off by default, unverified against a live tenant.
        if (UseDeltaScan)
        {
            await foreach (var f in EnumerateFilesForCopyViaDeltaAsync(driveId, rootItemId, basePath, ct))
                yield return f;
            yield break;
        }

        var channel = System.Threading.Channels.Channel.CreateUnbounded<SourceFileEntry>();
        var walkTask = Task.Run(async () =>
        {
            try { await WalkFilesForCopyAsync(driveId, rootItemId, basePath, controller, channel.Writer, ct); }
            finally { channel.Writer.Complete(); }
        }, ct);

        await foreach (var f in channel.Reader.ReadAllAsync(ct))
            yield return f;

        await walkTask; // propagate walk exceptions (channel completion alone would swallow them)
    }

    // ── C2: delta-based source scan (opt-in — see UseDeltaScan) ────────────────────────────────
    //
    // WalkFilesForCopyAsync above issues one Graph /children listing PER FOLDER — cheap for a shallow
    // tree, but on a 100k-file/5,000-folder library that's ~5,000 round trips paid purely to discover
    // the tree shape, before any file transfer starts. /drives/{id}/items/{id}/delta returns the WHOLE
    // subtree (files AND folders) in large pages (~1,000-2,000 items/page) regardless of how many
    // folders that spans — the same 100k/5,000-folder library costs ~20-50 delta calls instead of
    // ~5,000 listings, because delta's page boundary is item count, not folder boundary.
    //
    // The catch: /delta does NOT support $expand=listItem (confirmed by a Microsoft maintainer — it
    // 400s with "One of the provided arguments is not acceptable"), and Document Set detection (see
    // IsSpecialContainer) depends on listItem.contentType. This method still gets OneNote-style
    // detection for free (the "package" facet rides along on the normal $select), but Document Sets
    // need a SEPARATE batched pass (BatchFetchContentTypeIdsAsync) over whatever plain folders survive
    // the package-facet filter — batched 20/call so this supplementary pass doesn't reintroduce the
    // "1 round trip per folder" cost delta was meant to eliminate.
    //
    // Path reconstruction: delta returns a flat list with parentReference.Id per item, not a
    // pre-built tree, so relative paths are resolved by walking each item's parent chain (memoized)
    // until hitting either the scan root itself or an item absent from the fetched set (which can
    // only be the scan root's own external parent, since delta only returns items inside the
    // requested subtree).
    //
    // UNVERIFIED AGAINST A LIVE TENANT. In particular: (a) that /delta on a non-root itemId behaves as
    // documented for an arbitrary folder (not just the drive root), (b) that an initial (no stored
    // deltaLink) call reliably returns a complete snapshot in one paginated sweep without a `deleted`
    // facet appearing among live items, (c) the actual round-trip count on a real large tenant. Kept
    // off by default (UseDeltaScan) until someone runs it against one.
    private async IAsyncEnumerable<SourceFileEntry> EnumerateFilesForCopyViaDeltaAsync(
        string driveId, string rootItemId, string basePath,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
    {
        var resolvedRootId = rootItemId == "root" ? "root" : rootItemId;
        var allItems = new List<DriveItem>();

        var page = await Graph.Drives[driveId].Items[resolvedRootId].Delta.GetAsDeltaGetResponseAsync(cfg =>
        {
            cfg.QueryParameters.Select = ["id", "name", "file", "folder", "package", "size",
                "lastModifiedDateTime", "createdDateTime", "createdBy", "lastModifiedBy",
                "parentReference", "root", "deleted"];
        }, ct);

        while (page != null)
        {
            allItems.AddRange(page.Value ?? []);
            if (page.OdataNextLink == null) break;
            ct.ThrowIfCancellationRequested();
            page = await Graph.Drives[driveId].Items[resolvedRootId].Delta
                .WithUrl(page.OdataNextLink).GetAsDeltaGetResponseAsync(cancellationToken: ct);
        }
        // page.OdataDeltaLink (present on the final page) marks the end of the initial sync — this
        // scan is one-shot per run, so it's deliberately not persisted (see this method's header:
        // the deltaLink-persistence half is a separate, unimplemented follow-up, not this fix).

        var itemsById = new Dictionary<string, DriveItem>(StringComparer.Ordinal);
        foreach (var item in allItems)
        {
            if (item.Id == null || item.Deleted != null) continue;
            itemsById[item.Id] = item;
        }

        bool IsSelfRootEntry(DriveItem item) =>
            item.Root != null || (rootItemId != "root" && item.Id == resolvedRootId);

        // Package-facet special containers (OneNote-style) are known for free from this same fetch —
        // no extra call. Document Sets need the supplementary contentType pass below.
        var plainFolderCandidates = new List<(string driveId, string itemId)>();
        var specialIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in itemsById.Values)
        {
            if (IsSelfRootEntry(item) || !IsFolderLike(item)) continue;
            if (item.Package != null) specialIds.Add(item.Id!);
            else plainFolderCandidates.Add((driveId, item.Id!));
        }

        if (plainFolderCandidates.Count > 0)
        {
            var contentTypeIds = await BatchFetchContentTypeIdsAsync(plainFolderCandidates, maxConcurrency: 6, ct: ct);
            foreach (var (_, itemId) in plainFolderCandidates)
            {
                if (contentTypeIds.TryGetValue(itemId, out var ctId) && ctId != null &&
                    ctId.StartsWith(DocumentSetContentTypeIdPrefix, StringComparison.OrdinalIgnoreCase))
                    specialIds.Add(itemId);
            }
        }

        // Descendants of a special container are handled by the native special-folder copy (see
        // CopySpecialFolderAsync) — this walk must emit exactly one entry FOR the special folder
        // itself and nothing for anything inside it, matching WalkFilesForCopyAsync's behavior (it
        // physically never recurses into one). Delta, unlike a recursive walk, returns everything in
        // one flat list regardless, so descendants of a special folder have to be filtered out here.
        var childrenByParent = new Dictionary<string, List<DriveItem>>(StringComparer.Ordinal);
        foreach (var item in itemsById.Values)
        {
            if (IsSelfRootEntry(item)) continue;
            var parentId = item.ParentReference?.Id;
            if (parentId == null) continue;
            if (!childrenByParent.TryGetValue(parentId, out var list))
                childrenByParent[parentId] = list = [];
            list.Add(item);
        }
        var excludedIds = new HashSet<string>(StringComparer.Ordinal);
        void ExcludeDescendants(string parentId)
        {
            if (!childrenByParent.TryGetValue(parentId, out var kids)) return;
            foreach (var kid in kids)
            {
                if (kid.Id == null || !excludedIds.Add(kid.Id)) continue;
                ExcludeDescendants(kid.Id);
            }
        }
        foreach (var specialId in specialIds) ExcludeDescendants(specialId);

        var pathCache = new Dictionary<string, string>(StringComparer.Ordinal);
        string ResolvePath(DriveItem item)
        {
            if (item.Id != null && pathCache.TryGetValue(item.Id, out var cached)) return cached;
            var parentId = item.ParentReference?.Id;
            string path;
            if (parentId == null || !itemsById.TryGetValue(parentId, out var parent) || IsSelfRootEntry(parent))
                path = item.Name ?? string.Empty;
            else
            {
                var parentPath = ResolvePath(parent);
                path = string.IsNullOrEmpty(parentPath) ? (item.Name ?? string.Empty) : $"{parentPath}/{item.Name}";
            }
            if (item.Id != null) pathCache[item.Id] = path;
            return path;
        }
        string FullRelativePath(DriveItem item)
        {
            var p = ResolvePath(item);
            return string.IsNullOrEmpty(basePath) ? p : (string.IsNullOrEmpty(p) ? basePath : $"{basePath}/{p}");
        }

        foreach (var item in itemsById.Values)
        {
            ct.ThrowIfCancellationRequested();
            if (item.Id == null || item.Name == null) continue;
            if (IsSelfRootEntry(item) || excludedIds.Contains(item.Id)) continue;

            var relativePath = FullRelativePath(item);

            if (IsFolderLike(item))
            {
                if (specialIds.Contains(item.Id))
                {
                    yield return new SourceFileEntry(
                        driveId, item.Id, item.Name, relativePath, item.LastModifiedDateTime, IsSpecialFolder: true);
                }
                else
                {
                    bool hasChildren = childrenByParent.TryGetValue(item.Id, out var kids) && kids.Count > 0;
                    if (!hasChildren)
                    {
                        yield return new SourceFileEntry(
                            driveId, item.Id,
                            relativePath.Contains('/') ? relativePath[(relativePath.LastIndexOf('/') + 1)..] : relativePath,
                            relativePath, null, IsEmptyFolder: true);
                    }
                    else
                    {
                        yield return new SourceFileEntry(
                            driveId, item.Id, item.Name, relativePath, item.LastModifiedDateTime,
                            IsFolder: true, Created: item.CreatedDateTime,
                            CreatedByEmail: GetIdentityEmail(item.CreatedBy?.User),
                            ModifiedByEmail: GetIdentityEmail(item.LastModifiedBy?.User));
                    }
                }
            }
            else
            {
                yield return new SourceFileEntry(
                    driveId, item.Id, item.Name, relativePath, item.LastModifiedDateTime, item.Size);
            }
        }
    }

    private async Task WalkFilesForCopyAsync(
        string driveId, string itemId, string basePath, AdaptiveParallelismController controller,
        System.Threading.Channels.ChannelWriter<SourceFileEntry> writer, CancellationToken ct)
    {
        List<DriveItem> items;
        await controller.WaitAsync(ct);
        try { items = await GetChildrenWithMetadataAsync(driveId, itemId, ct); }
        finally { controller.Release(); }

        // A folder with no children at all can never otherwise be represented — every other entry
        // this walk produces is a file (or special folder) that drags its ancestor chain into
        // existence as a side effect of its own target path. A genuinely empty folder has no such
        // file, so without this it is silently never created at the destination (see the
        // Translational(1) case: it existed at the source with zero files and never showed up at
        // the target, with no error anywhere — because nothing ever attempted to create it).
        if (items.Count == 0)
        {
            await writer.WriteAsync(new SourceFileEntry(
                driveId, itemId, basePath.Contains('/') ? basePath[(basePath.LastIndexOf('/') + 1)..] : basePath,
                basePath, null, IsEmptyFolder: true), ct);
            return;
        }

        var subfolderWalks = new List<Task>();
        foreach (var item in items)
        {
            ct.ThrowIfCancellationRequested();
            if (item.Id == null || item.Name == null) continue;
            var childPath = string.IsNullOrEmpty(basePath) ? item.Name : $"{basePath}/{item.Name}";
            if (IsFolderLike(item))
            {
                // Special containers (OneNote notebooks, Document Sets) are detected from the fields
                // already in this listing — see IsSpecialContainer. This replaced a per-folder ProgID
                // probe that cost two extra round-trips on EVERY folder in the tree.
                if (IsSpecialContainer(item))
                    await writer.WriteAsync(new SourceFileEntry(
                        driveId, item.Id, item.Name, childPath, item.LastModifiedDateTime, IsSpecialFolder: true), ct);
                else
                {
                    // Metadata-only entry for this ordinary folder (see SourceFileEntry.IsFolder) —
                    // emitted from the PARENT's own listing, which already has these fields for free.
                    // Emitted unconditionally, before descending, regardless of whether the recursive
                    // call below finds this folder empty or not.
                    await writer.WriteAsync(new SourceFileEntry(
                        driveId, item.Id, item.Name, childPath, item.LastModifiedDateTime,
                        IsFolder: true, Created: item.CreatedDateTime,
                        CreatedByEmail: GetIdentityEmail(item.CreatedBy?.User),
                        ModifiedByEmail: GetIdentityEmail(item.LastModifiedBy?.User)), ct);
                    subfolderWalks.Add(WalkFilesForCopyAsync(driveId, item.Id, childPath, controller, writer, ct));
                }
            }
            else
                await writer.WriteAsync(new SourceFileEntry(
                    driveId, item.Id, item.Name, childPath, item.LastModifiedDateTime, item.Size), ct);
        }
        await Task.WhenAll(subfolderWalks);
    }

    // ── Enumerate all sub-folders under a folder ──────────────────────────────

    public async IAsyncEnumerable<(string driveId, string itemId, string relativePath)>
        EnumerateFoldersAsync(string driveId, string rootItemId, string basePath = "",
            [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
    {
        // Previously took no CancellationToken at all: cancelling mid-walk on a large (3,000+
        // folder) library left the button live but nothing actually stopped until the whole
        // recursive walk finished on its own.
        var children = await GetChildrenAsync(driveId, rootItemId, string.Empty);
        foreach (var child in children.Where(c => c.Type == NodeType.Folder))
        {
            ct.ThrowIfCancellationRequested();
            var childPath = string.IsNullOrEmpty(basePath) ? child.Name : $"{basePath}/{child.Name}";
            yield return (driveId, child.Id, childPath);
            await foreach (var item in EnumerateFoldersAsync(driveId, child.Id, childPath, ct))
                yield return item;
        }
    }

    // ── Enumerate all files with metadata, for independent post-copy verification ──────────
    // Separate from EnumerateFilesAsync/GetChildrenAsync (used by the copy engine) so a
    // verification re-scan never shares state/behavior with the copy path, and so its Graph
    // payload can be trimmed with an explicit $select for 100k+ file scale.
    //
    // Sibling subfolders are walked CONCURRENTLY (fanned out as tasks, funneled through a
    // channel), not one at a time — a plain recursive `await foreach` would serialize the entire
    // folder-tree walk to one Graph call in flight at a time regardless of how much concurrency
    // the controller allows, since each subfolder's whole subtree would have to finish before its
    // sibling starts. The controller still bounds how many of those concurrent walks are actually
    // making Graph calls at once (extra tasks just queue on controller.WaitAsync), so this doesn't
    // bypass throttle protection — it just lets the walk actually use the concurrency budget the
    // controller already grants, the same way the file-copy pipeline does with Parallel.ForEachAsync.
    internal async IAsyncEnumerable<ScannedFile> EnumerateFilesWithMetadataAsync(
        string driveId, string rootItemId, string basePath,
        AdaptiveParallelismController controller,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
    {
        var channel = System.Threading.Channels.Channel.CreateUnbounded<ScannedFile>();
        var walkTask = Task.Run(async () =>
        {
            try { await WalkFilesWithMetadataAsync(driveId, rootItemId, basePath, controller, channel.Writer, ct); }
            finally { channel.Writer.Complete(); }
        }, ct);

        await foreach (var f in channel.Reader.ReadAllAsync(ct))
            yield return f;

        await walkTask; // propagate any exception from the walk (channel completion alone would swallow it)
    }

    private async Task WalkFilesWithMetadataAsync(
        string driveId, string itemId, string basePath, AdaptiveParallelismController controller,
        System.Threading.Channels.ChannelWriter<ScannedFile> writer, CancellationToken ct)
    {
        List<DriveItem> items;
        await controller.WaitAsync(ct);
        try { items = await GetChildrenWithMetadataAsync(driveId, itemId, ct); }
        finally { controller.Release(); }

        var subfolderWalks = new List<Task>();
        foreach (var item in items)
        {
            ct.ThrowIfCancellationRequested();
            if (item.Id == null || item.Name == null) continue;
            var childPath = string.IsNullOrEmpty(basePath) ? item.Name : $"{basePath}/{item.Name}";
            if (IsFolderLike(item))
            {
                subfolderWalks.Add(WalkFilesWithMetadataAsync(driveId, item.Id, childPath, controller, writer, ct));
            }
            else
            {
                await writer.WriteAsync(new ScannedFile(driveId, item.Id, item.Name, childPath,
                    item.LastModifiedDateTime, item.File?.Hashes?.QuickXorHash, item.Size), ct);
            }
        }
        await Task.WhenAll(subfolderWalks);
    }

    // Read-only counterpart to GetOrCreateFolderPathAsync: walks an existing path segment-by-segment
    // (same per-segment escaping, never creates anything) to resolve the item id a relative path
    // points at — used by verification to navigate to the actual copied folder/file under a
    // library root, rather than assuming the path IS the item id. Returns null if any segment is
    // missing (e.g. the folder was deleted or renamed since the original copy).
    internal async Task<string?> ResolveItemIdByPathAsync(string driveId, string parentItemId, string relativePath)
    {
        var current = parentItemId;
        foreach (var part in relativePath.Split('/').Where(p => !string.IsNullOrEmpty(p)))
        {
            try
            {
                var item = await Graph.Drives[driveId].Items[current]
                    .ItemWithPath(Uri.EscapeDataString(part))
                    .GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
                if (item?.Id == null) return null;
                current = item.Id;
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
            {
                return null;
            }
        }
        return current;
    }

    // Fetches a single item directly (no recursion) for verification of a plain single-file root.
    internal async Task<ScannedFile?> GetFileForVerificationAsync(string driveId, string itemId, string relativePath)
    {
        var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "name", "file", "size", "lastModifiedDateTime"]);
        if (item?.Id == null || item.Name == null || item.File == null) return null;
        return new ScannedFile(driveId, item.Id, item.Name, relativePath, item.LastModifiedDateTime,
            item.File.Hashes?.QuickXorHash, item.Size);
    }

    // $select keeps the payload small at 100k+ files: only fields the comparison needs. The "file"
    // facet is requested for identity/type anyway, and Graph returns its Hashes sub-object (including
    // QuickXorHash) as part of that same facet at no extra cost — no additional field or round-trip
    // needed. See ComparisonStatus for how the hash and date fields are actually used: a genuine
    // content-hash comparison for most file types, and a date comparison for Office Open XML formats
    // specifically, since those get re-serialized by SharePoint's backend independently of content
    // changes (making size/hash unreliable for them, but not for anything else).
    // Document Set base content-type id. Document Sets present as ordinary folders (folder facet,
    // no package facet) but must be treated as special containers just like OneNote notebooks —
    // and their content-type id always begins with this prefix regardless of any custom Document
    // Set type derived from it. See IsSpecialContainer.
    private const string DocumentSetContentTypeIdPrefix = "0x0120D520";

    // Hidden list-item fields backing SharePoint's "Customize folder > color" feature.
    //
    // `_ColorHex` ({3BDAB9AC-9E5D-44D4-BDE9-13B37E170618}, Hidden) is the value holder: despite the
    // name it is Type="Text" holding a palette INDEX as a numeric string, "0".."15" (0/empty =
    // yellow, 1 = dark red, … 15 = light pink); out-of-range values fall back to the default.
    // `_ColorTag` ({76D13CD2-1BAE-45A5-8B74-545B87B65037}) is a related system column. Both are
    // ReadOnlyField="TRUE" on current SharePoint Online, so they are READ here but never written —
    // color is written through the supported /_api/foldercoloring/stampcolor endpoint instead (see
    // StampFolderColorAsync). Neither name is verified against this tenant, so every read is built to
    // be a no-op on failure rather than a regression:
    //   • the REST fast-path read falls back to the same call without the color fields — a SharePoint
    //     $select HARD-FAILS on an unknown column, and swallowing that would silently disable the
    //     whole folder fast-path optimization
    //   • the Graph read is best-effort in a try/catch; Graph's fieldValueSet is an OData open type,
    //     so an unknown name there is silently omitted rather than an error
    //   • color failures are reported separately from metadata failures so they don't pollute the
    //     "could not correct metadata" diagnostics
    //
    // NAMING: OData forbids a leading '_' in a property name, so SharePoint surfaces these as
    // `OData__ColorHex` in REST/Graph $select and JSON payloads, while CSOM SetFieldValue and
    // ValidateUpdateListItem's FieldName take the raw internal name. Both forms are needed.
    private const string FolderColorTagFieldName = "_ColorTag";
    private const string FolderColorHexFieldName = "_ColorHex";
    private const string FolderColorTagODataName = "OData__ColorTag";
    private const string FolderColorHexODataName = "OData__ColorHex";

    // Whether a folder-like DriveItem is a special container that must NOT be walked into as a
    // plain folder (its contents are handled by a server-side special-folder copy instead — see
    // SourceFileEntry.IsSpecialFolder / CopySpecialFolderAsync). Detected entirely from the fields
    // already returned by the children listing below — NO per-folder round-trip:
    //   • OneNote notebooks (and similar bundles) come back with a "package" facet.
    //   • Document Sets come back as folders carrying the Document Set content type, surfaced via
    //     the expanded listItem.contentType.
    // This replaces the old per-folder GetFolderProgIdAsync probe, which cost TWO extra round-trips
    // (a Graph sharepointIds call + a SharePoint REST ProgID read) on EVERY folder in the tree just
    // to find the rare special one — the dominant hidden cost of the source scan. NOTE: an exotic
    // folder carrying some OTHER non-empty ProgID (neither a package nor a Document Set) would no
    // longer be flagged special and would be walked as a plain folder. OneNote and Document Sets are
    // the only special containers this app has ever handled; if a tenant turns out to use another,
    // this is the spot to extend. Needs a live-tenant sanity check on a library that has Document Sets.
    private static bool IsSpecialContainer(DriveItem item)
    {
        if (item.Package != null) return true;
        var ctId = item.ListItem?.ContentType?.Id;
        return ctId != null && ctId.StartsWith(DocumentSetContentTypeIdPrefix, StringComparison.OrdinalIgnoreCase);
    }

    // Reads the two folder-color fields (see FolderColorTagFieldName) off an already-fetched
    // DriveItem's expanded listItem.fields — returns (null, null) for a plain, uncolored folder or a
    // tenant where the field doesn't exist, same "absent means no value, not an error" contract as
    // GetFolderProgIdAsync. Checks BOTH the OData-prefixed and raw internal names because it's
    // unconfirmed which form Graph's fields bag keys these on, and tries a case-insensitive sweep as
    // a last resort (AdditionalData uses an ordinal-comparer Dictionary). `_ColorHex` is an integer,
    // so values are normalized via ToString() rather than assuming a JSON string.
    private static (string? colorTag, string? colorHex) GetFolderColorValues(DriveItem? item)
    {
        var additional = item?.ListItem?.Fields?.AdditionalData;
        if (additional == null) return (null, null);

        string? Read(string odataName, string rawName)
        {
            if (additional.TryGetValue(odataName, out var v) && v != null) return v.ToString();
            if (additional.TryGetValue(rawName, out var v2) && v2 != null) return v2.ToString();
            foreach (var kv in additional)
                if ((string.Equals(kv.Key, odataName, StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(kv.Key, rawName, StringComparison.OrdinalIgnoreCase)) && kv.Value != null)
                    return kv.Value.ToString();
            return null;
        }

        var colorTag = Read(FolderColorTagODataName, FolderColorTagFieldName);
        var colorHex = Read(FolderColorHexODataName, FolderColorHexFieldName);
        return (string.IsNullOrEmpty(colorTag) ? null : colorTag, string.IsNullOrEmpty(colorHex) ? null : colorHex);
    }

    // Cheap special-container check for a folder that is itself the ROOT of a user's directly
    // selected copy source — the one case IsSpecialContainer's children-listing shortcut can't
    // cover, since the walk starts AT this item rather than encountering it as someone's child (see
    // CopyService.cs's scan loop). Previously this used GetFolderProgIdAsync unconditionally on
    // every folder-scope root — a REST round-trip that ran even for the overwhelming majority of
    // ordinary (non-special) roots, purely to learn "no." A single Graph GET for just the package
    // facet + listItem content type gives the same yes/no answer via IsSpecialContainer's existing
    // logic, at the same cost as the per-child check already used during the walk.
    internal async Task<bool> IsRootFolderSpecialContainerAsync(string driveId, string itemId)
    {
        var item = await Graph.Drives[driveId].Items[itemId].GetAsync(cfg =>
        {
            cfg.QueryParameters.Select = ["package"];
            cfg.QueryParameters.Expand = ["listItem($select=contentType)"];
        });
        return item != null && IsSpecialContainer(item);
    }

    private async Task<List<DriveItem>> GetChildrenWithMetadataAsync(
        string driveId, string itemId, CancellationToken ct)
    {
        var resolvedId = itemId == "root" ? "root" : itemId;
        var page = await Graph.Drives[driveId].Items[resolvedId].Children.GetAsync(cfg =>
        {
            cfg.QueryParameters.Top    = 1000;
            // createdDateTime/createdBy/lastModifiedBy ride along on this SAME listing call, already
            // paid for — this is what lets WalkFilesForCopyAsync emit a metadata-carrying entry for
            // every ordinary folder (see SourceFileEntry.IsFolder) with ZERO extra round trips,
            // eliminating both the separate per-folder metadata fetch this walk used to require
            // downstream and the ancestor "hopsUp" derivation SPMI used when no file lived directly
            // in a folder.
            cfg.QueryParameters.Select = ["id", "name", "file", "folder", "package", "size",
                "lastModifiedDateTime", "createdDateTime", "createdBy", "lastModifiedBy"];
            // Expand only the listItem's content type so Document Sets can be told apart from plain
            // folders without a per-folder round-trip (see IsSpecialContainer). If a tenant/library
            // doesn't return it, contentType is simply null and detection falls back to the package
            // facet — no failure, just the pre-existing OneNote-only behavior.
            cfg.QueryParameters.Expand = ["listItem($select=contentType)"];
        }, ct);

        var items = new List<DriveItem>();
        while (page != null)
        {
            items.AddRange(page.Value ?? []);
            if (page.OdataNextLink == null) break;
            page = await Graph.Drives[driveId].Items[resolvedId].Children
                .WithUrl(page.OdataNextLink).GetAsync(cancellationToken: ct);
        }
        return items;
    }

}
