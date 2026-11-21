using SharePointSmartCopy.Localization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Reads and recreates SharePoint document library structure (columns, versioning) across sites.
public class LibraryCopyService(SharePointService spService)
{
    // Reads all metadata needed to recreate a library at another site.
    public async Task<LibraryDefinition> ReadLibraryDefinitionAsync(string sourceSiteUrl, string sourceDriveId)
    {
        var serverRelUrl = await spService.GetLibraryServerRelativeUrlAsync(sourceDriveId);
        var listId       = await spService.GetListIdByServerRelativeUrlAsync(sourceSiteUrl, serverRelUrl);
        var definition   = await ReadListMetadataAsync(sourceSiteUrl, listId);
        definition.SourceDriveId = sourceDriveId;
        definition.SourceSiteUrl = sourceSiteUrl;
        definition.SourceListId  = listId;
        definition.Columns       = await spService.GetLibraryColumnsAsync(sourceSiteUrl, listId);
        return definition;
    }

    // Reads all library definitions for all document libraries on a site (parallel).
    // Also includes system libraries (Site Assets, Style Library) that the Graph Drives
    // API does not enumerate, fetching them by title via SharePoint REST.
    public async Task<List<LibraryDefinition>> ReadAllLibraryDefinitionsAsync(string sourceSiteId, string sourceSiteUrl)
    {
        var libraries = await spService.GetLibrariesAsync(sourceSiteId, sourceSiteUrl);

        // Append known system libraries missing from the Drives API, deduplicating by drive ID.
        string[] systemLibraryTitles = ["Site Assets", "Style Library"];
        foreach (var title in systemLibraryTitles)
        {
            var node = await spService.GetLibraryNodeByTitleAsync(sourceSiteId, sourceSiteUrl, title);
            if (node != null && !libraries.Any(l => l.DriveId == node.DriveId))
                libraries.Add(node);
        }

        var tasks   = libraries.Select(lib => ReadLibraryDefinitionAsync(sourceSiteUrl, lib.DriveId));
        var results = await Task.WhenAll(tasks);
        return [.. results];
    }

    // The title of the target document library that is this one: same title, else same URL name. Null when there is none
    // (or it cannot be told, in which case creating is still attempted and an "already exists" error is handled as before).
    private async Task<string?> FindExistingLibraryTitleAsync(string baseUrl, string targetSiteUrl, LibraryDefinition definition)
    {
        try
        {
            var url = $"{baseUrl}/_api/web/lists?$select=Id,Title,RootFolder/Name&$expand=RootFolder&$filter=BaseTemplate eq 101";
            using var resp = await spService.SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, targetSiteUrl);
            if (!resp.IsSuccessStatusCode) return null;
            using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
            if (!doc.RootElement.TryGetProperty("value", out var lists)) return null;
            string? byUrl = null;
            foreach (var l in lists.EnumerateArray())
            {
                var title = l.TryGetProperty("Title", out var t) ? t.GetString() : null;
                if (string.IsNullOrEmpty(title)) continue;
                if (title.Equals(definition.Title, StringComparison.OrdinalIgnoreCase)) return title;
                // A copy into the SAME site under a new title carries the source's URL name, so the source library itself would
                // match by URL name and the copy be skipped as "already exists". The source is never the library it is copied to.
                var isSourceLibrary = definition.SourceListId.Length > 0 && l.TryGetProperty("Id", out var lid) &&
                    string.Equals(lid.GetString()?.Trim('{', '}'), definition.SourceListId.Trim('{', '}'), StringComparison.OrdinalIgnoreCase);
                if (byUrl == null && !isSourceLibrary && definition.RootFolderName.Length > 0 &&
                    l.TryGetProperty("RootFolder", out var rf) && rf.ValueKind == JsonValueKind.Object &&
                    rf.TryGetProperty("Name", out var rn) &&
                    string.Equals(rn.GetString(), definition.RootFolderName, StringComparison.OrdinalIgnoreCase))
                    byUrl = title;
            }
            return byUrl;
        }
        catch (Exception ex) when (ex is not OperationCanceledException) { return null; }
    }

    // The identifiers of a target library that already exists, as the exception callers use to carry on with it.
    private async Task<LibraryAlreadyExistsException> ResolveExistingLibraryAsync(string targetSiteId, string targetSiteUrl, string title)
    {
        // Tier 1: Graph Drives API (normal libraries)
        var existing = (await spService.GetLibrariesAsync(targetSiteId, targetSiteUrl))
            .FirstOrDefault(d => d.Name.Equals(title, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            var existServerRel = await spService.GetLibraryServerRelativeUrlAsync(existing.DriveId);
            return new LibraryAlreadyExistsException(existing.DriveId, existServerRel, listId: null);
        }

        // Tier 2: REST lookup: handles system libraries (Site Assets, Style Library) not in the Drives API
        var node = await spService.GetLibraryNodeByTitleAsync(targetSiteId, targetSiteUrl, title);
        if (node != null)
        {
            var serverRelUrl = node.ServerRelativePath
                ?? (string.IsNullOrEmpty(node.DriveId) ? null
                    : await spService.GetLibraryServerRelativeUrlAsync(node.DriveId));
            return new LibraryAlreadyExistsException(node.DriveId, serverRelUrl, listId: null);
        }

        // Exists but unresolvable (e.g. Style Library with no Graph drive)
        return new LibraryAlreadyExistsException(driveId: null, serverRelativeUrl: null, listId: null);
    }

    // Creates a new document library at the target site matching the given definition.
    // Returns the new library's Graph drive ID and server-relative URL.
    // If a library with the same name already exists, throws LibraryAlreadyExistsException
    // carrying the existing library's identifiers so the caller can reuse it.
    public async Task<(string newDriveId, string newServerRelativeUrl)> CreateLibraryAsync(
        string targetSiteUrl, string targetSiteId,
        LibraryDefinition definition,
        IEnumerable<ColumnMapping>? columnMappings = null)
    {
        var baseUrl = targetSiteUrl.TrimEnd('/');

        // Look before creating, so running the copy again adds to the library that is already there instead of making
        // "Site Assets (2)". Matched by title, or by URL name when the title differs (a renamed or localized library).
        var existingTitle = await FindExistingLibraryTitleAsync(baseUrl, targetSiteUrl, definition);
        if (existingTitle != null)
            throw await ResolveExistingLibraryAsync(targetSiteId, targetSiteUrl, existingTitle);

        // Step 1: create the list
        var createBody = JsonSerializer.Serialize(new
        {
            BaseTemplate = 101,
            Title        = definition.Title,
            Description  = definition.Description,
        });

        using var createResp = await spService.SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/_api/web/lists");
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(createBody, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, targetSiteUrl);

        var createRespBody = await createResp.Content.ReadAsStringAsync();
        if (!createResp.IsSuccessStatusCode)
        {
            if (IsAlreadyExistsError(createRespBody))
                throw await ResolveExistingLibraryAsync(targetSiteId, targetSiteUrl, definition.Title);
            throw new Exception(Loc.T("Svc_Lib_CreateLibHttp", definition.Title, (int)createResp.StatusCode, createRespBody[..Math.Min(200, createRespBody.Length)]));
        }

        using var createDoc = JsonDocument.Parse(createRespBody);
        var listId = createDoc.RootElement.GetProperty("Id").GetString()!;

        // Step 2: set versioning settings
        if (definition.EnableVersioning)
        {
            var versionBody = JsonSerializer.Serialize(new
            {
                EnableVersioning    = true,
                EnableMinorVersions = definition.EnableMinorVersions,
                MajorVersionLimit   = definition.MajorVersionLimit > 0 ? definition.MajorVersionLimit : 500,
            });
            using var versionResp = await spService.SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/_api/web/lists('{listId}')");
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Headers.TryAddWithoutValidation("IF-MATCH", "*");
                r.Headers.TryAddWithoutValidation("X-HTTP-Method", "MERGE");
                r.Content = new StringContent(versionBody, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, targetSiteUrl);

            if (!versionResp.IsSuccessStatusCode &&
                versionResp.StatusCode != System.Net.HttpStatusCode.NoContent)
            {
                var errBody = await versionResp.Content.ReadAsStringAsync();
                throw new Exception(Loc.T("Svc_Lib_VersioningHttp", definition.Title, (int)versionResp.StatusCode, errBody[..Math.Min(200, errBody.Length)]));
            }
        }

        // Step 3: create custom columns.
        // Distinguish three cases per column:
        //  - dialog never opened at all (no mapping entries) → create everything;
        //  - column never SHOWN in the dialog (mappings were loaded from a different library in a
        //    Site/multi-library run) → create it, the user made no decision about it — the old
        //    logic silently created NO columns for every library after the first;
        //  - column shown and decided → honor the decision, including an explicit all-skip (which
        //    the old "empty dict means never opened" check inverted into create-all).
        var decidedColumns = columnMappings?
            .Select(m => m.SourceColumn.InternalName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var mappingEntries = columnMappings?
            .Where(m => m.TargetColumn != null || m.CreateNew)
            .ToDictionary(m => m.SourceColumn.InternalName, StringComparer.OrdinalIgnoreCase);

        var columnsToCreate = definition.Columns.Where(col =>
        {
            if (decidedColumns == null || decidedColumns.Count == 0) return true;
            if (!decidedColumns.Contains(col.InternalName)) return true;
            return mappingEntries != null &&
                   mappingEntries.TryGetValue(col.InternalName, out var mapping) && mapping.CreateNew;
        }).ToList();

        var columnFailures = new List<string>();
        foreach (var col in columnsToCreate)
        {
            try { await CreateColumnAsync(targetSiteUrl, listId, col, definition.SourceSiteUrl); }
            catch (Exception ex)
            {
                columnFailures.Add($"{col.DisplayName} ({ex.Message})");
                System.Diagnostics.Debug.WriteLine($"[LibraryCopy] Column '{col.InternalName}' create failed: {ex.Message}");
            }
        }
        LastColumnCreationFailures = columnFailures;

        // Views, after the columns they show exist.
        await CopyListStructureAsync(definition, targetSiteUrl, listId, isLibrary: true);

        // Step 4: discover new drive ID (retry — Graph lags ~2s after list provisioning)
        string newDriveId      = string.Empty;
        string newServerRelUrl = string.Empty;
        for (int attempt = 0; attempt < 3; attempt++)
        {
            if (attempt > 0) await Task.Delay(2000);
            var drives = await spService.GetLibrariesAsync(targetSiteId, targetSiteUrl);
            var match  = drives.FirstOrDefault(d =>
                d.Name.Equals(definition.Title, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                newDriveId      = match.DriveId;
                newServerRelUrl = await spService.GetLibraryServerRelativeUrlAsync(match.DriveId);
                break;
            }
        }

        if (string.IsNullOrEmpty(newDriveId))
        {
            // Tier 2: REST-based lookup — catches libraries not yet visible in the Graph Drives API
            // (system libraries and slow-propagating provisioning are common culprits).
            var node = await spService.GetLibraryNodeByTitleAsync(targetSiteId, targetSiteUrl, definition.Title);
            if (node != null && !string.IsNullOrEmpty(node.DriveId))
            {
                newDriveId      = node.DriveId;
                newServerRelUrl = node.ServerRelativePath
                    ?? await spService.GetLibraryServerRelativeUrlAsync(node.DriveId);
            }
        }

        if (string.IsNullOrEmpty(newDriveId))
            throw new Exception(Loc.T("Svc_Lib_NoDriveId", definition.Title));

        return (newDriveId, newServerRelUrl);
    }

    // Reads the schema for a custom list (non-document-library) identified by list GUID and title.
    // The title parameter is used as-is so locale variants don't overwrite the source display name.
    public async Task<LibraryDefinition> ReadListDefinitionAsync(
        string sourceSiteUrl, string listId, string title)
    {
        var definition       = await ReadListMetadataAsync(sourceSiteUrl, listId);
        definition.Title     = title;
        definition.SourceSiteUrl = sourceSiteUrl;
        definition.SourceListId  = listId;
        definition.Columns   = await spService.GetLibraryColumnsAsync(sourceSiteUrl, listId);
        return definition;
    }

    // Creates a generic (non-document-library) list at the target site and returns its GUID.
    // If a list with the same title already exists, throws LibraryAlreadyExistsException
    // carrying the existing list's ID so the caller can reuse it.
    public async Task<string> CreateCustomListAsync(
        string targetSiteUrl, string targetSiteId,
        LibraryDefinition definition, int baseTemplate)
    {
        var baseUrl = targetSiteUrl.TrimEnd('/');
        var createBody = JsonSerializer.Serialize(new
        {
            BaseTemplate = baseTemplate,
            Title        = definition.Title,
            Description  = definition.Description,
        });

        using var createResp = await spService.SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/_api/web/lists");
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(createBody, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, targetSiteUrl);

        string listId;
        var createRespBody = await createResp.Content.ReadAsStringAsync();

        if (!createResp.IsSuccessStatusCode)
        {
            if (IsAlreadyExistsError(createRespBody))
            {
                var existingId = await spService.GetListIdByTitleAsync(targetSiteUrl, definition.Title);
                throw new LibraryAlreadyExistsException(driveId: null, serverRelativeUrl: null, listId: existingId);
            }
            throw new Exception(Loc.T("Svc_Lib_CreateListHttp", definition.Title, (int)createResp.StatusCode, createRespBody[..Math.Min(200, createRespBody.Length)]));
        }
        else
        {
            using var doc = JsonDocument.Parse(createRespBody);
            listId = doc.RootElement.GetProperty("Id").GetString()!;
        }

        var existingCols          = await spService.GetLibraryColumnsAsync(targetSiteUrl, listId, skipCache: true);
        var existingInternalNames = existingCols.Select(c => c.InternalName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var existingDisplayNames  = existingCols.Select(c => c.DisplayName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var listColumnFailures = new List<string>();
        foreach (var col in definition.Columns.Where(c =>
            !existingInternalNames.Contains(c.InternalName) &&
            !existingDisplayNames.Contains(c.DisplayName)))
        {
            try { await CreateColumnAsync(targetSiteUrl, listId, col, definition.SourceSiteUrl); }
            catch (Exception ex) { listColumnFailures.Add($"{col.DisplayName} ({ex.Message})"); }
        }
        LastColumnCreationFailures = listColumnFailures;

        await CopyListStructureAsync(definition, targetSiteUrl, listId, isLibrary: baseTemplate == 101);

        return listId;
    }

    // ── Views ────────────────────────────────────────────────────────────────

    // Notes about views that could not be copied (or only partly), by list/library title. Read by the callers, which put
    // them in the run report next to the column failures.
    private readonly Dictionary<string, List<string>> _viewNotes = new(StringComparer.OrdinalIgnoreCase);
    public List<string> ViewFailuresFor(string title) => _viewNotes.TryGetValue(title, out var l) ? l : [];

    // Sends one SharePoint REST call; returns whether it succeeded and the body. Never throws on an HTTP error.
    private async Task<(bool Ok, string Body)> SpRestAsync(HttpMethod method, string url, string siteUrl, string? jsonBody = null, bool merge = false, string? xMethod = null)
    {
        using var resp = await spService.SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(method, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            if (merge || xMethod != null)
            {
                r.Headers.TryAddWithoutValidation("IF-MATCH", "*");
                r.Headers.TryAddWithoutValidation("X-HTTP-Method", xMethod ?? "MERGE");
            }
            if (jsonBody != null) r.Content = new StringContent(jsonBody, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl);
        var body = await resp.Content.ReadAsStringAsync();
        return (resp.IsSuccessStatusCode, body);
    }

    private static string Short(string s) => s.Length > 160 ? s[..160] : s;

    private static List<string> ReadStrings(JsonElement root)
    {
        var list = new List<string>();
        foreach (var name in new[] { "Items", "value" })
            if (root.TryGetProperty(name, out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var e in arr.EnumerateArray())
                    if (e.ValueKind == JsonValueKind.String && e.GetString() is { Length: > 0 } s) list.Add(s);
                break;
            }
        return list;
    }

    // Recreates the source list's shared views (filter, sorting, grouping, columns and their order, paging, default view) on a
    // list that was just created. Personal, hidden and non-standard (calendar, board, Gantt) views are skipped. A view whose
    // title the new list already has (its built-in "All Items"/"All Documents") is updated instead of duplicated. Columns the
    // target does not have are left out of the view. Never throws: problems become notes for the report.
    // Everything about a new list's structure beyond its columns: settings, column rules and formatting, content types,
    // then views. Each part is best effort; what could not be done becomes a note (read with ViewFailuresFor).
    // For a list or library that already exists on the target: only ADD what is missing (content types the list lacks, views
    // whose title it does not have). Existing views, settings, column rules, the template and default values are left alone,
    // unless Structure.UpdateExisting is on: then same-named views, the default view, list settings and column rules are replaced.
    public Task AddMissingStructureAsync(LibraryDefinition def, string targetSiteUrl, string targetListId, bool isLibrary) =>
        CopyListStructureAsync(def, targetSiteUrl, targetListId, isLibrary, additive: true);

    private async Task CopyListStructureAsync(LibraryDefinition def, string targetSiteUrl, string targetListId, bool isLibrary, bool additive = false)
    {
        var notes = new List<string>();
        _viewNotes[def.Title] = notes;
        if (string.IsNullOrEmpty(def.SourceSiteUrl) || string.IsNullOrEmpty(def.SourceListId)) return;
        var srcSite = def.SourceSiteUrl.TrimEnd('/');
        var srcBase = $"{srcSite}/_api/web/lists('{def.SourceListId}')";
        var tgtSite = targetSiteUrl.TrimEnd('/');
        var tgtBase = $"{tgtSite}/_api/web/lists('{targetListId}')";
        // Order matters: content types before column rules (columns may be recreated to put them on a type, which would lose a
        // rule applied earlier), and both before the views (which list the columns).
        var steps = new List<Func<Task>>();
        var update = additive && Structure.UpdateExisting;   // existing list, and the user asked for it to be brought in line
        if ((!additive || update) && Structure.ListSettings) steps.Add(() => CopyListSettingsAsync(def.Title, srcBase, srcSite, tgtBase, tgtSite, isLibrary, notes));
        if (Structure.ContentTypes) steps.Add(() => CopyContentTypesAsync(srcBase, srcSite, tgtBase, tgtSite, notes, recreateColumns: !additive, isLibrary: isLibrary));
        steps.Add(() => CopyProjectedFieldsAsync(srcBase, srcSite, tgtBase, tgtSite, notes));
        if ((!additive || update) && Structure.ColumnRules) steps.Add(() => CopyColumnRulesAsync(srcBase, srcSite, tgtBase, tgtSite, notes));
        if (Structure.Views) steps.Add(() => CopyViewsAsync(def, targetSiteUrl, targetListId, notes, additive, update));
        if (isLibrary && !additive && Structure.Templates)
        {
            steps.Add(() => CopyDocumentTemplateAsync(srcBase, srcSite, tgtBase, tgtSite, notes));
            steps.Add(() => CopyFolderDefaultsAsync(srcBase, srcSite, tgtBase, tgtSite, notes));
        }
        foreach (var step in steps)
        {
            try { await step(); }
            catch (Exception ex) when (ex is not OperationCanceledException) { notes.Add(Short(ex.Message)); }
        }
    }

    // Projected (additional) lookup fields: the extra columns a lookup column shows from the looked-up list. They are
    // read-only fields that point at their primary lookup column (FieldRef); at the target each one is added with
    // addDependentLookupField once the primary lookup column exists. One that is already there is left alone.
    private async Task CopyProjectedFieldsAsync(string srcBase, string srcSite, string tgtBase, string tgtSite, List<string> notes)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{srcBase}/fields?$select=Id,InternalName,Title,SchemaXml", srcSite);
        if (!ok) return;
        var srcFields = new List<(string Id, string Name, string Title, string Xml)>();
        using (var doc = JsonDocument.Parse(body))
            foreach (var f in doc.RootElement.GetProperty("value").EnumerateArray())
                srcFields.Add((f.GetProperty("Id").GetString() ?? "", f.GetProperty("InternalName").GetString() ?? "",
                               f.GetProperty("Title").GetString() ?? "", f.TryGetProperty("SchemaXml", out var x) ? x.GetString() ?? "" : ""));

        var projected = new List<(string Title, string PrimaryName, string ShowField)>();
        foreach (var f in srcFields)
        {
            if (!f.Xml.Contains("FieldRef=", StringComparison.Ordinal) || !f.Xml.Contains("Type=\"Lookup", StringComparison.Ordinal)) continue;
            try
            {
                var el = System.Xml.Linq.XElement.Parse(f.Xml);
                var primaryId = ((string?)el.Attribute("FieldRef"))?.Trim('{', '}');
                var show = (string?)el.Attribute("ShowField");
                if (string.IsNullOrEmpty(primaryId) || string.IsNullOrEmpty(show)) continue;
                var primary = srcFields.FirstOrDefault(p => string.Equals(p.Id.Trim('{', '}'), primaryId, StringComparison.OrdinalIgnoreCase));
                if (primary.Name != null) projected.Add((f.Title, primary.Name, show));
            }
            catch (System.Xml.XmlException) { }
        }
        if (projected.Count == 0) return;

        var (tok, tbody) = await SpRestAsync(HttpMethod.Get, $"{tgtBase}/fields?$select=Id,InternalName,Title", tgtSite);
        if (!tok) return;
        var tgtFields = new List<(string Id, string Name, string Title)>();
        using (var doc = JsonDocument.Parse(tbody))
            foreach (var f in doc.RootElement.GetProperty("value").EnumerateArray())
                tgtFields.Add((f.GetProperty("Id").GetString() ?? "", f.GetProperty("InternalName").GetString() ?? "", f.GetProperty("Title").GetString() ?? ""));

        foreach (var (title, primaryName, showField) in projected)
        {
            if (tgtFields.Any(t => string.Equals(t.Title, title, StringComparison.OrdinalIgnoreCase))) continue;
            var primary = tgtFields.FirstOrDefault(t => string.Equals(t.Name, primaryName, StringComparison.OrdinalIgnoreCase));
            if (primary.Id == null) { notes.Add(Loc.T("Svc_Lib_ProjectedNoPrimary", title, primaryName)); continue; }
            var url = $"{tgtBase}/fields/addDependentLookupField(displayName='{Uri.EscapeDataString(title.Replace("'", "''"))}'," +
                      $"primaryLookupFieldId='{primary.Id.Trim('{', '}')}',showField='{Uri.EscapeDataString(showField.Replace("'", "''"))}')";
            var (aok, abody) = await SpRestAsync(HttpMethod.Post, url, tgtSite, "{}");
            if (!aok) notes.Add(Loc.T("Svc_Lib_ProjectedFailed", title, Short(abody)));
        }
    }

    // List-level settings: content approval, check-out, attachments, open in browser, folders, search, draft visibility,
    // item-level security and the list's validation rule. One call each, so a setting that does not apply to this kind of
    // list (attachments on a library) or that the target refuses never blocks the others.
    private async Task CopyListSettingsAsync(string title, string srcBase, string srcSite, string tgtBase, string tgtSite, bool isLibrary, List<string> notes)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, srcBase, srcSite);
        if (!ok) { notes.Add(Loc.T("Svc_Set_ReadFail", Short(body))); return; }
        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;

        var settings = new List<(string Name, object Value)>();
        // Content types first (the content type step needs them on), approval before draft visibility (which depends on it).
        string[] bools = isLibrary
            ? new[] { "EnableContentTypes", "EnableModeration", "ForceCheckout", "DefaultItemOpenInBrowser", "EnableFolderCreation", "NoCrawl" }
            : ["EnableContentTypes", "EnableModeration", "EnableAttachments", "EnableFolderCreation", "NoCrawl"];   // a list has no check-out (refused on a live site)
        foreach (var n in bools)
            if (root.TryGetProperty(n, out var p) && p.ValueKind is JsonValueKind.True or JsonValueKind.False)
                settings.Add((n, p.GetBoolean()));
        foreach (var n in new[] { "DraftVersionVisibility", "ReadSecurity", "WriteSecurity" })
            if (root.TryGetProperty(n, out var p) && p.TryGetInt32(out var i)) settings.Add((n, i));
        foreach (var n in new[] { "ValidationFormula", "ValidationMessage" })
            if (root.TryGetProperty(n, out var p) && p.ValueKind == JsonValueKind.String && !string.IsNullOrEmpty(p.GetString()))
                settings.Add((n, p.GetString()!));

        // Require check-out and content approval change how files are written (checked-out files, drafts pending approval), so
        // they are held back until the content has been copied (ApplyPendingSettingsAsync), not set on the empty new list.
        string[] deferredNames = ["EnableModeration", "ForceCheckout", "DraftVersionVisibility"];
        var deferred = settings.Where(s => deferredNames.Contains(s.Name) && !(s.Value is bool b && !b) && !(s.Value is int i && i == 0)).ToList();
        if (deferred.Count > 0) _pendingSettings[title] = (tgtBase, tgtSite, deferred);
        settings = settings.Where(s => !deferredNames.Contains(s.Name)).ToList();

        var failed = new List<string>();
        foreach (var (name, value) in settings)
        {
            var (sok, _) = await SpRestAsync(HttpMethod.Post, tgtBase, tgtSite, JsonSerializer.Serialize(new Dictionary<string, object> { [name] = value }), merge: true);
            if (!sok) failed.Add(name);
        }
        if (failed.Count > 0) notes.Add(Loc.T("Svc_Set_Failed", string.Join(", ", failed)));
    }

    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (string TgtBase, string TgtSite, List<(string Name, object Value)> Props)> _pendingSettings = new(StringComparer.OrdinalIgnoreCase);

    // Applies the settings that were held back until the content was copied (content approval, require check-out, draft
    // visibility). Returns one note per list whose settings could not all be applied.
    public async Task<List<string>> ApplyPendingSettingsAsync()
    {
        var notes = new List<string>();
        foreach (var title in _pendingSettings.Keys.ToList())
        {
            if (!_pendingSettings.TryRemove(title, out var p)) continue;
            var failed = new List<string>();
            foreach (var (name, value) in p.Props.OrderBy(x => x.Name == "EnableModeration" ? 0 : 1))   // approval before draft visibility
            {
                try
                {
                    var (ok, _) = await SpRestAsync(HttpMethod.Post, p.TgtBase, p.TgtSite, JsonSerializer.Serialize(new Dictionary<string, object> { [name] = value }), merge: true);
                    if (!ok) failed.Add(name);
                }
                catch (Exception ex) when (ex is not OperationCanceledException) { failed.Add(name); }
            }
            if (failed.Count > 0) notes.Add($"{title}: {Loc.T("Svc_Set_Failed", string.Join(", ", failed))}");
        }
        return notes;
    }

    // ── Site level (Site scope): site columns, site content types, navigation, site settings ───────────

    // Copies the site-wide definitions the lists build on, before any list is created: custom site columns and site content
    // types (so a list's types can reuse them), the left and top navigation, and the site's own settings. Never throws; returns
    // notes for the report. Site permissions are copied by the permission step the Site scope already has.
    public async Task<List<string>> CopySiteStructureAsync(string sourceSiteUrl, string targetSiteUrl, CancellationToken ct = default)
    {
        var notes = new List<string>();
        var src = sourceSiteUrl.TrimEnd('/');
        var tgt = targetSiteUrl.TrimEnd('/');
        foreach (var step in new Func<Task>[]
        {
            () => Structure.SiteDefinitions ? CopySiteColumnsAsync(src, tgt, notes) : Task.CompletedTask,
            () => Structure.SiteDefinitions ? CopySiteContentTypesAsync(src, tgt, notes) : Task.CompletedTask,
            () => Structure.Navigation ? CopyNavigationAsync(src, tgt, "quicklaunch", notes) : Task.CompletedTask,
            () => Structure.Navigation ? CopyNavigationAsync(src, tgt, "topnavigationbar", notes) : Task.CompletedTask,
            () => Structure.SiteSettings ? CopySiteSettingsAsync(src, tgt, notes) : Task.CompletedTask,
        })
        {
            ct.ThrowIfCancellationRequested();
            try { await step(); }
            catch (Exception ex) when (ex is not OperationCanceledException) { notes.Add(Short(ex.Message)); }
        }
        return notes;
    }

    private async Task CopySiteColumnsAsync(string src, string tgt, List<string> notes)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{src}/_api/web/fields?$filter=Hidden eq false&$select=InternalName,Title,Group,TypeAsString,SchemaXml", src);
        if (!ok) { notes.Add(Loc.T("Svc_Site_ColumnsReadFail", Short(body))); return; }
        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var (tok, tBody) = await SpRestAsync(HttpMethod.Get, $"{tgt}/_api/web/fields?$select=InternalName", tgt);
        if (tok)
            using (var td = JsonDocument.Parse(tBody))
                foreach (var f in td.RootElement.GetProperty("value").EnumerateArray())
                    existing.Add(f.GetProperty("InternalName").GetString() ?? "");

        var failed = new List<string>();
        using var doc = JsonDocument.Parse(body);
        foreach (var f in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            var internalName = f.GetProperty("InternalName").GetString() ?? "";
            var xml = f.TryGetProperty("SchemaXml", out var sx) ? sx.GetString() : null;
            var group = f.TryGetProperty("Group", out var g) ? g.GetString() ?? "" : "";
            var type = f.TryGetProperty("TypeAsString", out var ty) ? ty.GetString() ?? "" : "";
            if (internalName.Length == 0 || string.IsNullOrEmpty(xml) || existing.Contains(internalName) || group.StartsWith("_Hidden", StringComparison.OrdinalIgnoreCase)) continue;
            // A column made on this site carries the site's own ID as its source; the built-in ones carry Microsoft's schema address.
            if (!System.Text.RegularExpressions.Regex.IsMatch(xml, "SourceID=\"\\{[0-9a-fA-F-]{36}\\}\"")) continue;
            // A lookup points at a list of the source site; managed metadata ids belong to the source tenant's term store.
            if (type.StartsWith("Lookup", StringComparison.OrdinalIgnoreCase)) continue;   // created later, once the lists exist (CopySiteLookupColumnsAsync)
            if (spService.IsCrossTenant && type.StartsWith("TaxonomyField", StringComparison.OrdinalIgnoreCase)) continue;
            try
            {
                var el = System.Xml.Linq.XElement.Parse(xml);
                // The ID is kept: the site content types below link to their columns by it.
                foreach (var attr in new[] { "SourceID", "ColName", "RowOrdinal", "Version", "WebId" }) el.Attribute(attr)?.Remove();
                foreach (var p in el.Descendants("Property").Where(p => (string?)p.Element("Name") == "TextField").ToList()) p.Remove();
                var json = JsonSerializer.Serialize(new { parameters = new { SchemaXml = el.ToString(System.Xml.Linq.SaveOptions.DisableFormatting) } });
                var (cok, _) = await SpRestAsync(HttpMethod.Post, $"{tgt}/_api/web/fields/createfieldasxml", tgt, json);
                if (!cok) failed.Add(internalName);
            }
            catch (System.Xml.XmlException) { failed.Add(internalName); }
        }
        if (failed.Count > 0) notes.Add(Loc.T("Svc_Site_ColumnsFailed", string.Join(", ", failed)));
    }

    // Site columns of type Lookup / LookupMulti, created AFTER the lists of the copy exist: a lookup column points at a list, and
    // at the target that list is only there once the lists have been copied. The list is found by title (the source GUID means
    // nothing at the target). Same-web lookups only; one whose list is not on the target is listed in the returned notes.
    public async Task<List<string>> CopySiteLookupColumnsAsync(string sourceSiteUrl, string targetSiteUrl, CancellationToken ct = default)
    {
        var notes = new List<string>();
        if (!Structure.SiteDefinitions) return notes;
        var src = sourceSiteUrl.TrimEnd('/');
        var tgt = targetSiteUrl.TrimEnd('/');
        try
        {
            var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{src}/_api/web/fields?$filter=Hidden eq false&$select=InternalName,Group,TypeAsString,SchemaXml", src);
            if (!ok) return notes;
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var (tok, tBody) = await SpRestAsync(HttpMethod.Get, $"{tgt}/_api/web/fields?$select=InternalName", tgt);
            if (tok)
                using (var td = JsonDocument.Parse(tBody))
                    foreach (var f in td.RootElement.GetProperty("value").EnumerateArray())
                        existing.Add(f.GetProperty("InternalName").GetString() ?? "");
            var tgtWebId = await GetWebIdAsync(tgt);
            var failed = new List<string>();
            var noList = new List<string>();
            using var doc = JsonDocument.Parse(body);
            foreach (var f in doc.RootElement.GetProperty("value").EnumerateArray())
            {
                ct.ThrowIfCancellationRequested();
                var internalName = f.GetProperty("InternalName").GetString() ?? "";
                var xml = f.TryGetProperty("SchemaXml", out var sx) ? sx.GetString() : null;
                var group = f.TryGetProperty("Group", out var g) ? g.GetString() ?? "" : "";
                var type = f.TryGetProperty("TypeAsString", out var ty) ? ty.GetString() ?? "" : "";
                if (internalName.Length == 0 || string.IsNullOrEmpty(xml) || existing.Contains(internalName) || group.StartsWith("_Hidden", StringComparison.OrdinalIgnoreCase)) continue;
                if (!type.StartsWith("Lookup", StringComparison.OrdinalIgnoreCase)) continue;
                if (!System.Text.RegularExpressions.Regex.IsMatch(xml, "SourceID=\"\\{[0-9a-fA-F-]{36}\\}\"")) continue;   // made on this site only
                try
                {
                    var el = System.Xml.Linq.XElement.Parse(xml);
                    if (el.Attribute("FieldRef") != null) continue;   // projected field: added below, once its primary lookup exists
                    var srcListGuid = (el.Attribute("List")?.Value ?? "").Trim('{', '}');
                    var lookupWeb = (el.Attribute("WebId")?.Value ?? "").Trim('{', '}');
                    if (!Guid.TryParse(srcListGuid, out _)) { failed.Add(internalName); continue; }   // "Self": nothing to point at
                    var (tgtListId, tgtOtherWebId) = await ResolveTargetLookupAsync(tgt, src, srcListGuid, lookupWeb);
                    if (string.IsNullOrEmpty(tgtListId)) { noList.Add(internalName); continue; }
                    el.SetAttributeValue("List", $"{{{tgtListId.Trim('{', '}')}}}");
                    var webForField = tgtOtherWebId ?? tgtWebId;
                    if (webForField != null) el.SetAttributeValue("WebId", webForField);
                    foreach (var attr in new[] { "SourceID", "ColName", "RowOrdinal", "Version" }) el.Attribute(attr)?.Remove();
                    var json = JsonSerializer.Serialize(new { parameters = new { SchemaXml = el.ToString(System.Xml.Linq.SaveOptions.DisableFormatting) } });
                    var (cok, _) = await SpRestAsync(HttpMethod.Post, $"{tgt}/_api/web/fields/createfieldasxml", tgt, json);
                    if (!cok) failed.Add(internalName);
                }
                catch (System.Xml.XmlException) { failed.Add(internalName); }
            }
            // The extra columns a site lookup shows from its list: added to the site (web) columns now that the primaries exist.
            await CopyProjectedFieldsAsync($"{src}/_api/web", src, $"{tgt}/_api/web", tgt, notes);
            if (noList.Count > 0) notes.Add(Loc.T("Svc_Site_LookupListMissing", string.Join(", ", noList)));
            if (failed.Count > 0) notes.Add(Loc.T("Svc_Site_ColumnsFailed", string.Join(", ", failed)));
        }
        catch (Exception ex) when (ex is not OperationCanceledException) { notes.Add(Short(ex.Message)); }
        return notes;
    }

    private async Task CopySiteContentTypesAsync(string src, string tgt, List<string> notes)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{src}/_api/web/contenttypes?$select=Name,Description,Group,Hidden,StringId", src);
        if (!ok) { notes.Add(Loc.T("Svc_Site_TypesReadFail", Short(body))); return; }
        var (tok, tBody) = await SpRestAsync(HttpMethod.Get, $"{tgt}/_api/web/contenttypes?$select=Name", tgt);
        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (tok)
            using (var td = JsonDocument.Parse(tBody))
                foreach (var t in td.RootElement.GetProperty("value").EnumerateArray())
                    existing.Add(t.GetProperty("Name").GetString() ?? "");

        var failed = new List<string>();
        using var doc = JsonDocument.Parse(body);
        // Parents before the types that derive from them (a derived ID is longer than its parent's).
        var types = doc.RootElement.GetProperty("value").EnumerateArray()
            .Where(t => !Bool(t, "Hidden") && !(t.TryGetProperty("Group", out var gg) && (gg.GetString() ?? "").StartsWith("_Hidden", StringComparison.OrdinalIgnoreCase)))
            .Where(t => !existing.Contains(t.GetProperty("Name").GetString() ?? ""))
            .Where(t => System.Text.RegularExpressions.Regex.IsMatch(t.GetProperty("StringId").GetString() ?? "", "00[0-9A-Fa-f]{32}"))   // made on this site, not built in
            .OrderBy(t => (t.GetProperty("StringId").GetString() ?? "").Length)
            .ToList();
        foreach (var t in types)
        {
            var name = t.GetProperty("Name").GetString() ?? "";
            var id = t.GetProperty("StringId").GetString() ?? "";
            try
            {
                // SharePoint ignores an ID sent with the request and always derives a new type from Item, so the new type's own ID
                // is read from the response. Its columns cannot be linked over REST (every call shape was refused on a live site).
                var create = new Dictionary<string, object>
                {
                    ["Name"] = name,
                    ["Description"] = t.TryGetProperty("Description", out var d) ? d.GetString() ?? "" : "",
                    ["Group"] = t.TryGetProperty("Group", out var g) ? g.GetString() ?? "Custom Content Types" : "Custom Content Types",
                };
                var (cok, cBody) = await SpRestAsync(HttpMethod.Post, $"{tgt}/_api/web/contenttypes", tgt, JsonSerializer.Serialize(create));
                if (!cok) { failed.Add(name); continue; }
                string newId;
                using (var cd = JsonDocument.Parse(cBody)) newId = cd.RootElement.GetProperty("StringId").GetString()!;
            }
            catch (Exception ex) when (ex is not OperationCanceledException) { failed.Add(name); }
        }
        if (failed.Count > 0) notes.Add(Loc.T("Svc_Site_TypesFailed", string.Join(", ", failed)));
    }

    // Copies one navigation (the left "quicklaunch" or the top bar) onto the target: nodes whose title the target does not
    // already have, with their child links, and addresses into the source site pointed at the target site.
    private async Task CopyNavigationAsync(string src, string tgt, string which, List<string> notes)
    {
        var srcPath = new Uri(src).AbsolutePath.TrimEnd('/');
        var tgtPath = new Uri(tgt).AbsolutePath.TrimEnd('/');
        string Rewrite(string url)
        {
            if (url.StartsWith(src, StringComparison.OrdinalIgnoreCase)) return tgt + url[src.Length..];
            if (srcPath.Length > 0 && url.StartsWith(srcPath + "/", StringComparison.OrdinalIgnoreCase)) return tgtPath + url[srcPath.Length..];
            if (srcPath.Length > 0 && url.Equals(srcPath, StringComparison.OrdinalIgnoreCase)) return tgtPath.Length == 0 ? "/" : tgtPath;
            return url;
        }

        async Task<List<(string Id, string Title, string Url, bool External, bool Visible)>> ReadNodesAsync(string site, string path)
        {
            var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{site}/_api/web/navigation/{path}", site);
            var list = new List<(string, string, string, bool, bool)>();
            if (!ok) return list;
            using var d = JsonDocument.Parse(body);
            foreach (var n in d.RootElement.GetProperty("value").EnumerateArray())
                list.Add((n.GetProperty("Id").GetInt32().ToString(),
                          n.TryGetProperty("Title", out var t) ? t.GetString() ?? "" : "",
                          n.TryGetProperty("Url", out var u) ? u.GetString() ?? "" : "",
                          Bool(n, "IsExternal"),
                          !n.TryGetProperty("IsVisible", out var v) || v.ValueKind != JsonValueKind.False));
            return list;
        }

        var existingTop = (await ReadNodesAsync(tgt, which)).Select(n => n.Title).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var failed = new List<string>();
        foreach (var node in await ReadNodesAsync(src, which))
        {
            if (node.Title.Length == 0 || existingTop.Contains(node.Title)) continue;
            var created = await AddNavNodeAsync(tgt, $"navigation/{which}", node.Title, Rewrite(node.Url), node.External, node.Visible);
            if (created is null) { failed.Add(node.Title); continue; }
            // The "Recent" node's links are generated by SharePoint from the site's recently used lists and cannot be created.
            if (node.Title.Equals("Recent", StringComparison.OrdinalIgnoreCase)) continue;
            foreach (var child in await ReadNodesAsync(src, $"getnodebyid({node.Id})/children"))
                if (await AddNavNodeAsync(tgt, $"navigation/getnodebyid({created})/children", child.Title, Rewrite(child.Url), child.External, child.Visible) is null)
                    failed.Add($"{node.Title} / {child.Title}");
        }
        if (failed.Count > 0) notes.Add(Loc.T("Svc_Site_NavFailed", which, string.Join(", ", failed)));
    }

    private async Task<string?> AddNavNodeAsync(string site, string path, string title, string url, bool external, bool visible)
    {
        var json = JsonSerializer.Serialize(new Dictionary<string, object> { ["Title"] = title, ["Url"] = url, ["IsExternal"] = external, ["IsVisible"] = visible });
        var (ok, body) = await SpRestAsync(HttpMethod.Post, $"{site}/_api/web/{path}", site, json);
        if (!ok) return null;
        using var d = JsonDocument.Parse(body);
        return d.RootElement.TryGetProperty("Id", out var id) ? id.GetInt32().ToString() : null;
    }

    // The site's own settings (not its title or address, and not its time zone, which the copy's date handling relies on).
    private async Task CopySiteSettingsAsync(string src, string tgt, List<string> notes)
    {
        var failed = new List<string>();
        async Task ApplyAsync(string path, string[] names)
        {
            var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{src}/_api/web{path}", src);
            if (!ok) return;
            using var d = JsonDocument.Parse(body);
            foreach (var n in names)
            {
                if (!d.RootElement.TryGetProperty(n, out var p)) continue;
                object? value = p.ValueKind switch
                {
                    JsonValueKind.True => true,
                    JsonValueKind.False => false,
                    JsonValueKind.Number => p.TryGetInt32(out var i) ? i : null,
                    JsonValueKind.String => p.GetString(),
                    _ => null,
                };
                if (value is null) continue;
                var (mok, _) = await SpRestAsync(HttpMethod.Post, $"{tgt}/_api/web{path}", tgt,
                    JsonSerializer.Serialize(new Dictionary<string, object> { [n] = value }), merge: true);
                if (!mok) failed.Add(n);
            }
        }
        // Confirmed on a live site: these web properties are accepted; NoCrawl needs a higher permission (403) and the regional
        // settings cannot be written over REST at all (400), so neither is attempted.
        await ApplyAsync("", ["Description", "QuickLaunchEnabled", "TreeViewEnabled", "HorizontalQuickLaunch", "MegaMenuEnabled", "EnableMinimalDownload"]);
        if (failed.Count > 0) notes.Add(Loc.T("Svc_Site_SettingsFailed", string.Join(", ", failed)));
    }

    // ── Library files that hold settings (document template, folder default values) ──────────────

    private async Task<(bool Ok, byte[]? Bytes)> SpGetBytesAsync(string url, string siteUrl)
    {
        using var resp = await spService.SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            return r;
        }, siteUrl);
        return resp.IsSuccessStatusCode ? (true, await resp.Content.ReadAsByteArrayAsync()) : (false, null);
    }

    private static string RestPathArg(string serverRelativePath) => $"'{Uri.EscapeDataString(serverRelativePath.Replace("'", "''"))}'";

    // Uploads (or replaces) a file in a target library's Forms folder.
    private Task<bool> UploadToFormsFolderAsync(string siteUrl, string libraryRoot, string fileName, byte[] data) =>
        UploadToFolderAsync(siteUrl, libraryRoot + "/Forms", fileName, data);

    private async Task<bool> UploadToFolderAsync(string siteUrl, string folderServerRelativePath, string fileName, byte[] data)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetFolderByServerRelativePath(decodedurl={RestPathArg(folderServerRelativePath)})" +
                  $"/Files/add(url='{Uri.EscapeDataString(fileName.Replace("'", "''"))}',overwrite=true)";
        using var resp = await spService.SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new ByteArrayContent(data);
            r.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            return r;
        }, siteUrl);
        return resp.IsSuccessStatusCode;
    }

    // A content type's own document template (the file new documents of that type start from): downloaded from the source type
    // and stored in the target LIBRARY's Forms folder, then set as the template of the library's copy of the type (a type's
    // template may point at any file; its own "_cts" folder refuses every REST write). srcCtUrl / tgtCtUrl are the types' REST
    // addresses (".../contenttypes('id')"). Best effort and quiet: a type with no template, or one that cannot be copied,
    // is simply left as it is.
    private async Task CopyContentTypeTemplateAsync(string srcCtUrl, string srcSite, string tgtCtUrl, string tgtSite, string tgtListBase, List<string> notes, string ctName)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{srcCtUrl}?$select=DocumentTemplate,DocumentTemplateUrl", srcSite);
        if (!ok) return;
        string? fileName, templateUrl;
        using (var d = JsonDocument.Parse(body))
        {
            fileName = d.RootElement.TryGetProperty("DocumentTemplate", out var t) ? t.GetString() : null;
            templateUrl = d.RootElement.TryGetProperty("DocumentTemplateUrl", out var u) ? u.GetString() : null;
        }
        if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(templateUrl)) return;
        // A server-relative address of a file on the source site (an absolute link has the host cut off).
        if (Uri.TryCreate(templateUrl, UriKind.Absolute, out var abs)) templateUrl = Uri.UnescapeDataString(abs.AbsolutePath);
        var (gok, bytes) = await SpGetBytesAsync($"{srcSite}/_api/web/GetFileByServerRelativePath(decodedurl={RestPathArg(templateUrl)})/$value", srcSite);
        if (!gok || bytes is null || bytes.Length == 0) return;

        var root = await RootFolderUrlAsync(tgtListBase, tgtSite);
        if (root is null) return;
        var name = System.IO.Path.GetFileName(templateUrl);
        if (!await UploadToFormsFolderAsync(tgtSite, root, name, bytes)) { notes.Add(Loc.T("Svc_CT_TemplateFailed", ctName)); return; }
        // Confirmed on a live site: the template is set by the file's server-relative address (a bare file name is refused).
        var (mok, _) = await SpRestAsync(HttpMethod.Post, tgtCtUrl, tgtSite,
            JsonSerializer.Serialize(new Dictionary<string, string> { ["DocumentTemplate"] = $"{root}/Forms/{name}" }), merge: true);
        if (!mok) notes.Add(Loc.T("Svc_CT_TemplateFailed", ctName));
    }

    private async Task<string?> RootFolderUrlAsync(string listBase, string siteUrl)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{listBase}/rootfolder?$select=ServerRelativeUrl", siteUrl);
        if (!ok) return null;
        using var d = JsonDocument.Parse(body);
        return d.RootElement.TryGetProperty("ServerRelativeUrl", out var u) ? u.GetString()?.TrimEnd('/') : null;
    }

    // The library's document template (the file new documents start from), copied into the target's Forms folder and set as
    // the target's template.
    private async Task CopyDocumentTemplateAsync(string srcBase, string srcSite, string tgtBase, string tgtSite, List<string> notes)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{srcBase}?$select=DocumentTemplateUrl", srcSite);
        if (!ok) return;
        string? templateUrl;
        using (var d = JsonDocument.Parse(body))
            templateUrl = d.RootElement.TryGetProperty("DocumentTemplateUrl", out var t) ? t.GetString() : null;
        if (string.IsNullOrEmpty(templateUrl)) return;
        var srcRoot = await RootFolderUrlAsync(srcBase, srcSite);
        var tgtRoot = await RootFolderUrlAsync(tgtBase, tgtSite);
        if (srcRoot is null || tgtRoot is null) return;
        // Only a template inside the library's own Forms folder is copied; a link to a template elsewhere is left as the target has it.
        var prefix = srcRoot + "/Forms/";
        if (!templateUrl.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return;
        var name = templateUrl[prefix.Length..];
        if (name.Length == 0 || name.Contains('/')) return;

        var (gok, bytes) = await SpGetBytesAsync($"{srcSite}/_api/web/GetFileByServerRelativePath(decodedurl={RestPathArg(templateUrl)})/$value", srcSite);
        if (!gok || bytes is null || bytes.Length == 0) return;
        if (!await UploadToFormsFolderAsync(tgtSite, tgtRoot, name, bytes)) { notes.Add(Loc.T("Svc_Lib_TemplateFailed")); return; }
        var (mok, _) = await SpRestAsync(HttpMethod.Post, tgtBase, tgtSite,
            JsonSerializer.Serialize(new Dictionary<string, string> { ["DocumentTemplateUrl"] = $"{tgtRoot}/Forms/{name}" }), merge: true);
        if (!mok) notes.Add(Loc.T("Svc_Lib_TemplateFailed"));
    }

    // Folder default column values: the library keeps them in Forms/client_LocationBasedDefaults.html (folder paths and values)
    // plus an item-added receiver that applies them. The file is copied with its folder paths pointed at the target library,
    // and the receiver is added to the target.
    private async Task CopyFolderDefaultsAsync(string srcBase, string srcSite, string tgtBase, string tgtSite, List<string> notes)
    {
        var srcRoot = await RootFolderUrlAsync(srcBase, srcSite);
        var tgtRoot = await RootFolderUrlAsync(tgtBase, tgtSite);
        if (srcRoot is null || tgtRoot is null) return;
        const string fileName = "client_LocationBasedDefaults.html";
        var (gok, bytes) = await SpGetBytesAsync($"{srcSite}/_api/web/GetFileByServerRelativePath(decodedurl={RestPathArg($"{srcRoot}/Forms/{fileName}")})/$value", srcSite);
        if (!gok || bytes is null || bytes.Length == 0) return;   // the library has no default values

        var text = System.Text.Encoding.UTF8.GetString(bytes);
        // Through placeholders, so a target whose path starts with the source path ("Docs" -> "Docs2") is not rewritten twice
        // (confirmed on a live site: a plain second Replace turned "ZZSLibCopy" into "ZZSLibCopyCopy").
        const string rawToken = "\u0001SRCROOT-RAW\u0001", encToken = "\u0001SRCROOT-ENC\u0001";
        var encSrc = srcRoot.Replace(" ", "%20");
        if (encSrc != srcRoot) text = text.Replace(encSrc, encToken, StringComparison.OrdinalIgnoreCase);
        text = text.Replace(srcRoot, rawToken, StringComparison.OrdinalIgnoreCase)
                   .Replace(rawToken, tgtRoot)
                   .Replace(encToken, tgtRoot.Replace(" ", "%20"));
        if (!await UploadToFormsFolderAsync(tgtSite, tgtRoot, fileName, System.Text.Encoding.UTF8.GetBytes(text)))
        {
            notes.Add(Loc.T("Svc_Lib_DefaultsFailed"));
            return;
        }

        // The receiver that applies the defaults to new files.
        var (rok, rBody) = await SpRestAsync(HttpMethod.Get, $"{srcBase}/eventreceivers?$select=ReceiverName,ReceiverClass,ReceiverAssembly,EventType,SequenceNumber,Synchronization", srcSite);
        if (!rok) return;
        var (tok, tBody) = await SpRestAsync(HttpMethod.Get, $"{tgtBase}/eventreceivers?$select=ReceiverClass", tgtSite);
        var targetHas = false;
        if (tok)
            using (var td = JsonDocument.Parse(tBody))
                targetHas = td.RootElement.GetProperty("value").EnumerateArray()
                    .Any(e => (e.GetProperty("ReceiverClass").GetString() ?? "").Contains("LocationBasedMetadataDefaultsReceiver", StringComparison.OrdinalIgnoreCase));
        if (targetHas) return;
        using var doc = JsonDocument.Parse(rBody);
        foreach (var e in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            if (!(e.GetProperty("ReceiverClass").GetString() ?? "").Contains("LocationBasedMetadataDefaultsReceiver", StringComparison.OrdinalIgnoreCase)) continue;
            var add = new Dictionary<string, object?>();
            foreach (var n in new[] { "ReceiverName", "ReceiverClass", "ReceiverAssembly", "EventType", "SequenceNumber", "Synchronization" })
                if (e.TryGetProperty(n, out var pv))
                    add[n] = pv.ValueKind == JsonValueKind.Number ? pv.GetInt32() : pv.GetString();
            var (aok, _) = await SpRestAsync(HttpMethod.Post, $"{tgtBase}/eventreceivers", tgtSite, JsonSerializer.Serialize(add));
            if (!aok) notes.Add(Loc.T("Svc_Lib_DefaultsReceiverFailed"));
        }
    }

    // Column validation rules (formula and message) and column formatting (JSON), copied onto the matching target column.
    private async Task CopyColumnRulesAsync(string srcBase, string srcSite, string tgtBase, string tgtSite, List<string> notes)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{srcBase}/fields?$filter=Hidden eq false&$select=InternalName,ValidationFormula,ValidationMessage,CustomFormatter", srcSite);
        if (!ok) return;   // column rules are an extra; a read problem here is not worth a note on its own
        var failed = new List<string>();
        using var doc = JsonDocument.Parse(body);
        foreach (var f in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            var props = new Dictionary<string, string>();
            foreach (var n in new[] { "ValidationFormula", "ValidationMessage", "CustomFormatter" })
                if (f.TryGetProperty(n, out var p) && p.ValueKind == JsonValueKind.String && !string.IsNullOrEmpty(p.GetString()))
                    props[n] = p.GetString()!;
            if (props.Count == 0) continue;
            var name = f.GetProperty("InternalName").GetString()!;
            var (mok, _) = await SpRestAsync(HttpMethod.Post, $"{tgtBase}/fields/getbyinternalnameortitle('{name.Replace("'", "''")}')", tgtSite, JsonSerializer.Serialize(props), merge: true);
            if (!mok) failed.Add(name);
        }
        if (failed.Count > 0) notes.Add(Loc.T("Svc_Col_RulesFailed", string.Join(", ", failed)));
    }

    // Content types the source list has (and the target does not): created on the target site from the same built-in parent
    // (Document, Folder, Document Set, Item...), added to the list, and given the same columns. Items and folders are NOT
    // re-bound to their types by this step. Best effort: a Document Set type needs the Document Sets feature on the target site
    // collection, and not every setup accepts column links, so failures are listed as notes.
    // REST cannot link an existing column to a content type (every call shape was refused on a live site), but a column created
    // with "add to all content types" lands on every type the list has. So on a list this copy has just created (it is still
    // empty), each column the source type has is deleted and created again from its own definition with that option. Managed-metadata
    // columns are not touched (they carry ties to the term store); lookup columns are recreated like the rest. Returns the columns that
    // are still not on the type.
    private async Task<List<string>> LinkColumnsToTypeAsync(string tgtBase, string tgtSite, string listCtId, List<string> columns)
    {
        var stillLacking = new List<string>();
        foreach (var col in columns)
        {
            try
            {
                var fieldUrl = $"{tgtBase}/fields/getbyinternalnameortitle('{col.Replace("'", "''")}')";
                var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{fieldUrl}?$select=SchemaXml,TypeAsString,Hidden,ReadOnlyField,Sealed", tgtSite);
                if (!ok) continue;   // not a column of the list (a built-in the type has)
                string xml, type; bool locked;
                using (var d = JsonDocument.Parse(body))
                {
                    var r = d.RootElement;
                    xml = r.GetProperty("SchemaXml").GetString() ?? "";
                    type = r.GetProperty("TypeAsString").GetString() ?? "";
                    locked = Bool(r, "Hidden") || Bool(r, "ReadOnlyField") || Bool(r, "Sealed");
                }
                // A lookup is safe to recreate from its own schema: the target column already points at the TARGET list (and the list
                // is still empty). A column another column depends on (a projected lookup field) is left for the report.
                if (locked || xml.Length == 0 || type.StartsWith("TaxonomyField", StringComparison.OrdinalIgnoreCase)
                    || xml.Contains("FieldRef=", StringComparison.OrdinalIgnoreCase))
                { stillLacking.Add(col); continue; }
                var (dok, _) = await SpRestAsync(HttpMethod.Post, fieldUrl, tgtSite, null, xMethod: "DELETE");
                if (!dok) { stillLacking.Add(col); continue; }
                // 28 = add to all content types (4) + keep the internal name (8) + add to the default view (16)
                var (cok, _) = await SpRestAsync(HttpMethod.Post, $"{tgtBase}/fields/createfieldasxml", tgtSite,
                    JsonSerializer.Serialize(new { parameters = new { SchemaXml = xml, Options = 28 } }));
                if (!cok) stillLacking.Add(col);
            }
            catch (Exception ex) when (ex is not OperationCanceledException) { stillLacking.Add(col); }
        }
        return stillLacking;
    }

    private async Task CopyContentTypesAsync(string srcBase, string srcSite, string tgtBase, string tgtSite, List<string> notes, bool recreateColumns = false, bool isLibrary = false)
    {
        var (ok, ctBody) = await SpRestAsync(HttpMethod.Get, $"{srcBase}/contenttypes?$select=Name,Description,Group,Hidden,StringId", srcSite);
        if (!ok) return;
        var (tok, tBody) = await SpRestAsync(HttpMethod.Get, $"{tgtBase}/contenttypes?$select=Name,StringId", tgtSite);
        var targetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (tok)
            using (var td = JsonDocument.Parse(tBody))
                foreach (var t in td.RootElement.GetProperty("value").EnumerateArray())
                    targetNames.Add(t.GetProperty("Name").GetString() ?? "");

        using var doc = JsonDocument.Parse(ctBody);
        foreach (var ct in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            var name = ct.GetProperty("Name").GetString() ?? "";
            if (name.Length == 0 || Bool(ct, "Hidden") || targetNames.Contains(name)) continue;
            var stringId = ct.GetProperty("StringId").GetString() ?? "";
            // A custom type's ID is its built-in parent + "00" + a GUID (+ more for the list copy of it).
            var m = System.Text.RegularExpressions.Regex.Match(stringId, "^(0x[0-9A-Fa-f]+?)00[0-9A-Fa-f]{32}");
            if (!m.Success) continue;   // a built-in type the new list already has
            try
            {
                string newSiteId;   // read from the response: SharePoint ignores an ID sent with the request
                var create = new Dictionary<string, object>
                {
                    ["Name"] = name,
                    ["Description"] = ct.TryGetProperty("Description", out var d) ? d.GetString() ?? "" : "",
                    ["Group"] = ct.TryGetProperty("Group", out var g) ? g.GetString() ?? "Custom Content Types" : "Custom Content Types",
                };
                var (cok, cBody) = await SpRestAsync(HttpMethod.Post, $"{tgtSite}/_api/web/contenttypes", tgtSite, JsonSerializer.Serialize(create));
                if (!cok)
                {
                    // Already a site content type of that name (an earlier copy): use it.
                    var (eok, eBody) = await SpRestAsync(HttpMethod.Get, $"{tgtSite}/_api/web/contenttypes?$select=StringId&$filter=Name eq '{name.Replace("'", "''")}'", tgtSite);
                    string? existing = null;
                    if (eok) using (var ed = JsonDocument.Parse(eBody))
                        existing = ed.RootElement.GetProperty("value").EnumerateArray().Select(e => e.GetProperty("StringId").GetString()).FirstOrDefault();
                    if (existing is null) { notes.Add(Loc.T("Svc_CT_Failed", name, Short(cBody))); continue; }
                    newSiteId = existing;
                }
                else
                {
                    using (var cd = JsonDocument.Parse(cBody)) newSiteId = cd.RootElement.GetProperty("StringId").GetString()!;
                }
                var (aok, aBody) = await SpRestAsync(HttpMethod.Post, $"{tgtBase}/contenttypes/addAvailableContentType(contentTypeId='{newSiteId}')", tgtSite, "{}");
                if (!aok) { notes.Add(Loc.T("Svc_CT_Failed", name, Short(aBody))); continue; }

                // The type's own columns cannot be linked over REST. Say which ones the source type has that the new one lacks.
                string? listCtId = null;
                using (var ad = JsonDocument.Parse(aBody)) listCtId = ad.RootElement.TryGetProperty("StringId", out var sid) ? sid.GetString() : null;
                // The type's document template goes into the library's own Forms folder (a site type's "_cts" folder refuses
                // every REST write, even for a site admin) and the library's copy of the type points at it.
                if (isLibrary && listCtId is not null)
                    await CopyContentTypeTemplateAsync($"{srcBase}/contenttypes('{stringId}')", srcSite, $"{tgtBase}/contenttypes('{listCtId}')", tgtSite, tgtBase, notes, name);
                var (fok, fBody) = await SpRestAsync(HttpMethod.Get, $"{srcBase}/contenttypes('{stringId}')/fieldlinks?$select=Name", srcSite);
                var (lok, lBody) = listCtId is null ? (false, "") : await SpRestAsync(HttpMethod.Get, $"{tgtBase}/contenttypes('{listCtId}')/fieldlinks?$select=Name", tgtSite);
                if (fok && lok)
                {
                    var have = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var ld = JsonDocument.Parse(lBody))
                        foreach (var l in ld.RootElement.GetProperty("value").EnumerateArray()) have.Add(l.GetProperty("Name").GetString() ?? "");
                    using var fd = JsonDocument.Parse(fBody);
                    var lacking = fd.RootElement.GetProperty("value").EnumerateArray()
                        .Select(fl => fl.GetProperty("Name").GetString() ?? "").Where(n => n.Length > 0 && !have.Contains(n)).ToList();
                    if (lacking.Count > 0 && recreateColumns && listCtId is not null)
                        lacking = await LinkColumnsToTypeAsync(tgtBase, tgtSite, listCtId, lacking);
                    if (lacking.Count > 0) notes.Add(Loc.T("Svc_CT_ColumnsFailed", name, string.Join(", ", lacking)));
                }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                notes.Add(Loc.T("Svc_CT_Failed", name, Short(ex.Message)));
            }
        }
    }

    // Points the column names in a view's query or totals at the target's internal names (source name -> target name).
    private static string RenameViewFields(string xml, Dictionary<string, string> renames)
    {
        if (renames.Count == 0 || string.IsNullOrEmpty(xml)) return xml;
        return System.Text.RegularExpressions.Regex.Replace(xml, "(FieldRef\\s[^>]*?Name=\")([^\"]+)(\")",
            m => renames.TryGetValue(m.Groups[2].Value, out var to) ? m.Groups[1].Value + to + m.Groups[3].Value : m.Value);
    }

    // Takes the parts of a view's query (filter, sort, grouping) that name a column the target does not have out of the query.
    // A sort or grouping column that is missing is removed from that clause (an empty clause goes); a filter that uses a missing
    // column is removed as a whole, since half a filter would change which items show. Returns the query to use and the missing
    // column names. A query that cannot be read is returned unchanged.
    private static string TrimViewQuery(string query, HashSet<string> targetFields, out List<string> missing)
    {
        missing = [];
        if (string.IsNullOrWhiteSpace(query)) return query;
        try
        {
            var root = System.Xml.Linq.XElement.Parse("<Q>" + query + "</Q>");
            string? NameOf(System.Xml.Linq.XElement e) => (string?)e.Attribute("Name");
            var changed = false;
            foreach (var clause in root.Elements().ToList())
            {
                var refs = clause.DescendantsAndSelf("FieldRef").ToList();
                var bad = refs.Where(r => NameOf(r) is { Length: > 0 } n && !targetFields.Contains(n)).ToList();
                if (bad.Count == 0) continue;
                changed = true;
                foreach (var b in bad) if (!missing.Contains(NameOf(b)!, StringComparer.OrdinalIgnoreCase)) missing.Add(NameOf(b)!);
                if (clause.Name.LocalName == "Where") clause.Remove();
                else
                {
                    foreach (var b in bad) b.Remove();
                    if (!clause.Elements("FieldRef").Any()) clause.Remove();
                }
            }
            return changed ? string.Concat(root.Nodes().Select(n => n.ToString(System.Xml.Linq.SaveOptions.DisableFormatting))) : query;
        }
        catch (System.Xml.XmlException) { return query; }
    }

    // A view's column totals (its Aggregations XML): totals on a column the target does not have are taken out.
    private static string TrimViewAggregations(string xml, HashSet<string> targetFields)
    {
        try
        {
            var root = System.Xml.Linq.XElement.Parse(xml);
            var bad = root.Descendants("FieldRef").Where(r => (string?)r.Attribute("Name") is { Length: > 0 } n && !targetFields.Contains(n)).ToList();
            if (bad.Count == 0) return xml;
            foreach (var b in bad) b.Remove();
            return root.Elements().Any() ? root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting) : "";
        }
        catch (System.Xml.XmlException) { return xml; }
    }

    private async Task CopyViewsAsync(LibraryDefinition def, string targetSiteUrl, string targetListId, List<string> notes, bool additive = false, bool update = false)
    {
        try
        {
            var srcSite = def.SourceSiteUrl.TrimEnd('/');
            var srcBase = $"{srcSite}/_api/web/lists('{def.SourceListId}')";
            var tgtBase = $"{targetSiteUrl.TrimEnd('/')}/_api/web/lists('{targetListId}')";

            var (ok, viewsBody) = await SpRestAsync(HttpMethod.Get, $"{srcBase}/views", srcSite);
            if (!ok) { notes.Add(Loc.T("Svc_View_ReadFail", Short(viewsBody))); return; }

            // What the target already has: its views (by title) and its columns.
            var targetViews = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            (ok, var tvBody) = await SpRestAsync(HttpMethod.Get, $"{tgtBase}/views?$select=Id,Title", targetSiteUrl);
            if (ok)
                using (var d = JsonDocument.Parse(tvBody))
                    foreach (var v in d.RootElement.GetProperty("value").EnumerateArray())
                        targetViews[v.GetProperty("Title").GetString() ?? ""] = v.GetProperty("Id").GetString() ?? "";

            var targetFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var targetByTitle = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            (ok, var tfBody) = await SpRestAsync(HttpMethod.Get, $"{tgtBase}/fields?$select=InternalName,Title", targetSiteUrl);
            if (ok)
                using (var d = JsonDocument.Parse(tfBody))
                    foreach (var f in d.RootElement.GetProperty("value").EnumerateArray())
                        if (f.TryGetProperty("InternalName", out var n) && n.GetString() is { } s)
                        {
                            targetFields.Add(s);
                            var t = f.TryGetProperty("Title", out var tt) ? tt.GetString() ?? "" : "";
                            if (t.Length > 0 && !targetByTitle.ContainsKey(t)) targetByTitle[t] = s;
                        }

            // A column the copy made under a different internal name than the source's (an older copy named new columns after
            // their display title) is matched by its title, so views, filters and sorts still find it.
            var renames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            (ok, var sfBody) = await SpRestAsync(HttpMethod.Get, $"{srcBase}/fields?$select=InternalName,Title", srcSite);
            if (ok)
                using (var d = JsonDocument.Parse(sfBody))
                    foreach (var f in d.RootElement.GetProperty("value").EnumerateArray())
                    {
                        var inName = f.TryGetProperty("InternalName", out var n) ? n.GetString() ?? "" : "";
                        var title = f.TryGetProperty("Title", out var tt) ? tt.GetString() ?? "" : "";
                        if (inName.Length > 0 && !targetFields.Contains(inName) && targetByTitle.TryGetValue(title, out var other))
                            renames[inName] = other;
                    }

            using var doc = JsonDocument.Parse(viewsBody);
            foreach (var v in doc.RootElement.GetProperty("value").EnumerateArray())
            {
                var title = v.TryGetProperty("Title", out var tEl) ? tEl.GetString() ?? "" : "";
                try
                {
                    if (title.Length == 0) continue;
                    if (Bool(v, "PersonalView") || Bool(v, "Hidden")) continue;
                    // An existing list keeps the views it has, unless one is broken (an older copy left its filter pointing at
                    // column names the target does not have, so it will not render): that one is rebuilt from the source.
                    var repairing = false;
                    if (additive && !update && targetViews.TryGetValue(title, out var existingViewId))
                    {
                        if (!await TargetViewIsBrokenAsync(tgtBase, existingViewId, targetSiteUrl, targetFields)) continue;
                        repairing = true;
                        notes.Add(Loc.T("Svc_View_Repaired", title));
                    }
                    var type =v.TryGetProperty("ViewType", out var vt) ? vt.GetString() : null;
                    // Calendar, Gantt and datasheet views are created with their own view kind (and the view's own XML data);
                    // modern board/calendar views carry a ViewType2 name. Any other kind is skipped with a note.
                    int? viewKind = type?.ToUpperInvariant() switch { "CALENDAR" => 524288, "GANTT" => 67108864, "GRID" => 2048, _ => null };
                    var type2 = v.TryGetProperty("ViewType2", out var vt2) && vt2.ValueKind == JsonValueKind.String ? vt2.GetString() : null;
                    if (string.Equals(type2, "HTML", StringComparison.OrdinalIgnoreCase)) type2 = null;
                    var exotic = viewKind != null || !string.IsNullOrEmpty(type2);
                    if (!string.IsNullOrEmpty(type) && !type.Equals("HTML", StringComparison.OrdinalIgnoreCase) && viewKind == null && string.IsNullOrEmpty(type2))
                    {
                        notes.Add(Loc.T("Svc_View_Skipped", title, type));
                        continue;
                    }
                    var srcViewId = v.GetProperty("Id").GetString();
                    var query = v.TryGetProperty("ViewQuery", out var q) ? q.GetString() ?? "" : "";
                    query = RenameViewFields(query, renames);
                    // A filter, sort or grouping on a column the target lacks is accepted by SharePoint but then the view will not
                    // render ("one or more field types are not installed properly"). Those parts are left out, with a note.
                    query = TrimViewQuery(query, targetFields, out var trimmed);
                    if (trimmed.Count > 0) notes.Add(Loc.T("Svc_View_QueryTrimmed", title, string.Join(", ", trimmed)));
                    var isDefault = Bool(v, "DefaultView") && (!additive || update);   // never changes an existing list's default view
                    var rowLimit = v.TryGetProperty("RowLimit", out var rl) && rl.TryGetInt32(out var rlv) ? rlv : 30;

                    // The view's columns, in order.
                    var (fok, fBody) = await SpRestAsync(HttpMethod.Get, $"{srcBase}/views('{srcViewId}')/viewfields", srcSite);
                    var fields = new List<string>();
                    if (fok) using (var fd = JsonDocument.Parse(fBody)) fields = ReadStrings(fd.RootElement);
                    fields = fields.Select(f => renames.TryGetValue(f, out var other) ? other : f).ToList();

                    // The properties that go onto the view (the ones the source has).
                    var props = new Dictionary<string, object?> { ["RowLimit"] = rowLimit, ["Paged"] = !v.TryGetProperty("Paged", out var pg) || pg.ValueKind != JsonValueKind.False };
                    foreach (var name in new[] { "Aggregations", "AggregationsStatus", "Formats", "CustomFormatter" })
                        if (v.TryGetProperty(name, out var pv) && pv.ValueKind == JsonValueKind.String && !string.IsNullOrEmpty(pv.GetString()))
                        {
                            // Every view reports AggregationsStatus "Off" (no totals), and SharePoint refuses to have that value set back
                            // ("Value does not fall within the expected range"), which made it refuse the whole batch of view settings.
                            if (name == "AggregationsStatus" && string.Equals(pv.GetString(), "Off", StringComparison.OrdinalIgnoreCase)) continue;
                            props[name] = name == "Aggregations" ? TrimViewAggregations(RenameViewFields(pv.GetString()!, renames), targetFields) : pv.GetString();
                        }
                    if (isDefault) props["DefaultView"] = true;

                    var queryDropped = false;
                    if (!targetViews.TryGetValue(title, out var targetViewId))
                    {
                        // New view. Created with its filter; if the filter names a column the target lacks, without it.
                        // views/add takes its settings inside a "parameters" object (ViewCreationInformation), and the filter is
                        // called Query there. Confirmed on a live site: flat properties are rejected.
                        var add = new Dictionary<string, object?> { ["Title"] = title, ["PersonalView"] = false, ["RowLimit"] = rowLimit, ["SetAsDefaultView"] = isDefault };
                        if (query.Length > 0) add["Query"] = query;
                        if (viewKind != null) add["ViewTypeKind"] = viewKind.Value;
                        (ok, var addBody) = await SpRestAsync(HttpMethod.Post, $"{tgtBase}/views/add", targetSiteUrl, JsonSerializer.Serialize(new { parameters = add }));
                        if (!ok && query.Length > 0)
                        {
                            add.Remove("Query");
                            queryDropped = true;
                            (ok, addBody) = await SpRestAsync(HttpMethod.Post, $"{tgtBase}/views/add", targetSiteUrl, JsonSerializer.Serialize(new { parameters = add }));
                        }
                        // What the add cannot carry is set on the new view afterwards: a calendar's or Gantt's own data, and a
                        // modern view type (board / calendar). The type is accepted; that it renders as a board is not confirmed.
                        if (exotic && v.TryGetProperty("ViewData", out var vd) && vd.ValueKind == JsonValueKind.String && !string.IsNullOrEmpty(vd.GetString()))
                            props["ViewData"] = vd.GetString();
                        if (!string.IsNullOrEmpty(type2)) props["ViewType2"] = type2;
                        if (!ok) { notes.Add(Loc.T("Svc_View_CreateFail", title, Short(addBody))); continue; }
                        using var created = JsonDocument.Parse(addBody);
                        targetViewId = created.RootElement.GetProperty("Id").GetString()!;
                    }
                    else if (query.Length > 0 || update || repairing)
                        props["ViewQuery"] = query;   // an existing view of that name (the built-in one) takes the source's filter

                    var mergeUrl = $"{tgtBase}/views('{targetViewId}')";
                    // One bad setting makes SharePoint refuse the whole batch ("Value does not fall within the expected range"), which
                    // used to take the filter down with it. The basics go first, then the filter, then every other setting on its own,
                    // so a failure costs only that setting and the note says which one.
                    var optional = new[] { "Aggregations", "AggregationsStatus", "Formats", "CustomFormatter", "ViewData", "ViewType2" };
                    var separate = new List<KeyValuePair<string, object?>>();
                    foreach (var name in new[] { "ViewQuery" }.Concat(optional))
                        if (props.TryGetValue(name, out var pval)) { separate.Add(new(name, pval)); props.Remove(name); }
                    (ok, var mergeBody) = await SpRestAsync(HttpMethod.Post, mergeUrl, targetSiteUrl, JsonSerializer.Serialize(props), merge: true);
                    if (!ok) notes.Add(Loc.T("Svc_View_SettingsFail", title, Short(mergeBody)));
                    foreach (var (name, pval) in separate)
                    {
                        var (pok, pbody) = await SpRestAsync(HttpMethod.Post, mergeUrl, targetSiteUrl, JsonSerializer.Serialize(new Dictionary<string, object?> { [name] = pval }), merge: true);
                        if (pok) continue;
                        if (name == "ViewQuery") notes.Add(Loc.T("Svc_View_QueryDropped", title));
                        else notes.Add(Loc.T("Svc_View_PropFail", title, name, Short(pbody)));
                    }
                    if (queryDropped) notes.Add(Loc.T("Svc_View_QueryDropped", title));   // a new view whose filter had to be left out at creation

                    // Columns: only ones the target has, in the source's order.
                    if (fields.Count > 0)
                    {
                        var keep = fields.Where(targetFields.Contains).ToList();
                        var missing = fields.Where(f => !targetFields.Contains(f)).ToList();
                        if (keep.Count > 0)
                        {
                            await SpRestAsync(HttpMethod.Post, $"{mergeUrl}/viewfields/removeallviewfields", targetSiteUrl, "{}");
                            foreach (var f in keep)
                                await SpRestAsync(HttpMethod.Post, $"{mergeUrl}/viewfields/addviewfield('{f.Replace("'", "''")}')", targetSiteUrl, "{}");
                        }
                        if (missing.Count > 0) notes.Add(Loc.T("Svc_View_FieldsMissing", title, string.Join(", ", missing)));
                    }
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    notes.Add(Loc.T("Svc_View_CreateFail", title, Short(ex.Message)));
                }
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            notes.Add(Loc.T("Svc_View_ReadFail", Short(ex.Message)));
        }
    }

    // True when an existing target view cannot render: its filter, sort or grouping names a column the list does not have
    // (checked from the XML, so it works in any display language), or SharePoint itself says its field types are not installed
    // properly. Any failure to find out counts as "not broken": a view is never rebuilt on a guess.
    private async Task<bool> TargetViewIsBrokenAsync(string tgtBase, string viewId, string siteUrl, HashSet<string> targetFields)
    {
        try
        {
            var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{tgtBase}/views('{viewId}')?$select=ViewQuery", siteUrl);
            if (ok)
            {
                using var d = JsonDocument.Parse(body);
                var q = d.RootElement.TryGetProperty("ViewQuery", out var qe) ? qe.GetString() ?? "" : "";
                if (q.Length > 0 && System.Xml.Linq.XElement.Parse("<Q>" + q + "</Q>").Descendants("FieldRef")
                        .Any(r => (string?)r.Attribute("Name") is { Length: > 0 } n && !targetFields.Contains(n)))
                    return true;
            }
            (ok, body) = await SpRestAsync(HttpMethod.Get, $"{tgtBase}/views('{viewId}')/renderashtml", siteUrl);
            return ok && body.Contains("not installed properly", StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception ex) when (ex is not OperationCanceledException) { return false; }
    }

    private static bool Bool(JsonElement e, string name) =>
        e.TryGetProperty(name, out var p) && p.ValueKind == JsonValueKind.True;

    // Column-creation failures from the most recent CreateLibraryAsync — surfaced by callers
    // into the run report (previously Debug-output only, so the user was never told).
    public List<string> LastColumnCreationFailures { get; private set; } = [];

    private async Task CreateColumnAsync(string siteUrl, string listId, ColumnDefinition col, string? sourceSiteUrl = null)
    {
        // Lookup columns bind to a target list, not a field type — the old fallthrough to the
        // JSON path created them as plain TEXT columns, after which every file's lookup value
        // failed with "Lookup unresolved".
        if (ColumnDefinition.IsLookupType(col.FieldType))
        {
            await CreateLookupColumnAsync(siteUrl, listId, col, sourceSiteUrl);
            return;
        }

        // Cross-tenant: a managed-metadata column is bound to the SOURCE term store's SspId/TermSetId,
        // which don't exist in another tenant — creating it would leave a broken column. Skipped, and
        // its values are skipped too (see SharePointService.SkipsTaxonomyValue).
        if (ColumnDefinition.IsTaxonomyType(col.FieldType) && spService.Router.IsMultiTenant)
        {
            await CreateTaxonomyColumnWithTermCopyAsync(siteUrl, listId, col, sourceSiteUrl);
            return;
        }

        // User, taxonomy, and choice columns with options need schema XML.
        // The plain JSON /fields endpoint rejects typed properties (Choices, __metadata)
        // regardless of the odata=nometadata accept header.
        bool isChoiceWithOptions = (col.FieldType == SupportedFieldType.Choice || col.FieldType == SupportedFieldType.MultiChoice)
                                   && col.Choices is { Length: > 0 };
        if (ColumnDefinition.IsUserType(col.FieldType) || ColumnDefinition.IsTaxonomyType(col.FieldType) || isChoiceWithOptions)
        {
            await CreateColumnFromSchemaXmlAsync(siteUrl, listId, col);
            return;
        }

        var fieldTypeKind = col.FieldType switch
        {
            SupportedFieldType.Text        => 2,
            SupportedFieldType.Note        => 3,
            SupportedFieldType.DateTime    => 4,
            SupportedFieldType.Choice      => 6,
            SupportedFieldType.Boolean     => 8,
            SupportedFieldType.Number      => 9,
            SupportedFieldType.Currency    => 10,
            SupportedFieldType.Hyperlink   => 11,
            SupportedFieldType.MultiChoice => 15,
            _ => 2,
        };
        // Create with Title = INTERNAL name: SharePoint derives InternalName from the creation
        // Title (StaticName is ignored for that), so creating with the display name broke every
        // renamed column — the target's internal name became the encoded display name while all
        // value writers still wrote to the source internal name. Create-then-rename yields the
        // exact source internal name with the right display title.
        var json = JsonSerializer.Serialize(new
        {
            FieldTypeKind = fieldTypeKind,
            Title         = col.InternalName,
            StaticName    = col.InternalName,
            Required      = false,
        });

        var createUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/fields";
        string actualInternalName = col.InternalName;

        using var createResp = await spService.SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, createUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl);

        var createBody = await createResp.Content.ReadAsStringAsync();
        if (!createResp.IsSuccessStatusCode)
            throw new HttpRequestException(
                Loc.T("Svc_Lib_ColumnRejected", (int)createResp.StatusCode, createBody));

        using (var doc = JsonDocument.Parse(createBody))
        {
            if (doc.RootElement.TryGetProperty("InternalName", out var n))
                actualInternalName = n.GetString() ?? col.InternalName;
        }

        // Rename to the display title (MERGE only changes Title; InternalName is now fixed).
        if (!string.Equals(col.DisplayName, actualInternalName, StringComparison.Ordinal))
        {
            var renameBody = JsonSerializer.Serialize(new { Title = col.DisplayName });
            using var renameResp = await spService.SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post,
                    $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/fields/getbyinternalnameortitle('{Uri.EscapeDataString(actualInternalName.Replace("'", "''"))}')");
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Headers.TryAddWithoutValidation("IF-MATCH", "*");
                r.Headers.TryAddWithoutValidation("X-HTTP-Method", "MERGE");
                r.Content = new StringContent(renameBody, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
            // Non-critical: a failed rename leaves the internal name as the display title,
            // which is cosmetic — values still land in the right field.
        }

        // Add to default view (non-critical)
        await AddColumnToDefaultViewAsync(siteUrl, listId, actualInternalName);
    }

    // Finds the TARGET list (and its web, when it is not the target site's own web) that equals a source lookup list. The source
    // column stores the source list's GUID, which means nothing at the target, so the list is matched by title. A lookup can point
    // at a list in another web of the same site collection: the source web is found from the column's WebId and the same
    // relative web is used at the target. Returns nulls when the list cannot be resolved.
    private async Task<(string? ListId, string? WebId)> ResolveTargetLookupAsync(string siteUrl, string? sourceSiteUrl, string? lookupListId, string? lookupWebId)
    {
        string? targetLookupListId = null;
        string? targetLookupWebId  = null;
        if (!string.IsNullOrEmpty(lookupListId) && !string.IsNullOrEmpty(sourceSiteUrl))
        {
            var srcLookupWeb = sourceSiteUrl.TrimEnd('/');
            var tgtLookupWeb = siteUrl.TrimEnd('/');
            if (!string.IsNullOrEmpty(lookupWebId))
            {
                var ownWebId = await GetWebIdAsync(srcLookupWeb);
                if (ownWebId != null && !string.Equals(ownWebId, lookupWebId, StringComparison.OrdinalIgnoreCase))
                {
                    var (ok, body) = await SpRestAsync(HttpMethod.Get,
                        $"{srcLookupWeb}/_api/site/openWebById('{lookupWebId}')?$select=Url", srcLookupWeb);
                    if (ok)
                    {
                        using var d = JsonDocument.Parse(body);
                        var webUrl = d.RootElement.TryGetProperty("Url", out var u) ? u.GetString() : null;
                        if (!string.IsNullOrEmpty(webUrl) &&
                            webUrl.StartsWith(srcLookupWeb, StringComparison.OrdinalIgnoreCase))
                        {
                            // A child web: same relative path under the target web.
                            tgtLookupWeb = tgtLookupWeb + webUrl[srcLookupWeb.Length..].TrimEnd('/');
                            srcLookupWeb = webUrl.TrimEnd('/');
                        }
                        else if (!string.IsNullOrEmpty(webUrl))
                        {
                            srcLookupWeb = webUrl.TrimEnd('/');   // sibling/parent web: only resolvable if it exists at the target by the same path
                        }
                    }
                }
            }

            var srcListTitle = await spService.GetListTitleByIdAsync(srcLookupWeb, lookupListId!);
            if (!string.IsNullOrEmpty(srcListTitle))
            {
                targetLookupListId = await spService.GetListIdByTitleAsync(tgtLookupWeb, srcListTitle!);
                if (!string.IsNullOrEmpty(targetLookupListId) &&
                    !string.Equals(tgtLookupWeb, siteUrl.TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
                    targetLookupWebId = await GetWebIdAsync(tgtLookupWeb);
            }
        }
        return (targetLookupListId, targetLookupWebId);
    }

    // Recreates a Lookup/LookupMulti column bound to the TARGET list equivalent to the source's
    // lookup list (matched by title — the source column stores the source list's GUID, which is
    // meaningless at the target). Throws with an actionable message when the list can't be
    // resolved, instead of silently degrading the column to Text as the old fallthrough did.
    private async Task CreateLookupColumnAsync(string siteUrl, string listId, ColumnDefinition col, string? sourceSiteUrl)
    {
        var (targetLookupListId, targetLookupWebId) = await ResolveTargetLookupAsync(siteUrl, sourceSiteUrl, col.LookupListId, col.LookupWebId);
        if (string.IsNullOrEmpty(targetLookupListId))
            throw new Exception(
                Loc.T("Svc_Lib_LookupNotCreated", col.DisplayName));

        var el = new System.Xml.Linq.XElement("Field",
            new System.Xml.Linq.XAttribute("Type", col.FieldType == SupportedFieldType.LookupMulti ? "LookupMulti" : "Lookup"),
            new System.Xml.Linq.XAttribute("DisplayName", col.DisplayName),
            new System.Xml.Linq.XAttribute("Name", col.InternalName),
            new System.Xml.Linq.XAttribute("StaticName", col.InternalName),
            new System.Xml.Linq.XAttribute("List", $"{{{targetLookupListId!.Trim('{', '}')}}}"),
            new System.Xml.Linq.XAttribute("ShowField", string.IsNullOrEmpty(col.LookupShowField) ? "Title" : col.LookupShowField!),
            new System.Xml.Linq.XAttribute("Required", "FALSE"));
        if (col.FieldType == SupportedFieldType.LookupMulti)
            el.Add(new System.Xml.Linq.XAttribute("Mult", "TRUE"));
        if (!string.IsNullOrEmpty(targetLookupWebId))
            el.Add(new System.Xml.Linq.XAttribute("WebId", $"{{{targetLookupWebId}}}"));

        await PostCreateFieldAsXmlAsync(siteUrl, listId, el.ToString(System.Xml.Linq.SaveOptions.DisableFormatting), col);
    }

    // Which kinds of structure a run copies (the Options screen's "Structure to copy"). All on by default.
    public sealed class StructureOptions
    {
        public bool Views { get; set; } = true;
        public bool ListSettings { get; set; } = true;
        public bool ContentTypes { get; set; } = true;
        public bool ColumnRules { get; set; } = true;
        public bool Templates { get; set; } = true;
        // Existing lists: also overwrite same-named views, the default view, list settings and column rules/formatting.
        public bool UpdateExisting { get; set; } = false;
        public bool SiteDefinitions { get; set; } = true;
        public bool Navigation { get; set; } = true;
        public bool SiteSettings { get; set; } = true;
    }
    public StructureOptions Structure { get; set; } = new();

    // Set by the caller before a run: when present (and CopyTermSets is on) a managed-metadata column whose term set
    // does not exist at the target gets that set copied first, so the column can be bound to it.
    public TermStoreCopyService? TermStore { get; set; }
    public bool CopyTermSets { get; set; } = true;
    // What the term store copy did (sets created/reused, terms skipped), for the run report.
    public List<string> TermStoreNotes { get; } = [];
    private readonly Dictionary<string, TermStoreCopyService.CopyOutcome> _copiedTermSets = new(StringComparer.OrdinalIgnoreCase);

    public void ResetTermStoreRun()
    {
        _copiedTermSets.Clear();
        TermStoreNotes.Clear();
    }

    // Cross-tenant managed metadata: copy the term set (once per run), then create the column bound to the copy.
    // Throws with a clear reason when it cannot, so the column shows up as a failure instead of vanishing.
    private async Task CreateTaxonomyColumnWithTermCopyAsync(string siteUrl, string listId, ColumnDefinition col, string? sourceSiteUrl)
    {
        if (!CopyTermSets || TermStore == null || string.IsNullOrEmpty(sourceSiteUrl))
            throw new Exception(Loc.T("Svc_Lib_TaxonomyNotCopied", col.DisplayName));
        var srcSetId = TermStoreCopyService.ReadTermSetId(col.SchemaXml);
        if (string.IsNullOrEmpty(srcSetId))
            throw new Exception(Loc.T("Svc_Lib_NoTaxonomySchema", col.DisplayName));

        if (!_copiedTermSets.TryGetValue(srcSetId, out var outcome))
        {
            var tree = await TermStore.ReadSetAsync(sourceSiteUrl, srcSetId);
            outcome = await TermStore.EnsureSetAsync(siteUrl, tree, spService.TermMap);
            _copiedTermSets[srcSetId] = outcome;
            TermStoreNotes.Add(Loc.T("Svc_Lib_TermSetCopied", outcome.SetName, outcome.Created, outcome.Existing));
            TermStoreNotes.AddRange(outcome.Notes);
        }

        var rewritten = TermStoreCopyService.RewriteSchema(col.SchemaXml!, outcome.TargetStoreId, outcome.TargetSetId, srcSetId, spService.TermMap);
        var copy = new ColumnDefinition
        {
            InternalName = col.InternalName, DisplayName = col.DisplayName, FieldType = col.FieldType,
            Choices = col.Choices, SchemaXml = rewritten,
            LookupListId = col.LookupListId, LookupShowField = col.LookupShowField, LookupWebId = col.LookupWebId,
        };
        await CreateColumnFromSchemaXmlAsync(siteUrl, listId, copy);
    }

    // The GUID of a web (without braces), or null when it cannot be read.
    private async Task<string?> GetWebIdAsync(string webUrl)
    {
        var (ok, body) = await SpRestAsync(HttpMethod.Get, $"{webUrl.TrimEnd('/')}/_api/web?$select=Id", webUrl);
        if (!ok) return null;
        using var d = JsonDocument.Parse(body);
        return d.RootElement.TryGetProperty("Id", out var i) ? i.GetString()?.Trim('{', '}') : null;
    }

    // Creates a User or Taxonomy column via fields/createfieldasxml.
    // User columns get a clean handwritten schema. Taxonomy columns reuse the source
    // field's schema — within-tenant the term store is shared, so the SspId/TermSetId
    // binding is valid as-is. Instance-specific attributes are regenerated or stripped.
    private async Task CreateColumnFromSchemaXmlAsync(string siteUrl, string listId, ColumnDefinition col)
    {
        string schemaXml;
        if (col.FieldType == SupportedFieldType.Choice || col.FieldType == SupportedFieldType.MultiChoice)
        {
            var type = col.FieldType == SupportedFieldType.MultiChoice ? "MultiChoice" : "Choice";
            var el = new System.Xml.Linq.XElement("Field",
                new System.Xml.Linq.XAttribute("Type", type),
                new System.Xml.Linq.XAttribute("DisplayName", col.DisplayName),
                new System.Xml.Linq.XAttribute("Name", col.InternalName),
                new System.Xml.Linq.XAttribute("StaticName", col.InternalName),
                new System.Xml.Linq.XAttribute("Required", "FALSE"),
                new System.Xml.Linq.XAttribute("FillInChoice", "FALSE"));
            if (col.Choices is { Length: > 0 })
            {
                var choicesEl = new System.Xml.Linq.XElement("CHOICES");
                foreach (var choice in col.Choices)
                    choicesEl.Add(new System.Xml.Linq.XElement("CHOICE", choice));
                el.Add(choicesEl);
            }
            schemaXml = el.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        else if (ColumnDefinition.IsUserType(col.FieldType))
        {
            var el = new System.Xml.Linq.XElement("Field",
                new System.Xml.Linq.XAttribute("Type", col.FieldType == SupportedFieldType.UserMulti ? "UserMulti" : "User"),
                new System.Xml.Linq.XAttribute("DisplayName", col.DisplayName),
                new System.Xml.Linq.XAttribute("Name", col.InternalName),
                new System.Xml.Linq.XAttribute("StaticName", col.InternalName),
                new System.Xml.Linq.XAttribute("List", "UserInfo"),
                new System.Xml.Linq.XAttribute("ShowField", "ImnName"),
                new System.Xml.Linq.XAttribute("UserSelectionMode", "PeopleAndGroups"),
                new System.Xml.Linq.XAttribute("Required", "FALSE"));
            if (col.FieldType == SupportedFieldType.UserMulti)
                el.Add(new System.Xml.Linq.XAttribute("Mult", "TRUE"));
            schemaXml = el.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        else
        {
            if (string.IsNullOrEmpty(col.SchemaXml))
                throw new Exception(Loc.T("Svc_Lib_NoTaxonomySchema", col.DisplayName));

            var el = System.Xml.Linq.XElement.Parse(col.SchemaXml);
            // Fresh field ID; source-instance attributes don't apply at the target.
            el.SetAttributeValue("ID", $"{{{Guid.NewGuid()}}}");
            foreach (var attr in new[] { "SourceID", "ColName", "RowOrdinal", "Version", "WebId" })
                el.Attribute(attr)?.Remove();
            // TextField points at the source's hidden companion note field by GUID —
            // it doesn't exist at the target, so let SharePoint regenerate it.
            var textFieldProps = el.Descendants("Property")
                .Where(p => (string?)p.Element("Name") == "TextField")
                .ToList();
            foreach (var p in textFieldProps)
                p.Remove();
            schemaXml = el.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }

        await PostCreateFieldAsXmlAsync(siteUrl, listId, schemaXml, col);
    }

    private async Task PostCreateFieldAsXmlAsync(string siteUrl, string listId, string schemaXml, ColumnDefinition col)
    {
        var url  = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/fields/createfieldasxml";
        // Options 8 = use the Name in the XML as the internal name. Without it SharePoint derives the internal name from the
        // DisplayName ("AssignedTo" became "Assigned_x0020_to"), and every view, filter and value writer that uses the
        // source's internal name then misses the column.
        var json = JsonSerializer.Serialize(new { parameters = new { SchemaXml = schemaXml, Options = 8 } });

        using var resp = await spService.SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl);

        var body = await resp.Content.ReadAsStringAsync();
        if (!resp.IsSuccessStatusCode)
            throw new Exception(Loc.T("Svc_Lib_CreateFieldXmlHttp", col.DisplayName, (int)resp.StatusCode, body[..Math.Min(200, body.Length)]));

        var actualInternalName = col.InternalName;
        using (var doc = JsonDocument.Parse(body))
        {
            if (doc.RootElement.TryGetProperty("InternalName", out var n))
                actualInternalName = n.GetString() ?? col.InternalName;
        }
        await AddColumnToDefaultViewAsync(siteUrl, listId, actualInternalName);
    }

    private async Task AddColumnToDefaultViewAsync(string siteUrl, string listId, string internalName)
    {
        try
        {
            var viewUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/defaultView/viewfields/addViewField('{internalName}')";
            using var _ = await spService.SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, viewUrl);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent("{}", System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl);
        }
        catch { }
    }

    // Adds any source columns that are absent from the existing target list.
    // Returns (created display names, failed display names with reason).
    public async Task<(List<string> Created, List<string> Failed)> AddMissingColumnsAsync(
        string targetSiteUrl, string targetListId,
        IEnumerable<ColumnDefinition> sourceColumns,
        string? sourceSiteUrl = null)
    {
        // Skip cache so we always see the live column list — a stale cached snapshot
        // would cause columns that were just created to appear "missing" and get duplicated.
        var existing              = await spService.GetLibraryColumnsAsync(targetSiteUrl, targetListId, skipCache: true);
        var existingInternalNames = existing.Select(c => c.InternalName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var existingDisplayNames  = existing.Select(c => c.DisplayName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var created = new List<string>();
        var failed  = new List<string>();
        foreach (var col in sourceColumns.Where(c =>
            !existingInternalNames.Contains(c.InternalName) &&
            !existingDisplayNames.Contains(c.DisplayName)))
        {
            try
            {
                await CreateColumnAsync(targetSiteUrl, targetListId, col, sourceSiteUrl);
                created.Add(col.DisplayName);
            }
            catch (Exception ex)
            {
                failed.Add($"{col.DisplayName} ({ex.Message})");
            }
        }
        return (created, failed);
    }

    private static bool IsAlreadyExistsError(string body) =>
        body.Contains("-2130575342", StringComparison.Ordinal) ||
        body.Contains("already exists", StringComparison.OrdinalIgnoreCase);

    private async Task<LibraryDefinition> ReadListMetadataAsync(string siteUrl, string listId)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')" +
                  "?$select=Title,Description,EnableVersioning,EnableMinorVersions,MajorVersionLimit,RootFolder/Name&$expand=RootFolder";

        using var response = await spService.SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl);

        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"ReadListMetadata HTTP {(int)response.StatusCode}: {body[..Math.Min(200, body.Length)]}");

        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;
        return new LibraryDefinition
        {
            Title               = root.TryGetProperty("Title",               out var t)   ? t.GetString()   ?? "" : "",
            Description         = root.TryGetProperty("Description",         out var d)   ? d.GetString()   ?? "" : "",
            RootFolderName      = root.TryGetProperty("RootFolder", out var rf) && rf.ValueKind == JsonValueKind.Object &&
                                  rf.TryGetProperty("Name", out var rfn) ? rfn.GetString() ?? "" : "",
            EnableVersioning    = root.TryGetProperty("EnableVersioning",    out var ev)  && ev.GetBoolean(),
            EnableMinorVersions = root.TryGetProperty("EnableMinorVersions", out var emv) && emv.GetBoolean(),
            MajorVersionLimit   = root.TryGetProperty("MajorVersionLimit",   out var mvl) && mvl.TryGetInt32(out var mvlv) ? mvlv : 0,
        };
    }
}

internal sealed class LibraryAlreadyExistsException(
    string? driveId, string? serverRelativeUrl, string? listId)
    : Exception(Loc.T("Svc_Lib_AlreadyExists"))
{
    public string? DriveId           { get; } = driveId;
    public string? ServerRelativeUrl { get; } = serverRelativeUrl;
    public string? ListId            { get; } = listId;
}
