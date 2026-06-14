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
            {
                // Tier 1: Graph Drives API (normal libraries)
                var existing = (await spService.GetLibrariesAsync(targetSiteId, targetSiteUrl))
                    .FirstOrDefault(d => d.Name.Equals(definition.Title, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    var existServerRel = await spService.GetLibraryServerRelativeUrlAsync(existing.DriveId);
                    throw new LibraryAlreadyExistsException(existing.DriveId, existServerRel, listId: null);
                }

                // Tier 2: REST lookup — handles system libraries (Site Assets, Style Library) not in Drives API
                var node = await spService.GetLibraryNodeByTitleAsync(targetSiteId, targetSiteUrl, definition.Title);
                if (node != null)
                {
                    var serverRelUrl = node.ServerRelativePath
                        ?? (string.IsNullOrEmpty(node.DriveId) ? null
                            : await spService.GetLibraryServerRelativeUrlAsync(node.DriveId));
                    throw new LibraryAlreadyExistsException(node.DriveId, serverRelUrl, listId: null);
                }

                // Exists but unresolvable (e.g. Style Library with no Graph drive)
                throw new LibraryAlreadyExistsException(driveId: null, serverRelativeUrl: null, listId: null);
            }
            throw new Exception($"Create library '{definition.Title}' HTTP {(int)createResp.StatusCode}: {createRespBody[..Math.Min(200, createRespBody.Length)]}");
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
                throw new Exception($"Set versioning on '{definition.Title}' HTTP {(int)versionResp.StatusCode}: {errBody[..Math.Min(200, errBody.Length)]}");
            }
        }

        // Step 3: create custom columns
        // Build a lookup only when there are actual mapping decisions.
        // An empty mapping list means the user never opened the dialog — treat as "create all".
        var mappingEntries = columnMappings?
            .Where(m => m.TargetColumn != null || m.CreateNew)
            .ToDictionary(m => m.SourceColumn.InternalName);
        var mappingLookup = mappingEntries?.Count > 0 ? mappingEntries : null;

        var columnsToCreate = definition.Columns.Where(col =>
        {
            if (mappingLookup == null) return true;
            if (!mappingLookup.TryGetValue(col.InternalName, out var mapping)) return false;
            return mapping.CreateNew;
        }).ToList();

        foreach (var col in columnsToCreate)
        {
            try { await CreateColumnAsync(targetSiteUrl, listId, col); }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LibraryCopy] Column '{col.InternalName}' create failed: {ex.Message}");
            }
        }

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
            throw new Exception($"Created library '{definition.Title}' but could not find its drive ID in Graph.");

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
            throw new Exception($"Create list '{definition.Title}' HTTP {(int)createResp.StatusCode}: {createRespBody[..Math.Min(200, createRespBody.Length)]}");
        }
        else
        {
            using var doc = JsonDocument.Parse(createRespBody);
            listId = doc.RootElement.GetProperty("Id").GetString()!;
        }

        var existingCols          = await spService.GetLibraryColumnsAsync(targetSiteUrl, listId, skipCache: true);
        var existingInternalNames = existingCols.Select(c => c.InternalName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var existingDisplayNames  = existingCols.Select(c => c.DisplayName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var col in definition.Columns.Where(c =>
            !existingInternalNames.Contains(c.InternalName) &&
            !existingDisplayNames.Contains(c.DisplayName)))
        {
            try { await CreateColumnAsync(targetSiteUrl, listId, col); }
            catch { }
        }

        return listId;
    }

    private async Task CreateColumnAsync(string siteUrl, string listId, ColumnDefinition col)
    {
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
            SupportedFieldType.MultiChoice => 15,
            _ => 2,
        };
        var json = JsonSerializer.Serialize(new
        {
            FieldTypeKind = fieldTypeKind,
            Title         = col.DisplayName,
            StaticName    = col.InternalName,
            Required      = false,
        });

        // Create the field and read back the actual InternalName (SP may mangle spaces)
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
                $"SharePoint rejected column creation ({(int)createResp.StatusCode}): {createBody}");

        using var doc = JsonDocument.Parse(createBody);
        if (doc.RootElement.TryGetProperty("InternalName", out var n))
            actualInternalName = n.GetString() ?? col.InternalName;

        // Add to default view (non-critical)
        await AddColumnToDefaultViewAsync(siteUrl, listId, actualInternalName);
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
                throw new Exception($"No source schema captured for taxonomy column '{col.DisplayName}'");

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

        var url  = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/fields/createfieldasxml";
        var json = JsonSerializer.Serialize(new { parameters = new { SchemaXml = schemaXml } });

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
            throw new Exception($"createfieldasxml '{col.DisplayName}' HTTP {(int)resp.StatusCode}: {body[..Math.Min(200, body.Length)]}");

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
        IEnumerable<ColumnDefinition> sourceColumns)
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
                await CreateColumnAsync(targetSiteUrl, targetListId, col);
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
                  "?$select=Title,Description,EnableVersioning,EnableMinorVersions,MajorVersionLimit";

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
            EnableVersioning    = root.TryGetProperty("EnableVersioning",    out var ev)  && ev.GetBoolean(),
            EnableMinorVersions = root.TryGetProperty("EnableMinorVersions", out var emv) && emv.GetBoolean(),
            MajorVersionLimit   = root.TryGetProperty("MajorVersionLimit",   out var mvl) && mvl.TryGetInt32(out var mvlv) ? mvlv : 0,
        };
    }
}

internal sealed class LibraryAlreadyExistsException(
    string? driveId, string? serverRelativeUrl, string? listId)
    : Exception("Already exists — skipped")
{
    public string? DriveId           { get; } = driveId;
    public string? ServerRelativeUrl { get; } = serverRelativeUrl;
    public string? ListId            { get; } = listId;
}
