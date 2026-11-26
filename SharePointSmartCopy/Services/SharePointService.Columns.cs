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

// Partial class SharePointService: Columns. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
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
        { 10, SupportedFieldType.Currency },
        { 11, SupportedFieldType.Hyperlink },
        { 15, SupportedFieldType.MultiChoice },
        { 20, SupportedFieldType.User },
    };

    // Types that FieldTypeKind alone cannot distinguish (taxonomy fields report kind 0,
    // multi-user fields report kind 20 like single-user).
    private static SupportedFieldType? ResolveFieldType(int typeKind, string? typeAsString) => typeAsString switch
    {
        "UserMulti"              => SupportedFieldType.UserMulti,
        "TaxonomyFieldType"      => SupportedFieldType.Taxonomy,
        "TaxonomyFieldTypeMulti" => SupportedFieldType.TaxonomyMulti,
        "Lookup"                 => SupportedFieldType.Lookup,
        "LookupMulti"            => SupportedFieldType.LookupMulti,
        "URL"                    => SupportedFieldType.Hyperlink,
        _ => _fieldTypeMap.TryGetValue(typeKind, out var t) ? t : null,
    };

    // Keyed by listId. Concurrent — read/written from parallel copy tasks.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, List<SpColumnDef>> _columnCache = new();

    // Keyed by "{listId}|{showField}|{displayValue}" — caches resolved target lookup item IDs (hits only).
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int> _lookupValueCache = new();

    // Same key — remembers a genuine "no such value" for a short time only. A miss is NOT permanent:
    // the user may copy the lookup list's items and re-run (or, within one Site-scope run, the list
    // may be copied after the value was first asked for), and a forever-cached miss would blank that
    // value on every file that references it. Expiry bounds the repeated queries without that risk.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, DateTime> _lookupMissCache = new();
    private static readonly TimeSpan LookupMissTtl = TimeSpan.FromSeconds(90);

    // Called at the start of every copy run so a value resolved (or missed) in an earlier run — against
    // a target list that has since changed — is never reused.
    // Source item ID -> target item ID for the items this run copied, per TARGET list: a lookup value is resolved through it
    // first, so two items with the same title keep pointing at the right copy and a value does not depend on matching text.
    // Only items copied in this run are in it; anything else falls back to the display-text match below.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int> _lookupIdMap = new(StringComparer.OrdinalIgnoreCase);
    private int _lookupResolvedById, _lookupResolvedByText, _lookupUnresolved;

    public void RecordCopiedItem(string targetListId, int sourceItemId, int targetItemId) =>
        _lookupIdMap[$"{targetListId}|{sourceItemId}"] = targetItemId;

    // How lookup values were resolved this run (by copied item ID, by display text, not at all), for the report summary.
    public (int ById, int ByText, int Unresolved) LookupStats => (_lookupResolvedById, _lookupResolvedByText, _lookupUnresolved);

    public void ClearLookupCache()
    {
        _lookupValueCache.Clear();
        _lookupMissCache.Clear();
        _lookupIdMap.Clear();
        _lookupResolvedById = _lookupResolvedByText = _lookupUnresolved = 0;
        TermMap.Clear();
        _taxonomyRemapped = _taxonomySkipped = 0;
    }

    // Every editable, non-hidden column of a list, built-in ones such as Title included. GetLibraryColumnsAsync (below)
    // returns only the custom columns the copy tool can recreate; this is the full picture for matching flow columns.
    // Throws when SharePoint refuses, so a caller can say why the list of columns is empty.
    public async Task<List<(string InternalName, string Title, string Type, bool Required)>> GetListFieldsAsync(string siteUrl, string listId, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/fields" +
                  "?$filter=Hidden eq false and ReadOnlyField eq false&$select=InternalName,Title,TypeAsString,Required";
        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl, cancellationToken: ct);

        var body = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
            throw new Exception($"Reading the columns failed: HTTP {(int)response.StatusCode}. {(body.Length > 160 ? body[..160] : body)}");

        using var doc = JsonDocument.Parse(body);
        var result = new List<(string, string, string, bool)>();
        if (!doc.RootElement.TryGetProperty("value", out var values)) return result;
        foreach (var f in values.EnumerateArray())
        {
            var internalName = f.TryGetProperty("InternalName", out var ip) ? ip.GetString() ?? "" : "";
            if (internalName.Length == 0 || internalName == "ContentType") continue;   // not a column a flow stores values in
            var title = f.TryGetProperty("Title", out var tp) ? tp.GetString() ?? internalName : internalName;
            var type = f.TryGetProperty("TypeAsString", out var ts) ? ts.GetString() ?? "" : "";
            var required = f.TryGetProperty("Required", out var rq) && rq.ValueKind == JsonValueKind.True;
            result.Add((internalName, title, type, required));
        }
        return result;
    }
    // Returns the custom (non-built-in) columns for a library.
    public async Task<List<SpColumnDef>> GetLibraryColumnsAsync(string siteUrl, string listId, bool skipCache = false, CancellationToken ct = default)
    {
        if (!skipCache && _columnCache.TryGetValue(listId, out var cached))
            return cached;

        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/fields" +
                  "?$filter=Hidden eq false and ReadOnlyField eq false and FromBaseType eq false" +
                  "&$select=InternalName,Title,FieldTypeKind,TypeAsString,Choices,SchemaXml";

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl, cancellationToken: ct);

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
                var typeAsString = field.TryGetProperty("TypeAsString", out var tas) ? tas.GetString() : null;
                if (ResolveFieldType(typeKind, typeAsString) is not { } fieldType) continue;

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

                var schemaXml = field.TryGetProperty("SchemaXml", out var sx) ? sx.GetString() : null;

                string? lookupListId    = null;
                string? lookupShowField = null;
                string? lookupWebId     = null;
                if (SpColumnDef.IsLookupType(fieldType) && schemaXml != null)
                {
                    var webMatch = System.Text.RegularExpressions.Regex.Match(schemaXml, @"WebId=""(\{?[0-9A-Fa-f\-]+\}?)""");
                    if (webMatch.Success) lookupWebId = webMatch.Groups[1].Value.Trim('{', '}');
                    var listMatch = System.Text.RegularExpressions.Regex.Match(schemaXml, @"List=""(\{?[0-9A-Fa-f\-]+\}?)""");
                    if (listMatch.Success)
                        lookupListId = listMatch.Groups[1].Value.Trim('{', '}');
                    var showMatch = System.Text.RegularExpressions.Regex.Match(schemaXml, @"ShowField=""([^""]+)""");
                    if (showMatch.Success)
                        lookupShowField = showMatch.Groups[1].Value;
                }

                result.Add(new SpColumnDef
                {
                    InternalName    = internalName,
                    DisplayName     = field.GetProperty("Title").GetString() ?? internalName,
                    FieldType       = fieldType,
                    Choices         = choices,
                    SchemaXml       = schemaXml,
                    LookupListId    = lookupListId,
                    LookupShowField = lookupShowField,
                    LookupWebId     = lookupWebId,
                    LookupListTitle = lookupListId != null ? await GetListTitleByIdAsync(siteUrl, lookupListId, ct) : null,
                });
            }
        }

        _columnCache[listId] = result;
        return result;
    }

    // Bulk-reads custom field values for all items in a library in one paginated pass.
    // Returns a dictionary keyed by SP list item integer ID (as string, e.g. "42").
    // This key matches the listItemId returned by GetSharePointIdsAsync, enabling O(1) cache lookup per file.
    //
    // `systemMetadataOut`, when given, is populated with each item's Modified/Created/Editor/Author
    // in the SAME paginated scan — riding along on chunk 0's request rather than costing a separate
    // per-file Graph call later. This is what the Migration API post-import custom-fields pass uses
    // to restamp Modified/Editor without a GetFileMetadataAsync call per file (that call previously
    // ran once per file with custom columns, on top of the custom-field write itself).
    public async Task<Dictionary<string, Dictionary<string, object?>>> BulkReadCustomFieldsAsync(
        string siteUrl, string listId, IEnumerable<SpColumnDef> columns,
        IProgress<int>? progress = null, IProgress<string>? warningLog = null,
        Dictionary<string, FileMetadata>? systemMetadataOut = null,
        CancellationToken ct = default)
    {
        var cols = columns.ToList();

        // Chunk columns into groups of 20 to stay within URL length limits. A library with no custom columns still gets one
        // (empty) pass, which reads each item's content type.
        var result = new Dictionary<string, Dictionary<string, object?>>();
        var chunks  = cols.Count == 0 ? [Array.Empty<SpColumnDef>()] : cols.Chunk(20).ToList();

        for (int chunk = 0; chunk < chunks.Count; chunk++)
        {
            // Mutable per-chunk column list: some fields exist in the fields collection (not
            // Hidden, not ReadOnly) yet SharePoint's item-level REST rejects them with "The field
            // or property 'X' does not exist" — e.g. '_ExtendedDescription', a phantom backing
            // field observed 2026-08-26. A single such field previously poisoned the WHOLE chunk's
            // request with an HTTP 400, silently emptying the bulk field cache for every item and
            // every OTHER column in that chunk too. Now: on that specific error, drop the offending
            // field from this chunk and retry, instead of losing the entire chunk.
            var chunkCols = chunks[chunk].ToList();
            var excludedFields = new List<string>();
            // Set when the lookup sub-select (Id + show field) is rejected: lookups then read Id only.
            var lookupIdOnly = false;

            while (chunkCols.Count > 0 || cols.Count == 0)   // the no-column pass runs once: it never asks for a retry
            {
                // User fields are lookups: they must be $expand-ed and read via {name}/Name
                // (the claims login). Other field types are selected directly.
                var selectParts = new List<string> { "ID" };
                var expandParts = new List<string>();
                // System metadata only needs to ride on ONE chunk's request, not every chunk —
                // chunk 0 is as good as any, and every item appears in every chunk's page anyway.
                if (chunk == 0 && systemMetadataOut != null)
                {
                    // Author/Name and Editor/Name (the claims login) ride in the same request: an account without an email has
                    // an empty EMail, and its UPN, taken from the login, is then used instead of dropping the person.
                    selectParts.AddRange(["Modified", "Editor/EMail", "Editor/Name", "Created", "Author/EMail", "Author/Name"]);
                    expandParts.AddRange(["Editor", "Author"]);
                }
                // Each item's content type name rides along on chunk 0 too (see DefaultContentTypeNames below).
                if (chunk == 0)
                {
                    selectParts.Add("ContentType/Name");
                    expandParts.Add("ContentType");
                }
                foreach (var c in chunkCols)
                {
                    if (SpColumnDef.IsUserType(c.FieldType))
                    {
                        selectParts.Add($"{RestFieldName(c.InternalName)}/Name");
                        expandParts.Add(RestFieldName(c.InternalName));
                    }
                    else if (SpColumnDef.IsLookupType(c.FieldType))
                    {
                        AppendLookupSelect(selectParts, expandParts, c, lookupIdOnly);
                    }
                    else
                    {
                        selectParts.Add(RestFieldName(c.InternalName));
                    }
                }
                var nextUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                              $"?$select={Uri.EscapeDataString(string.Join(",", selectParts))}" +
                              (expandParts.Count > 0 ? $"&$expand={Uri.EscapeDataString(string.Join(",", expandParts))}" : "") +
                              "&$top=1000";

                bool retryWithFieldExcluded = false;
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
                    if (!response.IsSuccessStatusCode)
                    {
                        var badField = TryExtractMissingFieldName(body);
                        if (badField != null &&
                            chunkCols.RemoveAll(c => string.Equals(c.InternalName, badField, StringComparison.OrdinalIgnoreCase)
                                                  || string.Equals(RestFieldName(c.InternalName), badField, StringComparison.OrdinalIgnoreCase)) > 0)
                        {
                            excludedFields.Add(badField);
                            retryWithFieldExcluded = true;
                            break;
                        }

                        // A lookup whose show field can't be projected (400): retry reading IDs only.
                        if ((int)response.StatusCode == 400 && !lookupIdOnly &&
                            chunkCols.Any(c => SpColumnDef.IsLookupType(c.FieldType)))
                        {
                            lookupIdOnly = true;
                            retryWithFieldExcluded = true;
                            break;
                        }

                        // A failed page silently dropped every item beyond it from the bulk cache —
                        // their files then copied with NO custom column values and reported Success.
                        throw new HttpRequestException(
                            Loc.T("Svc_Sp_BulkCustomFieldRead", (int)response.StatusCode, body[..Math.Min(body.Length, 300)]));
                    }

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

                            foreach (var col in chunkCols)
                            {
                                if (!item.TryGetProperty(RestFieldName(col.InternalName), out var valProp)) continue;
                                fields[col.InternalName] = ParseFieldValue(valProp, col.FieldType, col.LookupShowField);
                            }

                            // A custom content type is kept under "ContentType" so the write can put the copy on the same type.
                            // The built-in defaults are not kept: every file has one, and it would cost an extra write each.
                            if (chunk == 0 && item.TryGetProperty("ContentType", out var ctProp) && ctProp.ValueKind == JsonValueKind.Object &&
                                ctProp.TryGetProperty("Name", out var ctName) && ctName.GetString() is { Length: > 0 } ctn &&
                                !DefaultContentTypeNames.Contains(ctn))
                                fields[ContentTypeFieldKey] = ctn;

                            if (chunk == 0 && systemMetadataOut != null)
                            {
                                DateTimeOffset? modified = item.TryGetProperty("Modified", out var mp) &&
                                    mp.ValueKind == JsonValueKind.String &&
                                    DateTimeOffset.TryParse(mp.GetString(), System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None, out var mv) ? mv : null;
                                DateTimeOffset? created = item.TryGetProperty("Created", out var cp) &&
                                    cp.ValueKind == JsonValueKind.String &&
                                    DateTimeOffset.TryParse(cp.GetString(), System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None, out var cv) ? cv : null;
                                static string? PersonOf(JsonElement item, string field)
                                {
                                    if (!item.TryGetProperty(field, out var p) || p.ValueKind != JsonValueKind.Object) return null;
                                    var mail = p.TryGetProperty("EMail", out var m) && m.ValueKind == JsonValueKind.String ? m.GetString() : null;
                                    if (!string.IsNullOrWhiteSpace(mail)) return mail;
                                    return UpnFromLogin(p.TryGetProperty("Name", out var n) && n.ValueKind == JsonValueKind.String ? n.GetString() : null);
                                }
                                string? modifiedByEmail = PersonOf(item, "Editor");
                                string? createdByEmail  = PersonOf(item, "Author");

                                systemMetadataOut[itemId] = new FileMetadata
                                {
                                    ModifiedDateTime = modified,
                                    ModifiedByEmail  = modifiedByEmail,
                                    CreatedDateTime  = created,
                                    CreatedByEmail   = createdByEmail,
                                };
                            }
                        }
                    }

                    nextUrl = GetNextLink(root);
                }

                if (!retryWithFieldExcluded) break;
            }

            if (excludedFields.Count > 0)
                warningLog?.Report(Loc.T("Svc_Sp_CustomColumnSCould", excludedFields.Count) + " "
                    + Loc.T("Svc_Sp_NotValidAtThe", string.Join(", ", excludedFields)));

            progress?.Report((chunk + 1) * 100 / chunks.Count);
        }

        return result;
    }

    // Parses "The field or property 'X' does not exist." out of a ValidateUpdateListItem/items()
    // error body — used to drop a single phantom field from a bulk read instead of failing it whole.
    private static string? TryExtractMissingFieldName(string body)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("odata.error", out var err) &&
                err.TryGetProperty("message", out var msg) &&
                msg.TryGetProperty("value", out var val))
            {
                var text = val.GetString() ?? "";
                var m = System.Text.RegularExpressions.Regex.Match(text, @"The field or property '([^']+)' does not exist\.");
                if (m.Success) return m.Groups[1].Value;
            }
        }
        catch { /* not the expected error shape */ }
        return null;
    }

    // Parses "Column 'X' does not exist. It may have been deleted by another user." out of a
    // ValidateUpdateListItem error body — a different message shape than TryExtractMissingFieldName's,
    // thrown when a submitted FieldName has no matching column on the TARGET list at all (e.g. source
    // and target columns share a display name but not an internal name). Used to drop just that field
    // and retry instead of failing every field on the item.
    private static string? TryExtractNonexistentColumnName(string body)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("odata.error", out var err) &&
                err.TryGetProperty("message", out var msg) &&
                msg.TryGetProperty("value", out var val))
            {
                var text = val.GetString() ?? "";
                var m = System.Text.RegularExpressions.Regex.Match(text, @"Column '([^']+)' does not exist\.");
                if (m.Success) return m.Groups[1].Value;
            }
        }
        catch { /* not the expected error shape */ }
        return null;
    }

    // The property a lookup's display text is projected under: its ShowField, with the computed
    // title variants (LinkTitle/LinkTitleNoMenu) and an unset ShowField reading as "Title".
    private static string LookupQueryField(string? showField) =>
        string.IsNullOrEmpty(showField) || showField is "LinkTitle" or "LinkTitleNoMenu" ? "Title" : showField;

    // Adds a lookup column to an items() $select/$expand. With $expand, "{name}/X" addresses field X
    // on the lookup's TARGET list — "LookupId"/"LookupValue" are not fields there (HTTP 400), so the
    // entry is projected as {name}/Id plus the show field. idOnly drops the show field (fallback for
    // a show field SharePoint refuses to project); the display text is then empty.
    private static void AppendLookupSelect(List<string> selectParts, List<string> expandParts, SpColumnDef col, bool idOnly = false)
    {
        var name = RestFieldName(col.InternalName);
        selectParts.Add($"{name}/Id");
        if (!idOnly) selectParts.Add($"{name}/{LookupQueryField(col.LookupShowField)}");
        expandParts.Add(name);
    }

    // SharePoint's REST layer names a field whose internal name starts with an underscore with an "OData_" prefix: a column
    // titled "5 YearsService?" has the internal name "_x0035_YearsService_x003f_" but is selected, expanded and returned as
    // "OData__x0035_YearsService_x003f_". The bare name is answered with "The field or property '...' does not exist", which
    // used to stop the whole read. (Writes through ValidateUpdateListItem use the plain internal name.)
    internal static string RestFieldName(string internalName) =>
        internalName.StartsWith('_') ? "OData_" + internalName : internalName;

    private static object? ParseFieldValue(JsonElement el, SupportedFieldType type, string? lookupShowField = null)
    {
        if (el.ValueKind == JsonValueKind.Null || el.ValueKind == JsonValueKind.Undefined)
            return null;

        if (SpColumnDef.IsUserType(type))
        {
            var logins = EnumerateComplexValues(el)
                .Select(u => u.TryGetProperty("Name", out var n) ? n.GetString() : null)
                .Where(s => !string.IsNullOrEmpty(s))
                .Select(s => s!)
                .ToArray();
            return logins.Length > 0 ? new PersonFieldValue(logins) : null;
        }

        if (SpColumnDef.IsTaxonomyType(type))
        {
            var terms = EnumerateComplexValues(el)
                .Select(t => (
                    Label:    t.TryGetProperty("Label",    out var l) ? l.GetString() ?? "" : "",
                    TermGuid: t.TryGetProperty("TermGuid", out var g) ? g.GetString() ?? "" : ""))
                .Where(t => t.TermGuid.Length > 0)
                .ToArray();
            return terms.Length > 0 ? new TaxonomyFieldValue(terms) : null;
        }

        if (SpColumnDef.IsLookupType(type))
        {
            var showProp = LookupQueryField(lookupShowField);
            var entries = EnumerateComplexValues(el)
                .Select(e => (
                    Id:           e.TryGetProperty("Id", out var id) && id.ValueKind == JsonValueKind.Number ? id.GetInt32() : 0,
                    DisplayValue: e.TryGetProperty(showProp, out var dv) && dv.ValueKind != JsonValueKind.Null
                                      ? (dv.ValueKind == JsonValueKind.String ? dv.GetString() ?? "" : dv.ToString())
                                      : ""))
                .Where(e => e.Id > 0)
                .ToArray();
            return entries.Length > 0 ? new LookupFieldValue(entries) : null;
        }

        if (type == SupportedFieldType.Hyperlink)
        {
            // REST returns { "Url": "...", "Description": "..." } (an object even in nometadata mode).
            if (el.ValueKind != JsonValueKind.Object) return null;
            var url  = el.TryGetProperty("Url", out var u) && u.ValueKind == JsonValueKind.String ? u.GetString() : null;
            var desc = el.TryGetProperty("Description", out var d) && d.ValueKind == JsonValueKind.String ? d.GetString() : null;
            return string.IsNullOrWhiteSpace(url) ? null : new HyperlinkFieldValue(url!, desc);
        }

        return type switch
        {
            SupportedFieldType.MultiChoice when el.ValueKind == JsonValueKind.Array =>
                string.Join(";#", el.EnumerateArray().Select(v => v.GetString() ?? "")),
            SupportedFieldType.MultiChoice when el.ValueKind == JsonValueKind.Object &&
                el.TryGetProperty("results", out var r) =>
                string.Join(";#", r.EnumerateArray().Select(v => v.GetString() ?? "")),
            SupportedFieldType.Boolean =>
                el.ValueKind == JsonValueKind.True,
            SupportedFieldType.DateTime when el.ValueKind == JsonValueKind.String =>
                el.GetString(),
            _ => el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString()
        };
    }

    // Complex field values (user, taxonomy) arrive as a single object, a bare array,
    // or a verbose { "results": [...] } wrapper depending on OData mode and multiplicity.
    private static IEnumerable<JsonElement> EnumerateComplexValues(JsonElement el)
    {
        if (el.ValueKind == JsonValueKind.Array)
            return el.EnumerateArray();
        if (el.ValueKind == JsonValueKind.Object && el.TryGetProperty("results", out var results) &&
            results.ValueKind == JsonValueKind.Array)
            return results.EnumerateArray();
        if (el.ValueKind == JsonValueKind.Object)
            return [el];
        return [];
    }

    // Finds the ID of an item in a lookup list by its display value.
    // ShowField values "LinkTitle"/"LinkTitleNoMenu" are queried as "Title".
    // Returns null when the display value cannot be matched. Several target items can share a display
    // value; the lowest ID wins (ordered, so the choice is stable rather than arbitrary).
    private async Task<int?> ResolveLookupValueAsync(
        string siteUrl, string lookupListId, string showField, string displayValue,
        CancellationToken ct = default)
    {
        var queryField = LookupQueryField(showField);
        var cacheKey   = $"{lookupListId}|{queryField}|{displayValue}";
        if (_lookupValueCache.TryGetValue(cacheKey, out var cached)) return cached;
        if (_lookupMissCache.TryGetValue(cacheKey, out var missedAt) && DateTime.UtcNow - missedAt < LookupMissTtl)
            return null;

        var escaped = Uri.EscapeDataString(displayValue.Replace("'", "''"));
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{lookupListId}')/items" +
                  $"?$select=ID&$filter={Uri.EscapeDataString(queryField)} eq '{escaped}'&$orderby=ID&$top=1";

        using var response = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Get, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            return r;
        }, siteUrl, cancellationToken: ct);

        int? result = null;
        if (response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var items) &&
                items.GetArrayLength() > 0 &&
                items[0].TryGetProperty("ID", out var id))
            {
                result = id.GetInt32();
            }
            // Cache only on a SUCCESSFUL request. A match is cached for the run; a genuine miss
            // (the query succeeded but found nothing) is remembered briefly (see _lookupMissCache). A
            // failed request (429/500/etc., post-retry) is NOT cached at all: it could be transient
            // rather than "this value doesn't exist", and caching it would blank that lookup value for
            // every remaining file that references it.
            if (result.HasValue) _lookupValueCache[cacheKey] = result.Value;
            else                 _lookupMissCache[cacheKey]  = DateTime.UtcNow;
        }
        return result;
    }

    // The web that holds a lookup column's list: the column's own web unless its schema names another one (WebId).
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _lookupWebUrls = new(StringComparer.OrdinalIgnoreCase);
    private async Task<string> ResolveLookupWebUrlAsync(string siteUrl, SpColumnDef tgtCol, CancellationToken ct)
    {
        var webId = tgtCol.LookupWebId;
        if (string.IsNullOrEmpty(webId)) return siteUrl;
        var key = $"{siteUrl.TrimEnd('/')}|{webId}";
        if (_lookupWebUrls.TryGetValue(key, out var known)) return known;
        string result = siteUrl;
        try
        {
            using var own = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, $"{siteUrl.TrimEnd('/')}/_api/web?$select=Id,Url");
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            var ownId = own.IsSuccessStatusCode ? JsonDocument.Parse(await own.Content.ReadAsStringAsync(ct)).RootElement.GetProperty("Id").GetString()?.Trim('{', '}') : null;
            if (ownId != null && !string.Equals(ownId, webId, StringComparison.OrdinalIgnoreCase))
            {
                using var other = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, $"{siteUrl.TrimEnd('/')}/_api/site/openWebById('{webId}')?$select=Url");
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl, cancellationToken: ct);
                if (other.IsSuccessStatusCode &&
                    JsonDocument.Parse(await other.Content.ReadAsStringAsync(ct)).RootElement.TryGetProperty("Url", out var u) && u.GetString() is { Length: > 0 } url)
                    result = url;
            }
        }
        catch (Exception ex) when (ex is HttpRequestException or JsonException) { /* fall back to the column's own web */ }
        _lookupWebUrls[key] = result;
        return result;
    }

    // Message for a lookup field that can't be written because the TARGET column is missing, isn't a
    // Lookup column at all (e.g. a same-named Text column already there), or has no bound list.
    private static string DescribeUnusableLookupTarget(string targetName, SpColumnDef? tgtCol) =>
        tgtCol != null && !SpColumnDef.IsLookupType(tgtCol.FieldType)
            ? Loc.T("Svc_Sp_TargetNotLookup", targetName, tgtCol.FieldType)
            : Loc.T("Svc_Sp_LookupUnresolved", targetName);

    // Resolves a source lookup value to the ValidateUpdateListItem wire format against the TARGET
    // column's bound list: single "3", multi "3;#;#5;#" (IDs only — display text is not accepted).
    // Returns null when nothing resolved. Total and PARTIAL failures (e.g. 3 of 4 entries of a
    // multi-value lookup) are appended to lookupErrors; previously a partial resolution submitted
    // only the resolved entries with no warning that one was dropped.
    private async Task<string?> ResolveLookupFormValueAsync(
        string siteUrl, SpColumnDef tgtCol, LookupFieldValue lookup, string targetName,
        List<string> lookupErrors, CancellationToken ct)
    {
        var showField   = string.IsNullOrEmpty(tgtCol.LookupShowField) ? "Title" : tgtCol.LookupShowField;
        var resolvedIds = new List<int>();
        // A lookup to a list in a child web: the items are in that web, not in the column's own web.
        var lookupWeb   = await ResolveLookupWebUrlAsync(siteUrl, tgtCol, ct);
        foreach (var (sourceItemId, display) in lookup.Entries)
        {
            // The item this run copied wins over a text match (duplicate titles, items whose text changed).
            if (_lookupIdMap.TryGetValue($"{tgtCol.LookupListId}|{sourceItemId}", out var mappedId))
            {
                resolvedIds.Add(mappedId);
                Interlocked.Increment(ref _lookupResolvedById);
                continue;
            }
            var targetId = await ResolveLookupValueAsync(lookupWeb, tgtCol.LookupListId!, showField, display, ct);
            if (targetId.HasValue) { resolvedIds.Add(targetId.Value); Interlocked.Increment(ref _lookupResolvedByText); }
            else Interlocked.Increment(ref _lookupUnresolved);
        }

        if (resolvedIds.Count < lookup.Entries.Length)
            lookupErrors.Add(resolvedIds.Count == 0
                ? Loc.T("Svc_Sp_OfValuesResolved", targetName, lookup.Entries.Length)
                : Loc.T("Svc_Sp_OfValuesUnresolved", targetName, lookup.Entries.Length - resolvedIds.Count, lookup.Entries.Length));
        if (resolvedIds.Count == 0) return null;

        return resolvedIds.Count == 1
            ? resolvedIds[0].ToString()
            : string.Join(";#;#", resolvedIds.Select(id => id.ToString())) + ";#";
    }

    // The key under which an item's (custom) content type name travels in the custom-field dictionary. "ContentType" is on
    // the excluded-columns list, so it can never collide with a real column.
    internal const string ContentTypeFieldKey = "ContentType";
    internal static readonly HashSet<string> DefaultContentTypeNames = new(StringComparer.OrdinalIgnoreCase) { "Document", "Folder", "Item" };

    // The content type names of a target list, loaded once per list. Used to put a copied file on its type only when the
    // target list has that type, so a type that was not copied never turns into an error on every file.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, Task<Dictionary<string, HashSet<string>>>> _targetContentTypeFields = new();

    // The content types of a target list with the columns each one carries (loaded once per list). Columns cannot be linked to
    // a content type over REST, so a copied file is put on a type only when that type carries every column being written to it:
    // a column the item's type lacks is silently not written.
    internal Task<Dictionary<string, HashSet<string>>> GetTargetContentTypeFieldsAsync(string siteUrl, string listId, CancellationToken ct) =>
        _targetContentTypeFields.GetOrAdd($"{siteUrl.TrimEnd('/')}|{listId}", async _ =>
        {
            var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            async Task<JsonDocument?> GetAsync(string url)
            {
                using var resp = await SendSharePointRequestAsync(token =>
                {
                    var r = new HttpRequestMessage(HttpMethod.Get, url);
                    r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                    return r;
                }, siteUrl, cancellationToken: ct);
                return resp.IsSuccessStatusCode ? JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct)) : null;
            }
            try
            {
                var baseUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/contenttypes";
                using var list = await GetAsync($"{baseUrl}?$select=Name,StringId");
                if (list is null) return result;
                foreach (var e in list.RootElement.GetProperty("value").EnumerateArray())
                {
                    var name = e.TryGetProperty("Name", out var n) ? n.GetString() : null;
                    var id = e.TryGetProperty("StringId", out var i) ? i.GetString() : null;
                    if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(id)) continue;
                    var links = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using var fl = await GetAsync($"{baseUrl}('{id}')/fieldlinks?$select=Name");
                    if (fl is not null)
                        foreach (var l in fl.RootElement.GetProperty("value").EnumerateArray())
                            if (l.TryGetProperty("Name", out var ln) && ln.GetString() is { Length: > 0 } s) links.Add(s);
                    result[name] = links;
                }
            }
            catch (Exception ex) when (ex is HttpRequestException or JsonException) { /* no types known: nothing is bound */ }
            return result;
        });

    internal async Task<HashSet<string>> GetTargetContentTypeNamesAsync(string siteUrl, string listId, CancellationToken ct) =>
        new((await GetTargetContentTypeFieldsAsync(siteUrl, listId, ct)).Keys, StringComparer.OrdinalIgnoreCase);

    // Applies custom field values to a list item via ValidateUpdateListItem.
    // mappings translates source InternalName → target InternalName.
    public async Task<string?> ApplyFileCustomFieldsAsync(
        string driveId, string itemId,
        Dictionary<string, object?> fields,
        IEnumerable<SpColumnMap> mappings,
        CancellationToken ct = default)
    {
        if (fields.Count == 0) return null;

        // Resolve target SharePoint IDs first so we can look up target column definitions
        // (needed to find the target lookup list GUID for Lookup/LookupMulti fields).
        var ids = await GetSharePointIdsAsync(driveId, itemId, ct)
            ?? throw new Exception(Loc.T("Svc_Sp_CouldNotResolveSharePoint", driveId, itemId));

        return await ApplyFileCustomFieldsCoreAsync(ids.siteUrl, ids.listId, ids.listItemId, fields, mappings, null, ct);
    }

    // Same as ApplyFileCustomFieldsAsync, but resolves the target item via its REST server-relative
    // path instead of Graph driveId/itemId. Used for the Migration API post-import custom-fields
    // pass: SPMI-imported files have no Graph item ID to resolve (the migration API doesn't surface
    // one — see the permissions post-pass in CopyService for the same constraint), only the
    // server-relative path the file was imported to.
    //
    // `restamp`, when given, folds a Modified/Created/Editor/Author correction into the SAME
    // ValidateUpdateListItem call as the custom-field write, instead of CopyService issuing a
    // separate restamp call afterward. Migration API mode needs this restamp because the custom-
    // field write itself bumps Modified/Editor to the migrating account — see the caller. Combining
    // the two avoids resolving `serverRelativeUrl`'s list item ID twice and posting twice.
    public async Task<string?> ApplyFileCustomFieldsByPathAsync(
        string siteUrl, string listId, string serverRelativeUrl,
        Dictionary<string, object?> fields,
        IEnumerable<SpColumnMap> mappings,
        FileMetadata? restamp = null,
        CancellationToken ct = default)
    {
        if (fields.Count == 0 && restamp == null) return null;

        var listItemId = await GetFileListItemIdAsync(siteUrl, serverRelativeUrl, ct)
            ?? throw new Exception(Loc.T("Svc_Sp_CouldNotResolveList", serverRelativeUrl));

        return await ApplyFileCustomFieldsCoreAsync(siteUrl, listId, listItemId.ToString(), fields, mappings, restamp, ct);
    }

    // Returns a file's list item ID (not its UniqueId GUID) by server-relative path via REST.
    // Mirrors the GetFolderById(...)/ListItemAllFields/Id pattern used for folders (see
    // PatchFolderMetadataAsync) but resolves directly from a path in one round trip, matching
    // GetFileUniqueIdAsync's *ByServerRelativePath approach so '#'/'%'/'+' in names are unaffected.
    private async Task<int?> GetFileListItemIdAsync(string siteUrl, string serverRelativeUrl, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/GetFileByServerRelativePath({ServerRelativePathArg(serverRelativeUrl)})/ListItemAllFields/Id?{ServerRelativePathAliasQuery(serverRelativeUrl)}";
        try
        {
            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);

            if (!response.IsSuccessStatusCode) return null;
            var body = await response.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.ValueKind == JsonValueKind.Number
                ? doc.RootElement.GetInt32()
                : doc.RootElement.TryGetProperty("value", out var v) ? v.GetInt32() : null;
        }
        catch { return null; }
    }

    private async Task<string?> ApplyFileCustomFieldsCoreAsync(
        string siteUrl, string listId, string listItemId,
        Dictionary<string, object?> fields,
        IEnumerable<SpColumnMap> mappings,
        FileMetadata? restamp,
        CancellationToken ct)
    {
        var mappingLookup = SpColumnMap.BuildTargetNameMap(mappings);

        // Fetch target column defs (cached) to resolve lookup list GUIDs.
        List<SpColumnDef> targetCols;
        try { targetCols = await GetLibraryColumnsAsync(siteUrl, listId); }
        catch { targetCols = []; }
        var targetColsByName = targetCols.ToDictionary(c => c.InternalName, StringComparer.OrdinalIgnoreCase);

        var formValues      = new List<(string FieldName, string FieldValue)>();
        var submittedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var lookupErrors    = new List<string>();
        foreach (var (srcName, value) in fields)
        {
            if (value == null || srcName == ContentTypeFieldKey) continue;
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

            // Lookup fields: resolve each display value to a target item ID.
            if (value is LookupFieldValue lookup)
            {
                SpColumnDef? tgtCol = null;
                if (!targetColsByName.TryGetValue(targetName, out tgtCol) ||
                    string.IsNullOrEmpty(tgtCol.LookupListId))
                {
                    lookupErrors.Add(DescribeUnusableLookupTarget(targetName, tgtCol));
                    continue;
                }

                var formatted = await ResolveLookupFormValueAsync(siteUrl, tgtCol, lookup, targetName, lookupErrors, ct);
                if (formatted == null) continue; // nothing resolved — skip field
                formValues.Add((targetName, formatted));
                submittedFields.Add(targetName);
                continue;
            }

            var taxValue = RemapTaxonomyValue(value);
            if (taxValue == null) continue;

            // Cross-tenant: leave out (and record) people who do not exist in the target.
            var checkedValue = await DropUnresolvedPeopleAsync(siteUrl, taxValue, ct);
            if (checkedValue == null) continue;

            var formattedValue = FormatFieldValueForValidate(checkedValue);
            // A date column is read as a UTC ISO string ("…T…Z"), which ValidateUpdateListItem rejects ("Enter a date like
            // 2/23/2012"); it wants the site-local form. Found by the live harness: the rejected date made the whole
            // all-or-nothing write fail, so EVERY custom column value on that file was lost, not only the date.
            if (IsoUtcDate.IsMatch(formattedValue)) formattedValue = await ToSiteLocalFormDateAsync(siteUrl, formattedValue);
            formValues.Add((targetName, formattedValue));
            submittedFields.Add(targetName);
        }

        // The content type goes FIRST, so the other values are written against it (confirmed on a live site: setting it by name
        // through ValidateUpdateListItem works on files, folders and list items). Only when the target list has a type of that
        // name AND that type carries every column being written; otherwise the type is left off rather than lose a value.
        if (fields.TryGetValue(ContentTypeFieldKey, out var ctValue) && ctValue is string ctName &&
            (await GetTargetContentTypeFieldsAsync(siteUrl, listId, ct)).TryGetValue(ctName, out var ctColumns) &&
            submittedFields.All(ctColumns.Contains))
        {
            formValues.Insert(0, (ContentTypeFieldKey, ctName));
            submittedFields.Add(ContentTypeFieldKey);
        }

        // Fold a Modified/Created/Editor/Author correction into this SAME ValidateUpdateListItem
        // call instead of a separate call after this one returns — the write itself bumps
        // Modified/Editor to the calling account (see the Migration API caller), so the correction
        // must ride along in this call to land in the version this write actually creates, and doing
        // so in one call instead of two saves one full round trip (and, via ApplyFileCustomFieldsByPathAsync,
        // one duplicate list-item-ID resolution) per file.
        if (restamp != null)
        {
            var dateAndPersonValues = await BuildDateAndPersonFormValuesAsync(
                siteUrl, restamp.ModifiedDateTime, restamp.CreatedDateTime, restamp.CreatedByEmail, restamp.ModifiedByEmail);
            foreach (var (name, val) in dateAndPersonValues)
            {
                formValues.Add((name, val));
                submittedFields.Add(name);
            }
        }

        if (formValues.Count == 0) return lookupErrors.Count > 0 ? string.Join(", ", lookupErrors) : null;

        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({listItemId})/ValidateUpdateListItem()";

        // A field name that's correct on SOURCE but doesn't exist on TARGET (e.g. columns that
        // only look identical — same display name/values in the UI but a different internal name,
        // observed 2026-08-26 with a "Notes" column) makes ValidateUpdateListItem reject the WHOLE
        // request with a request-level 400 — unlike a content-type mismatch, which still returns
        // 200 with per-field results. One bad field previously blocked every other valid field on
        // the same item. Retry with that field dropped instead, same resilience pattern as
        // BulkReadCustomFieldsAsync's phantom-field handling above.
        var droppedFields = new List<string>();
        string body;
        while (true)
        {
            var payload = JsonSerializer.Serialize(new
            {
                formValues = formValues.Select(f => new { f.FieldName, f.FieldValue }),
                bNewDocumentUpdate = true,
            });

            using var response = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Post, url);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                r.Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");
                return r;
            }, siteUrl, cancellationToken: ct);

            body = await response.Content.ReadAsStringAsync(ct);
            if (response.IsSuccessStatusCode) break;

            var badColumn = TryExtractNonexistentColumnName(body);
            if (badColumn != null &&
                formValues.RemoveAll(f => string.Equals(f.FieldName, badColumn, StringComparison.OrdinalIgnoreCase)) > 0)
            {
                droppedFields.Add(badColumn);
                submittedFields.Remove(badColumn);
                if (formValues.Count == 0)
                    return Loc.T("Svc_Sp_CustomFieldErrors", string.Join(", ", droppedFields.Select(n => Loc.T("Svc_Sp_ColumnMissingOnTarget", n))))
                        + (lookupErrors.Count > 0 ? Loc.T("Svc_Sp_LookupUnresolved2", string.Join(", ", lookupErrors)) : "");
                continue;
            }

            // Body included (not just the status code) so an unrecognized 400 still names the actual
            // rejected field/reason in the activity log instead of just "HTTP 400". Payload/values are
            // deliberately NOT echoed here — the response body is SharePoint's own error text, but the
            // outgoing payload can contain the submitted field values themselves.
            return $"ApplyCustomFields HTTP {(int)response.StatusCode} — {body[..Math.Min(body.Length, 500)]}";
        }

        // ValidateUpdateListItem returns an entry for every field in the list definition,
        // including read-only system fields (Created, Modified, etc.) that always report
        // HasException=true.  Only surface errors for fields we actually submitted.
        var fieldErrors = new List<string>(lookupErrors);
        fieldErrors.AddRange(droppedFields.Select(n => Loc.T("Svc_Sp_ColumnMissingOnTarget", n)));
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var vals))
            {
                var returnedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var v in vals.EnumerateArray())
                {
                    if (!v.TryGetProperty("FieldName", out var fn)) continue;
                    var name = fn.GetString() ?? "";
                    returnedFields.Add(name);
                    if (!v.TryGetProperty("HasException", out var ex) || !ex.GetBoolean()) continue;
                    if (submittedFields.Contains(name))
                        fieldErrors.Add(name);
                }

                // A field we submitted can be silently absent from the response — e.g. the item's
                // content type doesn't include that field (common when the target library assigns a
                // different content type to certain file types, such as PDFs vs. Office documents).
                // ValidateUpdateListItem never reports this as HasException, so without this check
                // the copy is reported as fully successful while the value was never written.
                foreach (var name in submittedFields)
                {
                    if (!returnedFields.Contains(name))
                        fieldErrors.Add(Loc.T("Svc_Sp_NotApplicableToItem", name));
                }
            }
        }
        catch { /* ignore parse errors */ }

        return fieldErrors.Count > 0 ? Loc.T("Svc_Sp_CustomFieldErrors", string.Join(", ", fieldErrors)) : null;
    }

    // Managed-metadata values carry the SOURCE term store's term GUIDs, which don't exist in another
    // tenant's term store — writing them fails the field. In a cross-tenant copy they are skipped (the
    // copy reports this once per run; see CopyService) rather than failing every item.
    // With a term map (see TermStoreCopyService) the values are remapped instead: each source term GUID is
    // replaced by the target's, and a term the map does not know is dropped (cross-tenant) or kept (same
    // tenant, where the GUID still exists). Returns null when nothing is left to write.
    private object? RemapTaxonomyValue(object? value)
    {
        if (value is not TaxonomyFieldValue tax) return value;
        if (TermMap.IsEmpty)
        {
            if (!_router.IsMultiTenant) return value;
            Interlocked.Add(ref _taxonomySkipped, tax.Terms.Length);
            return null;
        }
        var kept = new List<(string Label, string TermGuid)>();
        foreach (var t in tax.Terms)
        {
            if (TermMap.TryGetValue(t.TermGuid.Trim('{', '}'), out var mapped))
            {
                kept.Add((t.Label, mapped));
                Interlocked.Increment(ref _taxonomyRemapped);
            }
            else if (!_router.IsMultiTenant) kept.Add(t);
            else Interlocked.Increment(ref _taxonomySkipped);
        }
        return kept.Count > 0 ? new TaxonomyFieldValue(kept.ToArray()) : null;
    }

    // Source term GUID -> target term GUID, filled by TermStoreCopyService when a term set is copied.
    public System.Collections.Concurrent.ConcurrentDictionary<string, string> TermMap { get; } = new(StringComparer.OrdinalIgnoreCase);
    private int _taxonomyRemapped, _taxonomySkipped;
    public (int Remapped, int Skipped) TaxonomyStats => (_taxonomyRemapped, _taxonomySkipped);

    private string FormatFieldValueForValidate(object value)
    {
        // ValidateUpdateListItem person format: JSON array of claims keys.
        // Cross-tenant: the logins were read from the source, so map them to target identities here.
        if (value is PersonFieldValue person)
            return JsonSerializer.Serialize(person.Logins.Select(l => new { Key = Identity.MapLogin(l) }));

        // ValidateUpdateListItem taxonomy format: "Label|guid" (";"-joined for multi).
        // SharePoint resolves WssId and the hidden note field server-side.
        if (value is TaxonomyFieldValue taxonomy)
            return string.Join(";", taxonomy.Terms.Select(t => $"{t.Label}|{t.TermGuid}"));

        // ValidateUpdateListItem hyperlink format: "url, description" (description falls back to the URL).
        if (value is HyperlinkFieldValue link)
            return $"{link.Url}, {(string.IsNullOrWhiteSpace(link.Description) ? link.Url : link.Description)}";

        // Lookups never reach here: callers resolve them to target item IDs first
        // (ResolveLookupFormValueAsync) — ValidateUpdateListItem doesn't accept display text.

        if (value is bool b)
            return b ? "1" : "0";

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
            _router.RegisterDriveForSite(driveId, siteUrl);

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

}
