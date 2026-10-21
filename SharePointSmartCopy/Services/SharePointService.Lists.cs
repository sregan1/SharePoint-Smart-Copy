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

// Partial class SharePointService: Lists. The shared fields, constructor and core setup live in SharePointService.cs.
public partial class SharePointService
{
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

    // Returns the display title for a list looked up by GUID, or null if not found.
    // Used to re-bind Lookup columns at the target: the source column stores the SOURCE list's
    // GUID, so the equivalent target list is resolved by matching title.
    public async Task<string?> GetListTitleByIdAsync(string siteUrl, string listId)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')?$select=Title";
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
            return doc.RootElement.TryGetProperty("Title", out var tp) ? tp.GetString() : null;
        }
        catch { return null; }
    }

    // Reads all items from a list with the given custom field names plus Created/Modified.
    // Returns each item as a string-keyed dictionary. Handles pagination automatically.
    public async Task<List<Dictionary<string, object?>>> GetListItemsAsync(
        string siteUrl, string listId,
        IEnumerable<SpColumnDef> customColumns,
        IEnumerable<string>? extraFieldNames = null,
        CancellationToken ct = default)
    {
        var cols       = customColumns.ToList();
        var colsByName = cols.ToDictionary(c => c.InternalName, StringComparer.OrdinalIgnoreCase);

        // User fields are lookups: $expand them and read {name}/Name (claims login).
        // Lookup fields must ALSO be $expand-ed ({name}/LookupId + LookupValue, same as
        // BulkReadCustomFieldsAsync) — selecting a lookup by bare internal name is a 400,
        // which used to abort the read and silently "copy" zero items.
        var selectParts = new List<string> { "Id", "Title", "Created", "Modified" };
        var expandParts = new List<string>();
        foreach (var c in cols)
        {
            if (SpColumnDef.IsUserType(c.FieldType))
            {
                selectParts.Add($"{c.InternalName}/Name");
                expandParts.Add(c.InternalName);
            }
            else if (SpColumnDef.IsLookupType(c.FieldType))
            {
                selectParts.Add($"{c.InternalName}/LookupId");
                selectParts.Add($"{c.InternalName}/LookupValue");
                expandParts.Add(c.InternalName);
            }
            else
            {
                selectParts.Add(c.InternalName);
            }
        }
        if (extraFieldNames != null)
            selectParts.AddRange(extraFieldNames);

        var baseUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                      $"?$select={Uri.EscapeDataString(string.Join(",", selectParts))}" +
                      (expandParts.Count > 0 ? $"&$expand={Uri.EscapeDataString(string.Join(",", expandParts))}" : "") +
                      "&$top=5000";

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

            if (!resp.IsSuccessStatusCode)
            {
                // Never truncate silently: a failed page used to `break`, returning a partial
                // (or empty) list the caller then reported as a successful copy of 0 items.
                var err = await resp.Content.ReadAsStringAsync(ct);
                throw new HttpRequestException(
                    Loc.T("Svc_Sp_ReadingListItemsFailed", (int)resp.StatusCode, err[..Math.Min(err.Length, 300)]));
            }
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);

            if (doc.RootElement.TryGetProperty("value", out var values))
                foreach (var el in values.EnumerateArray())
                {
                    var item = new Dictionary<string, object?>();
                    foreach (var prop in el.EnumerateObject())
                        item[prop.Name] = colsByName.TryGetValue(prop.Name, out var col)
                            ? ParseFieldValue(prop.Value, col.FieldType)
                            : ExtractJsonValue(prop.Value);
                    result.Add(item);
                }

            next = GetNextLink(doc.RootElement);
        }
        return result;
    }

    // Returns (Id, Title, Modified) for each item in a list — lightweight query used for
    // tree display and for skip / copy-if-newer decisions during list item copies.
    public async Task<List<(string Id, string Title, DateTimeOffset? Modified)>> GetListItemTitlesAsync(
        string siteUrl, string listId, CancellationToken ct = default)
    {
        var url = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items" +
                  "?$select=Id,Title,Modified&$top=5000&$orderby=Id";

        var result = new List<(string, string, DateTimeOffset?)>();
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

            if (!resp.IsSuccessStatusCode)
            {
                // A partial titles list corrupts skip/overwrite matching (existing items would
                // look absent and duplicate) — fail loudly instead.
                var err = await resp.Content.ReadAsStringAsync(ct);
                throw new HttpRequestException(
                    Loc.T("Svc_Sp_ReadingTargetListItems", (int)resp.StatusCode, err[..Math.Min(err.Length, 300)]));
            }
            var body = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(body);

            if (doc.RootElement.TryGetProperty("value", out var values))
                foreach (var el in values.EnumerateArray())
                {
                    var id    = el.TryGetProperty("Id",    out var ip) ? ip.GetInt32().ToString() : string.Empty;
                    var title = el.TryGetProperty("Title", out var tp) && tp.ValueKind == JsonValueKind.String
                                    ? tp.GetString() ?? string.Empty
                                    : string.Empty;
                    DateTimeOffset? modified = el.TryGetProperty("Modified", out var mp) &&
                                               mp.ValueKind == JsonValueKind.String &&
                                               DateTimeOffset.TryParse(mp.GetString(), System.Globalization.CultureInfo.InvariantCulture,
                                                   System.Globalization.DateTimeStyles.None, out var m)
                                                   ? m : null;
                    if (!string.IsNullOrEmpty(id))
                        result.Add((id, title, modified));
                }

            next = GetNextLink(doc.RootElement);
        }
        return result;
    }

    // Creates a list item and optionally back-fills Created/Modified timestamps.
    // Returns (new item ID or null, field-write error or null).
    public async Task<(string? Id, string? FieldError)> CreateListItemAsync(
        string siteUrl, string listId,
        Dictionary<string, object?> fields,
        string? createdDate, string? modifiedDate,
        CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();

        // Person/taxonomy values cannot be set through the plain item POST body — they
        // go through the ValidateUpdateListItem pass below alongside the date back-fill.
        var (simpleFields, complexFields) = SplitComplexFields(fields);

        var createUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items";
        var body      = JsonSerializer.Serialize(simpleFields);

        string? newItemId = null;
        using (var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, createUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct))
        {
            if (!resp.IsSuccessStatusCode)
            {
                var errBody = await resp.Content.ReadAsStringAsync(ct);
                throw new HttpRequestException(Loc.T("Svc_Sp_CreateItemFailed", (int)resp.StatusCode, resp.ReasonPhrase, errBody));
            }
            var respBody = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(respBody);
            newItemId = doc.RootElement.TryGetProperty("Id", out var ip) ? ip.GetRawText() : null;
        }

        if (newItemId == null) return (null, null);
        var fieldError = await ValidateUpdateItemFieldsAsync(siteUrl, listId, newItemId, complexFields, createdDate, modifiedDate, ct);
        return (newItemId, fieldError);
    }

    // Splits person/taxonomy values (which require ValidateUpdateListItem) from fields
    // that can be written through a plain REST POST/MERGE body.
    private static (Dictionary<string, object?> Simple, List<(string Name, object Value)> Complex)
        SplitComplexFields(Dictionary<string, object?> fields)
    {
        var simple  = new Dictionary<string, object?>();
        var complex = new List<(string, object)>();
        foreach (var (name, value) in fields)
        {
            if (value == null) continue;
            // Lookups are complex too: serializing a LookupFieldValue record into the plain POST
            // body produced JSON SharePoint rejects, failing the item create.
            if (value is PersonFieldValue or TaxonomyFieldValue or LookupFieldValue or HyperlinkFieldValue)
                complex.Add((name, value));
            else
                simple[name] = value;
        }
        return (simple, complex);
    }

    // Applies complex field values and/or Created/Modified back-fill in one
    // ValidateUpdateListItem call (bNewDocumentUpdate avoids creating a new version).
    // Returns an error string if the write failed, or null on success.
    private async Task<string?> ValidateUpdateItemFieldsAsync(
        string siteUrl, string listId, string itemId,
        List<(string Name, object Value)> complexFields,
        string? createdDate, string? modifiedDate,
        CancellationToken ct = default)
    {
        var formValues      = new List<object>();
        var submittedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var lookupErrors    = new List<string>();

        // Target column defs (cached) supply each lookup column's bound list for ID resolution.
        Dictionary<string, SpColumnDef>? targetColsByName = null;
        if (complexFields.Any(f => f.Value is LookupFieldValue))
        {
            List<SpColumnDef> targetCols;
            try { targetCols = await GetLibraryColumnsAsync(siteUrl, listId); }
            catch { targetCols = []; }
            targetColsByName = targetCols.ToDictionary(c => c.InternalName, StringComparer.OrdinalIgnoreCase);
        }

        foreach (var f in complexFields)
        {
            // Lookups must be submitted as TARGET item IDs, not display text.
            if (f.Value is LookupFieldValue lookup)
            {
                SpColumnDef? tgtCol = null;
                if (targetColsByName == null || !targetColsByName.TryGetValue(f.Name, out tgtCol) ||
                    string.IsNullOrEmpty(tgtCol.LookupListId))
                {
                    lookupErrors.Add(DescribeUnusableLookupTarget(f.Name, tgtCol));
                    continue;
                }
                var formatted = await ResolveLookupFormValueAsync(siteUrl, tgtCol, lookup, f.Name, lookupErrors, ct);
                if (formatted == null) continue;
                formValues.Add(new { FieldName = f.Name, FieldValue = formatted });
                submittedFields.Add(f.Name);
                continue;
            }

            if (SkipsTaxonomyValue(f.Value)) continue;
            // Cross-tenant: leave out (and record) people who do not exist in the target.
            var checkedValue = await DropUnresolvedPeopleAsync(siteUrl, f.Value, ct);
            if (checkedValue == null) continue;
            formValues.Add(new { FieldName = f.Name, FieldValue = FormatFieldValueForValidate(checkedValue) });
            submittedFields.Add(f.Name);
        }
        var lookupWarning = lookupErrors.Count > 0 ? string.Join(", ", lookupErrors) : null;
        // Created/Modified arrive as the raw REST ISO strings ("…T…Z"). That is the exact format this file
        // documents ValidateUpdateListItem failing to parse (or reading as site-local), so convert them
        // the same way the file/folder paths do (see BuildDateAndPersonFormValuesAsync).
        if (createdDate  != null) { formValues.Add(new { FieldName = "Created",  FieldValue = await ToSiteLocalFormDateAsync(siteUrl, createdDate) });  submittedFields.Add("Created"); }
        if (modifiedDate != null) { formValues.Add(new { FieldName = "Modified", FieldValue = await ToSiteLocalFormDateAsync(siteUrl, modifiedDate) }); submittedFields.Add("Modified"); }
        if (formValues.Count == 0) return lookupWarning;

        var payload = JsonSerializer.Serialize(new { formValues, bNewDocumentUpdate = true });
        var url     = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({itemId})/ValidateUpdateListItem()";

        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, url);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct);

        var body = await resp.Content.ReadAsStringAsync(ct);
        if (!resp.IsSuccessStatusCode)
            return Loc.T("Svc_Sp_FieldWriteFailed", (int)resp.StatusCode, body[..Math.Min(200, body.Length)]);

        // ValidateUpdateListItem reports an entry for every field in the list, including read-only
        // system fields that always carry HasException=true — only surface fields we submitted.
        string? fieldErrors = null;
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("value", out var vals))
            {
                var errors = vals.EnumerateArray()
                    .Where(v => v.TryGetProperty("HasException", out var ex) && ex.GetBoolean())
                    .Select(v => v.TryGetProperty("FieldName", out var fn) ? fn.GetString() ?? "?" : "?")
                    .Where(n => submittedFields.Contains(n))
                    .ToList();
                if (errors.Count > 0)
                    fieldErrors = Loc.T("Svc_Sp_FieldErrorsOn", string.Join(", ", errors));
            }
        }
        catch { }

        return (fieldErrors, lookupWarning) switch
        {
            (null, null)   => null,
            (null, var lw) => lw,
            (var fe, null) => fe,
            var (fe, lw)   => $"{fe} {lw}",
        };
    }

    // Updates an existing list item's fields via MERGE/PATCH and optionally back-fills Created/Modified.
    // Returns a field-write error string if person/taxonomy fields could not be applied, or null on success.
    public async Task<string?> UpdateListItemAsync(
        string siteUrl, string listId, string itemId,
        Dictionary<string, object?> fields,
        string? createdDate, string? modifiedDate,
        CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();

        var (simpleFields, complexFields) = SplitComplexFields(fields);

        var updateUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists('{listId}')/items({itemId})";
        var body      = JsonSerializer.Serialize(simpleFields);

        using var resp = await SendSharePointRequestAsync(token =>
        {
            var r = new HttpRequestMessage(HttpMethod.Post, updateUrl);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            r.Headers.Add("X-HTTP-Method", "MERGE");
            r.Headers.Add("If-Match", "*");
            r.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
            return r;
        }, siteUrl, cancellationToken: ct);

        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException(Loc.T("Svc_Sp_UpdateItemFailed", (int)resp.StatusCode, resp.ReasonPhrase, errBody));
        }

        return await ValidateUpdateItemFieldsAsync(siteUrl, listId, itemId, complexFields, createdDate, modifiedDate, ct);
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
            using var resp = await SendSharePointRequestAsync(token =>
            {
                var r = new HttpRequestMessage(HttpMethod.Get, next);
                r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                r.Headers.Accept.ParseAdd("application/json;odata=nometadata");
                return r;
            }, siteUrl, cancellationToken: ct);
            if (!resp.IsSuccessStatusCode)
            {
                // Never truncate silently: a failed page used to `break` and return the PARTIAL
                // dictionary built so far. A missing key reads identically to "no unique
                // permissions" at every call site (`if (!permissionFlags.TryGetValue(...) || !hu)
                // return;`), so every item past the failed page silently copied with inherited
                // permissions and no row was ever logged. All four call sites already wrap this in
                // a try/catch ("that library's permissions are skipped"), so throwing here just
                // routes a genuine failure to the existing, visible skip path instead of an
                // indistinguishable partial success.
                var err = await resp.Content.ReadAsStringAsync(ct);
                throw new HttpRequestException(
                    Loc.T("Svc_Sp_ReadingPermissionFlagsFailed", (int)resp.StatusCode, err[..Math.Min(err.Length, 300)]));
            }
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
        return result;
    }

    // Composite key for permission-flag dictionaries. GUID braces are stripped so keys
    // match whether the list ID came from SP REST or Graph sharepointIds.
    public static string PermissionFlagKey(string listId, string itemId) =>
        $"{listId.Trim('{', '}')}:{itemId}";

}
