namespace SharePointSmartCopy.Models;

public enum SupportedFieldType { Text, Note, Number, Boolean, DateTime, Choice, MultiChoice, User, UserMulti, Taxonomy, TaxonomyMulti, Lookup, LookupMulti }

// Typed field values let the write side format per destination: ValidateUpdateListItem
// (REST mode) and the SPMI manifest (Migration API mode) need different encodings.

// Claims logins (i:0#.f|membership|user@x.com). Within-tenant these resolve directly.
public sealed record PersonFieldValue(string[] Logins);

// Term references. Within-tenant the term store is shared, so TermGuids are valid as-is.
public sealed record TaxonomyFieldValue((string Label, string TermGuid)[] Terms);

// Lookup values carry the source item ID (ignored at write time) and the display value,
// which is used to resolve the matching item in the target lookup list.
public sealed record LookupFieldValue((int Id, string DisplayValue)[] Entries);

public class ColumnDefinition
{
    public string InternalName { get; set; } = string.Empty;
    public string DisplayName  { get; set; } = string.Empty;
    public SupportedFieldType FieldType { get; set; }
    public string[]? Choices  { get; set; }

    // Raw field schema captured from the source — used to recreate Taxonomy columns
    // at the target (carries the SspId/TermSetId binding, valid within-tenant).
    public string? SchemaXml { get; set; }

    // For Lookup/LookupMulti columns: the GUID of the list being looked up,
    // and the internal name of the field displayed as the lookup value.
    public string? LookupListId    { get; set; }
    public string? LookupShowField { get; set; }

    public static bool IsUserType(SupportedFieldType t)     => t is SupportedFieldType.User or SupportedFieldType.UserMulti;
    public static bool IsTaxonomyType(SupportedFieldType t) => t is SupportedFieldType.Taxonomy or SupportedFieldType.TaxonomyMulti;
    public static bool IsLookupType(SupportedFieldType t)   => t is SupportedFieldType.Lookup or SupportedFieldType.LookupMulti;
}
