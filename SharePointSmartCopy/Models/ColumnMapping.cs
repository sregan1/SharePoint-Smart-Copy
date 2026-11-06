namespace SharePointSmartCopy.Models;

public enum MappingStatus { AutoMatched, ManuallyMapped, Unmatched, WillCreate, Skipped }

// Column-mapping state for ONE source list in an item-level copy. Each selected list can have a different
// column set and a different destination, so the mapping dialog edits one of these per list.
public class ListMappingContext
{
    public string SourceListId { get; init; } = string.Empty;
    public string Title        { get; init; } = string.Empty;

    // True when the destination list is created by the copy (or its columns couldn't be read): the dialog then
    // offers Create/Skip only, since there is no existing target column to map to.
    public bool IsNewDestination { get; init; }

    public List<ColumnDefinition> SourceColumns { get; init; } = [];
    public List<ColumnDefinition> TargetColumns { get; init; } = [];

    // Saved decisions for this list. Empty until the user opens and saves the dialog.
    public List<ColumnMapping> Mappings { get; } = [];
}

public class ColumnMapping
{
    public ColumnDefinition  SourceColumn { get; set; } = new();
    public ColumnDefinition? TargetColumn { get; set; }
    public bool              CreateNew    { get; set; }
    public MappingStatus     Status       { get; set; }

    // Canonical mapping resolution shared by every value writer (REST custom fields,
    // SPMI manifest, list item copy). Semantics:
    //   key absent          → no decision recorded; writers fall back to the source name
    //   value is null       → explicitly skipped by the user; do not copy
    //   value is a name     → write to that target column (create-new keeps the source name)
    public static Dictionary<string, string?> BuildTargetNameMap(IEnumerable<ColumnMapping> mappings)
    {
        var map = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in mappings)
        {
            map[m.SourceColumn.InternalName] =
                m.Status is MappingStatus.Skipped or MappingStatus.Unmatched
                    ? null
                    : m.TargetColumn?.InternalName ?? m.SourceColumn.InternalName;
        }
        return map;
    }

    // Whether a source column's values can be meaningfully written into a target
    // column of the given type. Used by auto-match so a Person column is never
    // silently matched to a Text column that happens to share its name.
    public static bool AreTypesCompatible(SupportedFieldType source, SupportedFieldType target)
    {
        if (source == target) return true;

        static bool IsTexty(SupportedFieldType t) =>
            t is SupportedFieldType.Text or SupportedFieldType.Note
              or SupportedFieldType.Choice or SupportedFieldType.MultiChoice;

        if (IsTexty(source) && IsTexty(target)) return true;
        // A currency amount is a number with a currency format.
        if (source is SupportedFieldType.Number or SupportedFieldType.Currency &&
            target is SupportedFieldType.Number or SupportedFieldType.Currency) return true;
        if (ColumnDefinition.IsUserType(source) && ColumnDefinition.IsUserType(target)) return true;
        if (ColumnDefinition.IsTaxonomyType(source) && ColumnDefinition.IsTaxonomyType(target)) return true;
        return false;
    }
}
