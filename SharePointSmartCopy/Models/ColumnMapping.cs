namespace SharePointSmartCopy.Models;

public enum MappingStatus { AutoMatched, ManuallyMapped, Unmatched, WillCreate, Skipped }

public class ColumnMapping
{
    public ColumnDefinition  SourceColumn { get; set; } = new();
    public ColumnDefinition? TargetColumn { get; set; }
    public bool              CreateNew    { get; set; }
    public MappingStatus     Status       { get; set; }
}
