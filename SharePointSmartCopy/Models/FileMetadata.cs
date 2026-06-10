namespace SharePointSmartCopy.Models;

public class FileMetadata
{
    public DateTimeOffset? CreatedDateTime  { get; init; }
    public string?         CreatedByEmail   { get; init; }
    public DateTimeOffset? ModifiedDateTime { get; init; }
    public string?         ModifiedByEmail  { get; init; }
    public Dictionary<string, object?> CustomFields { get; init; } = [];
}
