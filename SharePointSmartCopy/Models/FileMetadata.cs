namespace SharePointSmartCopy.Models;

public class FileMetadata
{
    public DateTimeOffset? CreatedDateTime  { get; init; }
    public string?         CreatedByEmail   { get; init; }
    public DateTimeOffset? ModifiedDateTime { get; init; }
    public string?         ModifiedByEmail  { get; init; }
    public long?           Size             { get; init; }
    public string?         ProgId           { get; init; }

    // SharePoint's hidden "Customize folder > color" fields — see SharePointService's
    // FolderColorTagFieldName/FolderColorHexFieldName constants for the internal names used
    // to read/write these (unconfirmed against a live tenant; best-guess names).
    public string?         ColorTag         { get; init; }
    public string?         ColorHex         { get; init; }

    // A FOLDER's custom content type name (a Document Set or other folder-based type); null for the plain Folder type.
    public string?         ContentTypeName  { get; init; }

    public Dictionary<string, object?> CustomFields { get; init; } = [];
}
