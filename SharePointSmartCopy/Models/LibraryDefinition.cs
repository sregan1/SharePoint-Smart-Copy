namespace SharePointSmartCopy.Models;

public class LibraryDefinition
{
    public string Title                { get; set; } = string.Empty;
    public string Description          { get; set; } = string.Empty;
    public string SourceDriveId        { get; set; } = string.Empty;
    public string SourceSiteUrl        { get; set; } = string.Empty;
    public string SourceListId         { get; set; } = string.Empty;
    public bool   EnableVersioning     { get; set; }
    public bool   EnableMinorVersions  { get; set; }
    public int    MajorVersionLimit    { get; set; }
    public List<ColumnDefinition> Columns { get; set; } = [];
}
