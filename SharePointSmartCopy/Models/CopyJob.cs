namespace SharePointSmartCopy.Models;

public class CopyJob
{
    public string SourceDriveId { get; set; } = string.Empty;
    public string SourceItemId { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;
    public string SourceDisplayPath { get; set; } = string.Empty;

    // Captured during the source enumeration walk (the Children listing returns it anyway) so the
    // If Newer decision compares dates already in hand instead of re-fetching them from Graph —
    // per-file date lookups at 100k+ scale are what exhausted the tenant throttle budget and, when
    // they failed, misclassified up-to-date files as needing a copy. Null for jobs built outside
    // the walk (e.g. single-file picks); those fall back to the bulk Graph fetch.
    public DateTimeOffset? SourceModified { get; set; }

    public string TargetDriveId { get; set; } = string.Empty;
    public string TargetParentItemId { get; set; } = string.Empty;
    public string TargetSiteId { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string TargetDisplayPath { get; set; } = string.Empty;

    public string SourceSiteUrl { get; set; } = string.Empty;
    public bool IsFolder  { get; set; }
    public bool IsLibrary { get; set; }
    public bool IsPage    { get; set; }
    public List<ColumnMapping> ColumnMappings { get; set; } = [];
    public string TargetSubFolderPath { get; set; } = string.Empty;

    // Server-relative URL of the target document library root (e.g. "/sites/target/Shared Documents")
    // Required for Migration API mode manifest generation.
    public string TargetLibraryServerRelativeUrl { get; set; } = string.Empty;
}
