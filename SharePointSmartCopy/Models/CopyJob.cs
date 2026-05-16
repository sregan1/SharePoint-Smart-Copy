namespace SharePointSmartCopy.Models;

public class CopyJob
{
    public string SourceDriveId { get; set; } = string.Empty;
    public string SourceItemId { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;
    public string SourceDisplayPath { get; set; } = string.Empty;

    public string TargetDriveId { get; set; } = string.Empty;
    public string TargetParentItemId { get; set; } = string.Empty;
    public string TargetSiteId { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string TargetDisplayPath { get; set; } = string.Empty;

    public bool IsFolder { get; set; }
    public string TargetSubFolderPath { get; set; } = string.Empty;

    // Server-relative URL of the target document library root (e.g. "/sites/target/Shared Documents")
    // Required for Migration API mode manifest generation.
    public string TargetLibraryServerRelativeUrl { get; set; } = string.Empty;
}
