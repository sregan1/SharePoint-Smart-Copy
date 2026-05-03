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
    public string TargetDisplayPath { get; set; } = string.Empty;

    // When true, this job represents an entire folder subtree to be recursively copied
    public bool IsFolder { get; set; }

    // Sub-folder path to create within TargetParentItemId (for files inside copied folders)
    public string TargetSubFolderPath { get; set; } = string.Empty;
}
