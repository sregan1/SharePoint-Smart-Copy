namespace SharePointSmartCopy.Models;

// One file discovered by an independent post-copy verification re-scan.
// RelativePath is the join key between source and target scans (case-insensitive).
// QuickXorHash powers content comparison for non-Office files; LastModified powers the
// date-based comparison used for Office Open XML files instead (see VerificationReportService).
public sealed record ScannedFile(
    string DriveId,
    string ItemId,
    string Name,
    string RelativePath,
    DateTimeOffset? LastModified,
    string? QuickXorHash);
