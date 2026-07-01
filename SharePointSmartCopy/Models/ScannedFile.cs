namespace SharePointSmartCopy.Models;

// One file discovered by an independent post-copy verification re-scan.
// RelativePath is the join key between source and target scans (case-insensitive).
public sealed record ScannedFile(
    string DriveId,
    string ItemId,
    string Name,
    string RelativePath,
    long? Size,
    DateTimeOffset? LastModified);
