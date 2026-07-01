namespace SharePointSmartCopy.Models;

// The subset of CopyJob fields needed to restart a recursive verification scan, decoupled from
// CopyJob so the same scan logic can be driven from either a live run (current CopyJobs) or a
// historical run (a SavedReport loaded from disk, possibly from a previous app session).
public sealed record VerificationRoot(
    string SourceDriveId,
    string SourceItemId,
    string SourceName,
    bool IsFolder,
    bool IsLibrary,
    string TargetDriveId,
    string TargetParentItemId,
    string TargetSubFolderPath)
{
    public static List<VerificationRoot> FromCopyJobs(IReadOnlyList<CopyJob> jobs) =>
        jobs.Select(j => new VerificationRoot(
                j.SourceDriveId, j.SourceItemId, j.SourceName, j.IsFolder, j.IsLibrary,
                j.TargetDriveId, j.TargetParentItemId, j.TargetSubFolderPath))
            .ToList();

    public static List<VerificationRoot> FromSavedReport(SavedReport report) =>
        report.Roots
            .Select(r => new VerificationRoot(
                r.SourceDriveId, r.SourceItemId, r.SourceName, r.IsFolder, r.IsLibrary,
                r.TargetDriveId, r.TargetParentItemId, r.TargetSubFolderPath))
            .ToList();
}
