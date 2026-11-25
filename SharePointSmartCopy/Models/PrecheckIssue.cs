namespace SharePointSmartCopy.Models;

// One item the pre-copy check refused: a name or path SharePoint would reject at the destination.
// Collected during the source scan; shown to the user (with the option to cancel) before anything is
// transferred, and kept on the run's result rows (CopyResult.IsPrecheckFlagged) so the list can be
// reopened and exported after the run.
public sealed record PrecheckIssue(string SourcePath, string TargetPath, string Reason);
