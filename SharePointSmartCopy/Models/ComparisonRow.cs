namespace SharePointSmartCopy.Models;

// Split comparison strategy by file type. SharePoint routinely re-serializes Office/OLE
// compound-document containers — both modern OOXML (.docx/.xlsx/.pptx, ZIP-based) and legacy binary
// formats (.doc/.xls/.ppt, .msg) — for indexing, thumbnail generation, and co-authoring readiness,
// which changes both size and hash for files whose content is genuinely fine. So size/hash are too
// noisy to trust for those formats. Two signals are used instead, chosen per file type:
//   - Everything else: QuickXorHash content comparison (ContentMismatch on a genuine difference).
//     Reliable because these formats aren't silently rewritten by SharePoint's backend.
//   - Office/OLE compound formats: modified-date comparison (DateMismatch on a genuine difference).
//     The re-serialization happens below SharePoint's normal save pipeline (it doesn't bump version
//     count or change the editor), so the item's official Modified date stays stable even as the
//     underlying bytes drift — and the app is already responsible for preserving that date onto
//     the target, so this checks something already guaranteed rather than a new heuristic.
// Either signal missing on one side (e.g. a Graph anomaly) falls back to existence-only rather than
// manufacturing a false mismatch from absent data.
public enum ComparisonStatus { Match, ContentMismatch, DateMismatch, OnlyInSource, OnlyInTarget }

public sealed class ComparisonRow
{
    public required string RelativePath { get; init; }
    public required ComparisonStatus Status { get; init; }

    // Raw signals from each side, whenever that side's file exists — populated regardless of which
    // signal actually decided Status, so the Comparison sheet can show source/target values for
    // every row (blank on whichever side has no file, e.g. Only in Source/Target).
    public string? SourceHash { get; init; }
    public string? TargetHash { get; init; }
    public DateTimeOffset? SourceModified { get; init; }
    public DateTimeOffset? TargetModified { get; init; }
}
