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
// Either signal missing on one side (e.g. a Graph anomaly — quickXorHash is observed missing for a
// nontrivial fraction of items in children listings) first falls back to a size comparison; when
// size is missing too, the row is reported as Unverified — never as a fabricated Match. Office/OLE
// rows additionally short-circuit to Match when both quickXorHashes are present and EQUAL (equal
// hashes are always trustworthy; they only stop being meaningful when they differ).
//
// The opt-in "Deep verify Office files" pass (see VerificationReportService's deep-verify pass and
// OpcDeepComparer) can additionally resolve a ContentMismatch/DateMismatch to Match when the OOXML
// package's content parts are byte-identical and only the parts SharePoint rewrites (docProps,
// customXml, etc.) differ — reported as an ordinary Match with an explanatory Note, not a separate
// status, so a user doesn't need to know deep verify exists to read the summary correctly.
public enum ComparisonStatus { Match, ContentMismatch, DateMismatch, OnlyInSource, OnlyInTarget, Unverified }

public sealed class ComparisonRow
{
    public required string RelativePath { get; init; }
    // Settable: the deep-verify pass revisits a row's cheap-tier status after downloading and
    // comparing the actual file content (see VerificationReportService).
    public required ComparisonStatus Status { get; set; }

    // Raw signals from each side, whenever that side's file exists — populated regardless of which
    // signal actually decided Status, so the Comparison sheet can show source/target values for
    // every row (blank on whichever side has no file, e.g. Only in Source/Target).
    public string? SourceHash { get; init; }
    public string? TargetHash { get; init; }
    public DateTimeOffset? SourceModified { get; init; }
    public DateTimeOffset? TargetModified { get; init; }
    // Populated for every row (not just ones the size fallback actually used) so the report can
    // show what it compared even for rows where hash is unavailable — a blank Source/Target
    // Value column previously gave no indication a size comparison had even run, let alone what
    // it found, for the "hash missing on one side" fallback path.
    public long? SourceSize { get; init; }
    public long? TargetSize { get; init; }

    // Set by the deep-verify pass to explain what it found or why it couldn't run for this row
    // (e.g. which content part differed, or "skipped — cap reached"). Null for rows deep verify
    // never touched.
    public string? Note { get; set; }
}
