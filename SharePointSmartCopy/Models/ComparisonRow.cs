namespace SharePointSmartCopy.Models;

// Existence-only comparison. Byte-level signals (size, content hash) were tried and dropped:
// SharePoint routinely re-serializes the ZIP container behind .docx/.xlsx/.pptx (indexing,
// thumbnail generation, co-authoring readiness) which changes both size and hash for files whose
// content is genuinely fine, making either signal too noisy to trust. Whether a file made it to
// the target at all is the thing that can be verified reliably and cheaply.
public enum ComparisonStatus { Match, OnlyInSource, OnlyInTarget }

public sealed class ComparisonRow
{
    public required string RelativePath { get; init; }
    public required ComparisonStatus Status { get; init; }
}
