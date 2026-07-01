namespace SharePointSmartCopy.Models;

public enum ComparisonStatus { Match, SizeMismatch, ModifiedMismatch, OnlyInSource, OnlyInTarget }

public sealed class ComparisonRow
{
    public required string RelativePath { get; init; }
    public long? SourceSize { get; init; }
    public long? TargetSize { get; init; }
    public DateTimeOffset? SourceModified { get; init; }
    public DateTimeOffset? TargetModified { get; init; }
    public required ComparisonStatus Status { get; init; }
}
