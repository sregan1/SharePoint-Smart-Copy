namespace SharePointSmartCopy.Models;

public record FileMetadata(
    DateTimeOffset? CreatedDateTime,
    string? CreatedByEmail,
    DateTimeOffset? ModifiedDateTime,
    string? ModifiedByEmail);
