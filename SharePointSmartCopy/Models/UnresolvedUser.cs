namespace SharePointSmartCopy.Models;

// A person (or group) from the source that could not be found in the target tenant during a
// cross-tenant run. Source is their email (or the group's name) as the SOURCE knows them;
// AttemptedTarget is what the mapping turned it into and the target then failed to resolve;
// Where lists where it was needed ("author", "permission", "person column", ...); Count is how many
// times it came up, so the people affecting the most items sort first.
public sealed record UnresolvedUser(string Source, string AttemptedTarget, string Where, int Count);
