using System.Collections.Concurrent;

namespace SharePointSmartCopy.Teams;

// What the copy has learned about people, shared by all the teams of one batch (and safe for teams copied at the same time): a
// person who is in ten teams is looked up in the target once, not ten times.
public sealed class TeamsPeopleCache
{
    public ConcurrentDictionary<string, string> EmailBySourceId { get; } = new(StringComparer.OrdinalIgnoreCase);
    public ConcurrentDictionary<string, string?> TargetIdByEmail { get; } = new(StringComparer.OrdinalIgnoreCase);
}
