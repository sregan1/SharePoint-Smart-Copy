using System.Collections.Concurrent;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Collects the people a cross-tenant run could not find in the target.
//
// Before this, such a person was silently skipped (their grant dropped, their authorship left as the
// copying account) and the only trace was a count in one report row. Recording them here lets the report
// say who they are, and lets the user map them to the right target account and run again.
//
// Written from many copy tasks at once, so every method is thread-safe. De-duplicated by source identity
// (case-insensitive); the same person coming up for 5,000 files is one entry with a count of 5,000.
public sealed class UnresolvedIdentityLog
{
    private sealed class Entry
    {
        public required string Source { get; init; }
        public string AttemptedTarget = string.Empty;
        public readonly SortedSet<string> Where = new(StringComparer.OrdinalIgnoreCase);
        public int Count;
    }

    private readonly ConcurrentDictionary<string, Entry> _entries = new(StringComparer.OrdinalIgnoreCase);

    // source: the identity as the source tenant knows it. attemptedTarget: what the mapping produced and
    // the target could not resolve. where: a short label for the report ("author", "permission", ...).
    public void Record(string? source, string? attemptedTarget, string where, int count = 1)
    {
        if (string.IsNullOrWhiteSpace(source)) return;
        var entry = _entries.GetOrAdd(source.Trim(), key => new Entry { Source = key });
        lock (entry)
        {
            entry.Count += Math.Max(1, count);
            entry.Where.Add(where);
            if (!string.IsNullOrWhiteSpace(attemptedTarget)) entry.AttemptedTarget = attemptedTarget.Trim();
        }
    }

    public int DistinctCount => _entries.Count;

    // "Have I already checked this account in this run?" — true the first time only. Lives here (not on the
    // caller) so Clear() resets it with the rest of the run state: after the user fixes a mapping and runs
    // again, every account must be checked afresh.
    private readonly ConcurrentDictionary<string, byte> _checked = new(StringComparer.OrdinalIgnoreCase);
    public bool TryMarkChecked(string key) => _checked.TryAdd(key, 0);

    // Most-affecting people first; ties alphabetical.
    public IReadOnlyList<UnresolvedUser> Snapshot()
    {
        var list = new List<UnresolvedUser>(_entries.Count);
        foreach (var e in _entries.Values)
            lock (e) list.Add(new UnresolvedUser(e.Source, e.AttemptedTarget, string.Join(", ", e.Where), e.Count));
        return list.OrderByDescending(u => u.Count).ThenBy(u => u.Source, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public void Clear()
    {
        _entries.Clear();
        _checked.Clear();
    }
}
