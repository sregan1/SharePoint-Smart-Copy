using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Services;

// Translates source-tenant identities (emails / claims logins / group names) to their target-tenant
// equivalents for a cross-tenant copy. Applied only at the points where an identity is WRITTEN to the
// target (migration package users, Author/Editor stamping, person fields, permission principals) —
// never on the source read side.
//
// Rules, in order: an explicit override wins; otherwise a domain swap (user@old.com -> user@new.com);
// otherwise the identity is left unchanged. Mapping is idempotent — a value that is already a target
// value (an override result, or already on the target domain) is returned as-is — so it is safe if two
// layers both map the same identity.
//
// The default instance (None) is inactive and returns every input untouched, which is what
// single-tenant copies use.
public sealed class IdentityMapper
{
    private const string ClaimsPrefix = "i:0#.f|membership|";

    private readonly Dictionary<string, string> _overrides = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _overrideTargets = new(StringComparer.OrdinalIgnoreCase);
    private readonly string _sourceDomain;
    private readonly string _targetDomain;

    public static IdentityMapper None { get; } = new(null, null, null);

    public IdentityMapper(string? sourceDomain, string? targetDomain, IEnumerable<KeyValuePair<string, string>>? overrides)
    {
        _sourceDomain = NormalizeDomain(sourceDomain);
        _targetDomain = NormalizeDomain(targetDomain);
        if (overrides == null) return;
        foreach (var (source, target) in overrides)
        {
            if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(target)) continue;
            _overrides[source.Trim()] = target.Trim();
            _overrideTargets.Add(target.Trim());
        }
    }

    private bool HasDomainSwap => _sourceDomain.Length > 0 && _targetDomain.Length > 0;

    public bool IsActive => HasDomainSwap || _overrides.Count > 0;

    public string? MapEmail(string? email)
    {
        if (string.IsNullOrWhiteSpace(email) || !IsActive) return email;
        var trimmed = email.Trim();

        if (_overrides.TryGetValue(trimmed, out var mapped)) return mapped;
        if (_overrideTargets.Contains(trimmed)) return trimmed; // already a mapped value

        // Guest accounts (user_ext.com#EXT#@tenant.onmicrosoft.com) have no equivalent at the same
        // local part on another domain — swapping would invent a non-existent identity. Leave them
        // for an explicit override.
        if (trimmed.Contains("#EXT#", StringComparison.OrdinalIgnoreCase)) return email;

        var at = trimmed.LastIndexOf('@');
        if (HasDomainSwap && at > 0 &&
            trimmed[(at + 1)..].Equals(_sourceDomain, StringComparison.OrdinalIgnoreCase))
            return trimmed[..at] + "@" + _targetDomain;

        return email;
    }

    // Accepts a claims login ("i:0#.f|membership|user@x.com"), a bare email, or any other login /
    // group name (only an explicit override can remap those).
    public string MapLogin(string login)
    {
        if (string.IsNullOrEmpty(login) || !IsActive) return login;

        if (_overrides.TryGetValue(login.Trim(), out var direct)) return direct;
        if (login.StartsWith(ClaimsPrefix, StringComparison.OrdinalIgnoreCase))
            return ClaimsPrefix + MapEmail(login[ClaimsPrefix.Length..]);
        if (login.Contains('@') && !login.Contains('|'))
            return MapEmail(login) ?? login;
        return login;
    }

    // A permission principal: try its login first, then its display title (how an Entra or SharePoint
    // group is usually keyed in an override list).
    public string MapPrincipal(string login, string title)
    {
        if (!IsActive) return string.IsNullOrEmpty(login) ? title : login;
        var key    = string.IsNullOrEmpty(login) ? title : login;
        var mapped = MapLogin(key);
        if (!string.Equals(mapped, key, StringComparison.Ordinal)) return mapped;
        return !string.IsNullOrEmpty(title) && _overrides.TryGetValue(title.Trim(), out var byTitle) ? byTitle : key;
    }

    // Overrides text: one "source,target" pair per line (comma, semicolon or tab separated);
    // blank lines and lines starting with '#' are ignored.
    public static List<KeyValuePair<string, string>> ParseOverrides(string? text)
    {
        var result = new List<KeyValuePair<string, string>>();
        if (string.IsNullOrWhiteSpace(text)) return result;
        foreach (var raw in text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith('#')) continue;
            var parts = line.Split([',', ';', '\t'], 2, StringSplitOptions.TrimEntries);
            if (parts.Length == 2 && parts[0].Length > 0 && parts[1].Length > 0)
                result.Add(new(parts[0], parts[1]));
        }
        return result;
    }

    // 1-based line numbers (as shown in the editor) of lines that are neither blank, a comment, nor a
    // readable "source,target" pair. ParseOverrides skips these silently — the dialog uses this to tell
    // the user instead, so a typo can't quietly leave someone unmapped.
    public static List<int> FindInvalidLines(string? text)
    {
        var bad = new List<int>();
        if (string.IsNullOrEmpty(text)) return bad;
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (line.Length == 0 || line.StartsWith('#')) continue;
            var parts = line.Split([',', ';', '\t'], 2, StringSplitOptions.TrimEntries);
            if (parts.Length != 2 || parts[0].Length == 0 || parts[1].Length == 0) bad.Add(i + 1);
        }
        return bad;
    }

    public static IdentityMapper FromSettings(Models.IdentityMappingSettings? s)
        => s == null
            ? None
            : new IdentityMapper(s.SourceDomain, s.TargetDomain,
                s.Pairs.Select(p => new KeyValuePair<string, string>(p.Source, p.Target))
                    // A settings object built in code (not loaded) may still carry the legacy text.
                    .Concat(ParseOverrides(s.Overrides)));

    // How a source principal is shown to the user and keyed in the mapping table: a user is their email
    // (stripped of the claims prefix), a group is its name.
    public static string DisplayIdentity(string? login, string? title)
    {
        if (!string.IsNullOrEmpty(login))
        {
            if (login.StartsWith(ClaimsPrefix, StringComparison.OrdinalIgnoreCase)) return login[ClaimsPrefix.Length..];
            if (login.Contains('@') && !login.Contains('|')) return login;
        }
        return !string.IsNullOrEmpty(title) ? title : login ?? string.Empty;
    }

    // ── Pair editing helpers (used by the mapping table) ─────────────────────

    // Parses a CSV/text file of "source,target" pairs: quoted fields ("Doe, Jane"), comma / semicolon / tab
    // separators, an optional header row, and # comment lines. A line with no second field is returned with a
    // BLANK target, so the table shows it (with an error) instead of silently dropping it.
    public static List<(string Source, string Target)> ParseCsvPairs(string? text)
    {
        var result = new List<(string, string)>();
        if (string.IsNullOrWhiteSpace(text)) return result;
        char sep = text.Contains('\t') ? '\t' : (!text.Contains(',') && text.Contains(';')) ? ';' : ',';
        bool first = true;
        foreach (var raw in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith('#')) continue;
            var cells = SplitCsvLine(line, sep);
            var a = cells.Count > 0 ? cells[0].Trim() : string.Empty;
            var b = cells.Count > 1 ? cells[1].Trim() : string.Empty;
            bool isHeader = first && a.Equals("source", StringComparison.OrdinalIgnoreCase)
                                  && b.StartsWith("target", StringComparison.OrdinalIgnoreCase);
            first = false;
            if (isHeader || (a.Length == 0 && b.Length == 0)) continue;
            result.Add((a, b));
        }
        return result;
    }

    private static List<string> SplitCsvLine(string line, char sep)
    {
        var cells = new List<string>();
        var sb = new System.Text.StringBuilder();
        bool quoted = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (quoted)
            {
                if (c == '"' && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                else if (c == '"') quoted = false;
                else sb.Append(c);
            }
            else if (c == '"') quoted = true;
            else if (c == sep) { cells.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(c);
        }
        cells.Add(sb.ToString());
        return cells;
    }

    // Quotes a value for CSV output only when it needs it.
    public static string CsvQuote(string value) =>
        value.IndexOfAny([',', ';', '"', '\t', '\n', '\r']) >= 0
            ? "\"" + value.Replace("\"", "\"\"") + "\""
            : value;

    // Row-by-row problems in an edited table, by row index. A row with both cells blank is ignored (not an
    // error). Messages are localized; an empty result means every row is usable.
    public static Dictionary<int, string> ValidatePairs(IReadOnlyList<(string Source, string Target)> rows)
    {
        var problems = new Dictionary<int, string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < rows.Count; i++)
        {
            var source = rows[i].Source.Trim();
            var target = rows[i].Target.Trim();
            if (source.Length == 0 && target.Length == 0) continue;
            if (source.Length == 0)      problems[i] = Loc.T("Err_IdentityBlankSource");
            else if (target.Length == 0) problems[i] = Loc.T("Err_IdentityBlankTarget");
            else if (source.Equals(target, StringComparison.OrdinalIgnoreCase)) problems[i] = Loc.T("Err_IdentitySame");
            else if (!seen.Add(source))  problems[i] = Loc.T("Err_IdentityDuplicate");
        }
        return problems;
    }

    // "@contoso.com" and "contoso.com" both mean contoso.com.
    private static string NormalizeDomain(string? domain)
        => (domain ?? string.Empty).Trim().TrimStart('@');
}
