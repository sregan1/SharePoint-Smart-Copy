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

    public static IdentityMapper FromSettings(Models.IdentityMappingSettings? s)
        => s == null
            ? None
            : new IdentityMapper(s.SourceDomain, s.TargetDomain, ParseOverrides(s.Overrides));

    // "@contoso.com" and "contoso.com" both mean contoso.com.
    private static string NormalizeDomain(string? domain)
        => (domain ?? string.Empty).Trim().TrimStart('@');
}
