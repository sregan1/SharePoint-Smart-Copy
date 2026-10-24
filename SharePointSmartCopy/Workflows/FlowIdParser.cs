using System.Text.RegularExpressions;

namespace SharePointSmartCopy.Workflows;

// Lets the user paste what they can easily copy instead of a bare ID:
//   list   -> a GUID (with or without braces), or the list-settings page URL (...listedit.aspx?List=%7B...%7D)
//   form   -> the form ID, or the form's edit-page URL (...DesignPageV2.aspx?...&id=<formId>...)
public static partial class FlowIdParser
{
    // Returns the list GUID (lowercase), or null when the text contains none.
    public static string? ParseListId(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        var text = SafeUnescape(input.Trim());
        // Prefer the GUID that follows "List=" (a settings URL can also carry a web or site id)
        var m = ListParam().Match(text);
        if (!m.Success) m = AnyGuid().Match(text);
        return m.Success ? m.Groups.Count > 1 && m.Groups[1].Success ? m.Groups[1].Value.ToLowerInvariant() : m.Value.ToLowerInvariant() : null;
    }

    // Returns the form ID, or null when the text is neither an ID nor a URL carrying one.
    public static string? ParseFormId(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        var text = input.Trim();
        if (Uri.TryCreate(text, UriKind.Absolute, out var uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps))
        {
            foreach (var pair in (uri.Query.TrimStart('?') + "&" + uri.Fragment.TrimStart('#')).Split('&', StringSplitOptions.RemoveEmptyEntries))
            {
                var eq = pair.IndexOf('=');
                if (eq <= 0) continue;
                var key = pair[..eq];
                if (key.Equals("id", StringComparison.OrdinalIgnoreCase) || key.Equals("formid", StringComparison.OrdinalIgnoreCase))
                {
                    var value = SafeUnescape(pair[(eq + 1)..]);
                    if (FormIdShape().IsMatch(value)) return value;
                }
            }
            return null;
        }
        return FormIdShape().IsMatch(text) ? text : null;
    }

    private static string SafeUnescape(string s)
    {
        try { return Uri.UnescapeDataString(s); } catch (UriFormatException) { return s; }
    }

    [GeneratedRegex(@"List=\{?([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})", RegexOptions.IgnoreCase)] private static partial Regex ListParam();
    [GeneratedRegex(@"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")] private static partial Regex AnyGuid();
    // Forms IDs are long base64url strings (~80 chars).
    [GeneratedRegex(@"^[A-Za-z0-9_\-]{40,}$")] private static partial Regex FormIdShape();
}
