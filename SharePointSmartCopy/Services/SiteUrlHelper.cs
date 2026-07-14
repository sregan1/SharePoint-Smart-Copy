using System.IO;

namespace SharePointSmartCopy.Services;

public static class SiteUrlHelper
{
    // Extracts the site path segment used to prefix default report filenames, e.g.
    // "https://contoso.sharepoint.com/sites/Marketing" -> "Marketing".
    public static string ExtractSiteName(string? siteUrl)
    {
        if (string.IsNullOrWhiteSpace(siteUrl)) return "";

        var idx = siteUrl.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);
        if (idx < 0) idx = siteUrl.IndexOf("/teams/", StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return "";

        var name = siteUrl[(idx + "/sites/".Length)..].Trim('/');
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }

    // "Source-Target-" prefix for default report filenames. Omitted entirely if either side's
    // site name can't be determined, rather than producing a lopsided "Source--" or "-Target-".
    public static string ReportFilenamePrefix(string? sourceUrl, string? targetUrl, bool enabled = true)
    {
        if (!enabled) return "";
        var source = ExtractSiteName(sourceUrl);
        var target = ExtractSiteName(targetUrl);
        return source.Length > 0 && target.Length > 0 ? $"{source}-{target}-" : "";
    }
}
