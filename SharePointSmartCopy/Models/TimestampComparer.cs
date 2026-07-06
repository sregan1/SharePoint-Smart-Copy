namespace SharePointSmartCopy.Models;

// SharePoint stores Modified with whole-second precision, and the app's own date stamping
// (manifest TimeLastModified, ValidateUpdateListItem) truncates to seconds — while Graph
// returns source dates with fractional seconds. Comparing raw values makes an already-copied
// file look "newer" by up to 999 ms, so every Copy-If-Newer re-run deleted and re-copied
// identical files. All If-Newer decisions compare at second granularity through this helper.
public static class TimestampComparer
{
    public static bool IsUpToDate(DateTimeOffset source, DateTimeOffset target) =>
        Truncate(source) <= Truncate(target);

    private static DateTimeOffset Truncate(DateTimeOffset d)
    {
        var utc = d.ToUniversalTime();
        return utc.AddTicks(-(utc.Ticks % TimeSpan.TicksPerSecond));
    }
}
