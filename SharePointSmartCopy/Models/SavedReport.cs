using SharePointSmartCopy.Localization;
using System.Text.Json.Serialization;

namespace SharePointSmartCopy.Models;

public class SavedReportItem
{
    public string FileName { get; set; } = string.Empty;
    public string SourcePath { get; set; } = string.Empty;
    public string TargetPath { get; set; } = string.Empty;

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public CopyStatus Status { get; set; }

    public int VersionsCopied { get; set; }
    public int VersionsTotal { get; set; }
    public string? ErrorMessage { get; set; }
    public bool IsPermissionResult { get; set; }

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public CopyStatus? PermissionStatus { get; set; }
    public string? PermissionDetails { get; set; }

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public CopyStatus? CustomFieldStatus { get; set; }
    public string? CustomFieldDetails { get; set; }

    [JsonIgnore]
    public string StatusDisplay => Status switch
    {
        CopyStatus.Success   => Loc.T("VM_StatusSuccess"),
        CopyStatus.Failed    => Loc.T("VM_StatusFailed"),
        CopyStatus.Skipped   => Loc.T("VM_StatusSkipped"),
        CopyStatus.Cancelled => Loc.T("VM_StatusCancelledBlocked"),
        _                    => Status.ToString()
    };

    [JsonIgnore]
    public string StatusColor => Status switch
    {
        CopyStatus.Success   => "#107C10",
        CopyStatus.Failed    => "#A4262C",
        CopyStatus.Skipped   => "#797775",
        CopyStatus.Cancelled => "#797775",
        _                    => "#323130"
    };
}

// Root scan scope for a saved run, captured at save time so a "Verify" re-scan can be launched
// later from History — even in a future app session where the live CopyJobs list no longer
// exists. Mirrors VerificationRoot's fields (kept as a separate, plain-data, JSON-serializable
// type rather than reusing VerificationRoot directly, consistent with how SavedReportItem
// mirrors CopyResult rather than being it).
public class SavedReportRoot
{
    public string SourceDriveId { get; set; } = string.Empty;
    public string SourceItemId { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;
    public bool IsFolder { get; set; }
    public bool IsLibrary { get; set; }
    public string TargetDriveId { get; set; } = string.Empty;
    public string TargetParentItemId { get; set; } = string.Empty;
    public string TargetSubFolderPath { get; set; } = string.Empty;
}

public class SavedReport
{
    public string Id { get; set; } = string.Empty;
    public DateTimeOffset Timestamp { get; set; }
    public string SourceUrl { get; set; } = string.Empty;
    public string TargetUrl { get; set; } = string.Empty;
    public int SuccessCount { get; set; }
    public int FailedCount { get; set; }
    public int SkippedCount { get; set; }
    // Still-Copying items swept up when the run was cancelled or the app closed mid-copy — never
    // actually attempted, so excluded from FailedCount (see CopyStatus.Cancelled).
    public int CancelledCount { get; set; }
    public int TotalCount { get; set; }
    public TimeSpan Duration { get; set; }
    // Sum of CopyResult.SourceSize across the run. 0 for runs where size wasn't known for any row
    // (e.g. Library/Site-scope or permission-only copies) — see MainViewModel.HasKnownTotalSize for
    // the same "don't show a misleading partial sum" reasoning applied to the live progress screens.
    public long TotalSize { get; set; }

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public CopyMode CopyMode { get; set; }

    public List<SavedReportItem> Items { get; set; } = [];

    // Populated from CopyJobs at save time. Empty for reports saved before verification support
    // shipped — that's the signal used to disable the History "Verify" button for old reports.
    public List<SavedReportRoot> Roots { get; set; } = [];

    [JsonIgnore]
    public string DisplayDate => Timestamp.LocalDateTime.ToString("MMM d, yyyy  h:mm tt");

    [JsonIgnore]
    public string DurationDisplay
    {
        get
        {
            if (Duration.TotalHours >= 1)   return Loc.T("VM_DurationHMS", (int)Duration.TotalHours, Duration.Minutes, Duration.Seconds);
            if (Duration.TotalMinutes >= 1) return Loc.T("VM_DurationMS", (int)Duration.TotalMinutes, Duration.Seconds);
            return Loc.T("VM_DurationS", Duration.Seconds);
        }
    }

    [JsonIgnore]
    public string SizeDisplay => TotalSize > 0 ? SizeFormatter.FormatBytes(TotalSize) : string.Empty;

    [JsonIgnore]
    public string Summary => CancelledCount > 0
        ? Loc.T("VM_SummaryWithCancelled", SuccessCount, FailedCount, SkippedCount, CancelledCount, DurationDisplay, SizeSummarySuffix)
        : Loc.T("VM_Summary", SuccessCount, FailedCount, SkippedCount, DurationDisplay, SizeSummarySuffix);

    private string SizeSummarySuffix => TotalSize > 0 ? $"   ·  {SizeDisplay}" : string.Empty;
}

// Everything the History list needs to show one row, deliberately WITHOUT Items. Deserializing a
// report's JSON into this type instead of SavedReport makes System.Text.Json skip over the (often
// huge) Items array token-by-token rather than materializing a SavedReportItem per file — on a
// tenant with several 100,000+-file runs in its history, decoding all 50 saved reports' full Items
// graphs just to show a one-line summary per run was the dominant cost of opening History. Roots is
// included (cheap — one entry per copy root, not per file) since the Verify button's
// enable/disable check and the verification re-scan itself both only need Roots, not Items. Load a
// specific report's Items lazily via ReportHistoryService.LoadFull only once the user actually
// selects that run or exports/verifies it.
public class SavedReportSummary
{
    public string Id { get; set; } = string.Empty;
    public DateTimeOffset Timestamp { get; set; }
    public string SourceUrl { get; set; } = string.Empty;
    public string TargetUrl { get; set; } = string.Empty;
    public int SuccessCount { get; set; }
    public int FailedCount { get; set; }
    public int SkippedCount { get; set; }
    public int CancelledCount { get; set; }
    public int TotalCount { get; set; }
    public TimeSpan Duration { get; set; }
    public long TotalSize { get; set; }

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public CopyMode CopyMode { get; set; }

    public List<SavedReportRoot> Roots { get; set; } = [];

    [JsonIgnore]
    public string DisplayDate => Timestamp.LocalDateTime.ToString("MMM d, yyyy  h:mm tt");

    [JsonIgnore]
    public string DurationDisplay
    {
        get
        {
            if (Duration.TotalHours >= 1)   return Loc.T("VM_DurationHMS", (int)Duration.TotalHours, Duration.Minutes, Duration.Seconds);
            if (Duration.TotalMinutes >= 1) return Loc.T("VM_DurationMS", (int)Duration.TotalMinutes, Duration.Seconds);
            return Loc.T("VM_DurationS", Duration.Seconds);
        }
    }

    [JsonIgnore]
    public string SizeDisplay => TotalSize > 0 ? SizeFormatter.FormatBytes(TotalSize) : string.Empty;

    [JsonIgnore]
    public string Summary => CancelledCount > 0
        ? Loc.T("VM_SummaryWithCancelled", SuccessCount, FailedCount, SkippedCount, CancelledCount, DurationDisplay, SizeSummarySuffix)
        : Loc.T("VM_Summary", SuccessCount, FailedCount, SkippedCount, DurationDisplay, SizeSummarySuffix);

    private string SizeSummarySuffix => TotalSize > 0 ? $"   ·  {SizeDisplay}" : string.Empty;
}
