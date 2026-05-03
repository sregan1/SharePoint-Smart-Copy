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

    [JsonIgnore]
    public string StatusDisplay => Status switch
    {
        CopyStatus.Success => "✅ Success",
        CopyStatus.Failed  => "❌ Failed",
        CopyStatus.Skipped => "⏭ Skipped",
        _                  => Status.ToString()
    };

    [JsonIgnore]
    public string StatusColor => Status switch
    {
        CopyStatus.Success => "#107C10",
        CopyStatus.Failed  => "#A4262C",
        CopyStatus.Skipped => "#797775",
        _                  => "#323130"
    };
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
    public int TotalCount { get; set; }
    public TimeSpan Duration { get; set; }
    public List<SavedReportItem> Items { get; set; } = [];

    [JsonIgnore]
    public string DisplayDate => Timestamp.LocalDateTime.ToString("MMM d, yyyy  h:mm tt");

    [JsonIgnore]
    public string DurationDisplay
    {
        get
        {
            if (Duration.TotalHours >= 1)  return $"{(int)Duration.TotalHours}h {Duration.Minutes}m {Duration.Seconds}s";
            if (Duration.TotalMinutes >= 1) return $"{(int)Duration.TotalMinutes}m {Duration.Seconds}s";
            return $"{Duration.Seconds}s";
        }
    }

    [JsonIgnore]
    public string Summary => $"✅ {SuccessCount}   ❌ {FailedCount}   ⏭ {SkippedCount}   ⏱ {DurationDisplay}";
}
