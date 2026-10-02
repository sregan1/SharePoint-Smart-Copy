using SharePointSmartCopy.Localization;
using System.ComponentModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointSmartCopy.Models;

// Cancelled: the item was still Copying (never resolved) when the run was cancelled or the app
// closed mid-copy — distinct from Failed, which means the item was actually attempted and a real
// error occurred. Conflating the two used to dump every in-flight item into Failed on shutdown,
// making an interrupted run's saved report look like a mass failure (e.g. 5,295 "failed" out of
// 5,295 remaining, when in fact none of them had even been attempted yet).
public enum CopyStatus { Pending, Copying, Success, Failed, Skipped, Cancelled }

// Which rows the copy-log grids display (chips above the log).
public enum ResultFilterKind { All, Success, Failed, Skipped }

public partial class CopyResult : ObservableObject
{
    // Skip reason for Copy-if-newer: file exists at the target and is not older.
    // Compared against ErrorMessage to decide whether permissions still refresh.
    public const string UpToDate = "Up to date";

    // Exposed so ProcessDiagnostics can report the current BeginInvoke backlog — if background
    // threads enqueue UI updates faster than the dispatcher drains them, this grows unbounded
    // with no other visible symptom before a crash.
    private static int _pendingUiDispatches;
    public static int PendingUiDispatches => _pendingUiDispatches;

    protected override void OnPropertyChanged(PropertyChangedEventArgs e)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            // BeginInvoke, not Invoke: this fires from many concurrent background upload/download
            // threads (one per in-flight file) during large migration jobs. A blocking Invoke here
            // forces every one of those threads to stall on the UI thread's dispatcher queue, which
            // under sustained high-concurrency load is a known trigger for WPF's composition engine
            // to fail with UCEERR_RENDERTHREADFAILURE. The backing field is already set by the time
            // this runs, so nothing depends on the notification completing synchronously.
            Interlocked.Increment(ref _pendingUiDispatches);
            dispatcher.BeginInvoke(() =>
            {
                base.OnPropertyChanged(e);
                Interlocked.Decrement(ref _pendingUiDispatches);
            });
        }
        else
            base.OnPropertyChanged(e);
    }

    [ObservableProperty] private string _fileName = string.Empty;
    [ObservableProperty] private string _sourcePath = string.Empty;
    [ObservableProperty] private string _targetPath = string.Empty;
    // Captured from CopyJob.SourceSize at creation (byte-accurate ETA needs it) — null for rows
    // that don't represent a single sized file (permission-only rows, library/list metadata rows).
    // Always the CURRENT version's size only, even when version history is copied — ETA pacing
    // needs a size known up front, and every version's size isn't known until the copy itself
    // fetches the version list.
    public long? SourceSize { get; init; }

    // Sum of every version's byte size actually copied for this file, set by CopyService /
    // MigrationJobService once the version-replay loop finishes (before Status flips to Success) —
    // null when version copying isn't in play (single-version copy, permission-only rows), in which
    // case SourceSize is the whole story. Used for the post-copy "Total Size" figures (Step 5 tile,
    // History) so a run with version history reflects the bytes actually transferred rather than
    // just the current version — see MainViewModel's _bytesFinalTally.
    public long? VersionsBytesTotal { get; set; }
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusDisplay))]
    [NotifyPropertyChangedFor(nameof(StatusColor))]
    private CopyStatus _status = CopyStatus.Pending;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private int _versionsCopied;
    [ObservableProperty] private int _versionsTotal;
    [ObservableProperty] private bool _isLibraryCreation;
    [ObservableProperty] private bool _isPermissionResult;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PermissionStatusDisplay))]
    [NotifyPropertyChangedFor(nameof(PermissionStatusColor))]
    private CopyStatus? _permissionStatus;

    [ObservableProperty] private string? _permissionDetails;

    // Set only when a custom-column value was actually attempted for this row (copyCustomColumns
    // on AND the source item had at least one cached field value) — null otherwise, same "not
    // attempted" convention as PermissionStatus. Failed here means at least one submitted column
    // could not be written (mismatched/missing target column, content-type mismatch, etc.) — the
    // FILE itself still copied fine, so this is deliberately a separate signal from Status/Failed
    // rather than folded into it.
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CustomFieldStatusDisplay))]
    [NotifyPropertyChangedFor(nameof(CustomFieldStatusColor))]
    private CopyStatus? _customFieldStatus;

    [ObservableProperty] private string? _customFieldDetails;

    // Raised synchronously on whatever thread sets the property (often a background copy/permission
    // worker) so a listener can maintain O(1) incremental tallies instead of rescanning the whole
    // result set on every UI tick — see MainViewModel's SuccessCount/FailedCount/etc., which used to
    // each run a full CopyResults.Count(predicate) every 400 ms (~1.75M predicate calls/tick at 250k
    // rows). Listeners must only do cheap, thread-safe bookkeeping (e.g. Interlocked) here — this is
    // NOT marshaled to the UI thread the way OnPropertyChanged below is.
    public event Action<CopyResult, CopyStatus, CopyStatus>? StatusChanging;
    partial void OnStatusChanging(CopyStatus oldValue, CopyStatus newValue) => StatusChanging?.Invoke(this, oldValue, newValue);

    public event Action<CopyResult, CopyStatus?, CopyStatus?>? PermissionStatusChanging;
    partial void OnPermissionStatusChanging(CopyStatus? oldValue, CopyStatus? newValue) => PermissionStatusChanging?.Invoke(this, oldValue, newValue);

    public event Action<CopyResult, CopyStatus?, CopyStatus?>? CustomFieldStatusChanging;
    partial void OnCustomFieldStatusChanging(CopyStatus? oldValue, CopyStatus? newValue) => CustomFieldStatusChanging?.Invoke(this, oldValue, newValue);

    // FileFailedCount counts Status==Failed OR PermissionStatus==Failed as one thing; this remembers
    // which side of that OR last drove the count so a listener can detect the combined predicate's
    // transition without re-deriving it from both properties independently.
    internal bool CountedAsFileFailed;

    public string StatusDisplay => Status switch
    {
        CopyStatus.Pending   => Loc.T("VM_StatusPending"),
        CopyStatus.Copying   => Loc.T("VM_StatusProcessing"),
        CopyStatus.Success   => Loc.T("VM_StatusSuccess"),
        CopyStatus.Failed    => Loc.T("VM_StatusFailed"),
        CopyStatus.Skipped   => Loc.T("VM_StatusSkipped"),
        CopyStatus.Cancelled => Loc.T("VM_StatusCancelledSymbol"),
        _                    => string.Empty
    };

    public string StatusColor => Status switch
    {
        CopyStatus.Success   => "#107C10",
        CopyStatus.Failed    => "#A4262C",
        CopyStatus.Skipped   => "#797775",
        CopyStatus.Copying   => "#0078D4",
        CopyStatus.Cancelled => "#797775",
        _                    => "#323130"
    };

    public string PermissionStatusDisplay => PermissionStatus switch
    {
        CopyStatus.Success => Loc.T("VM_StatusSuccess"),
        CopyStatus.Failed  => Loc.T("VM_StatusFailed"),
        CopyStatus.Skipped => Loc.T("VM_StatusSkipped"),
        _                  => Loc.T("VM_Dash")
    };

    public string PermissionStatusColor => PermissionStatus switch
    {
        CopyStatus.Success => "#107C10",
        CopyStatus.Failed  => "#A4262C",
        CopyStatus.Skipped => "#797775",
        _                  => "#797775"
    };

    // "Warning" rather than "Failed" for the display text: a custom-field mismatch means the FILE
    // still copied successfully — only some metadata values didn't make it across — so labeling it
    // the same as a hard copy failure would overstate the severity.
    public string CustomFieldStatusDisplay => CustomFieldStatus switch
    {
        CopyStatus.Success => Loc.T("VM_StatusApplied"),
        CopyStatus.Failed  => Loc.T("VM_StatusWarning"),
        _                  => Loc.T("VM_Dash")
    };

    public string CustomFieldStatusColor => CustomFieldStatus switch
    {
        CopyStatus.Success => "#107C10",
        CopyStatus.Failed  => "#C19C00",
        _                  => "#797775"
    };
}
