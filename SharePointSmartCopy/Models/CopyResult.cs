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

    public string StatusDisplay => Status switch
    {
        CopyStatus.Pending  => "⏳ Pending",
        CopyStatus.Copying  => "⟳ Processing…",
        CopyStatus.Success  => "✅ Success",
        CopyStatus.Failed   => "❌ Failed",
        CopyStatus.Skipped  => "⏭ Skipped",
        _                   => string.Empty
    };

    public string StatusColor => Status switch
    {
        CopyStatus.Success  => "#107C10",
        CopyStatus.Failed   => "#A4262C",
        CopyStatus.Skipped  => "#797775",
        CopyStatus.Copying  => "#0078D4",
        _                   => "#323130"
    };

    public string PermissionStatusDisplay => PermissionStatus switch
    {
        CopyStatus.Success => "✅ Success",
        CopyStatus.Failed  => "❌ Failed",
        CopyStatus.Skipped => "⏭ Skipped",
        _                  => "—"
    };

    public string PermissionStatusColor => PermissionStatus switch
    {
        CopyStatus.Success => "#107C10",
        CopyStatus.Failed  => "#A4262C",
        CopyStatus.Skipped => "#797775",
        _                  => "#797775"
    };
}
