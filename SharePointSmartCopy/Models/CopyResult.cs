using System.ComponentModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointSmartCopy.Models;

public enum CopyStatus { Pending, Copying, Success, Failed, Skipped }

// Which rows the copy-log grids display (chips above the log).
public enum ResultFilterKind { All, Success, Failed, Skipped }

public partial class CopyResult : ObservableObject
{
    // Skip reason for Copy-if-newer: file exists at the target and is not older.
    // Compared against ErrorMessage to decide whether permissions still refresh.
    public const string UpToDate = "Up to date";

    protected override void OnPropertyChanged(PropertyChangedEventArgs e)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
            dispatcher.Invoke(() => base.OnPropertyChanged(e));
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
        CopyStatus.Copying  => "⟳ Copying…",
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
