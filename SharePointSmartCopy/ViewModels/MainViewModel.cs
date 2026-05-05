using System.Collections.ObjectModel;
using System.Text;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.ViewModels;

public partial class MainViewModel : ObservableObject
{
    public readonly AuthService AuthService;
    public readonly SharePointService SpService;
    private readonly CopyService _copyService;

    private CancellationTokenSource? _copyCts;
    private CancellationTokenSource? _connectSourceCts;
    private CancellationTokenSource? _connectTargetCts;

    public MainViewModel(AuthService? existingAuthService = null)
    {
        AuthService  = existingAuthService ?? new AuthService();
        SpService    = new SharePointService(AuthService);
        _copyService = new CopyService(SpService);
        Settings     = AppSettings.Load();

        if (Settings.IsConfigured)
        {
            // Only configure (and reset auth) when creating a fresh service
            if (existingAuthService == null)
                AuthService.Configure(Settings);
            SpService.Initialize();
        }

        SourceUrl = Settings.SourceUrl;
        TargetUrl = Settings.TargetUrl;
    }

    // ── Settings ──────────────────────────────────────────────────────────────

    [ObservableProperty] private AppSettings _settings;

    public void ApplySettings(AppSettings s)
    {
        Settings = s;
        AuthService.Configure(s);
        SpService.Initialize();
    }

    // ── Step navigation ───────────────────────────────────────────────────────
    // Steps: 0=Source  1=Browse  2=Target  3=Options  4=Copying  5=Report

    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(BackCommand))]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    private int _currentStep;

    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private bool _isBusy;

    // ── Step 0: Source ────────────────────────────────────────────────────────

    [ObservableProperty] private string _sourceUrl = string.Empty;
    [ObservableProperty] private string _sourceStatus = string.Empty;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] private bool _sourceConnected;
    [ObservableProperty] private string _sourceSiteId = string.Empty;
    [ObservableProperty] private string _signedInUser = string.Empty;
    [ObservableProperty] private bool _isConnectingSource;

    [RelayCommand]
    private async Task ConnectSourceAsync()
    {
        _connectSourceCts?.Cancel();
        _connectSourceCts = new CancellationTokenSource();
        var ct = _connectSourceCts.Token;

        SourceStatus       = "Connecting…";
        SourceConnected    = false;
        IsBusy             = true;
        IsConnectingSource = true;
        try
        {
            // Use silent auth when we already have a cached session (e.g. new copy reusing credentials)
            await AuthService.GetAccessTokenAsync(forceInteractive: !AuthService.IsAuthenticated, cancellationToken: ct);
            ct.ThrowIfCancellationRequested();
            SignedInUser = AuthService.UserName ?? string.Empty;
            SourceSiteId = await SpService.GetSiteIdAsync(SourceUrl.Trim());
            ct.ThrowIfCancellationRequested();
            SourceStatus    = $"✅ Connected as {SignedInUser}";
            SourceConnected = true;
            Settings.SourceUrl = SourceUrl.Trim();
            Settings.Save();
        }
        catch (OperationCanceledException)
        {
            SourceStatus = string.Empty;
        }
        catch (Exception ex)
        {
            SourceStatus = $"❌ {ex.Message}";
        }
        finally
        {
            IsBusy             = false;
            IsConnectingSource = false;
        }
    }

    [RelayCommand]
    private void CancelConnectSource()
    {
        _connectSourceCts?.Cancel();
        IsConnectingSource = false;
        SourceStatus       = string.Empty;
        IsBusy             = false;
    }

    [RelayCommand]
    private void DisconnectSource()
    {
        SourceConnected = false;
        SourceStatus    = string.Empty;
        SourceSiteId    = string.Empty;
    }

    // ── Step 1: Browse ────────────────────────────────────────────────────────

    [ObservableProperty] private ObservableCollection<SharePointNode> _sourceLibraries = [];

    public async Task LoadLibrariesAsync()
    {
        IsBusy = true;
        StatusMessage = "Loading libraries…";
        try
        {
            var libs = await SpService.GetLibrariesAsync(SourceSiteId);
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                SourceLibraries.Clear();
                foreach (var lib in libs)
                    SourceLibraries.Add(lib);
            });
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error loading libraries: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
            StatusMessage = string.Empty;
        }
    }

    public async Task LoadNodeChildrenAsync(SharePointNode node)
    {
        if (!node.HasChildren) return;
        if (!node.Children.Any(c => c.IsPlaceholder)) return;

        node.IsLoading = true;
        try
        {
            var children = await SpService.GetChildrenAsync(node.DriveId, node.Id, node.SiteId);
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                node.Children.Clear();
                foreach (var child in children)
                {
                    child.Parent = node;
                    node.Children.Add(child);
                }
                if (!node.Children.Any())
                    node.HasChildren = false;
            });
        }
        catch (Exception ex)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() => node.Children.Clear());
            StatusMessage = $"Error loading folder: {ex.Message}";
        }
        finally { node.IsLoading = false; }
    }

    public void SelectAllSource(bool value)
    {
        foreach (var lib in SourceLibraries)
            lib.IsChecked = value;
    }

    public int SelectedSourceCount
    {
        get
        {
            int count = 0;
            foreach (var lib in SourceLibraries)
                count += lib.GetCheckedNodes().Count();
            return count;
        }
    }

    // ── Step 2: Target ────────────────────────────────────────────────────────

    [ObservableProperty] private string _targetUrl = string.Empty;
    [ObservableProperty] private string _targetStatus = string.Empty;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] private bool _targetConnected;
    [ObservableProperty] private string _targetSiteId = string.Empty;
    [ObservableProperty] private ObservableCollection<SharePointNode> _targetLibraries = [];
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] private SharePointNode? _selectedTargetFolder;
    [ObservableProperty] private bool _isConnectingTarget;

    [RelayCommand]
    private async Task ConnectTargetAsync()
    {
        _connectTargetCts?.Cancel();
        _connectTargetCts = new CancellationTokenSource();
        var ct = _connectTargetCts.Token;

        TargetStatus       = "Connecting…";
        TargetConnected    = false;
        IsBusy             = true;
        IsConnectingTarget = true;
        try
        {
            TargetSiteId = await SpService.GetSiteIdAsync(TargetUrl.Trim());
            ct.ThrowIfCancellationRequested();
            var libs = await SpService.GetLibrariesAsync(TargetSiteId);
            ct.ThrowIfCancellationRequested();
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                TargetLibraries.Clear();
                foreach (var lib in libs)
                    TargetLibraries.Add(lib);
            });
            TargetStatus    = "✅ Connected";
            TargetConnected = true;
            Settings.TargetUrl = TargetUrl.Trim();
            Settings.Save();

            // Pre-warm the SharePoint REST token so the consent dialog (if needed) happens
            // here — on the UI thread, while the user is already waiting — rather than
            // unexpectedly mid-copy on a background thread.
            try { await AuthService.GetSharePointTokenAsync(TargetUrl.Trim(), ct); }
            catch
            {
                TargetStatus = "✅ Connected · Note: additional consent needed for metadata — reconnect to grant";
            }
        }
        catch (OperationCanceledException)
        {
            TargetStatus = string.Empty;
        }
        catch (Exception ex)
        {
            TargetStatus = $"❌ {ex.Message}";
        }
        finally
        {
            IsBusy             = false;
            IsConnectingTarget = false;
        }
    }

    [RelayCommand]
    private void CancelConnectTarget()
    {
        _connectTargetCts?.Cancel();
        IsConnectingTarget = false;
        TargetStatus       = string.Empty;
        IsBusy             = false;
    }

    [RelayCommand]
    private void DisconnectTarget()
    {
        TargetConnected      = false;
        TargetStatus         = string.Empty;
        TargetSiteId         = string.Empty;
        SelectedTargetFolder = null;
        TargetLibraries.Clear();
    }

    public async Task LoadTargetNodeChildrenAsync(SharePointNode node)
    {
        if (!node.HasChildren) return;
        if (!node.Children.Any(c => c.IsPlaceholder)) return;

        node.IsLoading = true;
        try
        {
            var children = await SpService.GetChildrenAsync(node.DriveId, node.Id, node.SiteId, foldersOnly: true);
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                node.Children.Clear();
                foreach (var child in children)
                {
                    child.Parent = node;
                    node.Children.Add(child);
                }
            });
        }
        catch { System.Windows.Application.Current.Dispatcher.Invoke(() => node.Children.Clear()); }
        finally { node.IsLoading = false; }
    }

    // ── Step 3: Options ───────────────────────────────────────────────────────

    [ObservableProperty] private bool _overwriteFiles = false;
    [ObservableProperty] private bool _copyVersions = true;
    [ObservableProperty] private bool _copyAllVersions = true;
    [ObservableProperty] private int _maxVersions = 10;
    [ObservableProperty] private int _maxParallelCopies = 4;
    // true  → preserve per-version metadata via PATCH+delete (version numbers become non-sequential)
    // false → keep version numbers sequential, only the latest version's metadata is synced
    [ObservableProperty] private bool _preserveVersionMetadata = true;
    [ObservableProperty] private ObservableCollection<CopyJob> _copyJobs = [];

    public void BuildCopyJobs()
    {
        CopyJobs.Clear();

        if (SelectedTargetFolder == null) return;

        foreach (var lib in SourceLibraries)
        {
            foreach (var node in lib.GetCheckedNodes())
            {
                var job = new CopyJob
                {
                    SourceDriveId      = node.DriveId,
                    SourceItemId       = node.Id,
                    SourceName         = node.Name,
                    SourceDisplayPath  = BuildPath(node),
                    TargetDriveId      = SelectedTargetFolder.DriveId,
                    TargetParentItemId = SelectedTargetFolder.Id,
                    TargetSiteId       = SelectedTargetFolder.SiteId,
                    TargetDisplayPath  = node.Name,
                    IsFolder           = node.Type != NodeType.File
                };
                CopyJobs.Add(job);
            }
        }
    }

    // ── Step 4: Copying ───────────────────────────────────────────────────────

    private readonly object _copyResultsLock = new();
    [ObservableProperty] private ObservableCollection<CopyResult> _copyResults = [];

    partial void OnCopyResultsChanged(ObservableCollection<CopyResult> value)
        => BindingOperations.EnableCollectionSynchronization(value, _copyResultsLock);
    [ObservableProperty] private double _totalProgress;
    [ObservableProperty] private int _completedCount;
    [ObservableProperty] private int _totalCount;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] [NotifyCanExecuteChangedFor(nameof(BackCommand))] private bool _isCopying;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] private bool _isCopyComplete;
    [ObservableProperty] private string _copyDuration = string.Empty;
    [ObservableProperty] private string _elapsedTime = string.Empty;

    private DateTimeOffset _copyStartTime;

    public int SuccessCount => CopyResults.Count(r => r.Status == CopyStatus.Success);
    public int FailedCount  => CopyResults.Count(r => r.Status == CopyStatus.Failed);
    public int SkippedCount => CopyResults.Count(r => r.Status == CopyStatus.Skipped);

    [RelayCommand]
    private async Task StartCopyAsync()
    {
        IsCopying      = true;
        IsCopyComplete = false;
        CopyDuration   = string.Empty;
        CopyResults.Clear();
        CompletedCount = 0;
        TotalProgress  = 0;
        _copyStartTime = DateTimeOffset.Now;

        // Pre-populate results for file jobs (folder jobs add results dynamically)
        foreach (var job in CopyJobs.Where(j => !j.IsFolder))
        {
            CopyResults.Add(new CopyResult
            {
                FileName   = job.SourceName,
                SourcePath = job.SourceDisplayPath,
                TargetPath = job.TargetDisplayPath
            });
        }

        _copyCts = new CancellationTokenSource();
        TotalCount = CopyJobs.Count;

        try
        {
            var progressTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(400)
            };
            progressTimer.Tick += (_, _) => UpdateProgress();
            progressTimer.Start();

            // 0 = all versions; positive number = copy only the N most recent
            int versionLimit = CopyVersions && !CopyAllVersions ? MaxVersions : 0;
            await _copyService.ExecuteAsync(
                CopyJobs,
                CopyResults,
                OverwriteFiles,
                CopyVersions,
                MaxParallelCopies,
                versionLimit,
                PreserveVersionMetadata,
                _copyCts.Token);

            progressTimer.Stop();
            UpdateProgress();
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Copy cancelled.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Copy error: {ex.Message}";
        }
        finally
        {
            IsCopying      = false;
            IsCopyComplete = true;
            TotalCount     = CopyResults.Count;
            CopyDuration   = FormatDuration(DateTimeOffset.Now - _copyStartTime);
            UpdateProgress();
            OnPropertyChanged(nameof(SuccessCount));
            OnPropertyChanged(nameof(FailedCount));
            OnPropertyChanged(nameof(SkippedCount));
            SaveReport();
        }
    }

    [RelayCommand]
    private void CancelCopy() => _copyCts?.Cancel();

    private void UpdateProgress()
    {
        var done = CopyResults.Count(r => r.Status is CopyStatus.Success or CopyStatus.Failed or CopyStatus.Skipped);
        CompletedCount = done;
        TotalCount     = CopyResults.Count;
        TotalProgress  = TotalCount > 0 ? done * 100.0 / TotalCount : 0;
        ElapsedTime    = FormatDuration(DateTimeOffset.Now - _copyStartTime);
    }

    public void RefreshCopyStats()
    {
        OnPropertyChanged(nameof(SuccessCount));
        OnPropertyChanged(nameof(FailedCount));
        OnPropertyChanged(nameof(SkippedCount));
    }

    // ── Step 5: Report ────────────────────────────────────────────────────────

    [RelayCommand]
    private void ExportReport()
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt",
            FileName = $"CopyReport_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true) return;

        var sb = new StringBuilder();
        sb.AppendLine("File Name,Source Path,Target Path,Status,Versions Copied,Error");
        foreach (var r in CopyResults)
        {
            sb.AppendLine($"\"{r.FileName}\",\"{r.SourcePath}\",\"{r.TargetPath}\"," +
                          $"{r.Status},{r.VersionsCopied},\"{r.ErrorMessage}\"");
        }
        System.IO.File.WriteAllText(dlg.FileName, sb.ToString());
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    [RelayCommand(CanExecute = nameof(CanGoBack))]
    private void Back() => CurrentStep--;

    private bool CanGoBack() => CurrentStep > 0 && !IsCopying;

    [RelayCommand(CanExecute = nameof(CanGoNext))]
    private void Next()
    {
        switch (CurrentStep)
        {
            case 0 when SourceConnected:
                CurrentStep = 1;
                _ = LoadLibrariesAsync();
                break;
            case 1:
                CurrentStep = 2;
                break;
            case 2 when TargetConnected && SelectedTargetFolder != null:
                BuildCopyJobs();
                CurrentStep = 3;
                break;
            case 3 when CopyJobs.Count > 0:
                CurrentStep = 4;
                _ = StartCopyAsync();
                break;
            case 4 when IsCopyComplete:
                CurrentStep = 5;
                break;
        }
    }

    private bool CanGoNext() => CurrentStep switch
    {
        0 => SourceConnected,
        1 => SourceLibraries.Any(l => l.GetCheckedNodes().Any()),
        2 => TargetConnected && SelectedTargetFolder != null,
        3 => CopyJobs.Count > 0,
        4 => IsCopyComplete,
        _ => false
    };

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string FormatDuration(TimeSpan ts)
    {
        if (ts.TotalHours >= 1)
            return $"{(int)ts.TotalHours}h {ts.Minutes}m {ts.Seconds}s";
        if (ts.TotalMinutes >= 1)
            return $"{(int)ts.TotalMinutes}m {ts.Seconds}s";
        return $"{ts.Seconds}s";
    }

    private void SaveReport()
    {
        try
        {
            var report = new SavedReport
            {
                Id           = _copyStartTime.ToString("yyyyMMdd_HHmmss"),
                Timestamp    = _copyStartTime,
                Duration     = DateTimeOffset.Now - _copyStartTime,
                SourceUrl    = SourceUrl,
                TargetUrl    = TargetUrl,
                SuccessCount = SuccessCount,
                FailedCount  = FailedCount,
                SkippedCount = SkippedCount,
                TotalCount   = TotalCount,
                Items        = CopyResults.Select(r => new SavedReportItem
                {
                    FileName       = r.FileName,
                    SourcePath     = r.SourcePath,
                    TargetPath     = r.TargetPath,
                    Status         = r.Status,
                    VersionsCopied = r.VersionsCopied,
                    VersionsTotal  = r.VersionsTotal,
                    ErrorMessage   = r.ErrorMessage
                }).ToList()
            };
            ReportHistoryService.Save(report);
        }
        catch { /* non-critical */ }
    }

    private static string BuildPath(SharePointNode node)
    {
        var parts = new List<string>();
        var current = (SharePointNode?)node;
        while (current != null)
        {
            parts.Insert(0, current.Name);
            current = current.Parent;
        }
        return string.Join("/", parts);
    }
}
