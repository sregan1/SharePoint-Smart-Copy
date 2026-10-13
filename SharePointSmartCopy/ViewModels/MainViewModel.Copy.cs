using SharePointSmartCopy.Localization;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Text;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;
using SpColumnDef = SharePointSmartCopy.Models.ColumnDefinition;

namespace SharePointSmartCopy.ViewModels;

// Partial class MainViewModel: Copy. The shared fields, constructor and core setup live in MainViewModel.cs.
public partial class MainViewModel : ObservableObject
{
    // ── Step 4: Copying ───────────────────────────────────────────────────────

    private readonly object _copyResultsLock = new();
    [ObservableProperty] private ObservableCollection<CopyResult> _copyResults = [];

    // Incremental tallies for SuccessCount/FailedCount/etc below — see CopyResult.StatusChanging.
    // Kept as plain ints updated via Interlocked (not [ObservableProperty]): they're written from
    // whatever thread sets a row's Status/PermissionStatus, and only ever READ from the UI thread
    // inside UpdateProgress's periodic tick, which is what actually raises the property-changed
    // notifications for the bound counter properties below.
    private int _successTally, _failedTally, _skippedTally, _cancelledTally;
    private int _fileTotalTally, _fileSuccessTally, _fileFailedTally, _fileSkippedTally, _fileCancelledTally;

    // Deliberately NOT folded into _fileFailedTally (unlike PermissionStatus==Failed — see
    // UpdateFileFailedTally): a custom-field mismatch means the file itself copied fine, only some
    // metadata didn't make it across, so it's tracked as its own independent count rather than
    // counting the whole file as failed.
    private int _customFieldWarningTally;

    // Byte-weighted counterparts of the tallies above, for ETA — see UpdateProgress. Kept alongside
    // the item counts (not replacing them) because several flows (library/site copy, permissions)
    // create CopyResult rows with no meaningful SourceSize; ETA falls back to item-count for those.
    private long _bytesTotalTally, _bytesDoneTally, _packedBytesTally;
    private int  _bytesMissingCount;

    // Sum of CopyResult.VersionsBytesTotal (falling back to SourceSize) across rows that finished —
    // unlike _bytesTotalTally/_bytesDoneTally above, this only grows as each row COMPLETES (the true
    // per-version total isn't known until then), so it can't be used for ETA pacing during the run.
    // It's the authoritative "Total Size" once the run is done — see TotalSizeDisplay — since it
    // reflects every version actually copied, not just each file's current version.
    private long _bytesFinalTally;

    partial void OnCopyResultsChanged(ObservableCollection<CopyResult> value)
    {
        BindingOperations.EnableCollectionSynchronization(value, _copyResultsLock);
        value.CollectionChanged += OnCopyResultsCollectionChanged;
        _copyResultsView = null;
        _fileResultsView = null;
        OnPropertyChanged(nameof(CopyResultsView));
        OnPropertyChanged(nameof(FileResultsView));
    }

    private void OnCopyResultsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        switch (e.Action)
        {
            case NotifyCollectionChangedAction.Add:
                if (e.NewItems != null)
                    foreach (CopyResult row in e.NewItems) AttachRowTallies(row);
                break;
            case NotifyCollectionChangedAction.Remove:
                if (e.OldItems != null)
                    foreach (CopyResult row in e.OldItems) DetachRowTallies(row);
                break;
            case NotifyCollectionChangedAction.Reset:
                _successTally = _failedTally = _skippedTally = _cancelledTally = 0;
                _fileTotalTally = _fileSuccessTally = _fileFailedTally = _fileSkippedTally = _fileCancelledTally = 0;
                _customFieldWarningTally = 0;
                _bytesTotalTally = _bytesDoneTally = 0;
                _bytesMissingCount = 0;
                _bytesFinalTally = 0;
                break;
        }
    }

    private void AttachRowTallies(CopyResult row)
    {
        row.StatusChanging            += OnRowStatusChanging;
        row.PermissionStatusChanging  += OnRowPermissionStatusChanging;
        row.CustomFieldStatusChanging += OnRowCustomFieldStatusChanging;
        if (!row.IsPermissionResult)
            Interlocked.Increment(ref _fileTotalTally);
        if (row.SourceSize is long size)
            Interlocked.Add(ref _bytesTotalTally, size);
        else
            Interlocked.Increment(ref _bytesMissingCount);
        // Reconcile against whatever Status/PermissionStatus the row already carries — an object
        // initializer (e.g. `new CopyResult { Status = CopyStatus.Copying }`) fires the Changing hook
        // before this subscription exists, so a non-default starting value would otherwise never get
        // counted. Pending/null are always the true starting point since these hooks only fire on the
        // FIRST real transition after subscribing.
        OnRowStatusChanging(row, CopyStatus.Pending, row.Status);
        if (row.PermissionStatus != null)
            OnRowPermissionStatusChanging(row, null, row.PermissionStatus);
        if (row.CustomFieldStatus != null)
            OnRowCustomFieldStatusChanging(row, null, row.CustomFieldStatus);
    }

    private void DetachRowTallies(CopyResult row)
    {
        row.StatusChanging            -= OnRowStatusChanging;
        row.PermissionStatusChanging  -= OnRowPermissionStatusChanging;
        row.CustomFieldStatusChanging -= OnRowCustomFieldStatusChanging;
        if (!row.IsPermissionResult)
            Interlocked.Decrement(ref _fileTotalTally);
        if (row.SourceSize is long size)
            Interlocked.Add(ref _bytesTotalTally, -size);
        else
            Interlocked.Decrement(ref _bytesMissingCount);
        OnRowStatusChanging(row, row.Status, CopyStatus.Pending);
        if (row.PermissionStatus != null)
            OnRowPermissionStatusChanging(row, row.PermissionStatus, null);
        if (row.CustomFieldStatus != null)
            OnRowCustomFieldStatusChanging(row, row.CustomFieldStatus, null);
    }

    private void OnRowStatusChanging(CopyResult row, CopyStatus oldStatus, CopyStatus newStatus)
    {
        AdjustStatusTally(oldStatus, -1);
        AdjustStatusTally(newStatus, +1);
        AdjustBytesDoneTally(oldStatus, newStatus, row.SourceSize ?? 0);
        AdjustFinalBytesTally(oldStatus, newStatus, row.VersionsBytesTotal ?? row.SourceSize ?? 0);
        if (!row.IsPermissionResult)
        {
            AdjustFileStatusTally(oldStatus, -1);
            AdjustFileStatusTally(newStatus, +1);
            UpdateFileFailedTally(row, newStatus, row.PermissionStatus);
        }
    }

    private void OnRowPermissionStatusChanging(CopyResult row, CopyStatus? oldValue, CopyStatus? newValue)
    {
        if (!row.IsPermissionResult)
            UpdateFileFailedTally(row, row.Status, newValue);
    }

    // Independent of UpdateFileFailedTally on purpose — see _customFieldWarningTally's doc comment.
    private void OnRowCustomFieldStatusChanging(CopyResult row, CopyStatus? oldValue, CopyStatus? newValue)
    {
        if (oldValue == CopyStatus.Failed) Interlocked.Decrement(ref _customFieldWarningTally);
        if (newValue == CopyStatus.Failed) Interlocked.Increment(ref _customFieldWarningTally);
    }

    // FileFailedCount is Status==Failed OR PermissionStatus==Failed — tracked as one combined flag
    // per row (CopyResult.CountedAsFileFailed) so a file that fails on both sides, or recovers on
    // one, is never double-counted or left stuck.
    private void UpdateFileFailedTally(CopyResult row, CopyStatus status, CopyStatus? permissionStatus)
    {
        bool isFailedNow = status == CopyStatus.Failed || permissionStatus == CopyStatus.Failed;
        if (isFailedNow == row.CountedAsFileFailed) return;
        Interlocked.Add(ref _fileFailedTally, isFailedNow ? 1 : -1);
        row.CountedAsFileFailed = isFailedNow;
    }

    private void AdjustStatusTally(CopyStatus status, int delta)
    {
        switch (status)
        {
            case CopyStatus.Success:   Interlocked.Add(ref _successTally, delta);   break;
            case CopyStatus.Failed:    Interlocked.Add(ref _failedTally, delta);    break;
            case CopyStatus.Skipped:   Interlocked.Add(ref _skippedTally, delta);   break;
            case CopyStatus.Cancelled: Interlocked.Add(ref _cancelledTally, delta); break;
        }
    }

    // Mirrors the done/total split used for ETA (UpdateProgress: done = success+failed+skipped,
    // deliberately excluding Cancelled — see CopyStatus.Cancelled).
    private static bool CountsTowardEtaDone(CopyStatus status) =>
        status is CopyStatus.Success or CopyStatus.Failed or CopyStatus.Skipped;

    private void AdjustBytesDoneTally(CopyStatus oldStatus, CopyStatus newStatus, long size)
    {
        bool wasDone = CountsTowardEtaDone(oldStatus);
        bool isDone  = CountsTowardEtaDone(newStatus);
        if (wasDone == isDone) return;
        Interlocked.Add(ref _bytesDoneTally, isDone ? size : -size);
    }

    private void AdjustFinalBytesTally(CopyStatus oldStatus, CopyStatus newStatus, long size)
    {
        bool wasDone = CountsTowardEtaDone(oldStatus);
        bool isDone  = CountsTowardEtaDone(newStatus);
        if (wasDone == isDone) return;
        Interlocked.Add(ref _bytesFinalTally, isDone ? size : -size);
    }

    private void AdjustFileStatusTally(CopyStatus status, int delta)
    {
        switch (status)
        {
            case CopyStatus.Success:   Interlocked.Add(ref _fileSuccessTally, delta);   break;
            case CopyStatus.Skipped:   Interlocked.Add(ref _fileSkippedTally, delta);   break;
            case CopyStatus.Cancelled: Interlocked.Add(ref _fileCancelledTally, delta); break;
            // Failed is handled by UpdateFileFailedTally's combined predicate, not here.
        }
    }

    // ── Result filter (All / Failed / Skipped chips above the log grids) ──────
    // The grids bind to this filtered view; CopyResults itself stays complete so
    // report counts and CSV export are unaffected.
    [ObservableProperty] private ResultFilterKind _resultFilter = ResultFilterKind.All;

    partial void OnResultFilterChanged(ResultFilterKind value) => CopyResultsView.Refresh();

    // ── Step 5 tab-specific filters ───────────────────────────────────────────
    [ObservableProperty] private ResultFilterKind _fileResultFilter = ResultFilterKind.All;

    partial void OnFileResultFilterChanged(ResultFilterKind value) => _fileResultsView?.Refresh();

    private System.ComponentModel.ICollectionView? _copyResultsView;
    public System.ComponentModel.ICollectionView CopyResultsView
    {
        get
        {
            if (_copyResultsView == null)
            {
                _copyResultsView = CollectionViewSource.GetDefaultView(CopyResults);
                _copyResultsView.Filter = o =>
                {
                    var r = (CopyResult)o;
                    return ResultFilter switch
                    {
                        ResultFilterKind.Success => r.Status == CopyStatus.Success,
                        // Cancelled items surface under Failed too — they represent unfinished
                        // work a user reviewing "what didn't succeed" needs to see, even though
                        // they're tallied separately (CancelledCount) so the summary stays honest.
                        ResultFilterKind.Failed  => r.Status == CopyStatus.Failed || r.Status == CopyStatus.Cancelled || r.PermissionStatus == CopyStatus.Failed,
                        ResultFilterKind.Skipped => r.Status == CopyStatus.Skipped,
                        _                        => true,
                    };
                };
            }
            return _copyResultsView;
        }
    }

    private System.ComponentModel.ICollectionView? _fileResultsView;
    public System.ComponentModel.ICollectionView FileResultsView
    {
        get
        {
            if (_fileResultsView == null)
            {
                _fileResultsView = new CollectionViewSource { Source = CopyResults }.View;
                _fileResultsView.Filter = o =>
                {
                    var r = (CopyResult)o;
                    if (r.IsPermissionResult) return false;
                    return FileResultFilter switch
                    {
                        ResultFilterKind.Success => r.Status == CopyStatus.Success,
                        // Cancelled items surface under Failed too — they represent unfinished
                        // work a user reviewing "what didn't succeed" needs to see, even though
                        // they're tallied separately (CancelledCount) so the summary stays honest.
                        ResultFilterKind.Failed  => r.Status == CopyStatus.Failed || r.Status == CopyStatus.Cancelled || r.PermissionStatus == CopyStatus.Failed,
                        ResultFilterKind.Skipped => r.Status == CopyStatus.Skipped,
                        _                        => true,
                    };
                };
            }
            return _fileResultsView;
        }
    }
    [ObservableProperty] private double _totalProgress;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(IsPackingInProgress))] private int _completedCount;
    [ObservableProperty] private int _totalCount;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] [NotifyCanExecuteChangedFor(nameof(BackCommand))] [NotifyCanExecuteChangedFor(nameof(ExportVerificationReportCommand))] [NotifyPropertyChangedFor(nameof(IsCancelable))] private bool _isCopying;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] [NotifyPropertyChangedFor(nameof(IsMetadataComplete))] [NotifyPropertyChangedFor(nameof(IsMetadataInProgress))] [NotifyPropertyChangedFor(nameof(IsReadyForReport))] private bool _isCopyComplete;
    [ObservableProperty] private string _copyDuration = string.Empty;
    [ObservableProperty] private string _elapsedTime = string.Empty;
    [ObservableProperty] private string _estimatedTimeRemaining = string.Empty;
    [ObservableProperty] private string _packagedEstimatedTimeRemaining = string.Empty;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] [NotifyPropertyChangedFor(nameof(IsMetadataComplete))] [NotifyPropertyChangedFor(nameof(IsMetadataInProgress))] [NotifyPropertyChangedFor(nameof(IsReadyForReport))] [NotifyPropertyChangedFor(nameof(IsCancelable))] private bool _isUpdatingMetadata;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(IsMetadataComplete))] private bool _isMetadataCancelled;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(MetadataFolderCountText))] private int _metadataFolderDone;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(MetadataFolderCountText))] private int _metadataFolderTotal;
    [ObservableProperty] private string _metadataElapsedTime = string.Empty;
    [ObservableProperty] private string _metadataEta = string.Empty;
    [ObservableProperty] private int _preflightChecked;
    [ObservableProperty] private int _preflightTotal;
    [ObservableProperty] private int _packedCount;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(IsCancelable))] [NotifyCanExecuteChangedFor(nameof(ExportVerificationReportCommand))] private bool _isVerifying;

    // Opt-in "Deep verify Office files" (see Docs/DEEP-VERIFY-PLAN.md). Seeded from Settings so the
    // checkbox remembers the user's last choice, but ExportVerificationReportAsync reads THIS live
    // property (not Settings) when it kicks off a run — this is the only gate for the whole feature.
    [ObservableProperty] private bool _deepVerifyOfficeFiles;

    // Frozen at copy-start (not a live binding to the Options step) so it always reflects what
    // this run actually started with, even if the Options step's fields could otherwise still
    // change — needed so mid-run/after-the-fact you can tell what a run was configured with
    // without having to trust memory or dig through the Options screen.
    [ObservableProperty] private string _runSettingsSummary = string.Empty;

    internal string BuildRunSettingsSummary()
    {
        var overwrite = OverwriteMode switch
        {
            Models.OverwriteMode.IfNewer => Loc.T("VM_OverwriteIfNewer"),
            var m => m.ToString()
        };
        // Record the EFFECTIVE values, not the selected ones: Pages scope forces Enhanced REST
        // with versions off, and this summary's whole purpose is "what this run actually used".
        bool pagesOverride = IsPagesScope;
        var copyMode = CopyMode == Models.CopyMode.MigrationApi && !pagesOverride ? Loc.T("VM_ModeMigrationApi") : Loc.T("VM_ModeEnhancedRest");
        var versions = !CopyVersions || pagesOverride ? Loc.T("VM_Off") : CopyAllVersions ? Loc.T("VM_All") : Loc.T("VM_MaxVersions", MaxVersions);

        return string.Join("  ·  ",
        [
            Loc.T("VM_SumOverwrite", overwrite),
            Loc.T("VM_SumCopyMode", copyMode),
            Loc.T("VM_SumVersions", versions),
            Loc.T("VM_SumParallel", MaxParallelCopies),
            Loc.T("VM_SumPreserveMetadata", PreserveMetadata ? Loc.T("VM_Yes") : Loc.T("VM_No")),
            Loc.T("VM_SumFolderRepair", PreserveMetadata && ReapplyFolderMetadataEveryRun ? Loc.T("VM_EveryRun") : Loc.T("VM_NewFoldersOnly")),
            Loc.T("VM_SumPermissions", CopyPermissions ? Loc.T("VM_Yes") : Loc.T("VM_No")),
            Loc.T("VM_SumCustomColumns", EffectiveCopyCustomColumns ? Loc.T("VM_Yes") : Loc.T("VM_No"))
        ]);
    }

    private readonly List<string> _activityLines = [];
    private string _activityText = "";
    public string ActivityText
    {
        get => _activityText;
        private set { _activityText = value; OnPropertyChanged(); }
    }

    // A bulk read from the source (custom column values / permission flags) failed. The copy continues
    // — but items from that source may then copy without custom fields or unique permissions, so say
    // so in the activity log instead of swallowing it. A user cancel is not a failure.
    private void ReportBulkReadFailure(string siteUrl, Exception ex)
    {
        if (ex is OperationCanceledException) return;
        PushActivity(Loc.T("VM_BulkReadFailed", siteUrl, ex.Message));
    }

    private void PushActivity(string message)
    {
        var now  = DateTime.Now;
        var line = $"{now:HH:mm:ss}  {message}";
        _activityLines.Insert(0, line);
        // Keep a generous window in the UI (newest-first, scrollable); the full history is on disk.
        while (_activityLines.Count > 300)
            _activityLines.RemoveAt(_activityLines.Count - 1);
        ScheduleActivityTextRefresh();

        // Hand the line to the background writer for the on-disk log (never blocks the copy, never
        // throws). The file gets the full date — a time-only stamp is ambiguous on a multi-day run.
        _activityWriter?.Write($"{now:yyyy-MM-dd HH:mm:ss}  {message}");
    }

    // The activity text is one string rebuilt from up to 300 lines; doing that on EVERY message cost
    // real UI-thread time at thousands of messages a minute. Coalesce to about four updates a second.
    private System.Windows.Threading.DispatcherTimer? _activityRefreshTimer;
    private void ScheduleActivityTextRefresh()
    {
        if (_activityRefreshTimer == null)
        {
            _activityRefreshTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(250)
            };
            _activityRefreshTimer.Tick += (_, _) =>
            {
                _activityRefreshTimer!.Stop();
                ActivityText = string.Join("\n", _activityLines);
            };
        }
        if (!_activityRefreshTimer.IsEnabled) _activityRefreshTimer.Start();
    }

    // Full activity log written to disk for the current run (one file per copy). Surfaced in the UI
    // so the user can open it after a long migration.
    private string? _activityLogPath;
    private ActivityLogWriter? _activityWriter;
    public string? ActivityLogPath
    {
        get => _activityLogPath;
        private set { _activityLogPath = value; OnPropertyChanged(); }
    }

    // Creates a fresh timestamped activity-log file for a new copy run. Returns silently on failure.
    private void StartActivityLogFile()
    {
        try
        {
            // Finish the previous run's file first (flushes anything still queued).
            _activityWriter?.Dispose();
            _activityWriter = null;
            System.IO.Directory.CreateDirectory(DiagnosticsLog.LogDirectory);
            var path = System.IO.Path.Combine(DiagnosticsLog.LogDirectory, $"activity-{DateTime.Now:yyyyMMdd-HHmmss}.log");
            _activityWriter = DiagnosticsLog.OpenActivityLog(path);
            ActivityLogPath = path;
        }
        catch { ActivityLogPath = null; }
    }

    public bool IsMetadataInProgress  => IsCopyComplete && IsUpdatingMetadata;
    public bool IsMetadataComplete    => IsCopyComplete && !IsUpdatingMetadata && !IsMetadataCancelled;
    public bool IsReadyForReport      => IsCopyComplete && !IsUpdatingMetadata;
    public bool IsCancelable          => IsCopying || IsUpdatingMetadata || IsVerifying;
    public string MetadataFolderCountText => MetadataFolderTotal > 0
        ? Loc.T("VM_MetadataFolderCount", MetadataFolderDone, MetadataFolderTotal)
        : string.Empty;
    public bool IsPreflightInProgress => IsCopying && PreflightTotal > 0 && PreflightChecked < PreflightTotal;
    public bool IsPackingInProgress   => IsCopying && PackedCount > 0 && CompletedCount < TotalCount;

    partial void OnPackedCountChanged(int value) => OnPropertyChanged(nameof(IsPackingInProgress));

    // Hold off system sleep while any long-running phase is active (see SleepPreventionService).
    // Property-changed hooks cover every code path that toggles these flags — both copy commands,
    // the metadata post-pass, and verification — without each needing its own Begin/End calls.
    partial void OnIsCopyingChanged(bool value)          => UpdateSleepPrevention();
    partial void OnIsVerifyingChanged(bool value)        => UpdateSleepPrevention();
    partial void OnIsUpdatingMetadataChanged(bool value) => UpdateSleepPrevention();

    // Persisted immediately (unlike the Step 3 Options, which save on Next) since this checkbox
    // lives on the Verify controls, not a step with its own commit point.
    partial void OnDeepVerifyOfficeFilesChanged(bool value)
    {
        Settings.DeepVerifyOfficeFiles = value;
        Settings.Save();
    }
    private void UpdateSleepPrevention()
    {
        if (IsCopying || IsVerifying || IsUpdatingMetadata) Services.SleepPreventionService.Begin();
        else Services.SleepPreventionService.End();
        RefreshTaskbar();
    }

    partial void OnPreflightCheckedChanged(int value)
    {
        OnPropertyChanged(nameof(IsPreflightInProgress));
        // Anchor the ETA clock the moment pre-flight finishes so the estimate
        // reflects actual copy throughput rather than existence-check time.
        // _etaStartDone captures files already finished (e.g. pre-flight skips)
        // so they don't inflate the rate — only post-preflight completions count.
        if (PreflightTotal > 0 && value >= PreflightTotal && _etaStartTime == null)
        {
            _etaStartTime = DateTimeOffset.Now;
            lock (_copyResultsLock)
            {
                int totalNow = CopyResults.Count;
                _etaStartDone = ShouldUseBytesEta(totalNow)
                    ? _bytesDoneTally
                    : CopyResults.Count(r => r.Status is CopyStatus.Success or CopyStatus.Failed or CopyStatus.Skipped);
            }
        }
    }

    partial void OnPreflightTotalChanged(int value) => OnPropertyChanged(nameof(IsPreflightInProgress));

    // Bytes-weighted ETA whenever we know most of the file sizes (see UpdateProgress) — shared so
    // the preflight-completion anchor above and the per-tick calculation agree on which unit is in use.
    private bool ShouldUseBytesEta(int total) =>
        total > 0 && _bytesTotalTally > 0 && _bytesMissingCount <= total / 10;

    private bool            _hasFolderJobs;
    private DateTimeOffset  _copyStartTime;
    private DateTimeOffset? _copyEndTime;
    private DateTimeOffset? _etaStartTime;
    private long            _etaStartDone;
    private DateTimeOffset? _packStartTime;
    private long            _packStartDone;

    // Recent-throughput samples for ETA calculation — a fixed lookback window rather than a
    // full-run average, so a slow start (cold start, throttling, a large first file) doesn't
    // drag the estimate down for the rest of the job and then snap to reality all at once.
    private static readonly TimeSpan EtaWindow = TimeSpan.FromSeconds(20);
    private readonly Queue<(DateTimeOffset Time, long Done)> _etaSamples = new();
    private readonly Queue<(DateTimeOffset Time, long Done)> _packSamples = new();

    // Damps tick-to-tick jitter in the displayed ETA — e.g. one large file finishing produces a
    // momentary rate spike (windowProgress jumps), and the very next tick with no completions can
    // fall back to the anchor-average branch in ComputeRemaining, which uses a completely different
    // rate. Without smoothing, that read as the estimate leaping from "1 minute" to "48 minutes" and
    // back on a byte-weighted, mixed-file-size run. See SmoothSeconds.
    private double?         _etaSmoothedSeconds;
    private DateTimeOffset? _etaSmoothedTime;
    private double?         _packEtaSmoothedSeconds;
    private DateTimeOffset? _packEtaSmoothedTime;
    private DateTimeOffset  _metadataStartTime;
    private DateTimeOffset? _metadataRateAnchorTime;
    private int             _metadataRateAnchorDone;

    // These used to each run a full CopyResults.Count(predicate) scan — at 250k rows, six such scans
    // every 400ms UI tick (~1.75M predicate calls/tick). Now backed by tallies maintained incrementally
    // as each row's Status/PermissionStatus changes — see AttachRowTallies/OnRowStatusChanging.
    public int SuccessCount   => _successTally;
    public int FailedCount    => _failedTally;
    public int SkippedCount   => _skippedTally;
    // Never actually attempted — the run was cancelled or the app closed while these were still
    // Copying. Kept separate from FailedCount so an interrupted run's summary doesn't read as a
    // mass failure (see CopyStatus.Cancelled).
    public int CancelledCount => _cancelledTally;

    public int FileTotalCount     => _fileTotalTally;
    public int FileSuccessCount   => _fileSuccessTally;
    public int FileFailedCount    => _fileFailedTally;
    public int FileSkippedCount   => _fileSkippedTally;
    public int FileCancelledCount => _fileCancelledTally;

    // Files that copied successfully but had at least one custom column value that could not be
    // applied (mismatched/missing target column, content-type mismatch, etc.) — see
    // CopyResult.CustomFieldStatus. Deliberately not part of FileFailedCount.
    public int CustomFieldWarningCount => _customFieldWarningTally;

    // Total/copied-so-far size for display, reusing the same byte tallies ShouldUseBytesEta already
    // maintains for the ETA calculation. Blank whenever too many rows are missing a SourceSize (same
    // threshold as ShouldUseBytesEta) — showing a partial byte sum next to a complete item count
    // would understate "Total Size" rather than just omitting it.
    public bool HasKnownTotalSize => ShouldUseBytesEta(TotalCount);
    // Once the run is complete, _bytesFinalTally (every version actually copied, per row) is the more
    // accurate figure — _bytesTotalTally only ever reflects each file's current version, since that's
    // all that's knowable before/during the copy (see CopyResult.SourceSize). Switching only after
    // completion keeps the in-progress Step 4 display (DoneSizeDisplay vs. this) internally consistent
    // — a mid-copy denominator that could grow past what ETA pacing expects would make "done" exceed
    // "total" and break the remaining-time math.
    public string TotalSizeDisplay  => HasKnownTotalSize
        ? SizeFormatter.FormatBytes(IsCopyComplete ? _bytesFinalTally : _bytesTotalTally)
        : string.Empty;
    public string DoneSizeDisplay   => HasKnownTotalSize ? SizeFormatter.FormatBytes(_bytesDoneTally)  : string.Empty;
    // Pre-formatted so the Step 4 progress line's Run can bind Text alone — Run isn't a UIElement
    // and has no Visibility property to hide a separate "of X" segment when size is unknown.
    public string SizeProgressDisplay => HasKnownTotalSize ? $"  ·  {DoneSizeDisplay} of {TotalSizeDisplay}" : string.Empty;

    public double PermColumnWidth        => CopyPermissions ? 100 : 0;
    public double PermDetailsColumnWidth => CopyPermissions ? 200 : 0;

    public double CustomFieldColumnWidth        => EffectiveCopyCustomColumns ? 100 : 0;
    public double CustomFieldDetailsColumnWidth => EffectiveCopyCustomColumns ? 200 : 0;

    [RelayCommand]
    private async Task StartCopyAsync()
    {
        _hasFolderJobs       = CopyJobs.Any(j => j.IsFolder);
        IsCopying            = true;
        IsCopyComplete       = false;
        CopyDuration         = string.Empty;
        IsUpdatingMetadata   = _hasFolderJobs;
        IsMetadataCancelled  = false;
        // Set before ExecuteAsync starts, not after it returns: Migration API mode reports
        // onMetadataDone synchronously from inside ExecuteAsync (no separate post-processing step),
        // which raced ahead of the old post-return assignment and computed elapsed time against
        // the field's uninitialized default (0001-01-01) — showing a ~2025-year duration.
        _metadataStartTime   = DateTimeOffset.Now;
        CopyResults.Clear();
        ResultFilter     = ResultFilterKind.All;
        FileResultFilter = ResultFilterKind.All;
        CompletedCount = 0;
        TotalProgress  = 0;
        _copyStartTime = DateTimeOffset.Now;
        _copyEndTime   = null;

        foreach (var job in CopyJobs.Where(j => !j.IsFolder))
        {
            CopyResults.Add(new CopyResult
            {
                FileName   = job.SourceName,
                SourcePath = job.SourceDisplayPath,
                TargetPath = job.TargetDisplayPath,
                SourceSize = job.SourceSize
            });
        }

        _copyCts?.Dispose();
        _copyCts      = new CancellationTokenSource();
        _etaStartTime = null;
        _etaStartDone = 0;
        _packStartTime = null;
        _packStartDone = 0;
        _etaSmoothedSeconds = null;
        _etaSmoothedTime    = null;
        _packEtaSmoothedSeconds = null;
        _packEtaSmoothedTime    = null;
        _etaSamples.Clear();
        _packSamples.Clear();
        TotalCount    = CopyJobs.Count;

        PreflightChecked = 0;
        PreflightTotal   = 0;
        PackedCount      = 0;
        _packedBytesTally = 0;
        _activityLines.Clear();
        ActivityText = "";
        StartActivityLogFile();
        if (ActivityLogPath != null)
            PushActivity(Loc.T("VM_ActivityLogPath", ActivityLogPath));
        BeginRunJournal();
        RunSettingsSummary = BuildRunSettingsSummary();
        PushActivity(Loc.T("VM_SettingsSummary", RunSettingsSummary));
        // Once per run (CopyService.ExecuteAsync runs once per library in Library/Site scopes).
        if (IsMultiTenant) PushActivity(Loc.T("Svc_CrossTenantNotes"));
        System.Windows.Threading.DispatcherTimer? metadataTimer = null;
        var onMetadataDone = new Progress<bool>(completed =>
        {
            metadataTimer?.Stop();
            if (completed)
                MetadataElapsedTime = FormatDuration(DateTimeOffset.Now - _metadataStartTime);
            else
                IsMetadataCancelled = true;
            MetadataEta        = string.Empty;
            IsUpdatingMetadata = false;
        });
        var onFolderProgress = new Progress<(int done, int total)>(p =>
        {
            MetadataFolderDone  = p.done;
            MetadataFolderTotal = p.total;
            OnPropertyChanged(nameof(MetadataFolderCountText));
        });
        var onPreflight    = new Progress<(int done, int total)>(p =>
        {
            PreflightTotal   = p.total;
            PreflightChecked = p.done;
        });
        var onActivity   = new Progress<string>(PushActivity);
        var onFilePacked = new Progress<long>(size =>
        {
            PackedCount++;
            Interlocked.Add(ref _packedBytesTally, size);
            OnPropertyChanged(nameof(IsPackingInProgress));
        });

        var progressTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(400)
        };
        progressTimer.Tick += (_, _) => UpdateProgress();
        progressTimer.Start();

        try
        {
            int versionLimit = CopyVersions && !CopyAllVersions ? MaxVersions : 0;

            // Build bulk field cache for custom columns (single paginated pass)
            // Also build permission flags when CopyPermissions is on.
            _bulkFieldCache = [];
            _sourceMetaCache = [];
            var _permissionFlags = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            if ((EffectiveCopyCustomColumns || CopyPermissions) && CopyJobs.Count > 0)
            {
                StatusMessage = Loc.T("VM_ReadingSourceMetadata");
                // Build bulk field cache for each source library. Keys are "{listId}:{itemId}"
                // so caches from multiple libraries can be merged without collision.
                if (EffectiveCopyCustomColumns)
                {
                    foreach (var driveId in CopyJobs.Select(j => j.SourceDriveId)
                                                    .Where(d => !string.IsNullOrEmpty(d))
                                                    .Distinct())
                    {
                        try
                        {
                            var srvUrl = await SpService.GetLibraryServerRelativeUrlAsync(driveId);
                            var listId = await SpService.GetListIdByServerRelativeUrlAsync(SourceUrl.TrimEnd('/'), srvUrl);
                            var cols   = await SpService.GetLibraryColumnsAsync(SourceUrl.TrimEnd('/'), listId);
                            // System metadata (Modified/Created/Editor/Author) is only needed by the
                            // Migration API's post-import restamp — riding it along on this SAME scan
                            // (rather than a per-file Graph call later) only when it'll actually be used.
                            var wantsSystemMeta = CopyMode == CopyMode.MigrationApi && PreserveMetadata;
                            var systemMeta = wantsSystemMeta ? new Dictionary<string, FileMetadata>() : null;
                            var cache  = await SpService.BulkReadCustomFieldsAsync(
                                SourceUrl.TrimEnd('/'), listId, cols, warningLog: onActivity,
                                systemMetadataOut: systemMeta, ct: _copyCts.Token);
                            foreach (var (itemId, fields) in cache)
                                _bulkFieldCache[$"{listId}:{itemId}"] = fields;
                            if (systemMeta != null)
                                foreach (var (itemId, meta) in systemMeta)
                                    _sourceMetaCache[$"{listId}:{itemId}"] = meta;
                        }
                        catch (Exception ex)
                        {
                            // Surfaced rather than silently swallowed — a failed bulk read here used to
                            // leave the whole custom-field cache empty with no visible reason, so every
                            // file in the run silently copied with no custom column values at all.
                            PushActivity(Loc.T("VM_ReadColumnValuesFailed", ex.Message));
                        }
                    }
                }

                // Permission flags must cover every source library in the selection —
                // jobs can span multiple libraries, and the flags use composite
                // "{listId}:{itemId}" keys so merging is collision-free.
                if (CopyPermissions)
                {
                    foreach (var driveId in CopyJobs.Select(j => j.SourceDriveId)
                                                    .Where(d => !string.IsNullOrEmpty(d))
                                                    .Distinct())
                    {
                        try
                        {
                            var listId = await SpService.GetListIdFromDriveAsync(driveId);
                            if (listId == null)
                            {
                                PushActivity(Loc.T("VM_PermissionFlagsNoList"));
                                continue;
                            }
                            var flags = await SpService.BulkReadPermissionFlagsAsync(
                                SourceUrl.TrimEnd('/'), listId, _copyCts.Token);
                            foreach (var (k, v) in flags)
                                _permissionFlags[k] = v;
                        }
                        catch (Exception ex) { ReportBulkReadFailure(SourceUrl, ex); }
                    }
                }
                StatusMessage = string.Empty;
            }

            if (CopyPermissions)
            {
                try { await _permissionCopyService.InitializeAsync(TargetUrl.TrimEnd('/'), _copyCts.Token); }
                catch { /* non-fatal */ }
            }

            // Create columns the user marked "Create in target" in the mapping dialog
            // before any values are written. Created columns keep the source internal
            // name, which is what BuildTargetNameMap resolves create-new mappings to.
            var columnsToCreate = ColumnMappings
                .Where(m => m.CreateNew && m.Status == MappingStatus.WillCreate)
                .Select(m => m.SourceColumn)
                .ToList();
            if (EffectiveCopyCustomColumns && columnsToCreate.Count > 0 && SelectedTargetFolder != null)
            {
                try
                {
                    StatusMessage = Loc.T("VM_CreatingTargetColumns");
                    var libraryNode = SelectedTargetFolder;
                    while (libraryNode.Parent != null)
                        libraryNode = libraryNode.Parent;
                    var targetServerRelUrl = libraryNode.ServerRelativePath
                        ?? await SpService.GetLibraryServerRelativeUrlAsync(libraryNode.DriveId);
                    var targetListId = await SpService.GetListIdByServerRelativeUrlAsync(
                        TargetUrl.TrimEnd('/'), targetServerRelUrl);
                    var created = await _libraryCopyService.AddMissingColumnsAsync(
                        TargetUrl.TrimEnd('/'), targetListId, columnsToCreate);
                    AddColumnCreationResult(libraryNode.Name, $"{TargetUrl.TrimEnd('/')}/{libraryNode.Name}", created);
                }
                catch { /* non-fatal — affected column values fall back to per-field errors */ }
                finally { StatusMessage = string.Empty; }
            }

            // Pages scope must use Enhanced REST — SPMI cannot import .aspx files into the Site
            // Pages list (WebPageLibrary/BaseTemplate=119) because the manifest declares the list
            // as DocumentLibrary (BaseTemplate=101), causing a fatal template-mismatch error.
            // Pages also skip version copy — GetVersionsAsync on .aspx files fails in Graph and
            // triggers 3 Kiota retries (4 ODataError entries in the debugger per page).
            var effectiveCopyMode     = IsPagesScope ? CopyMode.EnhancedRest : CopyMode;
            var effectiveCopyVersions = IsPagesScope ? false : CopyVersions;
            var effectiveVersionLimit = IsPagesScope ? 0 : versionLimit;

            await _copyService.ExecuteAsync(
                CopyJobs,
                CopyResults,
                new CopyRunOptions
                {
                    OverwriteMode         = OverwriteMode,
                    CopyVersions          = effectiveCopyVersions,
                    MaxParallel           = MaxParallelCopies,
                    MaxVersions           = effectiveVersionLimit,
                    CopyMode              = effectiveCopyMode,
                    OnMetadataDone        = onMetadataDone,
                    CopyCustomColumns     = EffectiveCopyCustomColumns,
                    ColumnMappings        = [.. ColumnMappings],
                    BulkFieldCache        = _bulkFieldCache,
                    CopyPages             = IsPagesScope,
                    RemapPageWebPartUrls  = RemapPageWebPartUrls,
                    PreserveMetadata      = PreserveMetadata,
                    CopyPermissions       = CopyPermissions,
                    PermissionService     = CopyPermissions ? _permissionCopyService : null,
                    PermissionFlags       = _permissionFlags,
                    PreflightProgress     = onPreflight,
                    ActivityLog           = onActivity,
                    OnFilePacked          = onFilePacked,
                    OnFolderProgress      = onFolderProgress,
                    SourceMetaCache       = _sourceMetaCache,
                    ReapplyFolderMetadata = ReapplyFolderMetadataEveryRun,
                },
                _copyCts.Token);
        }
        catch (OperationCanceledException) { StatusMessage = Loc.T("VM_CopyCancelled"); }
        catch (Exception ex)              { StatusMessage = Loc.T("VM_CopyError", ex.Message); }
        finally
        {
            _copyEndTime   = DateTimeOffset.Now;
            progressTimer.Stop();
            IsCopying      = false;
            IsCopyComplete = true;
            OnPropertyChanged(nameof(IsPreflightInProgress));
            OnPropertyChanged(nameof(IsPackingInProgress));
            TotalCount     = CopyResults.Count;
            CopyDuration   = FormatDuration(_copyEndTime.Value - _copyStartTime);
            UpdateProgress();
            OnPropertyChanged(nameof(SuccessCount));
            OnPropertyChanged(nameof(FailedCount));
            OnPropertyChanged(nameof(SkippedCount));
            OnPropertyChanged(nameof(CancelledCount));
            OnPropertyChanged(nameof(FileTotalCount));
            OnPropertyChanged(nameof(FileSuccessCount));
            OnPropertyChanged(nameof(FileFailedCount));
            OnPropertyChanged(nameof(FileSkippedCount));
            OnPropertyChanged(nameof(FileCancelledCount));
            OnPropertyChanged(nameof(CustomFieldWarningCount));
            SaveReport();

            if (IsUpdatingMetadata)
            {
                MetadataFolderDone      = 0;
                MetadataFolderTotal     = 0;
                MetadataElapsedTime     = string.Empty;
                MetadataEta             = string.Empty;
                _metadataStartTime      = DateTimeOffset.Now;
                _metadataRateAnchorTime = null;
                _metadataRateAnchorDone = 0;
                metadataTimer = new System.Windows.Threading.DispatcherTimer
                    { Interval = TimeSpan.FromMilliseconds(500) };
                metadataTimer.Tick += (_, _) => UpdateMetadataProgress();
                metadataTimer.Start();
            }
        }
    }

    [RelayCommand]
    private void CancelCopy()
    {
        // One stray click used to throw away hours of copying. Ask first while a run is in flight.
        if (IsRunInProgress && !Confirm("Dlg_StopCopyTitle", "Dlg_StopCopyBody")) return;
        _copyCts?.Cancel();
    }

    public void CancelAllOperations()
    {
        _copyCts?.Cancel();
        _connectSourceCts?.Cancel();
        _connectTargetCts?.Cancel();
        _verifyCts?.Cancel();
    }

    // Copies items from a source list into an already-resolved target list.
    // Fetches live target columns (bypassing cache) and resolves source InternalName →
    // target InternalName by direct match first, then display-name fallback. Per-item
    // errors are collected and surfaced on result rather than aborting the whole batch.
    private async Task CopyListItemsAsync(
        string targetListId,
        LibraryDefinition def,
        HashSet<string> selectedItemIds,
        bool isPartialSelection,
        CopyResult listResult)
    {
        var targetCols       = await SpService.GetLibraryColumnsAsync(TargetUrl.TrimEnd('/'), targetListId, skipCache: true);
        // GroupBy/First: duplicate display names are legal in SharePoint and a plain
        // ToDictionary would throw, aborting the whole list copy.
        var targetByInternal = targetCols
            .GroupBy(c => c.InternalName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        var targetByDisplay  = targetCols
            .GroupBy(c => c.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        // Honor mapping-dialog decisions (skip / rename) the same way the file writers do.
        var mappingLookup = ColumnMapping.BuildTargetNameMap(ColumnMappings);

        // Fetch existing target items to support skip/overwrite by Title.
        // Items are matched by Title, but a title is not unique (and may be blank), so existing target
        // items sharing a title are kept as a queue in Id order and each SOURCE item with that title
        // takes the next one. The old First()-only lookup sent two "Q1" source items to the SAME target
        // item (the second overwrote the first), and never matched a blank title at all — so blank-title
        // items were created again on every run, even in Skip mode. Pairing is positional, so it is exact
        // when the same items are re-run and best-effort for a partial selection.
        var existingTargetItems = await SpService.GetListItemTitlesAsync(
            TargetUrl.TrimEnd('/'), targetListId, _copyCts!.Token);
        var existingByTitle = new Dictionary<string, Queue<(string Id, DateTimeOffset? Modified)>>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in existingTargetItems.OrderBy(t => long.TryParse(t.Id, out var n) ? n : long.MaxValue))
        {
            if (!existingByTitle.TryGetValue(t.Title, out var queue))
                existingByTitle[t.Title] = queue = new();
            queue.Enqueue((t.Id, t.Modified));
        }

        var allItems = await SpService.GetListItemsAsync(
            SourceUrl.TrimEnd('/'), def.SourceListId ?? string.Empty, def.Columns,
            ["HasUniqueRoleAssignments"], _copyCts!.Token);

        var items = (isPartialSelection
            ? allItems.Where(i => i.TryGetValue("Id", out var id) && selectedItemIds.Contains(id?.ToString() ?? string.Empty)).ToList()
            : allItems)
            // Id order, to line duplicate titles up with the target queues built above.
            .OrderBy(i => i.TryGetValue("Id", out var oid) && long.TryParse(oid?.ToString(), out var on) ? on : long.MaxValue)
            .ToList();

        var targetListTitle = IsCreatingNewList ? NewListName.Trim()
                            : SelectedTargetList?.Title ?? def.Title;

        foreach (var item in items)
        {
            _copyCts.Token.ThrowIfCancellationRequested();

            var itemTitle = item.TryGetValue("Title", out var titleVal) ? titleVal?.ToString() : null;
            var itemId    = item.TryGetValue("Id",    out var iidVal)   ? iidVal?.ToString()   : "?";
            var rowLabel  = itemTitle ?? Loc.T("VM_ItemLabel", itemId);

            var itemResult = new CopyResult
            {
                FileName   = rowLabel,
                SourcePath = $"{SourceUrl.TrimEnd('/')}/{def.Title}",
                TargetPath = $"{TargetUrl.TrimEnd('/')}/{targetListTitle}",
                Status     = CopyStatus.Copying,
            };
            System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(itemResult));

            var fields = new Dictionary<string, object?>();
            if (itemTitle != null) fields["Title"] = itemTitle;

            foreach (var col in def.Columns)
            {
                if (!item.TryGetValue(col.InternalName, out var v) || v == null) continue;

                string? targetKey = null;
                if (mappingLookup.TryGetValue(col.InternalName, out var mapped))
                {
                    if (mapped == null) continue; // explicitly skipped in the mapping dialog
                    if (targetByInternal.ContainsKey(mapped))
                        targetKey = mapped;
                }
                if (targetKey == null)
                {
                    if (targetByInternal.ContainsKey(col.InternalName))
                        targetKey = col.InternalName;
                    else if (targetByDisplay.TryGetValue(col.DisplayName, out var tc))
                        targetKey = tc.InternalName;
                }
                if (targetKey == null) continue;
                fields[targetKey] = v;
            }

            string? createdDate  = PreserveMetadata && item.TryGetValue("Created",  out var cd) ? cd?.ToString() : null;
            string? modifiedDate = PreserveMetadata && item.TryGetValue("Modified", out var md) ? md?.ToString() : null;

            // A blank/missing source title matches the target's blank-title items (stored under "").
            (string Id, DateTimeOffset? Modified)? existing =
                existingByTitle.TryGetValue(itemTitle ?? string.Empty, out var matchQueue) && matchQueue.Count > 0
                    ? matchQueue.Dequeue() : null;
            var hasUniquePerms = item.TryGetValue("HasUniqueRoleAssignments", out var hurv) && hurv is true;
            try
            {
                string? resolvedTargetItemId = existing?.Id;
                if (existing != null)
                {
                    if (OverwriteMode == Models.OverwriteMode.Skip)
                    {
                        itemResult.Status       = CopyStatus.Skipped;
                        itemResult.ErrorMessage = Loc.T("VM_AlreadyExists");
                        continue;
                    }
                    if (OverwriteMode == Models.OverwriteMode.IfNewer)
                    {
                        var srcModified = item.TryGetValue("Modified", out var sm) && sm is string sms &&
                                          DateTimeOffset.TryParse(sms, out var smv) ? smv : (DateTimeOffset?)null;
                        if (srcModified is { } sv && existing.Value.Modified is { } tv &&
                            TimestampComparer.IsUpToDate(sv, tv))
                        {
                            itemResult.Status       = CopyStatus.Skipped;
                            itemResult.ErrorMessage = "Up to date";
                            continue;
                        }
                    }
                    var updateFieldError = await SpService.UpdateListItemAsync(
                        TargetUrl.TrimEnd('/'), targetListId, existing.Value.Id,
                        fields, createdDate, modifiedDate,
                        _copyCts.Token);
                    itemResult.Status       = CopyStatus.Success;
                    itemResult.ErrorMessage = updateFieldError != null ? Loc.T("VM_UpdatedFieldWarning", updateFieldError) : Loc.T("VM_Updated");
                }
                else
                {
                    var (newItemId, createFieldError) = await SpService.CreateListItemAsync(
                        TargetUrl.TrimEnd('/'), targetListId,
                        fields, createdDate, modifiedDate,
                        _copyCts.Token);
                    resolvedTargetItemId    = newItemId;
                    itemResult.Status       = CopyStatus.Success;
                    itemResult.ErrorMessage = createFieldError != null ? Loc.T("VM_FieldWarning", createFieldError) : null;
                }

                if (CopyPermissions && resolvedTargetItemId != null && hasUniquePerms)
                {
                    try
                    {
                        var srcItemId = item.TryGetValue("Id", out var srcId) ? srcId?.ToString() : null;
                        if (srcItemId != null)
                        {
                            var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                                SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                                $"web/lists('{def.SourceListId}')/items({srcItemId})",
                                $"web/lists('{targetListId}')/items({resolvedTargetItemId})",
                                hasUniquePermissions: true,
                                rowLabel, _copyCts.Token);
                            if (perm.HasActivity)
                                AddPermissionResult(perm, itemResult.TargetPath);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception permEx)
                    {
                        // The item itself copied, so this stays non-fatal — but show it on the row.
                        itemResult.PermissionStatus  = CopyStatus.Failed;
                        itemResult.PermissionDetails = Loc.T("Svc_PermException", permEx.Message);
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception itemEx)
            {
                itemResult.Status       = CopyStatus.Failed;
                itemResult.ErrorMessage = itemEx.Message;
                listResult.Status       = CopyStatus.Failed;
            }
        }
    }

    // Adds summary log rows for column creation outcomes on a target list.
    private void AddColumnCreationResult(string listTitle, string targetPath,
        (List<string> Created, List<string> Failed) result)
    {
        if (result.Created.Count > 0)
        {
            var detail = result.Created.Count == 1
                ? Loc.T("VM_ColumnCreated", result.Created[0])
                : Loc.T("VM_ColumnsCreated", result.Created.Count, string.Join(", ", result.Created));
            System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(new CopyResult
            {
                FileName          = Loc.T("VM_ColumnsTitle", listTitle),
                SourcePath        = string.Empty,
                TargetPath        = targetPath,
                Status            = CopyStatus.Success,
                ErrorMessage      = detail,
                IsLibraryCreation = true,
            }));
        }
        foreach (var failure in result.Failed)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(new CopyResult
            {
                FileName          = Loc.T("VM_ColumnFailedTitle", listTitle),
                SourcePath        = string.Empty,
                TargetPath        = targetPath,
                Status            = CopyStatus.Failed,
                ErrorMessage      = failure,
                IsLibraryCreation = true,
            }));
        }
    }

    // For file-level results: stamps the perm outcome onto the existing file row.
    // For site/library-level results (no matching file row): adds a structural summary row.
    private void AddPermissionResult(PermissionCopyResult perm, string targetPath)
    {
        string detail;
        CopyStatus status;
        if (perm.Error != null)
        {
            detail = perm.Error;
            status = CopyStatus.Failed;
        }
        else
        {
            detail = perm.Applied == 1
                ? Loc.T("VM_RoleAssignmentApplied1")
                : Loc.T("VM_RoleAssignmentsApplied", perm.Applied);
            if (perm.SkippedPrincipals.Count > 0)
                detail += Loc.T("VM_PermSkippedUnresolvable", perm.SkippedPrincipals.Count, string.Join(", ", perm.SkippedPrincipals));
            status = CopyStatus.Success;
        }

        // Match by full target path first — name-only matching stamps the wrong row when the
        // same item name exists in more than one place.
        var row = CopyResults.FirstOrDefault(r =>
                !r.IsPermissionResult &&
                string.Equals(r.TargetPath, targetPath, StringComparison.OrdinalIgnoreCase))
            ?? CopyResults.FirstOrDefault(r =>
                !r.IsPermissionResult &&
                string.Equals(r.FileName, perm.ItemName, StringComparison.OrdinalIgnoreCase));

        if (row != null)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                row.PermissionStatus  = status;
                row.PermissionDetails = detail;
            });
        }
        else
        {
            // Site/library-level result — no file row to stamp; add as a structural entry.
            var r = new CopyResult
            {
                FileName          = Loc.T("VM_PermissionsTitle", perm.ItemName),
                SourcePath        = string.Empty,
                TargetPath        = targetPath,
                Status            = status,
                ErrorMessage      = detail,
                IsLibraryCreation = true,
            };
            System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(r));
        }
    }

    // Eagerly loads source AND target library columns so the mapping dialog has data immediately.
    // Each section uses its own try/catch so a target failure does not prevent source columns loading.
    internal string? ColumnLoadError { get; private set; }
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsColumnsReady))]
    [NotifyPropertyChangedFor(nameof(CanConfigureMappings))]
    [NotifyPropertyChangedFor(nameof(MappingButtonLabel))]
    private bool _isColumnsLoading;
    public bool IsColumnsReady => !IsColumnsLoading;

    private async Task LoadTargetColumnsAsync()
    {
        IsColumnsLoading = true;
        ColumnLoadError  = null;

        // Load target columns — Files/Pages scope uses the selected folder; individual item
        // copies use the chosen destination list directly.
        if (SelectedTargetFolder != null)
        {
            try
            {
                var libraryNode = SelectedTargetFolder;
                while (libraryNode.Parent != null)
                    libraryNode = libraryNode.Parent;
                var targetServerRelUrl = libraryNode.ServerRelativePath
                    ?? await SpService.GetLibraryServerRelativeUrlAsync(libraryNode.DriveId);
                var targetListId = await SpService.GetListIdByServerRelativeUrlAsync(
                    TargetUrl.TrimEnd('/'), targetServerRelUrl);
                _targetColumns = await SpService.GetLibraryColumnsAsync(TargetUrl.TrimEnd('/'), targetListId);
                OnPropertyChanged(nameof(TargetColumns));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoadColumns] target load failed: {ex.Message}");
                ColumnLoadError = Loc.T("VM_TargetColumnsUnavailable", ex.Message);
            }
        }
        else if (IsItemSelectionActive && SelectedTargetList != null)
        {
            try
            {
                _targetColumns = await SpService.GetLibraryColumnsAsync(TargetUrl.TrimEnd('/'), SelectedTargetList.Id);
                OnPropertyChanged(nameof(TargetColumns));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoadColumns] target load failed: {ex.Message}");
                ColumnLoadError = Loc.T("VM_TargetColumnsUnavailable", ex.Message);
            }
        }
        else if (IsLibraryScope && !IsItemSelectionActive)
        {
            // Library full-copy: try to load columns from the matching library at the target.
            // The library may not exist yet — failure here is non-fatal.
            var firstLib = SourceLibraries.FirstOrDefault(l => l.IsChecked == true || l.IsChecked == null)
                ?? SourceLibraries.FirstOrDefault();
            if (firstLib != null)
            {
                try
                {
                    var srcServerRelUrl = firstLib.ServerRelativePath
                        ?? await SpService.GetLibraryServerRelativeUrlAsync(firstLib.DriveId);
                    // Map source server-relative URL to target by replacing the source site path prefix.
                    var srcSitePath    = new Uri(SourceUrl.TrimEnd('/')).AbsolutePath.TrimEnd('/');
                    var tgtSitePath    = new Uri(TargetUrl.TrimEnd('/')).AbsolutePath.TrimEnd('/');
                    var libRelPart     = srcServerRelUrl.StartsWith(srcSitePath, StringComparison.OrdinalIgnoreCase)
                        ? srcServerRelUrl[srcSitePath.Length..]
                        : srcServerRelUrl;
                    var tgtServerRelUrl = tgtSitePath + libRelPart;
                    var targetListId   = await SpService.GetListIdByServerRelativeUrlAsync(
                        TargetUrl.TrimEnd('/'), tgtServerRelUrl);
                    if (!string.IsNullOrEmpty(targetListId))
                    {
                        _targetColumns = await SpService.GetLibraryColumnsAsync(TargetUrl.TrimEnd('/'), targetListId);
                        OnPropertyChanged(nameof(TargetColumns));
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[LoadColumns] target load failed (library full copy): {ex.Message}");
                }
            }
        }

        // Load source columns — separate try/catch so a target failure cannot block this
        try
        {
            // Prefer first library with checked nodes; fall back to first library
            var firstLib = SourceLibraries.FirstOrDefault(l => l.GetCheckedNodes().Any())
                ?? SourceLibraries.FirstOrDefault();
            if (firstLib != null)
            {
                string sourceListId;
                if (firstLib.IsCustomList)
                {
                    // Custom lists already have the list GUID as their Id — no drive lookup needed
                    sourceListId = firstLib.Id;
                }
                else
                {
                    var sourceServerRelUrl = firstLib.ServerRelativePath
                        ?? await SpService.GetLibraryServerRelativeUrlAsync(firstLib.DriveId);
                    sourceListId = await SpService.GetListIdByServerRelativeUrlAsync(
                        SourceUrl.TrimEnd('/'), sourceServerRelUrl);
                }
                _sourceColumns = await SpService.GetLibraryColumnsAsync(SourceUrl.TrimEnd('/'), sourceListId);
                OnPropertyChanged(nameof(SourceColumns));
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[LoadColumns] source load failed: {ex.Message}");
            ColumnLoadError ??= Loc.T("VM_SourceColumnsUnavailable", ex.Message);
        }
        finally
        {
            IsColumnsLoading = false;
        }
    }


    // Estimates remaining time from a recent-window rate, falling back to the full-run average
    // since `anchorTime` when the window hasn't seen any progress yet (e.g. one huge file has been
    // in flight for the whole window) — so the ETA stays visible instead of blanking out between
    // infrequent completions. `samples` holds (time, cumulative-done) points, trimmed to EtaWindow.
    //
    // Uploads ramp up gradually — the adaptive concurrency controller (see AdaptiveParallelismController)
    // deliberately starts with few parallel PUTs and widens over time, so throughput in the first
    // stretch of a run is well below steady state. A rate measured during that ramp-up and projected
    // across the whole job used to read "3 hours" for what was actually a 20-30 minute job. Below,
    // EtaMinConfidenceSeconds/Fraction gate the estimate off entirely (caller shows "Calculating…")
    // until enough of the job has actually been observed to be past the ramp-up, rather than trusting
    // whatever rate the first few seconds happen to produce.
    private const double EtaMinConfidenceSeconds = 15;
    private const double EtaMinConfidenceFraction = 0.02;

    // Returns raw remaining seconds (unsmoothed) — see SmoothSeconds for the damping layer callers
    // apply before formatting for display.
    private static double? ComputeRemaining(
        long doneUnits, long totalUnits, DateTimeOffset now,
        DateTimeOffset anchorTime, long anchorDoneUnits,
        Queue<(DateTimeOffset Time, long Done)> samples)
    {
        samples.Enqueue((now, doneUnits));
        while (samples.Count > 1 && now - samples.Peek().Time > EtaWindow)
            samples.Dequeue();

        var (windowStart, windowDone) = samples.Peek();
        var windowElapsed  = (now - windowStart).TotalSeconds;
        var windowProgress = doneUnits - windowDone;

        double elapsedSecs;
        long   progressUnits;
        if (windowProgress > 0 && windowElapsed >= 5)
        {
            elapsedSecs   = windowElapsed;
            progressUnits = windowProgress;
        }
        else
        {
            elapsedSecs   = (now - anchorTime).TotalSeconds;
            progressUnits = doneUnits - anchorDoneUnits;
        }

        if (progressUnits <= 0 || elapsedSecs < EtaMinConfidenceSeconds) return null;

        // Require a real slice of the job to be done — not just a handful of files that may still
        // be inside the concurrency ramp-up — before trusting the extrapolation.
        var sinceAnchor = doneUnits - anchorDoneUnits;
        if (totalUnits > 0 && sinceAnchor < Math.Max(1, (long)(totalUnits * EtaMinConfidenceFraction)))
            return null;

        var remainingSecs = (totalUnits - doneUnits) * elapsedSecs / progressUnits;
        return remainingSecs >= 2 ? remainingSecs : null;
    }

    // Exponential moving average over the raw remaining-seconds estimate, with a time constant tied
    // to EtaWindow so the smoothing keeps pace with (rather than lagging well behind, or barely
    // damping) the rate calculation it's smoothing. `smoothed`/`lastTime` are the caller's persisted
    // state for one ETA stream (main vs. packaging) — each stream smooths independently.
    // A null raw value (not enough confidence yet) resets the state so a fresh run doesn't inherit a
    // stale average from a previous phase or job.
    private static double? SmoothSeconds(double? raw, ref double? smoothed, ref DateTimeOffset? lastTime, DateTimeOffset now)
    {
        if (raw == null) { smoothed = null; lastTime = null; return null; }
        if (smoothed == null || lastTime == null)
        {
            smoothed = raw;
            lastTime = now;
            return smoothed;
        }
        var dt = (now - lastTime.Value).TotalSeconds;
        lastTime = now;
        if (dt <= 0) return smoothed;
        var alpha = 1 - Math.Exp(-dt / EtaWindow.TotalSeconds);
        smoothed += alpha * (raw.Value - smoothed.Value);
        return smoothed;
    }

    private void UpdateProgress()
    {
        RefreshTaskbar();
        // done/total used to be a full CopyResults.Count(predicate) scan every tick — see the
        // SuccessCount/etc tallies above, which this reuses (Cancelled deliberately excluded, same
        // as before: a cancelled-but-never-attempted row shouldn't count as "done").
        int done  = _successTally + _failedTally + _skippedTally;
        int total = CopyResults.Count;
        CompletedCount = done;
        TotalCount     = total;
        TotalProgress  = TotalCount > 0 ? done * 100.0 / TotalCount : 0;

        var elapsed = (_copyEndTime ?? DateTimeOffset.Now) - _copyStartTime;
        ElapsedTime = FormatDuration(elapsed);

        // Weight the ETA by bytes instead of item count whenever we know most of the sizes — a
        // handful of huge files left after a burst of small ones otherwise reads as "almost done"
        // right up until those huge files start, which is what made estimates swing wildly. Falls
        // back to plain item count for flows that don't carry SourceSize (library/site copy,
        // permission-only rows) or when too many sizes are unknown to trust the byte total.
        bool useBytes = ShouldUseBytesEta(total);
        long totalUnits = useBytes ? _bytesTotalTally : total;
        long doneUnits  = useBytes ? _bytesDoneTally  : done;
        long packedUnits = useBytes ? _packedBytesTally : PackedCount;

        var now = DateTimeOffset.Now;

        // Packaging ETA — rate-based estimate for the download+encrypt+upload phase.
        // Anchors when the first file is packed; waits 5 s for a stable rate before showing.
        // Computed once per tick (not per consumer) since it enqueues into _packSamples — calling
        // it twice in the same tick would double-enqueue the same sample.
        var packed = PackedCount;
        string? packRemaining = null;
        if (packed > 0 && packed < total && _copyEndTime == null)
        {
            if (_packStartTime == null)
            {
                _packStartTime = now;
                _packStartDone = packedUnits;
            }
            var rawPackRemaining = ComputeRemaining(packedUnits, totalUnits, now, _packStartTime.Value, _packStartDone, _packSamples);
            var smoothedPackRemaining = SmoothSeconds(rawPackRemaining, ref _packEtaSmoothedSeconds, ref _packEtaSmoothedTime, now);
            packRemaining = smoothedPackRemaining != null ? FormatDuration(TimeSpan.FromSeconds(smoothedPackRemaining.Value)) : null;
            PackagedEstimatedTimeRemaining = packRemaining != null ? Loc.T("VM_EtaRemaining", packRemaining) : Loc.T("VM_EtaCalculating");
        }
        else
        {
            PackagedEstimatedTimeRemaining = string.Empty;
        }

        // Main ETA: two phases —
        //   Packaging phase (done==0): show packaging ETA labelled "to package" so the user
        //   knows it covers only download+encrypt+upload, not the subsequent SP import.
        //   Completion phase (done>0): switch to actual completion rate for "remaining".
        if (done > 0 && done < total && _copyEndTime == null)
        {
            // Anchor the completion clock the first time files start finishing.
            if (_etaStartTime == null)
            {
                _etaStartTime = now;
                _etaStartDone = doneUnits;
            }
            var rawRemaining = ComputeRemaining(doneUnits, totalUnits, now, _etaStartTime.Value, _etaStartDone, _etaSamples);
            var smoothedRemaining = SmoothSeconds(rawRemaining, ref _etaSmoothedSeconds, ref _etaSmoothedTime, now);
            var remaining = smoothedRemaining != null ? FormatDuration(TimeSpan.FromSeconds(smoothedRemaining.Value)) : null;
            EstimatedTimeRemaining = remaining != null ? Loc.T("VM_EtaRemaining", remaining) : Loc.T("VM_EtaCalculating");
        }
        else if (done == 0 && packed > 0 && packed < total && _copyEndTime == null)
        {
            // Packaging phase — reuse the packaging ETA computed above with a clearer label.
            EstimatedTimeRemaining = packRemaining != null ? Loc.T("VM_EtaToPackage", packRemaining) : Loc.T("VM_EtaCalculating");
        }
        else
        {
            EstimatedTimeRemaining = string.Empty;
        }

        // Keep the filter-chip counts live while the copy runs.
        OnPropertyChanged(nameof(SuccessCount));
        OnPropertyChanged(nameof(FailedCount));
        OnPropertyChanged(nameof(SkippedCount));
        OnPropertyChanged(nameof(FileSuccessCount));
        OnPropertyChanged(nameof(FileFailedCount));
        OnPropertyChanged(nameof(FileSkippedCount));
        OnPropertyChanged(nameof(CustomFieldWarningCount));
        OnPropertyChanged(nameof(HasKnownTotalSize));
        OnPropertyChanged(nameof(TotalSizeDisplay));
        OnPropertyChanged(nameof(DoneSizeDisplay));
        OnPropertyChanged(nameof(SizeProgressDisplay));
    }

    private void UpdateMetadataProgress()
    {
        MetadataElapsedTime = FormatDuration(DateTimeOffset.Now - _metadataStartTime);

        int done  = MetadataFolderDone;
        int total = MetadataFolderTotal;
        if (done > 0 && total > 0 && done < total)
        {
            if (_metadataRateAnchorTime == null)
            {
                _metadataRateAnchorTime = DateTimeOffset.Now;
                _metadataRateAnchorDone = done;
            }
            var anchorElapsed = (DateTimeOffset.Now - _metadataRateAnchorTime.Value).TotalSeconds;
            var anchorDone    = done - _metadataRateAnchorDone;
            if (anchorDone > 0 && anchorElapsed >= 3)
            {
                var remainingSecs = (total - done) * anchorElapsed / anchorDone;
                MetadataEta = remainingSecs >= 2
                    ? Loc.T("VM_EtaRemaining", FormatDuration(TimeSpan.FromSeconds(remainingSecs)))
                    : string.Empty;
            }
        }
        else
        {
            MetadataEta = string.Empty;
        }
    }

    public void RefreshCopyStats()
    {
        OnPropertyChanged(nameof(SuccessCount));
        OnPropertyChanged(nameof(FailedCount));
        OnPropertyChanged(nameof(SkippedCount));
        OnPropertyChanged(nameof(FileTotalCount));
        OnPropertyChanged(nameof(FileSuccessCount));
        OnPropertyChanged(nameof(FileFailedCount));
        OnPropertyChanged(nameof(FileSkippedCount));
        OnPropertyChanged(nameof(CustomFieldWarningCount));
        OnPropertyChanged(nameof(HasKnownTotalSize));
        OnPropertyChanged(nameof(TotalSizeDisplay));
        OnPropertyChanged(nameof(DoneSizeDisplay));
        OnPropertyChanged(nameof(SizeProgressDisplay));
    }

    // Public wrapper so code-behind and dialogs can fire property change notifications.
    public new void OnPropertyChanged(string propertyName) => base.OnPropertyChanged(propertyName);

}
