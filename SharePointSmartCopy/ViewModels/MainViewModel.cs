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

public partial class MainViewModel : ObservableObject
{
    public readonly AuthService AuthService;
    // Same instance as AuthService in single-tenant mode; a separate sign-in in multi-tenant mode.
    public AuthService TargetAuthService { get; private set; }
    public readonly TenantRouter Router;
    public readonly SharePointService SpService;
    private readonly CopyService _copyService;
    private readonly LibraryCopyService _libraryCopyService;
    private readonly PermissionCopyService _permissionCopyService;
    private readonly VerificationReportService _verificationReportService;

    private CancellationTokenSource? _copyCts;
    private CancellationTokenSource? _connectSourceCts;
    private CancellationTokenSource? _connectTargetCts;
    private CancellationTokenSource? _verifyCts;

    // Bulk custom field cache keyed by SP list item integer ID (populated before copy starts)
    private Dictionary<string, Dictionary<string, object?>> _bulkFieldCache = [];
    // Source Modified/Created/Editor/Author, bulk-read in the SAME pass as _bulkFieldCache (same
    // "{listId}:{itemId}" keys) — lets the Migration API custom-fields restamp use the already-paid-for
    // scan instead of a per-file GetFileMetadataAsync Graph call.
    private Dictionary<string, FileMetadata> _sourceMetaCache = [];
    // Source and target library columns for the column mapping dialog
    private List<SpColumnDef> _sourceColumns = [];
    private List<SpColumnDef> _targetColumns = [];
    // Tracked so ConfigureMappings_Click can await completion before opening dialog
    internal Task? _columnLoadTask;

    public MainViewModel(AuthService? existingAuthService = null, AppSettings? settings = null, AuthService? existingTargetAuthService = null)
    {
        AuthService          = existingAuthService ?? new AuthService();
        TargetAuthService    = AuthService;
        Router               = new TenantRouter(AuthService);
        SpService            = new SharePointService(AuthService, Router);
        SpService.Throttled += OnThrottled;
        var migrationJobService = new MigrationJobService(SpService);
        _copyService                    = new CopyService(SpService, migrationJobService);
        _copyService.ParallelismChanged += OnParallelismChanged;
        _libraryCopyService       = new LibraryCopyService(SpService);
        _permissionCopyService    = new PermissionCopyService(SpService);
        _verificationReportService                    = new VerificationReportService(SpService);
        _verificationReportService.ParallelismChanged += OnParallelismChanged;
        Settings             = settings ?? AppSettings.Load();

        if (Settings.IsConfigured)
        {
            if (existingAuthService == null)
                AuthService.Configure(Settings);
            ConfigureTenancy(existingTargetAuthService);
            SpService.Initialize();
        }

        SourceUrl         = Settings.SourceUrl;
        TargetUrl         = Settings.TargetUrl;
        CopyMode          = Settings.PreferredCopyMode;
        OverwriteMode     = Settings.OverwriteMode ?? Models.OverwriteMode.Skip;
        CopyVersions      = Settings.CopyVersions;
        CopyAllVersions   = Settings.CopyAllVersions;
        MaxVersions       = Settings.MaxVersions;
        ListItemVersions  = Settings.ListItemVersions;
        MaxParallelCopies = Settings.MaxParallelCopies;
        PreserveMetadata    = Settings.PreserveMetadata;
        ReapplyFolderMetadataEveryRun = Settings.ReapplyFolderMetadataEveryRun;
        CopyNavigation      = Settings.CopyNavigation;
        CopyLibraryContent  = Settings.CopyLibraryContent;
        RemapPageWebPartUrls = Settings.RemapPageWebPartUrls;
        CopyPermissions      = Settings.CopyPermissions;
        CopyCustomColumns    = Settings.CopyCustomColumns;
        DeepVerifyOfficeFiles = Settings.DeepVerifyOfficeFiles;

        SourceLibraries.CollectionChanged += (_, _) =>
        {
            OnPropertyChanged(nameof(LibrarySummaryCount));
            OnPropertyChanged(nameof(LibraryPreviewItems));
            OnPropertyChanged(nameof(ShowListItemVersionOption));
            OnPropertyChanged(nameof(CopyPreviewTotalSizeDisplay));
        };

        // The field initializer `= []` on CopyResults sets the backing field directly and does not
        // go through the generated property setter, so OnCopyResultsChanged never fires for it —
        // subscribe here explicitly so the very first run's rows are tallied too.
        CopyResults.CollectionChanged += OnCopyResultsCollectionChanged;

        // Every input that can change whether Next is enabled already raises this; piggyback on it so
        // the "why is Next disabled" text stays in step without a second list of dependencies.
        NextCommand.CanExecuteChanged += (_, _) => OnPropertyChanged(nameof(NextBlockedReason));
    }

    // ── Settings ──────────────────────────────────────────────────────────────

    [ObservableProperty] private AppSettings _settings;

    public void ApplySettings(AppSettings s)
    {
        // Saving Settings used to reconfigure sign-in every time — so changing only the theme or
        // language signed the user out (and, in multi-tenant mode, disconnected both sides). Only the
        // app registrations (and which one each side uses) affect sign-in; everything else applies
        // without touching the connections.
        bool registrationsChanged = !SameRegistrations(Settings, s);
        Settings = s;
        if (registrationsChanged || !AuthService.IsConfigured)
        {
            AuthService.Configure(s);
            // Reconfiguring drops both sign-ins and the router's host registrations, so a multi-tenant
            // session must reconnect — otherwise both sides stay marked connected while requests are
            // routed to the wrong tenant.
            if (IsMultiTenant && (SourceConnected || TargetConnected))
            {
                DisconnectSource();
                DisconnectTarget();
                SignedInUser       = string.Empty;
                TargetSignedInUser = string.Empty;
                ConnectionNotice   = Loc.T("Notice_RegistrationChanged");
                CurrentStep        = 0;
            }
            ConfigureTenancy();
            SpService.Initialize();
        }
        RefreshRegistrationProperties();
    }

    // Single tenant: the target shares the source's sign-in (TargetAuthService == AuthService) and the
    // router stays empty, so nothing about the original behavior changes. Multi-tenant: the target gets
    // its own AuthService configured from the target registration; the router learns which host belongs
    // to which side when each side connects.
    private void ConfigureTenancy(AuthService? existingTarget = null)
    {
        Router.Clear();
        if (Settings.ConnectionMode == ConnectionMode.MultiTenant)
        {
            if (existingTarget != null)
            {
                TargetAuthService = existingTarget;
            }
            else
            {
                if (ReferenceEquals(TargetAuthService, AuthService)) TargetAuthService = new AuthService();
                if (Settings.TargetRegistration is { } reg) TargetAuthService.Configure(reg);
            }
        }
        else
        {
            TargetAuthService = AuthService;
        }
        ApplyIdentityMapping();
    }

    // Identity translation only applies to cross-tenant copies; single-tenant always gets the no-op mapper.
    public void ApplyIdentityMapping()
        => SpService.Identity = IsMultiTenant
            ? IdentityMapper.FromSettings(Settings.IdentityMapping)
            : IdentityMapper.None;

    // Bound to the Single tenant / Multi-tenant tabs on the Source step. Switching resets both
    // connections, because routing is built per mode and hosts must be re-registered on connect.
    public bool IsMultiTenant
    {
        get => Settings.ConnectionMode == ConnectionMode.MultiTenant;
        set
        {
            if (value == IsMultiTenant) return;
            // Switching silently disconnected both sides. Ask first when something is connected.
            if ((SourceConnected || TargetConnected) && !Confirm("Dlg_SwitchModeTitle", "Dlg_SwitchModeBody"))
            {
                // The radio button already flipped on screen — re-read the real value once this
                // binding pass has finished so it snaps back.
                System.Windows.Application.Current?.Dispatcher.BeginInvoke(() =>
                {
                    OnPropertyChanged(nameof(IsMultiTenant));
                    OnPropertyChanged(nameof(IsSingleTenant));
                });
                return;
            }
            Settings.ConnectionMode = value ? ConnectionMode.MultiTenant : ConnectionMode.SingleTenant;
            try { Settings.Save(); } catch { /* the mode still applies for this session */ }
            DisconnectSource();
            DisconnectTarget();
            TargetSignedInUser = string.Empty;
            ConfigureTenancy();
            OnPropertyChanged(nameof(IsMultiTenant));
            OnPropertyChanged(nameof(IsSingleTenant));
            OnPropertyChanged(nameof(HeaderAccountsText));
        }
    }

    public bool IsSingleTenant
    {
        get => !IsMultiTenant;
        set => IsMultiTenant = !value;
    }

    // Signed-in account on the target side (multi-tenant only; empty otherwise).
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(HeaderAccountsText))]
    private string _targetSignedInUser = string.Empty;

    // ── Step navigation ───────────────────────────────────────────────────────
    // Steps: 0=Source  1=Browse  2=Target  3=Options  4=Copying  5=Report

    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(BackCommand))]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    private int _currentStep;

    [ObservableProperty] private string _statusMessage = string.Empty;

    // Shows a throttle notice while a retry wait is in progress, then clears it
    // (unless another status message has replaced it in the meantime).
    private void OnThrottled(TimeSpan delay, int attempt, int maxAttempts, string? reason)
    {
        // maxAttempts == 0: notification-only source (GraphThrottleNotifyHandler observes while
        // Kiota retries internally) — no meaningful attempt counter to show.
        var msg = maxAttempts > 0
            ? Loc.T("VM_ThrottleRetryAttempt", Math.Ceiling(delay.TotalSeconds), attempt, maxAttempts)
            : Loc.T("VM_ThrottleRetry", Math.Ceiling(delay.TotalSeconds));
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher == null) return;
        dispatcher.BeginInvoke(() => StatusMessage = msg);
        _ = Task.Delay(delay + TimeSpan.FromSeconds(2)).ContinueWith(_ =>
            dispatcher.BeginInvoke(() =>
            {
                if (StatusMessage == msg) StatusMessage = string.Empty;
            }));
    }

    // Shows a brief notice when adaptive throttling steps parallelism down or back up.
    private void OnParallelismChanged(int newLimit)
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher == null) return;
        var msg = newLimit < MaxParallelCopies
            ? Loc.T("VM_ParallelismReduced", newLimit)
            : Loc.T("VM_ParallelismRestored", newLimit);
        dispatcher.BeginInvoke(() => StatusMessage = msg);
        _ = Task.Delay(TimeSpan.FromSeconds(5)).ContinueWith(_ =>
            dispatcher.BeginInvoke(() =>
            {
                if (StatusMessage == msg) StatusMessage = string.Empty;
            }));
    }
    [ObservableProperty] private bool _isBusy;

}
