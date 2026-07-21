using System.Collections.ObjectModel;
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
    // Source and target library columns for the column mapping dialog
    private List<SpColumnDef> _sourceColumns = [];
    private List<SpColumnDef> _targetColumns = [];
    // Tracked so ConfigureMappings_Click can await completion before opening dialog
    internal Task? _columnLoadTask;

    public MainViewModel(AuthService? existingAuthService = null, AppSettings? settings = null)
    {
        AuthService          = existingAuthService ?? new AuthService();
        SpService            = new SharePointService(AuthService);
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
            SpService.Initialize();
        }

        SourceUrl         = Settings.SourceUrl;
        TargetUrl         = Settings.TargetUrl;
        CopyMode          = Settings.PreferredCopyMode;
        OverwriteMode     = Settings.OverwriteMode ?? Models.OverwriteMode.Skip;
        CopyVersions      = Settings.CopyVersions;
        CopyAllVersions   = Settings.CopyAllVersions;
        MaxVersions       = Settings.MaxVersions;
        MaxParallelCopies = Settings.MaxParallelCopies;
        PreserveMetadata    = Settings.PreserveMetadata;
        ReapplyFolderMetadataEveryRun = Settings.ReapplyFolderMetadataEveryRun;
        CopyNavigation      = Settings.CopyNavigation;
        CopyLibraryContent  = Settings.CopyLibraryContent;
        RemapPageWebPartUrls = Settings.RemapPageWebPartUrls;
        CopyPermissions      = Settings.CopyPermissions;
        DeepVerifyOfficeFiles = Settings.DeepVerifyOfficeFiles;

        SourceLibraries.CollectionChanged += (_, _) =>
        {
            OnPropertyChanged(nameof(LibrarySummaryCount));
            OnPropertyChanged(nameof(LibraryPreviewItems));
        };
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

    // Shows a throttle notice while a retry wait is in progress, then clears it
    // (unless another status message has replaced it in the meantime).
    private void OnThrottled(TimeSpan delay, int attempt, int maxAttempts, string? reason)
    {
        // maxAttempts == 0: notification-only source (GraphThrottleNotifyHandler observes while
        // Kiota retries internally) — no meaningful attempt counter to show.
        var msg = maxAttempts > 0
            ? $"SharePoint is throttling requests — retrying in {Math.Ceiling(delay.TotalSeconds):N0}s (attempt {attempt}/{maxAttempts})…"
            : $"SharePoint is throttling requests — retrying in {Math.Ceiling(delay.TotalSeconds):N0}s…";
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
            ? $"Parallelism reduced to {newLimit} thread{(newLimit == 1 ? "" : "s")} (throttle protection)…"
            : $"Parallelism restored to {newLimit} thread{(newLimit == 1 ? "" : "s")}.";
        dispatcher.BeginInvoke(() => StatusMessage = msg);
        _ = Task.Delay(TimeSpan.FromSeconds(5)).ContinueWith(_ =>
            dispatcher.BeginInvoke(() =>
            {
                if (StatusMessage == msg) StatusMessage = string.Empty;
            }));
    }
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
            await AuthService.GetAccessTokenAsync(forceInteractive: !AuthService.IsAuthenticated, cancellationToken: ct);
            ct.ThrowIfCancellationRequested();
            SignedInUser = AuthService.UserName ?? string.Empty;
            SourceSiteId = await SpService.GetSiteIdAsync(SourceUrl.Trim());
            ct.ThrowIfCancellationRequested();
            SourceStatus    = $"✅ Connected as {SignedInUser}";
            SourceConnected = true;
            Settings.SourceUrl = SourceUrl.Trim();
            Settings.Save();
            CurrentStep = 1;
            _ = LoadLibrariesAsync();
        }
        catch (OperationCanceledException) { SourceStatus = string.Empty; }
        catch (Exception ex)              { SourceStatus = $"❌ {ex.Message}"; }
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
            var libs = await SpService.GetLibrariesAsync(SourceSiteId, SourceUrl.Trim());

            List<(string Id, string Title, int BaseTemplate)> customLists = [];
            if (CopyScope == CopyScope.Library)
                customLists = await SpService.GetCustomListsAsync(SourceUrl.TrimEnd('/'));

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                SourceLibraries.Clear();
                foreach (var lib in libs)
                {
                    // Library scope: selecting whole libraries — prevent confusing child expansion
                    if (CopyScope == CopyScope.Library)
                    {
                        lib.HasChildren = false;
                        lib.Children.Clear();
                    }
                    SourceLibraries.Add(lib);
                }
                foreach (var (id, title, baseTemplate) in customLists)
                {
                    var listNode = new SharePointNode
                    {
                        Id               = id,
                        Name             = title,
                        DriveId          = string.Empty,
                        SiteId           = SourceSiteId,
                        SiteUrl          = SourceUrl,
                        Type             = NodeType.Library,
                        HasChildren      = true,
                        IsCustomList     = true,
                        ListBaseTemplate = baseTemplate,
                    };
                    listNode.Children.Add(new SharePointNode { Name = "__placeholder__", Id = "ph" });
                    SourceLibraries.Add(listNode);
                }
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

    // Loads the Site Pages library and immediately expands it to show page files.
    public async Task LoadPageLibraryAsync()
    {
        IsBusy = true;
        StatusMessage = "Loading Site Pages…";
        string? errorMessage = null;
        try
        {
            var sitePagesNode = await SpService.GetSitePagesLibraryAsync(SourceSiteId, SourceUrl.Trim());
            if (sitePagesNode == null)
            {
                errorMessage = "No Site Pages library found on this site.";
                return;
            }

            // Pre-load the pages so the user sees them immediately without expanding
            var pages = await SpService.GetChildrenAsync(
                sitePagesNode.DriveId, sitePagesNode.Id,
                sitePagesNode.SiteId, sitePagesNode.SiteUrl);

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                SourceLibraries.Clear();
                sitePagesNode.Children.Clear();
                sitePagesNode.IsExpanded = true;
                foreach (var page in pages)
                {
                    page.Parent = sitePagesNode;
                    page.IsPage = true;
                    sitePagesNode.Children.Add(page);
                }
                sitePagesNode.HasChildren = sitePagesNode.Children.Count > 0;
                SourceLibraries.Add(sitePagesNode);
            });
        }
        catch (Exception ex)
        {
            errorMessage = $"Error loading Site Pages: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
            StatusMessage = errorMessage ?? string.Empty;
        }
    }

    public async Task LoadNodeChildrenAsync(SharePointNode node)
    {
        if (!node.HasChildren) return;
        if (!node.Children.Any(c => c.IsPlaceholder)) return;

        node.IsLoading = true;
        try
        {
            if (node.IsCustomList)
            {
                var items = await SpService.GetListItemTitlesAsync(SourceUrl.TrimEnd('/'), node.Id);
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    node.Children.Clear();
                    foreach (var (id, title, _) in items)
                        node.Children.Add(new SharePointNode
                        {
                            Name         = string.IsNullOrWhiteSpace(title) ? $"(Item {id})" : title,
                            Id           = id,
                            Type         = NodeType.ListItem,
                            SourceListId = node.Id,
                            SiteId       = node.SiteId,
                            SiteUrl      = node.SiteUrl,
                            HasChildren  = false,
                            IsChecked    = false,
                            Parent       = node,
                        });
                    if (!node.Children.Any())
                        node.HasChildren = false;
                });
                return;
            }

            var children = await SpService.GetChildrenAsync(node.DriveId, node.Id, node.SiteId, node.SiteUrl);
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
    [ObservableProperty] private ObservableCollection<ListPickerItem> _targetCustomLists = [];
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    [NotifyPropertyChangedFor(nameof(IsCreatingNewList))]
    [NotifyPropertyChangedFor(nameof(EffectiveDestinationListName))]
    private ListPickerItem? _selectedTargetList;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    [NotifyPropertyChangedFor(nameof(EffectiveDestinationListName))]
    private string _newListName = string.Empty;

    private const string NewListSentinelId = "__new__";
    public bool IsCreatingNewList => SelectedTargetList?.Id == NewListSentinelId;
    public string EffectiveDestinationListName =>
        IsCreatingNewList ? NewListName : (SelectedTargetList?.Title ?? string.Empty);
    [ObservableProperty] private bool _isConnectingTarget;

    // Editing the target URL invalidates any existing connection state. Without this, retyping the
    // URL — or clicking "Connect to Target" again for a DIFFERENT site without an explicit Disconnect
    // first — left SelectedTargetFolder/TargetLibraries pointing at the PREVIOUSLY connected site
    // while TargetUrl's text (and hence the Review screen's "To:" banner) showed the new,
    // correct-looking destination. CanGoNext's Step 2 gate only checks
    // `TargetConnected && SelectedTargetFolder != null` — both already true from the stale
    // connection — so nothing forced a re-pick before Next. Observed 2026-07-20: a copy intended for
    // one site instead copied into an unrelated, previously-connected site because of exactly this gap.
    partial void OnTargetUrlChanged(string value) => DisconnectTarget();

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
            var libs = await SpService.GetLibrariesAsync(TargetSiteId, TargetUrl.Trim());
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

            // Pages scope: pages can only go into a Site Pages library (the SitePages API is
            // exclusive to BaseTemplate=119). Replace all target libraries with only Site Pages
            // and auto-select it — there is exactly one valid destination.
            if (IsPagesScope)
            {
                try
                {
                    var sitePagesTarget = await SpService.GetSitePagesLibraryAsync(TargetSiteId, TargetUrl.Trim());
                    if (sitePagesTarget != null)
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            TargetLibraries.Clear();
                            TargetLibraries.Add(sitePagesTarget);
                            SelectedTargetFolder = sitePagesTarget;
                        });
                }
                catch { /* non-critical */ }
            }

            // Library scope with individual item selection: load the target site's custom lists
            // so the user can pick a destination list in the Target step.
            if (CopyScope == CopyScope.Library)
            {
                try
                {
                    var lists = await SpService.GetCustomListsAsync(TargetUrl.TrimEnd('/'));
                    ct.ThrowIfCancellationRequested();
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        TargetCustomLists.Clear();
                        TargetCustomLists.Add(new ListPickerItem(NewListSentinelId, "[ Create New List ]"));
                        foreach (var (id, title, _) in lists)
                            TargetCustomLists.Add(new ListPickerItem(id, title));
                    });
                }
                catch { /* non-critical — list picker just stays empty */ }
            }

            try { await AuthService.GetSharePointTokenAsync(TargetUrl.Trim(), cancellationToken: ct); }
            catch
            {
                TargetStatus = "✅ Connected · Note: additional consent needed for metadata — reconnect to grant";
            }
        }
        catch (OperationCanceledException) { TargetStatus = string.Empty; }
        catch (Exception ex)              { TargetStatus = $"❌ {ex.Message}"; }
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
        SelectedTargetList   = null;
        NewListName          = string.Empty;
        TargetLibraries.Clear();
        TargetCustomLists.Clear();
    }

    public async Task LoadTargetNodeChildrenAsync(SharePointNode node)
    {
        if (!node.HasChildren) return;
        if (!node.Children.Any(c => c.IsPlaceholder)) return;

        node.IsLoading = true;
        try
        {
            var children = await SpService.GetChildrenAsync(node.DriveId, node.Id, node.SiteId, node.SiteUrl, foldersOnly: true);
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

    // ── Copy Scope ────────────────────────────────────────────────────────────

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsFilesScope))]
    [NotifyPropertyChangedFor(nameof(IsLibraryScope))]
    [NotifyPropertyChangedFor(nameof(IsSiteScope))]
    [NotifyPropertyChangedFor(nameof(IsPagesScope))]
    [NotifyPropertyChangedFor(nameof(IsLibraryOrSiteScope))]
    [NotifyPropertyChangedFor(nameof(IsFilesOrPagesScope))]
    [NotifyPropertyChangedFor(nameof(NeedsTargetFolder))]
    [NotifyPropertyChangedFor(nameof(IsItemSelectionActive))]
    [NotifyPropertyChangedFor(nameof(ShowLibrarySummaryLine))]
    [NotifyPropertyChangedFor(nameof(ShowLibraryFullCopyMapping))]
    [NotifyPropertyChangedFor(nameof(LibrarySummaryCount))]
    [NotifyPropertyChangedFor(nameof(LibraryPreviewItems))]
    [NotifyPropertyChangedFor(nameof(EffectiveCopyCustomColumns))]
    [NotifyPropertyChangedFor(nameof(ShowFileCopyOptions))]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    private CopyScope _copyScope = CopyScope.Files;

    public bool IsFilesScope         => CopyScope == CopyScope.Files;
    public bool IsLibraryScope       => CopyScope == CopyScope.Library;
    public bool IsSiteScope          => CopyScope == CopyScope.Site;
    public bool IsPagesScope         => CopyScope == CopyScope.Pages;
    public bool IsLibraryOrSiteScope => CopyScope is CopyScope.Library or CopyScope.Site;
    public bool IsFilesOrPagesScope  => CopyScope is CopyScope.Files or CopyScope.Pages;
    public bool NeedsTargetFolder      => CopyScope is CopyScope.Files or CopyScope.Pages;
    // True when item-level selection is active: either individual items are checked (partial),
    // or a list is in items-only mode (IsChecked == null).
    public bool IsItemSelectionActive  =>
        CopyScope == CopyScope.Library && (
            SourceLibraries.Any(lib => lib.IsCustomList && lib.IsChecked == null) ||
            SourceLibraries.Any(lib => lib.Children.Any(c => c.Type == NodeType.ListItem && c.IsChecked == true)));

    // True when the Site scope summary text should be shown (whole-site copy, not individual items).
    public bool ShowLibrarySummaryLine => IsSiteScope && !IsItemSelectionActive;
    // True when the Library/List full-copy mapping button should be shown.
    public bool ShowLibraryFullCopyMapping => IsLibraryScope && !IsItemSelectionActive;

    // For Library scope: count of checked library nodes. For Site scope: count of all loaded libraries.
    // Count of whole libraries/lists selected (not individual items).
    public int LibrarySummaryCount => IsSiteScope
        ? SourceLibraries.Count
        : SourceLibraries.SelectMany(l => l.GetCheckedNodes()).Count(n => n.Type == NodeType.Library);

    // Count of individually selected list items (partial or items-only selection).
    public int SelectedItemCount => SourceLibraries
        .SelectMany(l => l.Children)
        .Count(c => c.Type == NodeType.ListItem && c.IsChecked == true);

    // Libraries shown in the Options step preview for Library/Site scope.
    // The null (items-only) state only has meaning for custom lists — a stray null on
    // any other node type must not surface a phantom entry here.
    public IEnumerable<SharePointNode> LibraryPreviewItems => IsSiteScope
        ? SourceLibraries
        : SourceLibraries.Where(n => n.IsChecked == true
              || (n.IsCustomList && n.IsChecked == null)
              || (n.IsCustomList && n.Children.Any(c => c.Type == NodeType.ListItem && c.IsChecked == true)));

    // Called from code-behind when any source node check state changes.
    public void NotifySelectionChanged()
    {
        OnPropertyChanged(nameof(IsItemSelectionActive));
        OnPropertyChanged(nameof(ShowLibrarySummaryLine));
        OnPropertyChanged(nameof(ShowLibraryFullCopyMapping));
        OnPropertyChanged(nameof(LibrarySummaryCount));
        OnPropertyChanged(nameof(SelectedItemCount));
        OnPropertyChanged(nameof(LibraryPreviewItems));
        NextCommand.NotifyCanExecuteChanged();
    }

    // ── Step 3: Options ───────────────────────────────────────────────────────

    [ObservableProperty] private OverwriteMode _overwriteMode = OverwriteMode.Skip;
    [ObservableProperty] private bool _copyVersions = true;
    [ObservableProperty] private bool _copyAllVersions = true;
    [ObservableProperty] private int _maxVersions = 10;
    [ObservableProperty] private int _maxParallelCopies = 4;
    [ObservableProperty] private CopyMode _copyMode = CopyMode.MigrationApi;
    [ObservableProperty] private ObservableCollection<CopyJob> _copyJobs = [];

    // New options
    [ObservableProperty] private bool _preserveMetadata = true;
    // When false, folder date/author repair is limited to folders receiving new files this run (faster
    // repeated incremental runs). Only meaningful while PreserveMetadata is on. See AppSettings.
    [ObservableProperty] private bool _reapplyFolderMetadataEveryRun = true;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(EffectiveCopyCustomColumns))]
    [NotifyPropertyChangedFor(nameof(CanConfigureMappings))]
    private bool _copyCustomColumns = false;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowFileCopyOptions))]
    private bool _copyLibraryContent = true;
    [ObservableProperty] private bool _remapPageWebPartUrls = true;
    [ObservableProperty] private bool _copyNavigation = true;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PermColumnWidth))]
    [NotifyPropertyChangedFor(nameof(PermDetailsColumnWidth))]
    private bool _copyPermissions = false;
    public ObservableCollection<ColumnMapping> ColumnMappings { get; } = [];

    // Library, Site, and Pages scopes always copy all custom column values; Files scope uses the checkbox.
    public bool EffectiveCopyCustomColumns => IsLibraryOrSiteScope || IsPagesScope || CopyCustomColumns;

    // File-level options (overwrite, versions, permissions) apply in Files scope, and in
    // Library/Site scope only when library content is being copied.
    public bool ShowFileCopyOptions => IsFilesScope || (IsLibraryOrSiteScope && CopyLibraryContent);

    // Mapping configuration requires columns loaded and (in Files scope) the opt-in checkbox.
    public bool CanConfigureMappings => IsColumnsReady && CopyCustomColumns;
    public IReadOnlyList<SpColumnDef> SourceColumns => _sourceColumns;
    public IReadOnlyList<SpColumnDef> TargetColumns => _targetColumns;

    public string MappingButtonLabel
    {
        get
        {
            var total = _sourceColumns.Count;
            if (ColumnMappings.Count == 0)
                return total > 0 ? $"Configure mappings ({total} columns)" : "Configure mappings";

            var unmatched = ColumnMappings.Count(m => m.Status == MappingStatus.Unmatched);
            var mapped    = ColumnMappings.Count - unmatched;
            return unmatched > 0
                ? $"Configure mappings  ⚠ {unmatched} of {total} unmatched"
                : $"Configure mappings ({mapped} of {total} mapped)";
        }
    }

    public void BuildCopyJobs()
    {
        CopyJobs.Clear();
        if (SelectedTargetFolder == null) return;

        // Find the library ancestor to get its server-relative URL for Migration API mode
        var libraryNode = SelectedTargetFolder;
        while (libraryNode.Parent != null)
            libraryNode = libraryNode.Parent;

        // Compute the subfolder path relative to the library root by walking the parent chain.
        // ServerRelativePath is only populated on library root nodes, so we use node names instead.
        var subFolderRelPath = BuildRelativePath(SelectedTargetFolder, libraryNode);

        var sourceSiteUrl = SourceUrl.TrimEnd('/');
        var targetSiteUrl = SelectedTargetFolder.SiteUrl.TrimEnd('/');

        foreach (var lib in SourceLibraries)
        {
            foreach (var node in lib.GetCheckedNodes())
            {
                var isLibrary = node.Type == NodeType.Library;

                // When the source is a whole library, copy its contents directly into the
                // target folder — no wrapper folder with the library's name.
                var targetDisplayPath = isLibrary
                    ? (string.IsNullOrEmpty(subFolderRelPath)
                        ? $"{targetSiteUrl}/{libraryNode.Name}"
                        : $"{targetSiteUrl}/{libraryNode.Name}/{subFolderRelPath}")
                    : (string.IsNullOrEmpty(subFolderRelPath)
                        ? $"{targetSiteUrl}/{libraryNode.Name}/{node.Name}"
                        : $"{targetSiteUrl}/{libraryNode.Name}/{subFolderRelPath}/{node.Name}");

                var job = new CopyJob
                {
                    SourceDriveId                  = node.DriveId,
                    SourceItemId                   = node.Id,
                    SourceName                     = node.Name,
                    SourceSiteUrl                  = sourceSiteUrl,
                    SourceDisplayPath              = $"{sourceSiteUrl}/{BuildPath(node)}",
                    TargetDriveId                  = SelectedTargetFolder.DriveId,
                    TargetParentItemId             = libraryNode.Id,
                    TargetSubFolderPath            = subFolderRelPath,
                    TargetSiteId                   = SelectedTargetFolder.SiteId,
                    TargetSiteUrl                  = SelectedTargetFolder.SiteUrl,
                    TargetDisplayPath              = targetDisplayPath,
                    TargetLibraryServerRelativeUrl = libraryNode.ServerRelativePath ?? string.Empty,
                    IsFolder                       = node.Type != NodeType.File,
                    IsLibrary                      = isLibrary,
                    IsPage                         = node.IsPage,
                };
                CopyJobs.Add(job);
            }
        }
    }

    // ── Step 4: Copying ───────────────────────────────────────────────────────

    private readonly object _copyResultsLock = new();
    [ObservableProperty] private ObservableCollection<CopyResult> _copyResults = [];

    partial void OnCopyResultsChanged(ObservableCollection<CopyResult> value)
    {
        BindingOperations.EnableCollectionSynchronization(value, _copyResultsLock);
        _copyResultsView = null;
        _fileResultsView = null;
        OnPropertyChanged(nameof(CopyResultsView));
        OnPropertyChanged(nameof(FileResultsView));
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
            Models.OverwriteMode.IfNewer => "If Newer",
            var m => m.ToString()
        };
        // Record the EFFECTIVE values, not the selected ones: Pages scope forces Enhanced REST
        // with versions off, and this summary's whole purpose is "what this run actually used".
        bool pagesOverride = IsPagesScope;
        var copyMode = CopyMode == Models.CopyMode.MigrationApi && !pagesOverride ? "Migration API" : "Enhanced REST";
        var versions = !CopyVersions || pagesOverride ? "Off" : CopyAllVersions ? "All" : $"Max {MaxVersions}";

        return string.Join("  ·  ",
        [
            $"Overwrite: {overwrite}",
            $"Copy Mode: {copyMode}",
            $"Versions: {versions}",
            $"Parallel Copies: {MaxParallelCopies}",
            $"Preserve Metadata: {(PreserveMetadata ? "Yes" : "No")}",
            $"Folder Metadata Repair: {(PreserveMetadata && ReapplyFolderMetadataEveryRun ? "Every run" : "New folders only")}",
            $"Permissions: {(CopyPermissions ? "Yes" : "No")}",
            $"Custom Columns: {(EffectiveCopyCustomColumns ? "Yes" : "No")}"
        ]);
    }

    private readonly List<string> _activityLines = [];
    private string _activityText = "";
    public string ActivityText
    {
        get => _activityText;
        private set { _activityText = value; OnPropertyChanged(); }
    }

    private void PushActivity(string message)
    {
        var ts   = DateTime.Now.ToString("HH:mm:ss");
        var line = $"{ts}  {message}";
        _activityLines.Insert(0, line);
        // Keep a generous window in the UI (newest-first, scrollable); the full history is on disk.
        while (_activityLines.Count > 300)
            _activityLines.RemoveAt(_activityLines.Count - 1);
        ActivityText = string.Join("\n", _activityLines);

        // Append to the on-disk log so long runs can be inspected afterwards. Best-effort: never let
        // a logging error interrupt a copy.
        var path = _activityLogPath;
        if (path != null)
        {
            try { lock (_activityLogFileLock) System.IO.File.AppendAllText(path, line + Environment.NewLine); }
            catch { /* ignore disk/IO errors */ }
        }
    }

    // Full activity log written to disk for the current run (one file per copy). Surfaced in the UI
    // so the user can open it after a long migration.
    private string? _activityLogPath;
    private readonly object _activityLogFileLock = new();
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
            var dir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SharePointSmartCopy", "Logs");
            System.IO.Directory.CreateDirectory(dir);
            ActivityLogPath = System.IO.Path.Combine(dir, $"activity-{DateTime.Now:yyyyMMdd-HHmmss}.log");
        }
        catch { ActivityLogPath = null; }
    }

    public bool IsMetadataInProgress  => IsCopyComplete && IsUpdatingMetadata;
    public bool IsMetadataComplete    => IsCopyComplete && !IsUpdatingMetadata && !IsMetadataCancelled;
    public bool IsReadyForReport      => IsCopyComplete && !IsUpdatingMetadata;
    public bool IsCancelable          => IsCopying || IsUpdatingMetadata || IsVerifying;
    public string MetadataFolderCountText => MetadataFolderTotal > 0
        ? $" {MetadataFolderDone:N0} / {MetadataFolderTotal:N0} folders"
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
                _etaStartDone = CopyResults.Count(r => r.Status is CopyStatus.Success or CopyStatus.Failed or CopyStatus.Skipped);
        }
    }

    partial void OnPreflightTotalChanged(int value) => OnPropertyChanged(nameof(IsPreflightInProgress));

    private bool            _hasFolderJobs;
    private DateTimeOffset  _copyStartTime;
    private DateTimeOffset? _copyEndTime;
    private DateTimeOffset? _etaStartTime;
    private int             _etaStartDone;
    private DateTimeOffset? _packStartTime;
    private int             _packStartDone;
    private DateTimeOffset  _metadataStartTime;
    private DateTimeOffset? _metadataRateAnchorTime;
    private int             _metadataRateAnchorDone;

    public int SuccessCount   => CopyResults.Count(r => r.Status == CopyStatus.Success);
    public int FailedCount    => CopyResults.Count(r => r.Status == CopyStatus.Failed);
    public int SkippedCount   => CopyResults.Count(r => r.Status == CopyStatus.Skipped);
    // Never actually attempted — the run was cancelled or the app closed while these were still
    // Copying. Kept separate from FailedCount so an interrupted run's summary doesn't read as a
    // mass failure (see CopyStatus.Cancelled).
    public int CancelledCount => CopyResults.Count(r => r.Status == CopyStatus.Cancelled);

    public int FileTotalCount   => CopyResults.Count(r => !r.IsPermissionResult);
    public int FileSuccessCount => CopyResults.Count(r => !r.IsPermissionResult && r.Status == CopyStatus.Success);
    public int FileFailedCount  => CopyResults.Count(r => !r.IsPermissionResult && (r.Status == CopyStatus.Failed || r.PermissionStatus == CopyStatus.Failed));
    public int FileSkippedCount => CopyResults.Count(r => !r.IsPermissionResult && r.Status == CopyStatus.Skipped);
    public int FileCancelledCount => CopyResults.Count(r => !r.IsPermissionResult && r.Status == CopyStatus.Cancelled);

    public double PermColumnWidth        => CopyPermissions ? 100 : 0;
    public double PermDetailsColumnWidth => CopyPermissions ? 200 : 0;

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
                TargetPath = job.TargetDisplayPath
            });
        }

        _copyCts?.Dispose();
        _copyCts      = new CancellationTokenSource();
        _etaStartTime = null;
        _etaStartDone = 0;
        _packStartTime = null;
        _packStartDone = 0;
        TotalCount    = CopyJobs.Count;

        PreflightChecked = 0;
        PreflightTotal   = 0;
        PackedCount      = 0;
        _activityLines.Clear();
        ActivityText = "";
        StartActivityLogFile();
        if (ActivityLogPath != null)
            PushActivity($"Activity log: {ActivityLogPath}");
        RunSettingsSummary = BuildRunSettingsSummary();
        PushActivity($"Settings: {RunSettingsSummary}");
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
        var onFilePacked = new Progress<int>(_ =>
        {
            PackedCount++;
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
            var _permissionFlags = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            if ((EffectiveCopyCustomColumns || CopyPermissions) && CopyJobs.Count > 0)
            {
                StatusMessage = "Reading source metadata…";
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
                            var cache  = await SpService.BulkReadCustomFieldsAsync(SourceUrl.TrimEnd('/'), listId, cols, ct: _copyCts.Token);
                            foreach (var (itemId, fields) in cache)
                                _bulkFieldCache[$"{listId}:{itemId}"] = fields;
                        }
                        catch { /* non-critical — that library's custom columns are skipped */ }
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
                            if (listId == null) continue;
                            var flags = await SpService.BulkReadPermissionFlagsAsync(
                                SourceUrl.TrimEnd('/'), listId, _copyCts.Token);
                            foreach (var (k, v) in flags)
                                _permissionFlags[k] = v;
                        }
                        catch { /* non-critical — that library's permissions are skipped */ }
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
                    StatusMessage = "Creating target columns…";
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
                OverwriteMode,
                effectiveCopyVersions,
                MaxParallelCopies,
                effectiveVersionLimit,
                effectiveCopyMode,
                _copyCts.Token,
                onMetadataDone,
                EffectiveCopyCustomColumns,
                [.. ColumnMappings],
                _bulkFieldCache,
                IsPagesScope,
                RemapPageWebPartUrls,
                PreserveMetadata,
                copyPermissions: CopyPermissions,
                permissionService: CopyPermissions ? _permissionCopyService : null,
                permissionFlags: _permissionFlags,
                preflightProgress: onPreflight,
                activityLog: onActivity,
                onFilePacked: onFilePacked,
                onFolderProgress: onFolderProgress);
        }
        catch (OperationCanceledException) { StatusMessage = "Copy cancelled."; }
        catch (Exception ex)              { StatusMessage = $"Copy error: {ex.Message}"; }
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
    private void CancelCopy() => _copyCts?.Cancel();

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
        var existingTargetItems = await SpService.GetListItemTitlesAsync(
            TargetUrl.TrimEnd('/'), targetListId, _copyCts!.Token);
        var existingByTitle = existingTargetItems
            .GroupBy(t => t.Title, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => (g.First().Id, g.First().Modified), StringComparer.OrdinalIgnoreCase);

        var allItems = await SpService.GetListItemsAsync(
            SourceUrl.TrimEnd('/'), def.SourceListId ?? string.Empty, def.Columns,
            ["HasUniqueRoleAssignments"], _copyCts!.Token);

        var items = isPartialSelection
            ? allItems.Where(i => i.TryGetValue("Id", out var id) && selectedItemIds.Contains(id?.ToString() ?? string.Empty)).ToList()
            : allItems;

        var targetListTitle = IsCreatingNewList ? NewListName.Trim()
                            : SelectedTargetList?.Title ?? def.Title;

        foreach (var item in items)
        {
            _copyCts.Token.ThrowIfCancellationRequested();

            var itemTitle = item.TryGetValue("Title", out var titleVal) ? titleVal?.ToString() : null;
            var itemId    = item.TryGetValue("Id",    out var iidVal)   ? iidVal?.ToString()   : "?";
            var rowLabel  = itemTitle ?? $"Item {itemId}";

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

            (string Id, DateTimeOffset? Modified)? existing =
                itemTitle != null && existingByTitle.TryGetValue(itemTitle, out var ex) ? ex : null;
            var hasUniquePerms = item.TryGetValue("HasUniqueRoleAssignments", out var hurv) && hurv is true;
            try
            {
                string? resolvedTargetItemId = existing?.Id;
                if (existing != null)
                {
                    if (OverwriteMode == Models.OverwriteMode.Skip)
                    {
                        itemResult.Status       = CopyStatus.Skipped;
                        itemResult.ErrorMessage = "Already exists";
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
                    itemResult.ErrorMessage = updateFieldError != null ? $"Updated (field warning: {updateFieldError})" : "Updated";
                }
                else
                {
                    var (newItemId, createFieldError) = await SpService.CreateListItemAsync(
                        TargetUrl.TrimEnd('/'), targetListId,
                        fields, createdDate, modifiedDate,
                        _copyCts.Token);
                    resolvedTargetItemId    = newItemId;
                    itemResult.Status       = CopyStatus.Success;
                    itemResult.ErrorMessage = createFieldError != null ? $"field warning: {createFieldError}" : null;
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
                    catch { /* non-fatal */ }
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
                ? $"Column created: {result.Created[0]}"
                : $"{result.Created.Count} columns created: {string.Join(", ", result.Created)}";
            System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(new CopyResult
            {
                FileName          = $"Columns → {listTitle}",
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
                FileName          = $"Column failed → {listTitle}",
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
                ? "1 role assignment applied"
                : $"{perm.Applied} role assignments applied";
            if (perm.SkippedPrincipals.Count > 0)
                detail += $"; skipped {perm.SkippedPrincipals.Count} unresolvable: {string.Join(", ", perm.SkippedPrincipals)}";
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
                FileName          = $"Permissions — {perm.ItemName}",
                SourcePath        = string.Empty,
                TargetPath        = targetPath,
                Status            = status,
                ErrorMessage      = detail,
                IsLibraryCreation = true,
            };
            System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(r));
        }
    }

    private async Task StartLibraryCopyAsync()
    {
        IsCopying      = true;
        IsCopyComplete = false;
        CopyDuration   = string.Empty;
        CopyResults.Clear();
        ResultFilter     = ResultFilterKind.All;
        FileResultFilter = ResultFilterKind.All;
        _copyStartTime = DateTimeOffset.Now;
        _copyEndTime   = null;

        _copyCts?.Dispose();
        _copyCts = new CancellationTokenSource();

        var progressTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(400)
        };
        progressTimer.Tick += (_, _) => UpdateProgress();
        progressTimer.Start();

        try
        {
            // Pre-warm permission cache once before any copy loops.
            // For Site scope, also copy the web-level permissions.
            if (CopyPermissions)
            {
                try { await _permissionCopyService.InitializeAsync(TargetUrl.TrimEnd('/'), _copyCts.Token); }
                catch { /* non-fatal — proceed without permission copy */ }

                if (IsSiteScope)
                {
                    try
                    {
                        var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                            SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                            "web", "web",
                            hasUniquePermissions: true,
                            "Site", _copyCts.Token);
                        if (perm.HasActivity)
                            AddPermissionResult(perm, TargetUrl.TrimEnd('/'));
                    }
                    catch (OperationCanceledException) { throw; }
                    catch { /* non-fatal */ }
                }
            }

            List<LibraryDefinition> definitions;
            var listDefinitions = new List<(LibraryDefinition Def, int BaseTemplate, SharePointNode? SourceNode)>();
            if (IsSiteScope)
            {
                StatusMessage = "Reading site library structure…";
                definitions   = await _libraryCopyService.ReadAllLibraryDefinitionsAsync(
                    SourceSiteId, SourceUrl.TrimEnd('/'));
            }
            else
            {
                StatusMessage = "Reading library structure…";
                definitions   = [];
                foreach (var lib in SourceLibraries)
                {
                    foreach (var node in lib.GetCheckedNodes().Where(n => n.Type == NodeType.Library))
                    {
                        if (node.IsCustomList)
                        {
                            var def = await _libraryCopyService.ReadListDefinitionAsync(
                                SourceUrl.TrimEnd('/'), node.Id, node.Name);
                            if (!string.IsNullOrWhiteSpace(node.OverrideName))
                                def.Title = node.OverrideName.Trim();
                            listDefinitions.Add((def, node.ListBaseTemplate, node));
                        }
                        else
                        {
                            var def = await _libraryCopyService.ReadLibraryDefinitionAsync(
                                SourceUrl.TrimEnd('/'), node.DriveId);
                            if (!string.IsNullOrWhiteSpace(node.OverrideName))
                                def.Title = node.OverrideName.Trim();
                            definitions.Add(def);
                        }
                    }

                    // Items-only mode (IsChecked == null) or partial item selection (IsChecked == false with
                    // individual children checked): list node itself is not fully checked.
                    if (lib.Type == NodeType.Library && lib.IsCustomList && lib.IsChecked != true)
                    {
                        var isItemsOnly   = lib.IsChecked == null;
                        var hasSelected   = isItemsOnly || lib.Children.Any(c => c.Type == NodeType.ListItem && c.IsChecked == true);
                        if (hasSelected)
                        {
                            var def = await _libraryCopyService.ReadListDefinitionAsync(
                                SourceUrl.TrimEnd('/'), lib.Id, lib.Name);
                            if (!string.IsNullOrWhiteSpace(lib.OverrideName))
                                def.Title = lib.OverrideName.Trim();
                            listDefinitions.Add((def, lib.ListBaseTemplate, lib));
                        }
                    }
                }
            }

            StatusMessage = string.Empty;

            foreach (var def in definitions)
            {
                _copyCts.Token.ThrowIfCancellationRequested();

                // Emit a "library created" result row
                var libResult = new CopyResult
                {
                    FileName          = def.Title,
                    SourcePath        = $"{SourceUrl.TrimEnd('/')}/{def.Title}",
                    TargetPath        = $"{TargetUrl.TrimEnd('/')}/{def.Title}",
                    Status            = CopyStatus.Copying,
                    IsLibraryCreation = true,
                };
                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(libResult));

                try
                {
                    var (newDriveId, newServerRelUrl) = await _libraryCopyService.CreateLibraryAsync(
                        TargetUrl.TrimEnd('/'), TargetSiteId, def, ColumnMappings);
                    libResult.Status = CopyStatus.Success;
                    if (_libraryCopyService.LastColumnCreationFailures.Count > 0)
                        libResult.ErrorMessage =
                            $"{_libraryCopyService.LastColumnCreationFailures.Count} column(s) could not be created: "
                            + string.Join("; ", _libraryCopyService.LastColumnCreationFailures.Take(3));

                    if (CopyPermissions)
                    {
                        try
                        {
                            var newListId = await SpService.GetListIdByServerRelativeUrlAsync(TargetUrl.TrimEnd('/'), newServerRelUrl);
                            var srcHasUnique = await SpService.GetHasUniqueRoleAssignmentsAsync(
                                SourceUrl.TrimEnd('/'), $"web/lists('{def.SourceListId}')", _copyCts.Token);
                            var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                                SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                                $"web/lists('{def.SourceListId}')",
                                $"web/lists('{newListId}')",
                                hasUniquePermissions: srcHasUnique,
                                def.Title, _copyCts.Token);
                            if (perm.HasActivity)
                                AddPermissionResult(perm, $"{TargetUrl.TrimEnd('/')}/{def.Title}");
                        }
                        catch (OperationCanceledException) { throw; }
                        catch { /* non-fatal */ }
                    }

                    if (CopyLibraryContent)
                    {
                        StatusMessage = $"Copying files into '{def.Title}'…";

                        // Build file jobs targeting the new library
                        var newLibRoot = await SpService.GetLibraryRootItemIdAsync(newDriveId);
                        if (newLibRoot == null)
                        {
                            libResult.ErrorMessage = "Library created but root item ID could not be retrieved; file copy skipped.";
                        }
                        else
                        {
                            var fileJobs = new List<CopyJob>
                            {
                                new CopyJob
                                {
                                    SourceDriveId                  = def.SourceDriveId,
                                    SourceItemId                   = "root",
                                    SourceName                     = def.Title,
                                    SourceSiteUrl                  = def.SourceSiteUrl,
                                    SourceDisplayPath              = $"{def.SourceSiteUrl}/{def.Title}",
                                    TargetDriveId                  = newDriveId,
                                    TargetParentItemId             = newLibRoot,
                                    TargetSiteId                   = TargetSiteId,
                                    TargetSiteUrl                  = TargetUrl.TrimEnd('/'),
                                    TargetDisplayPath              = $"{TargetUrl.TrimEnd('/')}/{def.Title}",
                                    TargetLibraryServerRelativeUrl = newServerRelUrl,
                                    IsFolder                       = true,
                                    IsLibrary                      = true,
                                    ColumnMappings                 = [.. ColumnMappings],
                                }
                            };

                            // Build bulk field cache for this library; also permission flags if enabled
                            _bulkFieldCache = [];
                            var libPermFlags = new Dictionary<string, bool>();
                            try
                            {
                                if (EffectiveCopyCustomColumns && def.Columns.Count > 0)
                                {
                                    var raw = await SpService.BulkReadCustomFieldsAsync(
                                        def.SourceSiteUrl, def.SourceListId, def.Columns,
                                        ct: _copyCts.Token);
                                    _bulkFieldCache = raw.ToDictionary(
                                        kvp => $"{def.SourceListId}:{kvp.Key}", kvp => kvp.Value);
                                }
                                if (CopyPermissions)
                                    libPermFlags = await SpService.BulkReadPermissionFlagsAsync(
                                        def.SourceSiteUrl, def.SourceListId, _copyCts.Token);
                            }
                            catch { }

                            int versionLimit = CopyVersions && !CopyAllVersions ? MaxVersions : 0;
                            await _copyService.ExecuteAsync(
                                fileJobs, CopyResults,
                                OverwriteMode, CopyVersions, MaxParallelCopies, versionLimit,
                                CopyMode, _copyCts.Token,
                                copyCustomColumns: EffectiveCopyCustomColumns,
                                columnMappings: [.. ColumnMappings],
                                bulkFieldCache: _bulkFieldCache,
                                preserveMetadata: PreserveMetadata,
                                reapplyFolderMetadata: ReapplyFolderMetadataEveryRun,
                                copyPermissions: CopyPermissions,
                                permissionService: CopyPermissions ? _permissionCopyService : null,
                                permissionFlags: libPermFlags);
                        }
                        StatusMessage = string.Empty;
                    }
                }
                catch (LibraryAlreadyExistsException alreadyEx)
                {
                    libResult.Status       = CopyStatus.Skipped;
                    libResult.ErrorMessage = alreadyEx.Message;

                    // Sync schema: add any source columns missing from the existing target library
                    if (!string.IsNullOrEmpty(alreadyEx.ServerRelativeUrl) && def.Columns.Count > 0)
                    {
                        try
                        {
                            var existingListId = await SpService.GetListIdByServerRelativeUrlAsync(
                                TargetUrl.TrimEnd('/'), alreadyEx.ServerRelativeUrl);
                            var createdCols = await _libraryCopyService.AddMissingColumnsAsync(
                                TargetUrl.TrimEnd('/'), existingListId, def.Columns, def.SourceSiteUrl);
                            AddColumnCreationResult(def.Title, $"{TargetUrl.TrimEnd('/')}/{def.Title}", createdCols);
                        }
                        catch { }
                    }

                    if (CopyLibraryContent && !string.IsNullOrEmpty(alreadyEx.DriveId))
                    {
                        StatusMessage = $"Copying files into existing '{def.Title}'…";

                        var newLibRoot = await SpService.GetLibraryRootItemIdAsync(alreadyEx.DriveId);
                        if (newLibRoot == null)
                        {
                            libResult.ErrorMessage += " (file copy skipped — root item ID unavailable)";
                        }
                        else
                        {
                            var serverRelUrl = alreadyEx.ServerRelativeUrl ?? string.Empty;
                            var fileJobs = new List<CopyJob>
                            {
                                new CopyJob
                                {
                                    SourceDriveId                  = def.SourceDriveId,
                                    SourceItemId                   = "root",
                                    SourceName                     = def.Title,
                                    SourceSiteUrl                  = def.SourceSiteUrl,
                                    SourceDisplayPath              = $"{def.SourceSiteUrl}/{def.Title}",
                                    TargetDriveId                  = alreadyEx.DriveId,
                                    TargetParentItemId             = newLibRoot,
                                    TargetSiteId                   = TargetSiteId,
                                    TargetSiteUrl                  = TargetUrl.TrimEnd('/'),
                                    TargetDisplayPath              = $"{TargetUrl.TrimEnd('/')}/{def.Title}",
                                    TargetLibraryServerRelativeUrl = serverRelUrl,
                                    IsFolder                       = true,
                                    IsLibrary                      = true,
                                    ColumnMappings                 = [.. ColumnMappings],
                                }
                            };

                            _bulkFieldCache = [];
                            var existLibPermFlags = new Dictionary<string, bool>();
                            try
                            {
                                if (EffectiveCopyCustomColumns && def.Columns.Count > 0)
                                {
                                    var raw = await SpService.BulkReadCustomFieldsAsync(
                                        def.SourceSiteUrl, def.SourceListId, def.Columns,
                                        ct: _copyCts.Token);
                                    _bulkFieldCache = raw.ToDictionary(
                                        kvp => $"{def.SourceListId}:{kvp.Key}", kvp => kvp.Value);
                                }
                                if (CopyPermissions)
                                    existLibPermFlags = await SpService.BulkReadPermissionFlagsAsync(
                                        def.SourceSiteUrl, def.SourceListId, _copyCts.Token);
                            }
                            catch { }

                            int versionLimit = CopyVersions && !CopyAllVersions ? MaxVersions : 0;
                            await _copyService.ExecuteAsync(
                                fileJobs, CopyResults,
                                OverwriteMode, CopyVersions, MaxParallelCopies, versionLimit,
                                CopyMode, _copyCts.Token,
                                copyCustomColumns: EffectiveCopyCustomColumns,
                                columnMappings: [.. ColumnMappings],
                                bulkFieldCache: _bulkFieldCache,
                                preserveMetadata: PreserveMetadata,
                                reapplyFolderMetadata: ReapplyFolderMetadataEveryRun,
                                copyPermissions: CopyPermissions,
                                permissionService: CopyPermissions ? _permissionCopyService : null,
                                permissionFlags: existLibPermFlags);
                        }
                        StatusMessage = string.Empty;
                    }
                }
                catch (Exception ex)
                {
                    libResult.Status       = CopyStatus.Failed;
                    libResult.ErrorMessage = ex.Message;
                }
            }

            // For Library scope: copy any selected custom lists (Tasks, Calendars, etc.).
            foreach (var (def, baseTemplate, sourceNode) in listDefinitions)
            {
                _copyCts.Token.ThrowIfCancellationRequested();

                // Determine copy mode for this list node.
                // isItemsOnly: IsChecked == null — copy all items, skip structure creation, use SelectedTargetList.
                // isPartialSelection: IsChecked == false with individual children checked — copy those items only.
                var isItemsOnly       = sourceNode?.IsChecked == null;
                var selectedItemIds   = (!isItemsOnly && sourceNode != null)
                    ? sourceNode.Children
                        .Where(c => c.Type == NodeType.ListItem && c.IsChecked == true)
                        .Select(c => c.Id)
                        .ToHashSet()
                    : [];
                var isPartialSelection = selectedItemIds.Count > 0;
                var needsItemCopy      = CopyLibraryContent || isPartialSelection || isItemsOnly;

                var listResult = new CopyResult
                {
                    FileName          = def.Title,
                    SourcePath        = $"{SourceUrl.TrimEnd('/')}/{def.Title}",
                    TargetPath        = $"{TargetUrl.TrimEnd('/')}/{def.Title}",
                    Status            = CopyStatus.Copying,
                    IsLibraryCreation = true,
                };
                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(listResult));

                try
                {
                    StatusMessage = $"Copying list '{def.Title}'…";
                    string targetListId;
                    if ((isPartialSelection || isItemsOnly) && SelectedTargetList != null)
                    {
                        // Items-only or partial selection: create a new list or use the chosen destination list.
                        if (IsCreatingNewList)
                        {
                            // The Target step is the single authority for the destination list
                            // name; the preview's Destination Name column mirrors it read-only.
                            def.Title    = NewListName.Trim();
                            targetListId = await _libraryCopyService.CreateCustomListAsync(
                                TargetUrl.TrimEnd('/'), TargetSiteId, def, baseTemplate);
                        }
                        else
                        {
                            targetListId = SelectedTargetList.Id;
                        }
                        if (CopyCustomColumns && def.Columns.Count > 0)
                        {
                            var colTitle = IsCreatingNewList ? def.Title : SelectedTargetList.Title;
                            try
                            {
                                var createdCols = await _libraryCopyService.AddMissingColumnsAsync(
                                    TargetUrl.TrimEnd('/'), targetListId, def.Columns, def.SourceSiteUrl);
                                AddColumnCreationResult(colTitle, $"{TargetUrl.TrimEnd('/')}/{colTitle}", createdCols);
                            }
                            catch { /* non-fatal — column creation best-effort */ }
                        }
                    }
                    else
                        targetListId = await _libraryCopyService.CreateCustomListAsync(
                            TargetUrl.TrimEnd('/'), TargetSiteId, def, baseTemplate);

                    if (CopyPermissions)
                    {
                        try
                        {
                            var srcHasUnique = await SpService.GetHasUniqueRoleAssignmentsAsync(
                                SourceUrl.TrimEnd('/'), $"web/lists('{def.SourceListId}')", _copyCts.Token);
                            var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                                SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                                $"web/lists('{def.SourceListId}')",
                                $"web/lists('{targetListId}')",
                                hasUniquePermissions: srcHasUnique,
                                def.Title, _copyCts.Token);
                            if (perm.HasActivity)
                                AddPermissionResult(perm, $"{TargetUrl.TrimEnd('/')}/{def.Title}");
                        }
                        catch (OperationCanceledException) { throw; }
                        catch { /* non-fatal */ }
                    }

                    if (needsItemCopy)
                        await CopyListItemsAsync(targetListId, def, selectedItemIds, isPartialSelection, listResult);

                    if (listResult.Status != CopyStatus.Failed)
                        listResult.Status = CopyStatus.Success;
                }
                catch (OperationCanceledException) { throw; }
                catch (LibraryAlreadyExistsException alreadyEx)
                {
                    listResult.Status       = CopyStatus.Skipped;
                    listResult.ErrorMessage = alreadyEx.Message;

                    if (!string.IsNullOrEmpty(alreadyEx.ListId) && def.Columns.Count > 0)
                    {
                        try
                        {
                            var createdCols = await _libraryCopyService.AddMissingColumnsAsync(
                                TargetUrl.TrimEnd('/'), alreadyEx.ListId, def.Columns, def.SourceSiteUrl);
                            AddColumnCreationResult(def.Title, $"{TargetUrl.TrimEnd('/')}/{def.Title}", createdCols);
                        }
                        catch { }
                    }

                    if (needsItemCopy && !string.IsNullOrEmpty(alreadyEx.ListId))
                        await CopyListItemsAsync(alreadyEx.ListId, def, selectedItemIds, isPartialSelection, listResult);

                    if (listResult.Status != CopyStatus.Failed)
                        listResult.Status = CopyStatus.Skipped;
                }
                catch (Exception ex)
                {
                    listResult.Status       = CopyStatus.Failed;
                    listResult.ErrorMessage = ex.Message;
                }
            }
            StatusMessage = string.Empty;

            // For Site scope: also copy Site Pages (excluded from the Drives API, so not
            // captured by ReadAllLibraryDefinitionsAsync). The target Site Pages library
            // always exists — no creation step needed, just copy pages into it.
            if (IsSiteScope && CopyLibraryContent)
            {
                _copyCts.Token.ThrowIfCancellationRequested();
                StatusMessage = "Copying site pages…";
                try
                {
                    var srcSitePages = await SpService.GetSitePagesLibraryAsync(SourceSiteId, SourceUrl.TrimEnd('/'));
                    var tgtSitePages = await SpService.GetSitePagesLibraryAsync(TargetSiteId, TargetUrl.TrimEnd('/'));

                    if (srcSitePages != null && tgtSitePages != null)
                    {
                        var pages = await SpService.GetChildrenAsync(
                            srcSitePages.DriveId, srcSitePages.Id,
                            srcSitePages.SiteId, srcSitePages.SiteUrl);

                        var pageJobs = pages
                            .Where(p => p.Type == NodeType.File)
                            .Select(p => new CopyJob
                            {
                                SourceDriveId                  = srcSitePages.DriveId,
                                SourceItemId                   = p.Id,
                                SourceName                     = p.Name,
                                SourceSiteUrl                  = SourceUrl.TrimEnd('/'),
                                SourceDisplayPath              = $"{SourceUrl.TrimEnd('/')}/SitePages/{p.Name}",
                                TargetDriveId                  = tgtSitePages.DriveId,
                                TargetParentItemId             = tgtSitePages.Id,
                                TargetSubFolderPath            = string.Empty,
                                TargetSiteId                   = TargetSiteId,
                                TargetSiteUrl                  = TargetUrl.TrimEnd('/'),
                                TargetDisplayPath              = $"{TargetUrl.TrimEnd('/')}/SitePages/{p.Name}",
                                TargetLibraryServerRelativeUrl = tgtSitePages.ServerRelativePath ?? string.Empty,
                                IsFolder                       = false,
                                IsPage                         = true,
                            })
                            .ToList();

                        if (pageJobs.Count > 0)
                        {
                            // Pre-add result rows — non-folder jobs are not dynamically added during execution
                            foreach (var pj in pageJobs)
                            {
                                var pr = new CopyResult
                                {
                                    FileName   = pj.SourceName,
                                    SourcePath = pj.SourceDisplayPath,
                                    TargetPath = pj.TargetDisplayPath,
                                };
                                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(pr));
                            }

                            // Build field cache for any custom columns on the Site Pages library
                            var pageBulkCache   = new Dictionary<string, Dictionary<string, object?>>();
                            var pagePermFlags   = new Dictionary<string, bool>();
                            if (srcSitePages.ServerRelativePath != null)
                            {
                                try
                                {
                                    var pagesListId = await SpService.GetListIdByServerRelativeUrlAsync(
                                        SourceUrl.TrimEnd('/'), srcSitePages.ServerRelativePath);
                                    if (EffectiveCopyCustomColumns)
                                    {
                                        var pageCols = await SpService.GetLibraryColumnsAsync(
                                            SourceUrl.TrimEnd('/'), pagesListId);
                                        if (pageCols.Count > 0)
                                            pageBulkCache = await SpService.BulkReadCustomFieldsAsync(
                                                SourceUrl.TrimEnd('/'), pagesListId, pageCols,
                                                ct: _copyCts.Token);
                                    }
                                    if (CopyPermissions)
                                        pagePermFlags = await SpService.BulkReadPermissionFlagsAsync(
                                            SourceUrl.TrimEnd('/'), pagesListId, _copyCts.Token);
                                }
                                catch { /* non-critical */ }
                            }

                            await _copyService.ExecuteAsync(
                                pageJobs, CopyResults,
                                OverwriteMode,
                                copyVersions: false,
                                MaxParallelCopies,
                                maxVersions: 0,
                                CopyMode.EnhancedRest,
                                _copyCts.Token,
                                copyPages: true,
                                remapPageWebPartUrls: RemapPageWebPartUrls,
                                preserveMetadata: PreserveMetadata,
                                reapplyFolderMetadata: ReapplyFolderMetadataEveryRun,
                                copyCustomColumns: EffectiveCopyCustomColumns,
                                columnMappings: [.. ColumnMappings],
                                bulkFieldCache: pageBulkCache,
                                copyPermissions: CopyPermissions,
                                permissionService: CopyPermissions ? _permissionCopyService : null,
                                permissionFlags: pagePermFlags);
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    StatusMessage = $"Site pages warning: {ex.Message}";
                }
                StatusMessage = string.Empty;
            }

            // For Site scope: copy navigation (Quick Launch + Top Navigation Bar).
            if (IsSiteScope && CopyNavigation)
            {
                _copyCts.Token.ThrowIfCancellationRequested();
                StatusMessage = "Copying navigation…";
                var navResult = new CopyResult
                {
                    FileName          = "Navigation",
                    SourcePath        = $"{SourceUrl.TrimEnd('/')}/navigation",
                    TargetPath        = $"{TargetUrl.TrimEnd('/')}/navigation",
                    Status            = CopyStatus.Copying,
                    IsLibraryCreation = true,
                };
                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(navResult));
                try
                {
                    await SpService.CopyNavigationAsync(SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'), quickLaunch: true);
                    await SpService.CopyNavigationAsync(SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'), quickLaunch: false);
                    navResult.Status = CopyStatus.Success;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    navResult.Status       = CopyStatus.Failed;
                    navResult.ErrorMessage = ex.Message;
                }
                StatusMessage = string.Empty;
            }

            // For Site scope: copy custom lists (Tasks, Calendars, Announcements, etc.).
            if (IsSiteScope)
            {
                _copyCts.Token.ThrowIfCancellationRequested();
                StatusMessage = "Reading custom lists…";
                try
                {
                    var customLists = await SpService.GetCustomListsAsync(SourceUrl.TrimEnd('/'));

                    foreach (var (srcListId, listTitle, baseTemplate) in customLists)
                    {
                        _copyCts.Token.ThrowIfCancellationRequested();

                        var listResult = new CopyResult
                        {
                            FileName          = listTitle,
                            SourcePath        = $"{SourceUrl.TrimEnd('/')}/{listTitle}",
                            TargetPath        = $"{TargetUrl.TrimEnd('/')}/{listTitle}",
                            Status            = CopyStatus.Copying,
                            IsLibraryCreation = true,
                        };
                        System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(listResult));

                        LibraryDefinition? definition = null;
                        try
                        {
                            StatusMessage = $"Copying list '{listTitle}'…";
                            definition   = await _libraryCopyService.ReadListDefinitionAsync(
                                SourceUrl.TrimEnd('/'), srcListId, listTitle);
                            var targetListId = await _libraryCopyService.CreateCustomListAsync(
                                TargetUrl.TrimEnd('/'), TargetSiteId, definition, baseTemplate);

                            if (CopyPermissions)
                            {
                                try
                                {
                                    var srcHasUnique = await SpService.GetHasUniqueRoleAssignmentsAsync(
                                        SourceUrl.TrimEnd('/'), $"web/lists('{srcListId}')", _copyCts.Token);
                                    var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                                        SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                                        $"web/lists('{srcListId}')",
                                        $"web/lists('{targetListId}')",
                                        hasUniquePermissions: srcHasUnique,
                                        listTitle, _copyCts.Token);
                                    if (perm.HasActivity)
                                        AddPermissionResult(perm, $"{TargetUrl.TrimEnd('/')}/{listTitle}");
                                }
                                catch (OperationCanceledException) { throw; }
                                catch { }
                            }

                            if (CopyLibraryContent)
                                await CopyListItemsAsync(targetListId, definition, [], false, listResult);

                            if (listResult.Status != CopyStatus.Failed)
                                listResult.Status = CopyStatus.Success;
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (LibraryAlreadyExistsException alreadyEx)
                        {
                            listResult.Status       = CopyStatus.Skipped;
                            listResult.ErrorMessage = alreadyEx.Message;

                            if (!string.IsNullOrEmpty(alreadyEx.ListId) && definition?.Columns.Count > 0)
                            {
                                try
                                {
                                    var createdCols = await _libraryCopyService.AddMissingColumnsAsync(
                                        TargetUrl.TrimEnd('/'), alreadyEx.ListId, definition.Columns, definition.SourceSiteUrl);
                                    AddColumnCreationResult(listTitle, $"{TargetUrl.TrimEnd('/')}/{listTitle}", createdCols);
                                }
                                catch { }
                            }

                            if (CopyLibraryContent && !string.IsNullOrEmpty(alreadyEx.ListId) && definition != null)
                                await CopyListItemsAsync(alreadyEx.ListId, definition, [], false, listResult);

                            if (listResult.Status != CopyStatus.Failed)
                                listResult.Status = CopyStatus.Skipped;
                        }
                        catch (Exception ex)
                        {
                            listResult.Status       = CopyStatus.Failed;
                            listResult.ErrorMessage = ex.Message;
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) { StatusMessage = $"List copy warning: {ex.Message}"; }
                StatusMessage = string.Empty;
            }
        }
        catch (OperationCanceledException) { StatusMessage = "Copy cancelled."; }
        catch (Exception ex)              { StatusMessage = $"Library copy error: {ex.Message}"; }
        finally
        {
            _copyEndTime   = DateTimeOffset.Now;
            progressTimer.Stop();
            IsCopying      = false;
            IsCopyComplete = true;
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
            SaveReport();
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
                ColumnLoadError = $"Target columns unavailable: {ex.Message}";
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
                ColumnLoadError = $"Target columns unavailable: {ex.Message}";
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
            ColumnLoadError ??= $"Source columns unavailable: {ex.Message}";
        }
        finally
        {
            IsColumnsLoading = false;
        }
    }


    private void UpdateProgress()
    {
        int done, total;
        lock (_copyResultsLock)
        {
            done  = CopyResults.Count(r => r.Status is CopyStatus.Success or CopyStatus.Failed or CopyStatus.Skipped);
            total = CopyResults.Count;
        }
        CompletedCount = done;
        TotalCount     = total;
        TotalProgress  = TotalCount > 0 ? done * 100.0 / TotalCount : 0;

        var elapsed = (_copyEndTime ?? DateTimeOffset.Now) - _copyStartTime;
        ElapsedTime = FormatDuration(elapsed);

        // Packaging ETA — rate-based estimate for the download+encrypt+upload phase.
        // Anchors when the first file is packed; waits 5 s for a stable rate before showing.
        var packed = PackedCount;
        if (packed > 0 && packed < total && _copyEndTime == null)
        {
            if (_packStartTime == null)
            {
                _packStartTime = DateTimeOffset.Now;
                _packStartDone = packed;
            }
            var packElapsed = (DateTimeOffset.Now - _packStartTime.Value).TotalSeconds;
            var packDone    = packed - _packStartDone;
            if (packDone > 0 && packElapsed >= 5)
            {
                var remainingSecs = (total - packed) * packElapsed / packDone;
                PackagedEstimatedTimeRemaining = remainingSecs >= 2
                    ? $" · ~{FormatDuration(TimeSpan.FromSeconds(remainingSecs))} remaining"
                    : string.Empty;
            }
            else
            {
                PackagedEstimatedTimeRemaining = string.Empty;
            }
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
                _etaStartTime = DateTimeOffset.Now;
                _etaStartDone = done;
            }
            var etaElapsed = (DateTimeOffset.Now - _etaStartTime.Value).TotalSeconds;
            var etaDone    = done - _etaStartDone;
            if (etaDone > 0 && etaElapsed >= 5)
            {
                var remainingSecs = (total - done) * etaElapsed / etaDone;
                EstimatedTimeRemaining = remainingSecs >= 2
                    ? $" · ~{FormatDuration(TimeSpan.FromSeconds(remainingSecs))} remaining"
                    : string.Empty;
            }
            else
            {
                EstimatedTimeRemaining = string.Empty;
            }
        }
        else if (done == 0 && packed > 0 && packed < total && _copyEndTime == null)
        {
            // Packaging phase — reuse the packaging ETA with a clearer label.
            var packElapsed2 = _packStartTime == null ? 0
                : (DateTimeOffset.Now - _packStartTime.Value).TotalSeconds;
            var packDone2 = packed - _packStartDone;
            if (packDone2 > 0 && packElapsed2 >= 5)
            {
                var remainingSecs = (total - packed) * packElapsed2 / packDone2;
                EstimatedTimeRemaining = remainingSecs >= 2
                    ? $" · ~{FormatDuration(TimeSpan.FromSeconds(remainingSecs))} to package"
                    : string.Empty;
            }
            else
            {
                EstimatedTimeRemaining = string.Empty;
            }
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
                    ? $" · ~{FormatDuration(TimeSpan.FromSeconds(remainingSecs))} remaining"
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
    }

    // Public wrapper so code-behind and dialogs can fire property change notifications.
    public new void OnPropertyChanged(string propertyName) => base.OnPropertyChanged(propertyName);

    // ── Step 5: Report ────────────────────────────────────────────────────────

    [RelayCommand]
    private void ExportFilesReport()
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt",
            FileName = $"{SiteUrlHelper.ReportFilenamePrefix(SourceUrl, TargetUrl, Settings.PrefixReportFilenamesWithSiteNames)}CopyReport_Files_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true) return;

        static string Csv(string? s) => $"\"{(s ?? "").Replace("\"", "\"\"")}\"";
        var sb = new StringBuilder();
        sb.AppendLine("File Name,Source Path,Target Path,Status,Versions Copied,Error,Permissions Status,Permissions Details");
        foreach (var r in CopyResults.Where(r => !r.IsPermissionResult))
        {
            sb.AppendLine($"{Csv(r.FileName)},{Csv(r.SourcePath)},{Csv(r.TargetPath)}," +
                          $"{r.Status},{r.VersionsCopied},{Csv(r.ErrorMessage)}," +
                          $"{r.PermissionStatus?.ToString() ?? ""},{Csv(r.PermissionDetails)}");
        }
        System.IO.File.WriteAllText(dlg.FileName, sb.ToString());
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
    }

    private bool CanExportVerificationReport() => !IsCopying && !IsVerifying && CopyJobs.Count > 0;

    [RelayCommand(CanExecute = nameof(CanExportVerificationReport))]
    private async Task ExportVerificationReportAsync()
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = "Excel Workbook (*.xlsx)|*.xlsx",
            // "DeepVerificationReport_" when the deep Office-file pass is on, so the file is
            // distinguishable from a standard verification at a glance. Live checkbox state.
            FileName = $"{SiteUrlHelper.ReportFilenamePrefix(SourceUrl, TargetUrl, Settings.PrefixReportFilenamesWithSiteNames)}{(DeepVerifyOfficeFiles ? "Deep" : "")}VerificationReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
        };
        if (dlg.ShowDialog() != true) return;

        _verifyCts?.Dispose();
        _verifyCts   = new CancellationTokenSource();
        IsVerifying  = true;
        // Verification previously only appended to whatever log file a prior copy in this session had
        // already opened — running Verify on its own (no copy this session, or after reopening the app)
        // left zero on-disk record of the run, so a real failure had nothing to inspect afterward.
        // Give it its own log file, same pattern as StartCopyAsync.
        StartActivityLogFile();
        if (ActivityLogPath != null)
            PushActivity($"Activity log: {ActivityLogPath}");
        // StatusMessage previously only ever got updated by scan progress below — once the scan
        // finished and the deep-verify pass started, its own progress messages (which DO arrive here
        // via onActivity/PushActivity, appended to the scrolling log) never replaced the single
        // status line, which stayed frozen on the last "found X source, Y target" text for the
        // entire deep-verify phase, making an actively-running verification look stalled. "⚠"-
        // prefixed messages (throttle waits, scan errors) are excluded — those are transient blips
        // that belong in the log, not overwriting the main status line on every occurrence.
        // Live checkbox state, not a settings re-read — see the field's own doc comment. Read up front
        // so every status/activity line below can label the run as "Deep verification" when the deep
        // Office-file pass is enabled, matching the checkbox the user ticked.
        bool deepVerify   = DeepVerifyOfficeFiles;
        string verifyName = deepVerify ? "Deep verification" : "Verification";  // sentence-initial
        string verifyLow  = deepVerify ? "deep verification" : "verification";  // mid-sentence
        string verifyIng  = deepVerify ? "Deep verifying"    : "Verifying";

        var onActivity = new Progress<string>(msg =>
        {
            PushActivity(msg);
            if (!msg.StartsWith('⚠'))
                StatusMessage = msg;
        });
        var onScanned  = new Progress<VerificationReportService.ScanProgress>(p =>
            StatusMessage = $"{verifyIng}… found {p.SourceFilesFound:N0} source file(s), {p.TargetFilesFound:N0} target file(s)");

        try
        {
            StatusMessage = $"Re-scanning source and target for {verifyLow}…";
            var roots  = VerificationRoot.FromCopyJobs(CopyJobs);
            var result = await _verificationReportService.RunAsync(
                roots, MaxParallelCopies, onActivity, onScanned, _verifyCts.Token, deepVerify);
            StatusMessage = $"Writing {verifyLow} workbook…";
            // Off the UI thread: ClosedXML SaveAs is CPU-heavy on 100k+ row runs.
            await Task.Run(() => ExcelReportWriter.Write(dlg.FileName, result));

            int matched         = result.Comparison.Count(r => r.Status == ComparisonStatus.Match);
            int contentMismatch = result.Comparison.Count(r => r.Status == ComparisonStatus.ContentMismatch);
            int dateMismatch    = result.Comparison.Count(r => r.Status == ComparisonStatus.DateMismatch);
            int onlyInSource    = result.Comparison.Count(r => r.Status == ComparisonStatus.OnlyInSource);
            int onlyInTarget    = result.Comparison.Count(r => r.Status == ComparisonStatus.OnlyInTarget);
            int unverified      = result.Comparison.Count(r => r.Status == ComparisonStatus.Unverified);
            PushActivity($"✔ {verifyName} complete: {matched:N0} matched, {contentMismatch:N0} content mismatch, {dateMismatch:N0} date mismatch, {onlyInSource:N0} only in source, {onlyInTarget:N0} only in target"
                + (unverified > 0 ? $", {unverified:N0} unverified (no comparable signal)" : ""));
            PushActivity($"{verifyName} report written: {dlg.FileName}");
            if (result.ScanErrors.Count > 0)
                PushActivity($"⚠ {result.ScanErrors.Count} root(s) could not be scanned — see the Scan Errors tab");
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
            StatusMessage = string.Empty;
        }
        catch (OperationCanceledException) { StatusMessage = $"{verifyName} cancelled."; PushActivity($"{verifyName} cancelled."); }
        catch (Exception ex)
        {
            StatusMessage = $"{verifyName} error: {ex.Message}";
            PushActivity($"⚠ {verifyName} error: {ex.Message}");
        }
        finally { IsVerifying = false; }
    }

    [RelayCommand]
    private void CancelVerification() => _verifyCts?.Cancel();

    // ── Navigation ────────────────────────────────────────────────────────────

    // Set to true by MainViewModel.Demo.cs LoadDemoData(); always false in production.
    private bool _isDemoMode = false;
    public bool IsDemoMode => _isDemoMode;
    partial void AdvanceDemoToReport();  // implemented in MainViewModel.Demo.cs; no-op when absent

    [RelayCommand(CanExecute = nameof(CanGoBack))]
    private void Back()
    {
        if (CurrentStep == 3)
            ColumnMappings.Clear();
        // Leaving the progress/report steps invalidates the completed run: without this, Back →
        // change options → Next silently jumped to the OLD results ("case 3: if (IsCopyComplete)")
        // and the new options never ran, while the UI presented them as applied. Next now re-runs
        // the copy with whatever is currently selected.
        if (CurrentStep == 4 && IsCopyComplete && !IsUpdatingMetadata)
            IsCopyComplete = false;
        CurrentStep--;
    }

    private bool CanGoBack() => CurrentStep > 0 && !IsCopying;

    [RelayCommand(CanExecute = nameof(CanGoNext))]
    private void Next()
    {
        switch (CurrentStep)
        {
            case 0 when SourceConnected || IsDemoMode:
                CurrentStep = 1;
                if (!IsDemoMode) _ = LoadLibrariesAsync();
                break;
            case 1:
                OnPropertyChanged(nameof(IsItemSelectionActive));
                NextCommand.NotifyCanExecuteChanged();
                CurrentStep = 2;
                break;
            case 2 when (TargetConnected && (NeedsTargetFolder ? SelectedTargetFolder != null : true)) || IsDemoMode:
                if (!IsDemoMode)
                {
                    if (NeedsTargetFolder) BuildCopyJobs();
                    // Item-level copies write into the chosen destination list (or a new one) —
                    // reflect that name in the preview's Destination Name column so what the
                    // user typed at the Target step is what the preview shows.
                    if (IsItemSelectionActive && !string.IsNullOrWhiteSpace(EffectiveDestinationListName))
                    {
                        foreach (var lib in SourceLibraries.Where(l => l.IsCustomList &&
                                     (l.IsChecked == null ||
                                      l.Children.Any(c => c.Type == NodeType.ListItem && c.IsChecked == true))))
                            lib.OverrideName = EffectiveDestinationListName;
                        OnPropertyChanged(nameof(LibraryPreviewItems));
                    }
                    // Eagerly load target columns for mapping dialog; store task so dialog can await it
                    _columnLoadTask = LoadTargetColumnsAsync();
                }
                CurrentStep = 3;
                break;
            case 3 when IsLibraryOrSiteScope || CopyJobs.Count > 0 || IsDemoMode:
                if (IsCopyComplete)
                {
                    // Copy already ran — navigate forward without re-running.
                    CurrentStep = 4;
                    break;
                }
                if (!IsDemoMode)
                {
                    Settings.PreferredCopyMode    = CopyMode;
                    Settings.OverwriteMode        = OverwriteMode;
                    // Keep the legacy bool in sync for older builds reading the same file.
                    Settings.OverwriteFiles       = OverwriteMode == Models.OverwriteMode.Overwrite;
                    Settings.CopyVersions         = CopyVersions;
                    Settings.CopyAllVersions      = CopyAllVersions;
                    Settings.MaxVersions          = MaxVersions;
                    Settings.MaxParallelCopies    = MaxParallelCopies;
                    Settings.CopyCustomColumns    = CopyCustomColumns;
                    Settings.CopyLibraryContent   = CopyLibraryContent;
                    Settings.RemapPageWebPartUrls = RemapPageWebPartUrls;
                    Settings.PreserveMetadata     = PreserveMetadata;
                    Settings.ReapplyFolderMetadataEveryRun = ReapplyFolderMetadataEveryRun;
                    Settings.CopyNavigation       = CopyNavigation;
                    Settings.Scope                = CopyScope;
                    Settings.Save();
                    _ = IsLibraryOrSiteScope ? StartLibraryCopyAsync() : StartCopyAsync();
                }
                CurrentStep = 4;
                break;
            case 4 when (IsCopyComplete && !IsUpdatingMetadata) || IsDemoMode:
                if (IsDemoMode) AdvanceDemoToReport();
                CurrentStep = 5;
                break;
        }
    }

    private bool CanGoNext() => IsDemoMode || CurrentStep switch
    {
        0 => SourceConnected,
        1 => IsSiteScope
               ? SourceConnected
               : SourceLibraries.Any(l => l.GetCheckedNodes().Any()),
        2 => TargetConnected &&
             (NeedsTargetFolder ? SelectedTargetFolder != null : true) &&
             (!IsItemSelectionActive || (SelectedTargetList != null &&
                 (!IsCreatingNewList || !string.IsNullOrWhiteSpace(NewListName)))),
        3 => IsLibraryOrSiteScope || CopyJobs.Count > 0,
        4 => IsCopyComplete && !IsUpdatingMetadata,
        _ => false
    };

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string FormatDuration(TimeSpan ts)
    {
        if (ts.TotalHours >= 1)   return $"{(int)ts.TotalHours}h {ts.Minutes}m {ts.Seconds}s";
        if (ts.TotalMinutes >= 1) return $"{(int)ts.TotalMinutes}m {ts.Seconds}s";
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
                SuccessCount   = SuccessCount,
                FailedCount    = FailedCount,
                SkippedCount   = SkippedCount,
                CancelledCount = CancelledCount,
                TotalCount     = TotalCount,
                CopyMode     = CopyMode,
                Items        = CopyResults.Select(r => new SavedReportItem
                {
                    FileName           = r.FileName,
                    SourcePath         = r.SourcePath,
                    TargetPath         = r.TargetPath,
                    Status             = r.Status,
                    VersionsCopied     = r.VersionsCopied,
                    VersionsTotal      = r.VersionsTotal,
                    ErrorMessage       = r.ErrorMessage,
                    IsPermissionResult = r.IsPermissionResult,
                    PermissionStatus   = r.PermissionStatus,
                    PermissionDetails  = r.PermissionDetails,
                }).ToList(),
                Roots = CopyJobs.Select(j => new SavedReportRoot
                {
                    SourceDriveId       = j.SourceDriveId,
                    SourceItemId        = j.SourceItemId,
                    SourceName          = j.SourceName,
                    IsFolder            = j.IsFolder,
                    IsLibrary           = j.IsLibrary,
                    TargetDriveId       = j.TargetDriveId,
                    TargetParentItemId  = j.TargetParentItemId,
                    TargetSubFolderPath = j.TargetSubFolderPath,
                }).ToList()
            };
            ReportHistoryService.Save(report);
        }
        catch { /* non-critical */ }
    }

    private static string BuildPath(SharePointNode node)
    {
        var parts   = new List<string>();
        var current = (SharePointNode?)node;
        while (current != null)
        {
            parts.Insert(0, current.Name);
            current = current.Parent;
        }
        return string.Join("/", parts);
    }

    // Returns the path of `node` relative to `libraryRoot` by walking the parent chain.
    // Returns empty string when node IS the library root.
    private static string BuildRelativePath(SharePointNode node, SharePointNode libraryRoot)
    {
        var parts   = new List<string>();
        var current = node;
        while (current != null && current != libraryRoot)
        {
            parts.Insert(0, current.Name);
            current = current.Parent;
        }
        return string.Join("/", parts);
    }

}
