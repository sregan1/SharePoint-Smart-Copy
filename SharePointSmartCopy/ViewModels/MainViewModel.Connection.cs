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

// Partial class MainViewModel: Connection. The shared fields, constructor and core setup live in MainViewModel.cs.
public partial class MainViewModel : ObservableObject
{
    // ── Step 0: Source ────────────────────────────────────────────────────────

    [ObservableProperty] private string _sourceUrl = string.Empty;
    [ObservableProperty] private string _sourceStatus = string.Empty;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(NextCommand))] private bool _sourceConnected;
    [ObservableProperty] private string _sourceSiteId = string.Empty;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(HeaderAccountsText))]
    private string _signedInUser = string.Empty;
    [ObservableProperty] private bool _isConnectingSource;

    [RelayCommand]
    private async Task ConnectSourceAsync()
    {
        _connectSourceCts?.Cancel();
        _connectSourceCts = new CancellationTokenSource();
        var ct = _connectSourceCts.Token;

        SourceStatus       = Loc.T("VM_Connecting");
        ConnectionNotice   = string.Empty;
        SourceConnected    = false;
        IsBusy             = true;
        IsConnectingSource = true;
        try
        {
            await AuthService.GetAccessTokenAsync(forceInteractive: !AuthService.IsAuthenticated, cancellationToken: ct);
            ct.ThrowIfCancellationRequested();
            SignedInUser = AuthService.UserName ?? string.Empty;
            if (IsMultiTenant)
            {
                // Mirror of the check in ConnectTargetAsync: routing is by host, so registering the
                // source on the target's host would silently re-route the target to the source sign-in.
                if (TargetConnected && Uri.TryCreate(SourceUrl.Trim(), UriKind.Absolute, out var srcUri)
                    && Uri.TryCreate(TargetUrl.Trim(), UriKind.Absolute, out var tgtUri)
                    && srcUri.Host.Equals(tgtUri.Host, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(Loc.T("VM_SameTenantError"));
                Router.RegisterHost(SourceUrl.Trim(), AuthService);
            }
            SourceSiteId = await SpService.GetSiteIdAsync(SourceUrl.Trim());
            ct.ThrowIfCancellationRequested();
            SourceStatus    = Loc.T("VM_SourceConnectedAs", SignedInUser);
            SourceConnected = true;
            Settings.SourceUrl = SourceUrl.Trim();
            Settings.Save();
            CurrentStep = 1;
            _ = LoadLibrariesAsync();
        }
        catch (OperationCanceledException) { SourceStatus = string.Empty; }
        catch (Exception ex)              { SetSourceError(ex); }
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
        StatusMessage = Loc.T("VM_LoadingLibraries");
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
            StatusMessage = Loc.T("VM_ErrorLoadingLibraries", ex.Message);
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
        StatusMessage = Loc.T("VM_LoadingSitePages");
        string? errorMessage = null;
        try
        {
            var sitePagesNode = await SpService.GetSitePagesLibraryAsync(SourceSiteId, SourceUrl.Trim());
            if (sitePagesNode == null)
            {
                errorMessage = Loc.T("VM_NoSitePagesLibrary");
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
            errorMessage = Loc.T("VM_ErrorLoadingSitePages", ex.Message);
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
                            Name         = string.IsNullOrWhiteSpace(title) ? Loc.T("VM_ItemNumber", id) : title,
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
            StatusMessage = Loc.T("VM_ErrorLoadingFolder", ex.Message);
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

        TargetStatus       = Loc.T("VM_Connecting");
        TargetConnected    = false;
        IsBusy             = true;
        IsConnectingTarget = true;
        try
        {
            if (IsMultiTenant)
            {
                // Routing is by SharePoint host, so both sides on one host cannot use different sign-ins.
                if (Uri.TryCreate(SourceUrl.Trim(), UriKind.Absolute, out var srcUri)
                    && Uri.TryCreate(TargetUrl.Trim(), UriKind.Absolute, out var tgtUri)
                    && srcUri.Host.Equals(tgtUri.Host, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(Loc.T("VM_SameTenantError"));

                await TargetAuthService.GetAccessTokenAsync(forceInteractive: !TargetAuthService.IsAuthenticated, cancellationToken: ct);
                ct.ThrowIfCancellationRequested();
                TargetSignedInUser = TargetAuthService.UserName ?? string.Empty;
                Router.RegisterHost(TargetUrl.Trim(), TargetAuthService);
            }

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
            TargetStatus    = Loc.T("VM_Connected");
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
                        TargetCustomLists.Add(new ListPickerItem(NewListSentinelId, Loc.T("VM_CreateNewList")));
                        foreach (var (id, title, _) in lists)
                            TargetCustomLists.Add(new ListPickerItem(id, title));
                    });
                }
                catch { /* non-critical — list picker just stays empty */ }
            }

            try { await TargetAuthService.GetSharePointTokenAsync(TargetUrl.Trim(), cancellationToken: ct); }
            catch
            {
                TargetStatus = Loc.T("VM_ConnectedConsentNote");
            }
        }
        catch (OperationCanceledException) { TargetStatus = string.Empty; }
        catch (Exception ex)              { SetTargetError(ex); }
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

}
