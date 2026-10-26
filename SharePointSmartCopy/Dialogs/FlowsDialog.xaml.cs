using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Data;
using Microsoft.Win32;
using SharePointSmartCopy.Forms;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Services;
using SharePointSmartCopy.Workflows;

namespace SharePointSmartCopy.Dialogs;

// Rewrites exported Power Automate flow packages so they point at the target tenant, in six guided steps like the
// main copy wizard:
//   1 Packages (add + automatic scan)  2 Sites  3 Lists & columns (automatic matching, editable)
//   4 Forms & email  5 Write  6 Report
// Originals are never modified; patched packages go to a separate output folder together with a report.
//
// Source and target are shown the same way on every step: SOURCE (neutral) is where the flows come from, TARGET (accent)
// is where they go, each with its tenant address.
public partial class FlowsDialog : Window
{
    private readonly SharePointService _sp;
    private readonly Func<string, bool, CancellationToken, Task> _prepareSite;       // signs in to / registers a site: (siteUrl, isSource, ct)
    private readonly Func<bool, string> _signedInUser;                                // account of the source (true) or target (false) sign-in
    private readonly Func<bool, string?> _tenantId;                                   // tenant ID of the source (true) or target (false) sign-in
    private readonly Func<CancellationToken, bool, Task<string>> _targetFormsToken;   // Microsoft Forms token of the target account
    private readonly Func<CancellationToken, bool, Task<string>> _sourceFormsToken;   // Microsoft Forms token of the source account
    private readonly Func<CancellationToken, Task<string?>> _signInSource;            // signs in to the source again (always opens the browser); returns its SharePoint address
    private readonly Func<CancellationToken, Task<string?>> _signInTarget;            // the same for the target
    private readonly Func<bool, CancellationToken, Task<string?>> _connectedRoot;     // tenant address of an already signed-in side (never prompts)
    private string _sourceRoot = string.Empty;   // SharePoint address of the tenant connected as source, e.g. https://contoso.sharepoint.com
    private string _targetRoot = string.Empty;   // the same for the target
    private readonly string _connectedSource;
    private readonly string _connectedTarget;
    private readonly bool _isMultiTenant;
    private readonly FlowsDialogViewModel _vm = new();

    private readonly ObservableCollection<TargetFormOption> _targetForms = [];       // forms found in the target account (shared by every form row)
    private readonly ObservableCollection<ValueMapRow> _values = [];
    private readonly ICollectionView _valuesView;
    // A mapping loaded from a file; applied to the rows now and again whenever the packages are rescanned.
    private FlowMapping? _loaded;
    private Dictionary<string, Dictionary<string, string>> _loadedColumns = new(StringComparer.OrdinalIgnoreCase);

    private readonly List<string> _zips = [];
    // Packages in the list that this tool wrote earlier (recognized by content), by path.
    private readonly Dictionary<string, FlowOutputLog.Entry> _rewritten = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, List<FlowScan>> _zipScans = new(StringComparer.OrdinalIgnoreCase);   // the flows found in each package
    private readonly Dictionary<string, string> _saveAs = new(StringComparer.OrdinalIgnoreCase);        // package path -> the name the user chose to save it as
    private bool _formNamesAttempted;   // the form names load by themselves the first time the Forms step opens
    private List<FlowScan> _scans = [];
    private FlowResolveResult? _resolved;
    private readonly ObservableCollection<SiteRow> _sites = [];
    private readonly ObservableCollection<FormMapRow> _forms = [];
    // One card per list the flows use (however many flows use it): its target list and column mapping.
    private readonly ObservableCollection<ListCard> _cards = [];
    private readonly Dictionary<string, ObservableCollection<TargetListOption>> _targetListOptions = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, string> _lastReadMap = new(StringComparer.OrdinalIgnoreCase);   // flow site -> where its lists were read from
    private SharePointFlowLookup? _lookup;                 // kept so a card can read a newly chosen target list's columns
    private readonly List<string> _listIds = [];           // the distinct list IDs the flows use
    private CancellationTokenSource? _cts;
    private string? _lastOutput;
    private string? _lastReport;

    public FlowsDialog(SharePointService sp, Func<string, bool, CancellationToken, Task> prepareSite, Func<bool, string> signedInUser,
        Func<bool, string?> tenantId, Func<CancellationToken, bool, Task<string>> targetFormsToken,
        Func<CancellationToken, bool, Task<string>> sourceFormsToken,
        Func<CancellationToken, Task<string?>> signInSource, Func<CancellationToken, Task<string?>> signInTarget,
        Func<bool, CancellationToken, Task<string?>> connectedRoot,
        string sourceUrl, string targetUrl, bool isMultiTenant)
    {
        _sp = sp;
        _prepareSite = prepareSite;
        _signedInUser = signedInUser;
        _tenantId = tenantId;
        _targetFormsToken = targetFormsToken;
        _sourceFormsToken = sourceFormsToken;
        _connectedRoot = connectedRoot;
        _signInSource = signInSource;
        _signInTarget = signInTarget;
        _isMultiTenant = isMultiTenant;
        _vm.IsMultiTenant = isMultiTenant;
        _connectedSource = FlowMappingResolver.NormalizeSite(sourceUrl ?? "");
        _connectedTarget = FlowMappingResolver.NormalizeSite(targetUrl ?? "");
        // Each side's tenant address is read from its signed-in account (see RefreshRootsAsync), NOT taken from the site URLs typed
        // on the main window, which can be stale or belong to the other side.
        DataContext = _vm;
        InitializeComponent();

        SitesGrid.ItemsSource = _sites;
        FormsGrid.ItemsSource = _forms;
        _valuesView = CollectionViewSource.GetDefaultView(_values);
        _valuesView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(ValueMapRow.KindText)));
        ValuesGrid.ItemsSource = _valuesView;
        CardsList.ItemsSource = _cards;
        PackagesList.ItemsSource = _vm.PackageRows;
        PackagesGrid.ItemsSource = _vm.PackageResults;

        // The last output folder is remembered; the first time it is Documents\SharePointSmartCopy\Flows.
        _vm.OutputPath = FlowsUiState.Load().OutputPath is { Length: > 0 } saved
            ? saved
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SharePointSmartCopy", "Flows");
        _vm.PropertyChanged += OnViewModelChanged;
        _vm.TargetConn.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(SideConnection.State)) UpdateGates(); };
        ShowConnectionStatus();
        Loaded += async (_, _) => await RefreshRootsAsync();
    }

    private void OnViewModelChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(FlowsDialogViewModel.IsBusy)) UpdateGates();
        if (e.PropertyName == nameof(FlowsDialogViewModel.ShowReadFrom))
            ReadFromColumn.Visibility = _vm.ShowReadFrom ? Visibility.Visible : Visibility.Collapsed;
    }

    // ── Connections (shown on the Sites step) ─────────────────────────────────

    // "https://contoso.sharepoint.com" for any site URL on it, or "" when the text is not a URL.
    private static string RootOf(string url) =>
        Uri.TryCreate(url?.Trim(), UriKind.Absolute, out var u) ? $"{u.Scheme}://{u.Host}" : string.Empty;

    private static string HostOf(string url) =>
        Uri.TryCreate(url?.Trim(), UriKind.Absolute, out var u) ? u.Host : string.Empty;

    // The tenant part of a SharePoint host, ignoring OneDrive's "-my": contoso-my.sharepoint.com and
    // contoso.sharepoint.com are both "contoso".
    private static string TenantOf(string host)
    {
        const string sp = ".sharepoint.com", my = "-my";
        var h = host.EndsWith(sp, StringComparison.OrdinalIgnoreCase) ? host[..^sp.Length] : host;
        return h.EndsWith(my, StringComparison.OrdinalIgnoreCase) ? h[..^my.Length] : h;
    }

    // Reads the tenant address of each side that is already signed in (the main window's connections), so the Source and
    // Target lines show the tenant of the ACCOUNT, not a typed URL. Never opens a sign-in window; a side that is not signed
    // in simply shows no address yet.
    private async Task RefreshRootsAsync()
    {
        try
        {
            var source = await _connectedRoot(true, CancellationToken.None);
            if (!string.IsNullOrWhiteSpace(source)) _sourceRoot = RootOf(source);
            var target = await _connectedRoot(false, CancellationToken.None);
            if (!string.IsNullOrWhiteSpace(target)) _targetRoot = RootOf(target);
            ShowConnectionStatus();
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            // Only the address display depends on this; the tool works without it.
            Log(ex.Message);
        }
    }

    // What the dialog is signed in to right now: each side's tenant address and the account it is connected as. Single
    // tenant: both sides are the account from the main window. Multi-tenant: each side has its own sign-in, made on the
    // Connect step of the main window or with the buttons here.
    private void ShowConnectionStatus()
    {
        ApplySide(_vm.SourceConn, _sourceRoot, _signedInUser(true));
        ApplySide(_vm.TargetConn, _targetRoot, _signedInUser(false));
        if (!_isMultiTenant) return;
        FillReadFromSourceTenant();
        UpdateSourceWarning();
    }

    private void ApplySide(SideConnection side, string root, string account)
    {
        if (side.State == "Waiting") return;   // a sign-in window is open; its own code sets the result
        var signedIn = !string.IsNullOrWhiteSpace(account);
        side.Host = HostOf(root);
        side.Account = signedIn ? account : string.Empty;
        side.State = signedIn ? "Connected" : "None";
        side.Note = string.Empty;
        side.Suffix = !_isMultiTenant && signedIn ? Loc.T("Flows_ConnMainWindowSignIn") : string.Empty;
    }

    // Puts a side back to what it really is (signed in or not) after a sign-in attempt ends.
    private void Settle(SideConnection side)
    {
        side.State = "None";
        ShowConnectionStatus();
    }

    // Hosts of flow sites whose "read from" address was filled in automatically (see FillReadFromSourceTenant).
    private readonly HashSet<string> _autoFilledFlowHosts = new(StringComparer.OrdinalIgnoreCase);

    // The flows name sites on whichever tenant they were exported from. When that is not the tenant connected as source,
    // reading those sites from the flow's address can only fail, so "Read lists from" is filled in with the same site on the
    // source tenant (same path, "-my" kept for OneDrive). A row the user edited is left alone; a row whose flow site IS on
    // the source tenant goes back to the flow's own address.
    private void FillReadFromSourceTenant()
    {
        _autoFilledFlowHosts.Clear();
        if (!_isMultiTenant || _sourceRoot.Length == 0) return;
        var root = new Uri(_sourceRoot);
        var sourceTenant = TenantOf(root.Host);
        foreach (var row in _sites.Where(r => !r.ReadFromEdited))
        {
            if (!Uri.TryCreate(row.Source, UriKind.Absolute, out var src)) continue;
            if (TenantOf(src.Host).Equals(sourceTenant, StringComparison.OrdinalIgnoreCase))
            {
                if (!string.Equals(row.ReadFrom, row.Source, StringComparison.Ordinal)) row.SetReadFromAuto(row.Source);
            }
            else
            {
                row.SetReadFromAuto(TargetHostFor(src, root) + src.AbsolutePath.TrimEnd('/'));
                _autoFilledFlowHosts.Add(src.Host);
            }
        }
        // Show the advanced column when the tool changed something there, so the correction is visible.
        if (_autoFilledFlowHosts.Count > 0) _vm.ShowReadFrom = true;
    }

    // Under the connections: ONE message, chosen by priority. A site that is on the target tenant (a package this tool wrote);
    // a site still read from a tenant other than the source; or a note about what was filled in automatically.
    private void UpdateSourceWarning()
    {
        _vm.ShowRemoveTargetPackages = false;
        if (!_isMultiTenant) { _vm.SourceWarning = string.Empty; return; }
        var sourceHost = _sourceRoot.Length > 0 ? new Uri(_sourceRoot).Host : string.Empty;
        var sourceTenant = sourceHost.Length > 0 ? TenantOf(sourceHost) : string.Empty;
        var others = _sites
            .Select(r => HostOf(string.IsNullOrWhiteSpace(r.ReadFrom) ? r.Source : r.ReadFrom))
            .Where(h => h.Length > 0 && !TenantOf(h).Equals(sourceTenant, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        // Flows whose own sites are on the TARGET tenant were most likely written by this tool already (or exported from the
        // target): there is nothing left to move, and the source lookups cannot work on them. Before the target sign-in, the
        // target address saved on the main window is good enough for this check.
        var targetForCheck = _targetRoot.Length > 0 ? _targetRoot : RootOf(_connectedTarget);
        var onTarget = targetForCheck.Length == 0 ? [] : _sites
            .Select(r => HostOf(r.Source))
            .Where(h => h.Length > 0 && TenantOf(h).Equals(TenantOf(HostOf(targetForCheck)), StringComparison.OrdinalIgnoreCase)
                        && !TenantOf(h).Equals(sourceTenant, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        // Source and target on one tenant: Multi-tenant mode cannot route that.
        if (sourceTenant.Length > 0 && _targetRoot.Length > 0 && TenantOf(HostOf(_targetRoot)).Equals(sourceTenant, StringComparison.OrdinalIgnoreCase))
        {
            _vm.SourceWarning = Loc.T("Flows_SameTenantBoth", sourceHost);
            return;
        }
        if (onTarget.Count > 0)
        {
            _vm.SourceWarning = Loc.T("Flows_FlowsOnTarget", string.Join(", ", onTarget), sourceHost.Length > 0 ? sourceHost : Loc.T("Flows_State_None"));
            _vm.ShowRemoveTargetPackages = true;
            return;
        }

        if (sourceHost.Length == 0) { _vm.SourceWarning = string.Empty; return; }
        _vm.SourceWarning = others.Count > 0 ? Loc.T("Flows_SourceWarn", string.Join(", ", others), sourceHost)
            : _autoFilledFlowHosts.Count > 0 ? Loc.T("Flows_SourceAutoFilled", string.Join(", ", _autoFilledFlowHosts), sourceHost)
            : string.Empty;
    }

    private async void ConnectSource_Click(object sender, RoutedEventArgs e) => await ConnectSideAsync(isSource: true);
    private async void ConnectTarget_Click(object sender, RoutedEventArgs e) => await ConnectSideAsync(isSource: false);

    // "Change source…" / "Sign in to target…": always opens the sign-in window (so it also switches account or browser),
    // then reads that tenant's SharePoint address from the account and shows it. For the target it also fills the empty
    // target sites from that address. No site has to be typed first. While it waits, the button is replaced by Cancel.
    private async Task ConnectSideAsync(bool isSource)
    {
        var side = isSource ? _vm.SourceConn : _vm.TargetConn;
        _vm.ConnectingSide = isSource ? "source" : "target";
        side.State = "Waiting";
        side.Note = Loc.T("Flows_ConnWaiting");
        try
        {
            await RunAsync(Loc.T(isSource ? "Flows_ConnectingSource" : "Flows_ConnectingTarget"), async ct =>
            {
                try
                {
                    Log(Loc.T(isSource ? "Flows_LogConnectSource" : "Flows_LogConnectTarget"));
                    var root = await (isSource ? _signInSource(ct) : _signInTarget(ct));
                    Settle(side);
                    if (string.IsNullOrWhiteSpace(root)) return;   // single tenant: nothing to infer

                    var host = new Uri(root);
                    var who = _signedInUser(isSource);
                    if (isSource) _sourceRoot = RootOf(root); else _targetRoot = RootOf(root);
                    _vm.ShowManualHost = false;
                    Log(Loc.T(isSource ? "Flows_LogConnectedSource" : "Flows_LogConnectedTarget", who, root));
                    ShowConnectionStatus();

                    if (isSource)
                    {
                        _vm.StatusText = Loc.T("Flows_ConnSourceDone", host.Host);
                        return;
                    }

                    var filled = FillTargetsFromHost(host);
                    AutoFillFromSignIns();

                    // If the flows turn out to be on the target (or both sides are one tenant), the warning under the connections
                    // says so once; the sign-in itself worked, so the target stays "Connected".
                    _vm.StatusText = filled > 0 ? Loc.T("Flows_ConnTargetFilled", filled, host.Host) : Loc.T("Flows_ConnTargetNothingToFill");
                }
                catch (OperationCanceledException)
                {
                    Settle(side);
                    side.Note = Loc.T("Flows_ConnCancelled");
                    throw;
                }
                catch (Exception ex)
                {
                    // The sign-in failed, or it worked but the address could not be read: say why. For the target, offer the
                    // manual address box (unless the problem is "same tenant", which an address does not fix).
                    Settle(side);
                    side.State = "Error";
                    side.Note = Loc.T("Flows_ConnFailed", ex.Message);
                    if (!isSource) _vm.ShowManualHost = ex is not InvalidOperationException;
                    Log(ex.Message, important: true);
                    _vm.StatusText = Loc.T(isSource ? "Flows_ConnSourceFailed" : "Flows_ConnInferFailed");
                }
            });
        }
        finally
        {
            _vm.ConnectingSide = string.Empty;
            if (side.State == "Waiting") Settle(side);
        }
    }

    // ── Step 1: packages (scan runs automatically when files are added) ───────

    private async void AddZips_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Multiselect = true, Filter = Loc.T("Flows_ZipFilter"), Title = Loc.T("Flows_AddZips") };
        if (dlg.ShowDialog(this) == true) await AddAndScanAsync(dlg.FileNames);
    }

    private async void AddFolder_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog { Title = Loc.T("Flows_AddFolder") };
        if (dlg.ShowDialog(this) == true) await AddAndScanAsync(Directory.GetFiles(dlg.FolderName, "*.zip"));
    }

    private void ClearPackages_Click(object sender, RoutedEventArgs e) => ResetAll();

    // Dropping .zip files (or a folder of them) onto the Packages step adds them, same as the buttons.
    private void Packages_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = !_vm.IsBusy && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private async void Packages_Drop(object sender, DragEventArgs e)
    {
        if (_vm.IsBusy || e.Data.GetData(DataFormats.FileDrop) is not string[] items) return;
        var zips = new List<string>();
        foreach (var item in items)
        {
            if (Directory.Exists(item)) zips.AddRange(Directory.GetFiles(item, "*.zip"));
            else if (item.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)) zips.Add(item);
        }
        if (zips.Count > 0) await AddAndScanAsync(zips);
    }

    private async void RemovePackage_Click(object sender, RoutedEventArgs e)
    {
        if (_vm.IsBusy || (sender as FrameworkElement)?.DataContext is not PackageRow row) return;
        _zips.RemoveAll(z => string.Equals(z, row.FullPath, StringComparison.OrdinalIgnoreCase));
        _rewritten.Remove(row.FullPath);
        RefreshPackageWarning();
        if (_zips.Count == 0) { ResetAll(); return; }
        await ScanAsync();
    }

    // "Remove them" in the warning: drops every package whose sites are all on the TARGET tenant (a package this tool wrote),
    // so the original exports can be added instead.
    private async void RemoveTargetPackages_Click(object sender, RoutedEventArgs e)
    {
        if (_vm.IsBusy) return;
        var targetForCheck = _targetRoot.Length > 0 ? _targetRoot : RootOf(_connectedTarget);
        if (targetForCheck.Length == 0) return;
        var targetTenant = TenantOf(HostOf(targetForCheck));
        var drop = _zips.Where(z => _zipScans.TryGetValue(z, out var flows) && AllSitesOnTenant(flows, targetTenant)).ToList();
        if (drop.Count == 0) return;
        foreach (var z in drop) { _zips.Remove(z); _rewritten.Remove(z); }
        RefreshPackageWarning();
        Log(Loc.T("Flows_LogRemovedTargetPackages", drop.Count, string.Join(", ", drop.Select(Path.GetFileName))), important: true);
        if (_zips.Count == 0)
        {
            ResetAll();
            _vm.StatusText = Loc.T("Flows_RemovedAllAddOriginals");
            return;
        }
        await ScanAsync();
    }

    private static bool AllSitesOnTenant(List<FlowScan> flows, string tenant)
    {
        var hosts = flows.SelectMany(f => f.Findings).Where(f => f.Kind == FindingKind.SiteUrl && !f.Value.StartsWith('@'))
            .Select(f => HostOf(f.Value)).Where(h => h.Length > 0).ToList();
        return hosts.Count > 0 && hosts.All(h => TenantOf(h).Equals(tenant, StringComparison.OrdinalIgnoreCase));
    }

    private async Task AddAndScanAsync(IEnumerable<string> paths)
    {
        var added = new List<string>();
        foreach (var p in paths)
            if (!_zips.Contains(p, StringComparer.OrdinalIgnoreCase)) { _zips.Add(p); added.Add(p); }

        // A package this tool wrote keeps the original's file name but already holds the TARGET addresses, so it is easy to
        // add by mistake. Recognized by content, wherever it was copied to.
        var found = await Task.Run(() => added.Select(p => (Path: p, Entry: FlowOutputLog.Find(p))).Where(x => x.Entry is not null).ToList());
        foreach (var (path, entry) in found)
        {
            _rewritten[path] = entry!;
            Log(Loc.T("Flows_AlreadyRewritten", Path.GetFileName(path), entry!.WrittenAt.LocalDateTime.ToString("g")), important: true);
        }
        RefreshPackageWarning();
        await ScanAsync();
    }

    // The step-1 warning lists the packages in the list that this tool wrote.
    private void RefreshPackageWarning()
    {
        _vm.PackageWarning = string.Join("\n", _zips.Where(_rewritten.ContainsKey)
            .Select(z => Loc.T("Flows_AlreadyRewritten", Path.GetFileName(z), _rewritten[z].WrittenAt.LocalDateTime.ToString("g"))));
        UpdateGates();
    }

    private async Task ScanAsync()
    {
        if (_zips.Count == 0) { _vm.StatusText = Loc.T("Flows_NoPackages"); return; }
        await RunAsync(Loc.T("Flows_Scanning"), async ct =>
        {
            var scans = new List<FlowScan>();
            var perZip = new Dictionary<string, List<FlowScan>>(StringComparer.OrdinalIgnoreCase);
            var unreadable = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Log(Loc.T("Flows_LogScanStart", _zips.Count));
            await Task.Run(() =>
            {
                foreach (var zip in _zips)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        var found = FlowScanner.ScanPackage(zip).ToList();
                        perZip[zip] = found;
                        scans.AddRange(found);
                        Log(Loc.T("Flows_LogScanned", Path.GetFileName(zip), found.Count));
                    }
                    catch (Exception ex) when (ex is InvalidDataException or IOException or System.Text.Json.JsonException)
                    {
                        unreadable.Add(zip);
                        AppendLog(Loc.T("Flows_LogBadPackage", Path.GetFileName(zip), ex.Message));
                        Log(Loc.T("Flows_LogBadPackage", Path.GetFileName(zip), ex.Message), important: true);
                    }
                }
            }, ct);
            _scans = scans;
            _zipScans = perZip;
            _formNamesAttempted = false;
            _resolved = null;
            _cards.Clear();
            _vm.HasCards = false;
            _vm.MatchSummary.Text = string.Empty;
            BuildPackageRows(perZip, unreadable);
            FillRowsFromScan();
            ShowResources(scans);

            _vm.FlowCount = scans.Count;
            _vm.HasScan = scans.Count > 0;
            var f = scans.SelectMany(s => s.Findings).ToList();
            _vm.ScanSummaryText = scans.Count == 0
                ? Loc.T("Flows_NoFlowsFound")
                : Loc.T("Flows_ScanSummary", scans.Count,
                    f.Where(x => x.Kind == FindingKind.SiteUrl).Select(x => x.Value).Distinct().Count(),
                    f.Where(x => x.Kind == FindingKind.ListId).Select(x => x.Value).Distinct().Count(),
                    f.Where(x => x.Kind == FindingKind.FormId).Select(x => x.Value).Distinct().Count());
            _vm.StatusText = scans.Count == 0 ? Loc.T("Flows_NoFlowsFound") : string.Empty;
        });
    }

    // One row per added package: the flows found in it, how many sites and lists they use, and whether anything is off.
    private void BuildPackageRows(Dictionary<string, List<FlowScan>> perZip, HashSet<string> unreadable)
    {
        _vm.PackageRows.Clear();
        foreach (var zip in _zips)
        {
            var flows = perZip.GetValueOrDefault(zip) ?? [];
            var (status, level) =
                unreadable.Contains(zip) ? (Loc.T("Flows_Pkg_Unreadable"), "Error")
                : _rewritten.ContainsKey(zip) ? (Loc.T("Flows_Pkg_Rewritten"), "Look")
                : flows.Count == 0 ? (Loc.T("Flows_Pkg_NoFlows"), "Look")
                : (Loc.T("Flows_Pkg_Flows", flows.Count), "Ready");
            var findings = flows.SelectMany(s => s.Findings).Where(x => !x.Value.StartsWith('@')).ToList();
            _vm.PackageRows.Add(new PackageRow
            {
                FileName = Path.GetFileName(zip),
                FullPath = zip,
                FlowNames = string.Join(", ", flows.Select(s => s.FlowName)),
                SiteCount = findings.Where(x => x.Kind == FindingKind.SiteUrl).Select(x => FlowMappingResolver.NormalizeSite(x.Value)).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                ListCount = findings.Where(x => x.Kind == FindingKind.ListId).Select(x => x.Value).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                StatusText = status,
                StatusLevel = level,
            });
        }
    }

    private void FillRowsFromScan()
    {
        // Adding or removing a package scans again and rebuilds every row. Keep what was already entered for the
        // rows that still exist (site targets, "read from" edits, form IDs, replacement values) instead of
        // handing the user a blank table after they filled it in.
        var prevSites = new Dictionary<string, (string Target, string? ReadFrom)>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in _sites) prevSites.TryAdd(r.Source, (r.Target, r.ReadFromEdited ? r.ReadFrom : null));
        var prevForms = new Dictionary<string, (string NewId, string Title)>(StringComparer.Ordinal);
        foreach (var r in _forms) prevForms.TryAdd(r.SourceId, (r.NewId, r.SourceTitle));
        var prevValues = new Dictionary<string, (string Target, string Note)>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in _values) prevValues.TryAdd(r.Category + "|" + r.Source, (r.Target, r.Note));

        _sites.Clear();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in _scans.SelectMany(s => s.Findings).Where(f => f.Kind == FindingKind.SiteUrl && !f.Value.StartsWith('@')))
        {
            var norm = FlowMappingResolver.NormalizeSite(f.Value);
            if (!seen.Add(norm)) continue;
            // The site the user connected as source maps to the connected target by default.
            var target = _connectedTarget.Length > 0 && string.Equals(norm, _connectedSource, StringComparison.OrdinalIgnoreCase)
                ? _connectedTarget : string.Empty;
            prevSites.TryGetValue(norm, out var prevSite);
            if (!string.IsNullOrWhiteSpace(prevSite.Target)) target = prevSite.Target;
            var row = new SiteRow { Source = norm, Target = target };
            row.SetReadFromAuto(norm);
            if (prevSite.ReadFrom is { Length: > 0 } keptRead) row.ReadFrom = keptRead;   // a user edit stays an edit
            row.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(SiteRow.ReadFrom)) { _resolved = null; UpdateSourceWarning(); }
                else if (e.PropertyName == nameof(SiteRow.Target)) { _resolved = null; UpdateGates(); }   // lists were matched against the OLD target site
            };
            _sites.Add(row);
        }
        _vm.HasSites = _sites.Count > 0;
        UpdateGates();

        _listIds.Clear();
        _listIds.AddRange(_scans.SelectMany(s => s.Findings).Where(f => f.Kind == FindingKind.ListId && !f.Value.StartsWith('@'))
                     .Select(f => f.Value).Distinct(StringComparer.OrdinalIgnoreCase));

        _forms.Clear();
        var saved = FormsIdMapStore.Load();
        foreach (var id in _scans.SelectMany(s => s.Findings).Where(f => f.Kind == FindingKind.FormId).Select(f => f.Value).Distinct())
        {
            saved.TryGetValue(id, out var fromCopy);
            prevForms.TryGetValue(id, out var prevForm);
            var row = new FormMapRow
            {
                Options = _targetForms, SourceId = id,
                SourceTitle = prevForm.Title is { Length: > 0 } t ? t : fromCopy?.SourceTitle ?? "",
                NewId = prevForm.NewId is { Length: > 0 } n ? n : fromCopy?.NewId ?? "",
            };
            row.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(FormMapRow.NewId)) { ShowFormNote(row, fromCopy); UpdateGates(); }
                else if (e.PropertyName == nameof(FormMapRow.StatusLevel)) UpdateGates();
            };
            ShowFormNote(row, fromCopy);
            _forms.Add(row);
        }
        _vm.HasForms = _forms.Count > 0;

        // Other values the flows hold that the tool can map when told how: addresses, the tenant ID, library paths.
        _values.Clear();
        var findings = _scans.SelectMany(s => s.Findings).Where(f => !f.Value.StartsWith('@')).ToList();
        void AddValues(string category, IEnumerable<Finding> source)
        {
            foreach (var v in source.Select(f => f.Value).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v => v, StringComparer.OrdinalIgnoreCase))
                _values.Add(new ValueMapRow { Category = category, Source = v, KindText = Loc.T("Flows_ValueKind_" + category) });
        }
        AddValues("Email", findings.Where(f => f.Kind is FindingKind.Email or FindingKind.Mailbox));
        AddValues("Tenant", findings.Where(f => f.Kind == FindingKind.TenantId));
        AddValues("Path", findings.Where(f => f.Kind == FindingKind.FilePath && f.Detail is null));
        foreach (var v in _values)
            if (prevValues.TryGetValue(v.Category + "|" + v.Source, out var pv) && pv.Target.Length > 0) { v.Target = pv.Target; v.Note = pv.Note; }
        _vm.HasValues = _values.Count > 0;

        if (_loaded is not null) ApplyLoadedMapping(_loaded);
        AutoFillFromSignIns();   // after a loaded mapping, so loaded values win
        ShowConnectionStatus();  // refreshes the "flows point at another tenant" warning for the new site list
    }

    // What the scanned flows use: one row per service with counts and how the tool handles it.
    private void ShowResources(IReadOnlyList<FlowScan> scans)
    {
        var rows = FlowResources.Summarize(scans);
        ResourcesGrid.ItemsSource = rows.Select(r => new
        {
            r.Service, r.Flows, r.Operations, r.Detail,
            HandlingText = Loc.T("Flows_Handling_" + r.Handling),
        }).ToList();
        _vm.ResourcesHeaderText = Loc.T("Flows_ResourcesHeaderCount", rows.Count);
    }

    // ── Activity strip ────────────────────────────────────────────────────────

    private void ToggleActivity_Click(object sender, RoutedEventArgs e)
    {
        _vm.IsActivityOpen = !_vm.IsActivityOpen;
        if (_vm.IsActivityOpen) ActivityBox.ScrollToEnd();
    }

    // ── Form names: the flows' forms on the source, the forms in the target ───

    // The flow stores only a form's ID, so its title comes from Microsoft Forms. One action reads both sides:
    //   source: the title of each form the flows use (the account's own forms, then each remaining ID on its own, which
    //           covers a form shared with the account);
    //   target: the forms the target account owns, for the dropdown on each row.
    // Each side reports for itself (what it found, or why not), so a problem on one side never hides the other, and the
    // result is always shown under the button. Then every form whose title matches exactly one target form is matched.
    // Picking a form in a row's dropdown sets that row's form ID. Done here, on the selection itself, rather than relying on the
    // binding alone, so the ID box is always filled the moment a form is picked.
    private void FormCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (sender is System.Windows.Controls.ComboBox { DataContext: FormMapRow row, SelectedItem: TargetFormOption option } && !ReferenceEquals(row.Picked, option))
            row.Picked = option;
    }

    private async void LoadFormNames_Click(object sender, RoutedEventArgs e) => await LoadFormNamesAsync();

    private async Task LoadFormNamesAsync()
    {
        _formNamesAttempted = true;
        await RunAsync(Loc.T("Flows_LoadingFormNames"), async ct =>
        {
            var lines = new List<string>();
            try
            {
                lines.Add(await LoadSourceTitlesAsync(ct));
                lines.Add(await LoadTargetFormsAsync(ct));
                var matched = SuggestFormsByTitle();
                if (matched > 0) lines.Add(Loc.T("Flows_FormsMatchedByName", matched));
            }
            catch (OperationCanceledException)
            {
                lines.Add(Loc.T("Flows_TargetFormsCancelled"));
                throw;
            }
            finally { _vm.TargetFormsText = string.Join("\n", lines); }
        });
    }

    private async Task<string> LoadSourceTitlesAsync(CancellationToken ct)
    {
        var account = _signedInUser(true);
        try
        {
            Log(Loc.T("Flows_LogLoadSourceForms", account));
            var client = new FormsApiClient(_sourceFormsToken);
            var owned = (await client.ListFormsAsync(ct)).GroupBy(f => f.Id).ToDictionary(g => g.Key, g => g.First().Title);
            var found = 0;
            foreach (var row in _forms)
            {
                if (owned.TryGetValue(row.SourceId, out var title) && title.Length > 0) { row.SourceTitle = title; found++; continue; }
                // Not one of the account's own forms (shared with it, or owned by a group): ask for this one by its ID.
                if (await client.GetFormTitleAsync(row.SourceId, ct) is { Length: > 0 } byId) { row.SourceTitle = byId; found++; }
            }
            var text = Loc.T("Flows_SourceTitlesLoaded", found, _forms.Count, account);
            Log(text);
            return text;
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            Log(ex.Message, important: true);
            return Loc.T("Flows_SourceTitlesFailed", ex.Message) + (FormsDialog.HelpFor(ex) is { } h ? " " + h : "");
        }
    }

    private async Task<string> LoadTargetFormsAsync(CancellationToken ct)
    {
        var account = _signedInUser(false);
        try
        {
            Log(Loc.T("Flows_LogLoadTargetForms", account));
            var forms = await new FormsApiClient(_targetFormsToken).ListFormsAsync(ct);
            _targetForms.Clear();
            _targetForms.Add(new TargetFormOption(string.Empty, Loc.T("Flows_FormNone")));
            foreach (var f in forms.OrderBy(f => f.Title, StringComparer.CurrentCultureIgnoreCase))
                _targetForms.Add(new TargetFormOption(f.Id, f.Title.Length == 0 ? Loc.T("Flows_FormUntitled") : f.Title));
            foreach (var row in _forms) row.RefreshPicked();
            var text = forms.Count == 0 ? Loc.T("Flows_TargetFormsNone", account) : Loc.T("Flows_TargetFormsLoaded", forms.Count, account);
            Log(text);
            return text;
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            Log(ex.Message, important: true);
            return Loc.T("Flows_TargetFormsFailed", ex.Message) + (FormsDialog.HelpFor(ex) is { } h ? " " + h : "");
        }
    }

    // A form with no target yet is matched to the target form that has the same title (or the same title plus " (copy)",
    // which is what the Forms tool names a copy), when exactly one does. Never touches a row that already has a target.
    private int SuggestFormsByTitle()
    {
        var real = _targetForms.Where(o => o.Id.Length > 0).ToList();
        var matched = 0;
        foreach (var row in _forms.Where(r => string.IsNullOrWhiteSpace(r.NewId) && r.SourceTitle.Length > 0))
        {
            var hits = real.Where(o => string.Equals(o.Title, row.SourceTitle, StringComparison.OrdinalIgnoreCase)).ToList();
            if (hits.Count != 1) hits = real.Where(o => string.Equals(o.Title, row.SourceTitle + " (copy)", StringComparison.OrdinalIgnoreCase)).ToList();
            if (hits.Count != 1) continue;
            row.NewId = hits[0].Id;
            if (row.Note.Length == 0 || row.StatusLevel == "Ready") { row.Note = Loc.T("Flows_FormNoteMatchedByName"); row.StatusLevel = "Ready"; }   // keep a warning the new ID raised
            matched++;
        }
        return matched;
    }
    // ── Fill in what the sign-ins already tell us ─────────────────────────────

    // Fills EMPTY cells only (never anything you typed or loaded), and says so in the row, so each value can be checked:
    //   - the target tenant ID, from the target sign-in (Multi-tenant; in Single tenant the tenant doesn't change);
    //   - the source account's own address -> the target account's address;
    //   - the email domain, when the two accounts are on different domains and the flows use the source domain.
    private void AutoFillFromSignIns()
    {
        var sourceUser = _signedInUser(true).Trim();
        var targetUser = _signedInUser(false).Trim();

        if (_isMultiTenant && _tenantId(false) is { Length: > 0 } targetTenant)
        {
            foreach (var r in _values.Where(r => r.Category == "Tenant" && string.IsNullOrWhiteSpace(r.Target)
                                                 && !string.Equals(r.Source, targetTenant, StringComparison.OrdinalIgnoreCase)))
            {
                r.Target = targetTenant;
                r.Note = Loc.T("Flows_NoteTenant", _signedInUser(false));
            }
        }

        if (sourceUser.Length > 0 && targetUser.Length > 0 && !string.Equals(sourceUser, targetUser, StringComparison.OrdinalIgnoreCase))
        {
            foreach (var r in _values.Where(r => r.Category == "Email" && string.IsNullOrWhiteSpace(r.Target)
                                                 && string.Equals(r.Source, sourceUser, StringComparison.OrdinalIgnoreCase)))
            {
                r.Target = targetUser;
                r.Note = Loc.T("Flows_NoteAccount");
            }

            var sourceDomain = DomainOf(sourceUser);
            var targetDomain = DomainOf(targetUser);
            if (sourceDomain.Length > 0 && targetDomain.Length > 0 && !sourceDomain.Equals(targetDomain, StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrWhiteSpace(DomainFromBox.Text) && string.IsNullOrWhiteSpace(DomainToBox.Text)
                && _values.Any(r => r.Category == "Email" && DomainOf(r.Source).Equals(sourceDomain, StringComparison.OrdinalIgnoreCase)))
            {
                DomainFromBox.Text = sourceDomain;
                DomainToBox.Text = targetDomain;
                _vm.DomainNote = Loc.T("Flows_NoteDomain", sourceUser, targetUser);
            }
        }
        _vm.HasAutofill = _values.Any(v => v.HasNote) || _vm.HasDomainNote;
    }

    // "Clear values filled in automatically": empties every cell the tool filled from the sign-ins, and the email domain.
    private void ClearAutofill_Click(object sender, RoutedEventArgs e)
    {
        foreach (var r in _values.Where(v => v.HasNote)) r.Target = string.Empty;   // editing a value also clears its note
        if (_vm.HasDomainNote)
        {
            DomainFromBox.Text = string.Empty;
            DomainToBox.Text = string.Empty;
            _vm.DomainNote = string.Empty;
        }
        _vm.HasAutofill = false;
    }

    private static string DomainOf(string email)
    {
        var at = email.LastIndexOf('@');
        return at >= 0 && at < email.Length - 1 ? email[(at + 1)..].Trim() : string.Empty;
    }

    // What to tell the user about the form ID in a row: where it came from, or why it looks wrong.
    private void ShowFormNote(FormMapRow row, FormsMapEntry? fromCopy)
    {
        var (text, level) = FormNote(row, fromCopy);
        row.Note = text;
        row.StatusLevel = level;
    }

    private (string Text, string Level) FormNote(FormMapRow row, FormsMapEntry? fromCopy)
    {
        if (string.IsNullOrWhiteSpace(row.NewId)) return (string.Empty, "Info");
        var id = FlowIdParser.ParseFormId(row.NewId);
        if (id is null) return (Loc.T("Flows_FormNoteBadId"), "Look");
        if (string.Equals(id, row.SourceId, StringComparison.Ordinal)) return (Loc.T("Flows_FormNoteSameAsSource"), "Look");

        // A form's ID carries its tenant. The copy must be in the target's tenant: a different one from the source when
        // copying across tenants, the same one within a tenant.
        if (FlowScanner.TryGetFormTenant(row.SourceId) is { } sourceTenant && FlowScanner.TryGetFormTenant(id) is { } copyTenant)
        {
            if (_isMultiTenant && sourceTenant == copyTenant) return (Loc.T("Flows_FormNoteSameTenant"), "Look");
            if (!_isMultiTenant && sourceTenant != copyTenant) return (Loc.T("Flows_FormNoteOtherTenant"), "Look");
        }
        if (fromCopy is not null && string.Equals(fromCopy.NewId, id, StringComparison.Ordinal))
            return (fromCopy.CopiedAt is { } at ? Loc.T("Flows_FormNoteFromCopy", at.LocalDateTime.ToString("d"), fromCopy.TargetAccount ?? "") : Loc.T("Flows_FormNoteFromCopyOld"), "Ready");
        return (Loc.T("Flows_FormNoteOk"), "Ready");
    }

    // ── Step 2: sites ─────────────────────────────────────────────────────────

    private void FillHost_Click(object sender, RoutedEventArgs e)
    {
        // "contoso.sharepoint.com" without https:// is accepted too.
        var typed = HostBox.Text.Trim();
        if (typed.Length > 0 && !typed.Contains("://")) typed = "https://" + typed;
        if (!Uri.TryCreate(typed, UriKind.Absolute, out var host))
        {
            _vm.StatusText = Loc.T("Flows_BadHost");
            return;
        }

        var emptyBefore = _sites.Count(r => string.IsNullOrWhiteSpace(r.Target));
        var filled = FillTargetsFromHost(host);
        _vm.StatusText = filled > 0 ? Loc.T("Flows_HostFilledCount", filled)
            : emptyBefore == 0 ? Loc.T("Flows_HostNothingToFill")
            : Loc.T("Flows_BadHost");
    }

    // Gives every source site that has no target yet the same path on the given tenant address. Targets already there
    // (the connected site's default, or one you typed) are left alone. Returns how many were filled.
    private int FillTargetsFromHost(Uri host)
    {
        var filled = 0;
        foreach (var row in _sites.Where(r => string.IsNullOrWhiteSpace(r.Target)))
            if (Uri.TryCreate(row.Source, UriKind.Absolute, out var src))
            {
                row.Target = TargetHostFor(src, host) + src.AbsolutePath.TrimEnd('/');
                filled++;
            }
        return filled;
    }

    // OneDrive sites live on the tenant's "-my" host (contoso-my.sharepoint.com), so each source keeps its kind of
    // host on the target, whichever of the two the user typed.
    private static string TargetHostFor(Uri source, Uri host)
    {
        const string sp = ".sharepoint.com", my = "-my.sharepoint.com";
        var name = host.Host;
        if (name.EndsWith(sp, StringComparison.OrdinalIgnoreCase))
        {
            var tenant = name.EndsWith(my, StringComparison.OrdinalIgnoreCase) ? name[..^my.Length] : name[..^sp.Length];
            name = tenant + (source.Host.EndsWith(my, StringComparison.OrdinalIgnoreCase) ? my : sp);
        }
        return $"{host.Scheme}://{name}{(host.IsDefaultPort ? "" : ":" + host.Port)}";
    }

    private Dictionary<string, string> SiteMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in _sites.Where(r => !string.IsNullOrWhiteSpace(r.Target)))
            map[FlowMappingResolver.NormalizeSite(r.Source)] = FlowMappingResolver.NormalizeSite(r.Target);
        return map;
    }

    // ── Step 3: lists and columns (runs automatically when the step opens) ────

    private async void MatchAgain_Click(object sender, RoutedEventArgs e) => await ResolveAsync();

    private async Task ResolveAsync()
    {
        var siteMap = SiteMap();
        if (_listIds.Count == 0)
        {
            ShowMatchSummary(Loc.T("Flows_NoListsFound"), "Info");
            return;
        }
        if (siteMap.Count == 0)
        {
            ShowMatchSummary(Loc.T("Flows_NoSiteTargets"), "Look");
            return;
        }

        await RunAsync(Loc.T("Flows_Resolving"), async ct =>
        {
            _vm.ClearNotes();
            Log(Loc.T("Flows_LogResolveStart", _listIds.Count));
            _lookup = new SharePointFlowLookup(_sp, _prepareSite, AppendLog);

            // Targets the user chose or typed win over name matching (this is how lists with different names are mapped),
            // and so do the ones from a loaded mapping.
            var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (_loaded is not null)
                foreach (var (source, target) in _loaded.Lists)
                    if (FlowIdParser.ParseListId(target) is { } loadedId) overrides[source] = loadedId;
            foreach (var c in _cards.Where(c => c.UserEdited))
            {
                // A list the user picked lives on the target site it was picked from; if that site's target changed since,
                // the choice no longer applies and the list is matched by name again.
                var stillSameSite = c.SourceSite.Length == 0 || string.Equals(
                    FlowMappingResolver.NormalizeSite(siteMap.GetValueOrDefault(FlowMappingResolver.NormalizeSite(c.SourceSite)) ?? ""),
                    FlowMappingResolver.NormalizeSite(c.TargetSite), StringComparison.OrdinalIgnoreCase);
                if (stillSameSite && FlowIdParser.ParseListId(c.TargetListId) is { } id) overrides[c.SourceListId] = id;
            }

            // Column choices the user made are put back after the cards are rebuilt.
            var kept = _cards.SelectMany(c => c.Columns.Where(r => r.UserEdited).Select(r => (c.SourceListId, r.SourceColumn, r.Selected))).ToList();

            _lastReadMap = _sites.Where(r => !string.IsNullOrWhiteSpace(r.ReadFrom) && !string.Equals(FlowMappingResolver.NormalizeSite(r.ReadFrom), FlowMappingResolver.NormalizeSite(r.Source), StringComparison.OrdinalIgnoreCase))
                .ToDictionary(r => FlowMappingResolver.NormalizeSite(r.Source), r => FlowMappingResolver.NormalizeSite(r.ReadFrom), StringComparer.OrdinalIgnoreCase);
            _resolved = await FlowMappingResolver.ResolveAsync(_scans, siteMap, _lookup, ct, overrides, line => Log(line), _lastReadMap);
            ShowConnectionStatus();
            AutoFillFromSignIns();
            BuildCards(_resolved, kept);

            foreach (var n in _resolved.Notes) AppendLog(n);
            Log(Loc.T("Flows_LogResolveDone", _vm.MatchSummary.Text));
        });
    }

    private void ShowMatchSummary(string text, string level)
    {
        _vm.MatchSummary.Text = text;
        _vm.MatchSummary.StatusLevel = level;
    }

    // "3 lists · 2 ready · 1 needs a look", recomputed whenever a card changes so it never disagrees with the cards.
    private void UpdateMatchSummary()
    {
        UpdateGates();
        if (_cards.Count == 0) { _vm.MatchSummary.Text = string.Empty; return; }
        var look = _cards.Count(c => c.NeedsAttention);
        ShowMatchSummary(look == 0 ? Loc.T("Flows_Summary_AllReady", _cards.Count) : Loc.T("Flows_Summary_Cards", _cards.Count, _cards.Count - look, look),
                         look == 0 ? "Ready" : "Look");
    }

    // The flows that use a list, for the card's "Used by" line.
    private IReadOnlyList<string> FlowsUsing(string listId) => _scans
        .Where(s => s.Findings.Any(f => f.Kind == FindingKind.ListId && string.Equals(f.Value, listId, StringComparison.OrdinalIgnoreCase)))
        .Select(s => s.FlowName).Distinct().ToList();

    // The target lists of one site as dropdown choices, with a leading "(none)". One collection per site, shared by every
    // card on it.
    private ObservableCollection<TargetListOption> TargetOptionsFor(string? targetSite, FlowResolveResult resolved)
    {
        var none = new TargetListOption(string.Empty, Loc.T("Flows_TargetNone"));
        if (string.IsNullOrEmpty(targetSite)) return [none];
        if (!_targetListOptions.TryGetValue(targetSite, out var options)) _targetListOptions[targetSite] = options = [];
        options.Clear();
        options.Add(none);
        foreach (var l in (resolved.TargetListsBySite.GetValueOrDefault(targetSite) ?? []).OrderBy(l => l.Title, StringComparer.CurrentCultureIgnoreCase))
            options.Add(new TargetListOption(l.Id, l.Title));
        return options;
    }

    // One card per list: its name and columns on the source, the target list (the same-named one by default, a dropdown of
    // every list on the target site, or an ID to paste), and the column mapping between the two for the user to check.
    // Cards that need a look come first and start open; cards that matched cleanly start closed.
    private void BuildCards(FlowResolveResult resolved, List<(string ListId, string Column, string Selected)> keptChoices)
    {
        _cards.Clear();
        _targetListOptions.Clear();
        var made = new List<ListCard>();

        string? Preferred(string listId, string column, List<ColumnOption> options)
        {
            var kept = keptChoices.FirstOrDefault(k => string.Equals(k.ListId, listId, StringComparison.OrdinalIgnoreCase) && k.Column == column);
            if (kept.Column is not null && options.Any(o => o.Internal == kept.Selected)) return kept.Selected;
            return LoadedColumnChoice(listId, column, options);
        }

        void Hook(ListCard card, FlowResolveResult r)
        {
            card.Used = r.UsedColumnsBySourceList.GetValueOrDefault(card.SourceListId) ?? new HashSet<string>(StringComparer.Ordinal);
            // When the source list could not be read there are no source columns to list, but the flow itself says which columns
            // it uses: map those, so the user is never stuck with an empty grid.
            card.SourceColumns = r.SourceColumnsBySourceList.GetValueOrDefault(card.SourceListId)
                                 ?? card.Used.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).Select(n => new FlowColumnInfo(n, n)).ToList();
            if (_loaded?.RequiredValues.FirstOrDefault(kv => string.Equals(kv.Key, card.SourceListId, StringComparison.OrdinalIgnoreCase)).Value is { } savedValues)
                card.LoadFixValues(savedValues);
            card.CreateSteps = _scans.SelectMany(s => s.CreateSteps)
                .Where(s => string.Equals(s.ListId, card.SourceListId, StringComparison.OrdinalIgnoreCase)).ToList();
            card.FillColumns(r.TargetColumnsBySourceList.GetValueOrDefault(card.SourceListId) ?? [], Preferred);
            foreach (var row in card.Columns)
                if (keptChoices.Any(k => string.Equals(k.ListId, card.SourceListId, StringComparison.OrdinalIgnoreCase) && k.Column == row.SourceColumn)) row.MarkEdited();
            card.IsExpanded = card.NeedsAttention;
            card.TargetChanged += OnCardTargetChanged;
            card.RecheckRequested += async c => await ReloadCardAsync(c, recheck: true);
            card.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(ListCard.BlockerCount)) UpdateGates(); };
        }

        foreach (var l in resolved.Lists)
        {
            var readSite = _lastReadMap.GetValueOrDefault(FlowMappingResolver.NormalizeSite(l.SourceSite)) ?? l.SourceSite;
            var flows = FlowsUsing(l.SourceListId);
            var card = new ListCard
            {
                SourceSite = l.SourceSite, SourceListId = l.SourceListId,
                UsedBy = string.Join(", ", flows), UsedByCount = flows.Count,
                ReadSite = readSite, SourceHost = HostOf(readSite),
                TargetSite = l.TargetSite ?? string.Empty, TargetHost = HostOf(l.TargetSite ?? string.Empty),
                TargetLists = TargetOptionsFor(l.TargetSite, resolved),
                SourceTitle = l.Title ?? string.Empty,
                SourceReadable = l.Title is not null && l.Status != ListMatchStatus.AlreadyOnTarget,
                AlreadyOnTarget = l.Status == ListMatchStatus.AlreadyOnTarget,
                SourceNote = SourceNoteFor(l),
                TargetBaseNote = l.Status == ListMatchStatus.TargetListNotFound ? Loc.T("Flows_ListStatus_TargetListNotFound") : string.Empty,
            };
            card.SetTarget(l.TargetListId);   // the name match (or the user's earlier choice): not counted as an edit
            Hook(card, resolved);
            made.Add(card);
        }

        // A list the flow reaches through an expression (so no site is stated) can't be looked up: its card says so and takes a
        // pasted target list ID.
        var seen = new HashSet<string>(resolved.Lists.Select(l => l.SourceListId), StringComparer.OrdinalIgnoreCase);
        foreach (var id in _listIds.Where(i => !seen.Contains(i)))
        {
            var flows = FlowsUsing(id);
            var card = new ListCard
            {
                SourceListId = id, UsedBy = string.Join(", ", flows), UsedByCount = flows.Count,
                TargetLists = TargetOptionsFor(null, resolved),
                SourceReadable = false,
                SourceNote = Loc.T("Flows_ListStatus_NoSite"),
            };
            card.SetTarget(OverrideFor(id));
            Hook(card, resolved);
            made.Add(card);
        }

        foreach (var c in made.OrderBy(c => c.NeedsAttention ? 0 : 1).ThenBy(c => c.Title, StringComparer.CurrentCultureIgnoreCase)) _cards.Add(c);
        _vm.HasCards = _cards.Count > 0;
        UpdateMatchSummary();

        string? OverrideFor(string listId) =>
            _loaded?.Lists.FirstOrDefault(kv => string.Equals(kv.Key, listId, StringComparison.OrdinalIgnoreCase)).Value is { Length: > 0 } v
                ? FlowIdParser.ParseListId(v) : null;
    }

    // Why a list's source side is not simply "read fine", in words; empty when it was.
    private static string SourceNoteFor(ListResolution l) => l.Status switch
    {
        ListMatchStatus.AlreadyOnTarget => Loc.T("Flows_SourceNote_AlreadyOnTarget", l.Title ?? string.Empty),
        ListMatchStatus.SourceListNotFound => Loc.T("Flows_ListStatus_SourceListNotFound"),
        ListMatchStatus.SiteNotMapped => Loc.T("Flows_ListStatus_SiteNotMapped"),
        _ => l.Title is null ? Loc.T("Flows_ListStatus_SourceListNotFound") : string.Empty,
    };

    private async void OnCardTargetChanged(ListCard card) => await ReloadCardAsync(card);

    // The user chose another target list (or typed an ID): read its columns and match them again.
    private async Task ReloadCardAsync(ListCard card, bool recheck = false)
    {
        var id = FlowIdParser.ParseListId(card.TargetListId);
        if (id is null || _lookup is null || card.TargetSite.Length == 0)
        {
            // No target (or none that can be read): the columns have nothing to map to.
            card.TargetBaseNote = string.IsNullOrWhiteSpace(card.TargetListId) ? Loc.T("Flows_ListStatus_None")
                : id is null ? Loc.T("Flows_BadListId", card.TargetListId) : Loc.T("Flows_ListStatus_NoSite");
            card.FillColumns([], null);
            UpdateMatchSummary();
            return;
        }

        await RunAsync(Loc.T("Flows_LoadingTargetColumns", card.Title), async ct =>
        {
            Log(Loc.T("Flows_LogReloadCard", card.Title, card.TargetSite));
            var targetCols = await _lookup.GetTargetColumnsAsync(card.TargetSite, id, ct);
            card.TargetBaseNote = string.Empty;
            // "Check again" reads the same target list after it was changed in SharePoint: the user's column choices stay.
            var kept = recheck ? card.Columns.Where(r => r.UserEdited).ToDictionary(r => r.SourceColumn, r => r.Selected) : [];
            card.FillColumns(targetCols, (listId, column, options) =>
                kept.TryGetValue(column, out var chosen) && options.Any(o => o.Internal == chosen) ? chosen : LoadedColumnChoice(listId, column, options));
            foreach (var row in card.Columns.Where(r => kept.ContainsKey(r.SourceColumn))) row.MarkEdited();
            Log(card.SummaryText);
        });
        UpdateMatchSummary();
    }

    private string? LoadedColumnChoice(string sourceListId, string sourceColumn, List<ColumnOption> options) =>
        _loadedColumns.TryGetValue(sourceListId, out var perList)
        && perList.TryGetValue(sourceColumn, out var target)
        && options.Any(o => o.Internal == target)
            ? target : null;

    // Opens or closes every list card.
    private void ExpandAll_Click(object sender, RoutedEventArgs e) { foreach (var c in _cards) c.IsExpanded = true; }
    private void CollapseAll_Click(object sender, RoutedEventArgs e) { foreach (var c in _cards) c.IsExpanded = false; }

    // ── Load / save a mapping ─────────────────────────────────────────────────

    private void LoadMapping_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = Loc.T("Flows_FilterJson"), Title = Loc.T("Flows_LoadMapping") };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            _loaded = FlowMapping.Load(dlg.FileName);
        }
        catch (Exception ex) when (ex is IOException or System.Text.Json.JsonException or UnauthorizedAccessException)
        {
            _vm.StatusText = Loc.T("Flows_MappingLoadFailed", ex.Message);
            return;
        }
        _loadedColumns = new(StringComparer.OrdinalIgnoreCase);
        foreach (var (listId, columns) in _loaded.ListColumns) _loadedColumns[listId] = columns;
        if (_scans.Count > 0) ApplyLoadedMapping(_loaded);
        _vm.StatusText = _scans.Count > 0 ? Loc.T("Flows_MappingLoaded", Path.GetFileName(dlg.FileName)) : Loc.T("Flows_MappingLoadedPending", Path.GetFileName(dlg.FileName));
    }

    // Fills the rows with what the loaded mapping says. Only values the flows actually contain have a row, so a mapping
    // saved for a bigger set of flows applies cleanly to a smaller one. Rows already filled in are overwritten: a loaded
    // file is an explicit choice.
    private void ApplyLoadedMapping(FlowMapping m)
    {
        static string? Get(Dictionary<string, string> map, string key) =>
            map.FirstOrDefault(kv => string.Equals(kv.Key.TrimEnd('/'), key.TrimEnd('/'), StringComparison.OrdinalIgnoreCase)).Value is { Length: > 0 } v ? v : null;

        foreach (var r in _sites) if (Get(m.Sites, r.Source) is { } s) r.Target = s;
        foreach (var r in _forms) if (Get(m.Forms, r.SourceId) is { } f) r.NewId = f;
        foreach (var r in _values)
        {
            var map = r.Category switch { "Email" => m.Emails, "Tenant" => m.TenantIds, _ => m.Text };
            if (Get(map, r.Source) is { } v) r.Target = v;
        }
        if (m.EmailDomains.FirstOrDefault(kv => !string.IsNullOrWhiteSpace(kv.Value)) is { Key: not null } d && !string.IsNullOrWhiteSpace(d.Key))
        {
            DomainFromBox.Text = d.Key;
            DomainToBox.Text = d.Value;
        }
        _resolved = null;   // list targets changed: matching runs again when the lists step opens
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    private void Back_Click(object sender, RoutedEventArgs e)
    {
        if (_vm.CurrentStep is < 1 or > 5) return;
        var target = _vm.CurrentStep - 1;
        if (target == 4) PrepareWrite();   // back from the report to Write: refresh what is ready
        GoTo(target);
    }

    private async void Next_Click(object sender, RoutedEventArgs e)
    {
        switch (_vm.CurrentStep)
        {
            case 0:
                GoTo(1);
                break;
            case 1:
                GoTo(2);
                if (_resolved is null) await ResolveAsync();   // automatic matching the first time the step opens
                break;
            case 2:
                AutoFillFromSignIns();   // by now the sign-ins exist; fill the tenant ID, account and domain
                GoTo(3);
                // The form names load by themselves the first time this step opens (the button reloads them).
                if (_forms.Count > 0 && !_formNamesAttempted) await LoadFormNamesAsync();
                break;
            case 3:
                PrepareWrite();
                GoTo(4);
                break;
            case 4:
                if (await RewriteAsync()) GoTo(5);
                break;
        }
    }

    // The footer shows only what is happening right now, so moving to another step clears it.
    private void GoTo(int step)
    {
        _vm.CurrentStep = step;
        _vm.StatusText = string.Empty;
        UpdateGates();
    }

    // Decides whether the current step lets the user go on. The steps that check the flows will work (Sites, Lists & columns,
    // Forms & email) stay closed while something would make the imported flows fail, and say what.
    private void UpdateGates()
    {
        var reason = string.Empty;
        switch (_vm.CurrentStep)
        {
            case 0:
                // A package this tool wrote already holds the target's values, so it has nothing left to move.
                var written = _zips.Count(_rewritten.ContainsKey);
                if (written > 0) reason = Loc.T("Flows_Gate_Written", written);
                break;
            case 1:
                var noTarget = _sites.Count(s => string.IsNullOrWhiteSpace(s.Target));
                if (_vm.IsMultiTenant && _vm.TargetConn.State != "Connected") reason = Loc.T("Flows_Gate_SignInTarget");
                else if (_vm.IsMultiTenant && noTarget > 0) reason = Loc.T("Flows_Gate_Sites", noTarget);
                break;
            case 2:
                var failing = _cards.Count(c => c.HasBlockers);
                if (failing > 0) reason = Loc.T("Flows_Gate_Lists", failing);
                else if (_cards.Count == 0 && _listIds.Count > 0 && !_vm.IsBusy) reason = Loc.T("Flows_Gate_NoCards");
                break;
            case 3:
                // A form ID that is not valid, or points at the wrong tenant, fails. An empty one keeps the source form, which
                // only works inside one tenant.
                var badForms = _forms.Count(f => string.IsNullOrWhiteSpace(f.NewId)
                    ? _vm.IsMultiTenant
                    : FlowIdParser.ParseFormId(f.NewId) is null || f.StatusLevel == "Look");
                if (badForms > 0) reason = Loc.T("Flows_Gate_Forms", badForms);
                break;
        }
        _vm.BlockReason = reason;
    }

    // ── Step 4/5: build the mapping, then write ───────────────────────────────

    private FlowMapping BuildMapping()
    {
        var map = new FlowMapping();
        var siteMap = SiteMap();
        _resolved?.ApplyTo(map, siteMap);
        // The cards hold the final decisions (automatic matches, then the user's edits), so they replace the automatic
        // results. Only columns the flows use are renamed; the rest of a list's columns are shown for context.
        map.Lists.Clear();
        map.Columns.Clear();
        map.ListColumns.Clear();
        foreach (var note in map.ApplyColumnChoices(_cards.SelectMany(c => c.Columns.Where(r => r.IsUsed)
                     .Select(r => new ColumnChoice(c.SourceListId, r.SourceColumn, r.Selected)))))
            AppendLog(note);
        foreach (var (k, v) in siteMap) map.Sites[k] = v;   // user-entered sites always apply, even if no lists resolved
        foreach (var c in _cards.Where(c => !string.IsNullOrWhiteSpace(c.TargetListId)))
        {
            var id = FlowIdParser.ParseListId(c.TargetListId);
            if (id is null) { AppendLog(Loc.T("Flows_BadListId", c.TargetListId)); continue; }
            map.Lists[c.SourceListId] = id;
        }
        map.RequiredValues.Clear();
        foreach (var c in _cards.Where(c => c.FixValues.Count > 0))
            map.RequiredValues[c.SourceListId] = c.FixValues.ToDictionary(kv => kv.Key, kv => kv.Value);
        foreach (var f in _forms.Where(f => !string.IsNullOrWhiteSpace(f.NewId)))
        {
            var id = FlowIdParser.ParseFormId(f.NewId);
            if (id is null) { AppendLog(Loc.T("Flows_BadFormId", f.NewId)); continue; }
            map.Forms[f.SourceId] = id;
        }
        foreach (var v in _values.Where(v => !string.IsNullOrWhiteSpace(v.Target)))
        {
            var target = v.Target.Trim();
            switch (v.Category)
            {
                case "Email": map.Emails[v.Source] = target; break;
                case "Tenant": map.TenantIds[v.Source] = target; break;
                default: map.Text[v.Source] = target; break;
            }
        }
        if (!string.IsNullOrWhiteSpace(DomainFromBox.Text) && !string.IsNullOrWhiteSpace(DomainToBox.Text))
            map.EmailDomains[DomainFromBox.Text.Trim()] = DomainToBox.Text.Trim();
        return map;
    }

    // One row per package with the name its updated copy will be saved as. Starts as the original name; whatever the user
    // types is remembered for that package, also when they go back and forth.
    private void BuildOutputNameRows()
    {
        _vm.OutputNames.Clear();
        foreach (var zip in _zips)
        {
            var row = new OutputNameRow
            {
                Original = Path.GetFileName(zip), FullPath = zip,
                SaveAs = _saveAs.GetValueOrDefault(zip) ?? Path.GetFileNameWithoutExtension(zip),
            };
            row.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(OutputNameRow.SaveAs)) _saveAs[zip] = row.SaveAs; };
            _vm.OutputNames.Add(row);
        }
    }

    private void NamesKeep_Click(object sender, RoutedEventArgs e)
    {
        foreach (var row in _vm.OutputNames) row.SaveAs = Path.GetFileNameWithoutExtension(row.FullPath);
    }

    private void NamesSuffix_Click(object sender, RoutedEventArgs e)
    {
        foreach (var row in _vm.OutputNames) row.SaveAs = Path.GetFileNameWithoutExtension(row.FullPath) + " - migrated";
    }

    // The file names the updated packages will get, one per package in the same order: what was typed in "Save as" (".zip" is
    // optional), with characters Windows does not allow in file names turned into "_", an empty name falling back to the
    // original, and a name used twice getting " (2)", " (3)"... so packages never overwrite each other.
    private List<string> OutputFileNames()
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var names = new List<string>();
        foreach (var zip in _zips)
        {
            var typed = _vm.OutputNames.FirstOrDefault(r => string.Equals(r.FullPath, zip, StringComparison.OrdinalIgnoreCase))?.SaveAs?.Trim() ?? string.Empty;
            if (typed.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)) typed = typed[..^4];
            var baseName = SafeFileName(typed);
            if (baseName.Length == 0) baseName = SafeFileName(Path.GetFileNameWithoutExtension(zip));
            var fileName = baseName + ".zip";
            for (var n = 2; !used.Add(fileName); n++) fileName = $"{baseName} ({n}).zip";
            names.Add(fileName);
        }
        return names;
    }

    private static string SafeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
    }
    // The "Ready to write?" list: where each thing stands, and for anything not settled, the step that settles it.
    private void PrepareWrite()
    {
        BuildOutputNameRows();
        var map = BuildMapping();
        var saved = OutputFileNames();
        var list = new List<ChecklistItem>
        {
            new() { Text = Loc.T("Flows_Chk_Packages", _zips.Count, _vm.FlowCount), StatusLevel = "Ready" },
            new() { Text = Loc.T("Flows_Chk_SavedAs", string.Join(", ", saved.Take(3)) + (saved.Count > 3 ? ", …" : "")), StatusLevel = "Info" },
        };

        var siteTotal = _sites.Count;
        var siteMapped = _sites.Count(s => !string.IsNullOrWhiteSpace(s.Target));
        list.Add(siteTotal == 0 ? new() { Text = Loc.T("Flows_Chk_NoSites"), StatusLevel = "Info" }
            : new() { Text = Loc.T("Flows_Chk_Sites", siteMapped, siteTotal), StatusLevel = siteMapped == siteTotal ? "Ready" : "Look", Step = 1 });

        if (_listIds.Count == 0) list.Add(new() { Text = Loc.T("Flows_Chk_NoLists"), StatusLevel = "Info" });
        else if (_cards.Count == 0) list.Add(new() { Text = Loc.T("Flows_Chk_ListsNotMatched"), StatusLevel = "Look", Step = 2 });
        else
        {
            var withTarget = _cards.Count(c => FlowIdParser.ParseListId(c.TargetListId) is not null);
            list.Add(new() { Text = Loc.T("Flows_Chk_Lists", withTarget, _cards.Count), StatusLevel = withTarget == _cards.Count ? "Ready" : "Look", Step = 2 });
            var toCheck = _cards.Where(c => c.HasTarget).Sum(c => c.Columns.Count(r => r.NeedsLook));
            list.Add(toCheck == 0 ? new() { Text = Loc.T("Flows_Chk_ColumnsReady"), StatusLevel = "Ready" }
                : new() { Text = Loc.T("Flows_Chk_Columns", toCheck), StatusLevel = "Look", Step = 2 });
        }

        var formsMapped = _forms.Count(f => FlowIdParser.ParseFormId(f.NewId) is not null);
        list.Add(_forms.Count == 0 ? new() { Text = Loc.T("Flows_Chk_NoForms"), StatusLevel = "Info" }
            : new() { Text = Loc.T("Flows_Chk_Forms", formsMapped, _forms.Count), StatusLevel = formsMapped == _forms.Count ? "Ready" : "Look", Step = 3 });

        list.Add(map.EmailDomains.Count > 0
            ? new() { Text = Loc.T("Flows_Chk_Email", map.EmailDomains.Keys.First(), map.EmailDomains.Values.First()), StatusLevel = "Info" }
            : new() { Text = Loc.T("Flows_Chk_NoEmail"), StatusLevel = "Info" });
        list.Add(new() { Text = Loc.T("Flows_Chk_Connections"), StatusLevel = "Info" });

        _vm.Checklist.Clear();
        foreach (var item in list) _vm.Checklist.Add(item);
    }

    private void ChecklistGo_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is ChecklistItem { HasStep: true } item) GoTo(item.Step);
    }

    private void ReviewGo_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is ReviewItem { HasStep: true } item) GoTo(item.Step);
    }

    // Asks where to save each updated package (name pre-filled from the "Save as" card). Returns the chosen paths in package order,
    // or null when the user cancels. A path that is one of the originals is refused and asked again.
    private List<string>? AskSaveLocations(string folder)
    {
        var names = OutputFileNames();
        var originals = _zips.Select(Path.GetFullPath).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var chosen = new List<string>();
        for (var i = 0; i < _zips.Count; i++)
        {
            var dir = chosen.Count > 0 ? Path.GetDirectoryName(chosen[^1])! : folder;
            while (true)
            {
                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Title = Loc.T("Flows_SaveAsTitle", Path.GetFileName(_zips[i]), i + 1, _zips.Count),
                    Filter = Loc.T("Flows_SaveAsFilter"),
                    DefaultExt = ".zip",
                    AddExtension = true,
                    OverwritePrompt = true,
                    FileName = names[i],
                    InitialDirectory = Directory.Exists(dir) ? dir : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                };
                if (dlg.ShowDialog(this) != true) return null;
                var path = Path.GetFullPath(dlg.FileName);
                if (originals.Contains(path))
                {
                    MessageBox.Show(this, Loc.T("Flows_OutputSameAsSource"), Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                    continue;
                }
                if (chosen.Contains(path, StringComparer.OrdinalIgnoreCase))
                {
                    MessageBox.Show(this, Loc.T("Flows_SaveAsDuplicate"), Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                    continue;
                }
                chosen.Add(path);
                break;
            }
        }
        return chosen;
    }

    private async Task<bool> RewriteAsync()
    {
        var outDir = _vm.OutputPath.Trim();
        if (outDir.Length == 0) { _vm.StatusText = Loc.T("Flows_NoOutput"); return false; }

        // Every updated package is saved through a Save As dialog, so nothing is written without the user choosing where and
        // under what name, and an original can never be overwritten. Cancel stops before anything is written.
        var dests = AskSaveLocations(outDir);
        if (dests is null) return false;
        outDir = Path.GetDirectoryName(dests[0])!;

        var ok = false;
        await RunAsync(Loc.T("Flows_Rewriting"), async ct =>
        {
            var clock = Stopwatch.StartNew();
            var map = BuildMapping();
            Directory.CreateDirectory(outDir);

            // What the tool could not settle by itself: each is a line in the report's "To do after import" and on the report
            // step, with the step that can fix it.
            var review = new List<ReviewItem>();
            foreach (var s in _sites.Where(s => string.IsNullOrWhiteSpace(s.Target)))
                review.Add(new ReviewItem { Text = Loc.T("Flows_FollowUpSiteUnchanged", s.Source), Step = 1 });
            foreach (var c in _cards.Where(c => FlowIdParser.ParseListId(c.TargetListId) is null))
                review.Add(new ReviewItem { Text = Loc.T("Flows_FollowUpListNoTarget", c.Title, c.SourceListId, c.SourceSite), Step = 2 });
            foreach (var c in _cards.Where(c => c.HasTarget))   // a list with no target is already listed above
                foreach (var r in c.Columns.Where(r => r.IsUsed && string.IsNullOrEmpty(r.Selected) && r.Match == ColumnMatchStatus.NotFoundInTarget))
                    review.Add(new ReviewItem { Text = Loc.T("Flows_FollowUpColumnNoMatch", r.SourceColumn, c.Title), Step = 2 });
            var followUps = review.Select(r => r.Text).ToList();

            Log(Loc.T("Flows_LogWriteStart", _zips.Count, outDir));
            var report = new StringBuilder(Loc.T("Flows_ReportHeader"));
            report.Append(FlowReport.Overview(_scans, followUps));
            var rows = new List<PackageResultRow>();
            int replacements = 0, leftoverValues = 0, failed = 0;

            // The names come from the "Name of the updated package" setting (default: the original names), unique per package.
            // A patched package must never land on top of an original.

            await Task.Run(() =>
            {
                for (var i = 0; i < _zips.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    var zip = _zips[i];
                    var dest = dests[i];
                    var fileName = Path.GetFileName(dest);
                    Log(Loc.T("Flows_LogWriting", fileName));

                    // One unreadable package (it was already reported as skipped by the scan) must not stop the others.
                    List<(string Flow, List<Change> Changes)> results;
                    try { results = FlowRewriter.RewritePackage(zip, dest, map); }
                    // Also an unwritable folder (UnauthorizedAccessException) and a package whose JSON has an unexpected shape
                    // (InvalidOperationException, FormatException from GetValue<T>): those fail THIS package, not the whole run.
                    catch (Exception ex) when (ex is InvalidDataException or IOException or System.Text.Json.JsonException
                                                   or UnauthorizedAccessException or InvalidOperationException or FormatException)
                    {
                        try { File.Delete(dest); } catch (Exception delEx) when (delEx is IOException or UnauthorizedAccessException) { /* partial output: best effort */ }
                        AppendLog(Loc.T("Flows_LogBadPackage", Path.GetFileName(zip), ex.Message));
                        rows.Add(new PackageResultRow { Package = Path.GetFileName(zip), StatusText = "✗ " + Loc.T("Flows_PackageFailed") });
                        failed++;
                        continue;
                    }
                    try { FlowOutputLog.Record(dest); }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* the log is a convenience */ }
                    Log(Loc.T("Flows_LogChecking", fileName));
                    var leftovers = FlowRewriter.FindLeftovers(dest, map);
                    var changed = results.Sum(r => r.Changes.Sum(c => c.Count));
                    var still = leftovers.Select(l => l.Value).Distinct().Count();
                    replacements += changed;
                    leftoverValues += still;
                    rows.Add(new PackageResultRow
                    {
                        Package = fileName, Flows = results.Count, Replacements = changed, StillSource = still,
                        StatusText = still > 0 ? "⚠ " + Loc.T("Flows_PackageStillSource", still)
                                   : changed == 0 ? "– " + Loc.T("Flows_PackageNoChange")
                                   : "✓ " + Loc.T("Flows_PackageOk"),
                    });
                    report.Append(FlowRewriter.ToMarkdown(Path.GetFileNameWithoutExtension(fileName), results, leftovers));
                }
                File.WriteAllText(Path.Combine(outDir, "rewrite-report.md"), report.ToString());
                File.WriteAllText(Path.Combine(outDir, "rewrite-report.html"), FlowReport.ToHtml(report.ToString(), Loc.T("Flows_ReportTitle")), new UTF8Encoding(false));
            }, ct);
            clock.Stop();

            _lastOutput = outDir;
            _lastReport = Path.Combine(outDir, "rewrite-report.html");
            AppendLog(Loc.T("Flows_LogReport", _lastReport));
            FlowsUiState.SaveOutputPath(outDir);

            // Also to review: services the tool doesn't rewrite, values still pointing at the source, packages that failed.
            foreach (var r in FlowResources.Summarize(_scans).Where(r => r.Handling == ResourceHandling.ReviewManually))
                review.Add(new ReviewItem { Text = Loc.T("Flows_Rep_TodoReview", r.Service, r.Detail) });
            if (leftoverValues > 0) review.Add(new ReviewItem { Text = Loc.T("Flows_Review_StillSource", leftoverValues) });
            if (failed > 0) review.Add(new ReviewItem { Text = Loc.T("Flows_Review_PackagesFailed", failed) });

            _vm.PackageResults.Clear();
            foreach (var r in rows) _vm.PackageResults.Add(r);
            _vm.Finish(rows.Count - failed, replacements, review, clock.Elapsed, outDir);
            ok = true;
        });
        return ok;
    }

    private void SaveMapping_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog { Filter = Loc.T("Flows_FilterJson"), FileName = "flow-mapping.json" };
        if (dlg.ShowDialog(this) != true) return;
        BuildMapping().Save(dlg.FileName);
        _vm.StatusText = Loc.T("Flows_MappingSaved", dlg.FileName);
    }

    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog { Title = Loc.T("Flows_OutputFolder") };
        if (dlg.ShowDialog(this) == true) _vm.OutputPath = dlg.FolderName;
    }

    private void OpenOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = _vm.OutputPath.Trim();
        if (folder.Length > 0 && Directory.Exists(folder))
            Process.Start(new ProcessStartInfo(folder) { UseShellExecute = true });
        else
            _vm.StatusText = Loc.T("Flows_OutputFolderMissing");
    }

    // ── Step 6: report ────────────────────────────────────────────────────────

    private void OpenOutput_Click(object sender, RoutedEventArgs e)
    {
        if (_lastOutput is not null && Directory.Exists(_lastOutput))
            Process.Start(new ProcessStartInfo(_lastOutput) { UseShellExecute = true });
    }

    private void OpenReport_Click(object sender, RoutedEventArgs e)
    {
        if (_lastReport is not null && File.Exists(_lastReport))
            Process.Start(new ProcessStartInfo(_lastReport) { UseShellExecute = true });
    }

    private void OpenPowerAutomate_Click(object sender, RoutedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo("https://make.powerautomate.com/") { UseShellExecute = true }); }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException) { /* no default browser */ }
    }

    private void StartOver_Click(object sender, RoutedEventArgs e) => ResetAll();

    private void ResetAll()
    {
        _zips.Clear();
        _zipScans = new(StringComparer.OrdinalIgnoreCase);
        _saveAs.Clear();
        _formNamesAttempted = false;
        _vm.OutputNames.Clear();
        _rewritten.Clear();
        _scans = [];
        _resolved = null;
        _sites.Clear(); _forms.Clear(); _cards.Clear(); _listIds.Clear(); _targetListOptions.Clear(); _lookup = null; _values.Clear();
        _lastReadMap = new(StringComparer.OrdinalIgnoreCase);
        _loaded = null;
        _loadedColumns = new(StringComparer.OrdinalIgnoreCase);
        _vm.Reset();
        _vm.PackageWarning = string.Empty;
        DomainFromBox.Text = string.Empty;   // the previous batch's email-domain rewrite must not carry over
        DomainToBox.Text = string.Empty;
        ResourcesGrid.ItemsSource = null;
        UpdateGates();
    }

    // ── Plumbing ──────────────────────────────────────────────────────────────

    private async Task RunAsync(string busyText, Func<CancellationToken, Task> work)
    {
        _cts = new CancellationTokenSource();
        _vm.IsBusy = true;
        _vm.StatusText = busyText;
        try { await work(_cts.Token); }
        catch (OperationCanceledException)
        {
            _vm.StatusText = Loc.T("Flows_Cancelled");
            Log(Loc.T("Flows_LogCancelled"), important: true);
        }
        catch (Exception ex)
        {
            _vm.StatusText = Loc.T("Flows_Error", ex.Message);
            AppendLog(ex.Message);
            Log(Loc.T("Flows_LogFailed", ex.Message), important: true);
        }
        finally
        {
            _vm.IsBusy = false;
            // The footer shows what is happening right now; once it is done, don't leave the busy text behind.
            if (_vm.StatusText == busyText) _vm.StatusText = string.Empty;
        }
    }

    private void CancelRun_Click(object sender, RoutedEventArgs e)
    {
        if (_cts is { IsCancellationRequested: false })
        {
            _cts.Cancel();
            _vm.StatusText = Loc.T("Flows_Cancelling");   // a sign-in window can take a moment to close
        }
    }

    // The activity strip: plain-language lines saying what the tool is doing right now and why. A failure or cancellation
    // (important) opens the strip so it cannot be missed. Safe to call from any thread.
    private void Log(string line, bool important = false)
    {
        if (!Dispatcher.CheckAccess()) { Dispatcher.Invoke(() => Log(line, important)); return; }
        _vm.AddActivity(line, important);
        if (_vm.IsActivityOpen) ActivityBox.ScrollToEnd();
    }

    // Messages worth showing (a list that couldn't be read, a column mapped two ways...) appear in a notes box on the
    // matching and write steps. Repeats are collapsed so building the mapping twice doesn't duplicate them.
    private void AppendLog(string line)
    {
        if (!Dispatcher.CheckAccess()) { Dispatcher.Invoke(() => AppendLog(line)); return; }
        _vm.AddNote(line);
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    protected override void OnClosing(CancelEventArgs e)
    {
        // Close also answers to Esc: do not abandon a write, scan or sign-in that is still running without asking.
        if (_vm.IsBusy && MessageBox.Show(this, Loc.T("Flows_ClosePrompt"), Loc.T("Flows_CloseTitle"),
                MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) != MessageBoxResult.Yes)
        {
            e.Cancel = true;
            return;
        }
        _cts?.Cancel();
        base.OnClosing(e);
    }
}
