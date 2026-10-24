using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
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
public partial class FlowsDialog : Window
{
    private readonly SharePointService _sp;
    private readonly Func<string, bool, Task> _prepareSite;
    private readonly string _connectedSource;
    private readonly string _connectedTarget;
    private readonly bool _isMultiTenant;
    private readonly FlowsDialogViewModel _vm = new();
    private readonly ObservableCollection<ValueMapRow> _values = [];
    // A mapping loaded from a file; applied to the rows now and again whenever the packages are rescanned.
    private FlowMapping? _loaded;
    private Dictionary<string, Dictionary<string, string>> _loadedColumns = new(StringComparer.OrdinalIgnoreCase);

    private readonly List<string> _zips = [];
    private List<FlowScan> _scans = [];
    private FlowResolveResult? _resolved;
    private readonly ObservableCollection<SiteRow> _sites = [];
    private readonly ObservableCollection<FormMapRow> _forms = [];
    private readonly ObservableCollection<ColumnMapRow> _columnRows = [];
    private readonly ObservableCollection<ListMapRow> _listRows = [];
    private CancellationTokenSource? _cts;
    private string? _lastOutput;
    private string? _lastReport;

    public FlowsDialog(SharePointService sp, Func<string, bool, Task> prepareSite, string sourceUrl, string targetUrl, bool isMultiTenant)
    {
        _sp = sp;
        _prepareSite = prepareSite;
        _isMultiTenant = isMultiTenant;
        _connectedSource = FlowMappingResolver.NormalizeSite(sourceUrl ?? "");
        _connectedTarget = FlowMappingResolver.NormalizeSite(targetUrl ?? "");
        DataContext = _vm;
        InitializeComponent();
        SitesGrid.ItemsSource = _sites;
        FormsGrid.ItemsSource = _forms;
        ValuesGrid.ItemsSource = _values;
        ColumnsGrid.ItemsSource = _columnRows;
        ListsGrid.ItemsSource = _listRows;
        PackagesGrid.ItemsSource = _vm.PackageResults;
        _vm.OutputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SharePointSmartCopy", "Flows");
        _vm.StatusText = Loc.T("Flows_StatusStart");
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

    private async Task AddAndScanAsync(IEnumerable<string> paths)
    {
        foreach (var p in paths)
            if (!_zips.Contains(p, StringComparer.OrdinalIgnoreCase)) _zips.Add(p);
        _vm.PackagesText = string.Join("; ", _zips.Select(Path.GetFileName));
        await ScanAsync();
    }

    private async Task ScanAsync()
    {
        if (_zips.Count == 0) { _vm.StatusText = Loc.T("Flows_NoPackages"); return; }
        await RunAsync(Loc.T("Flows_Scanning"), async ct =>
        {
            var scans = new List<FlowScan>();
            await Task.Run(() =>
            {
                foreach (var zip in _zips)
                {
                    ct.ThrowIfCancellationRequested();
                    try { scans.AddRange(FlowScanner.ScanPackage(zip)); }
                    catch (Exception ex) when (ex is InvalidDataException or IOException or System.Text.Json.JsonException)
                    {
                        AppendLog(Loc.T("Flows_LogBadPackage", Path.GetFileName(zip), ex.Message));
                    }
                }
            }, ct);
            _scans = scans;
            _resolved = null;
            _columnRows.Clear();
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
            _vm.StatusText = scans.Count == 0 ? Loc.T("Flows_NoFlowsFound") : Loc.T("Flows_ScanDoneNext");
        });
    }

    private void FillRowsFromScan()
    {
        _sites.Clear();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in _scans.SelectMany(s => s.Findings).Where(f => f.Kind == FindingKind.SiteUrl && !f.Value.StartsWith('@')))
        {
            var norm = FlowMappingResolver.NormalizeSite(f.Value);
            if (!seen.Add(norm)) continue;
            // The site the user connected as source maps to the connected target by default.
            var target = _connectedTarget.Length > 0 && string.Equals(norm, _connectedSource, StringComparison.OrdinalIgnoreCase)
                ? _connectedTarget : string.Empty;
            _sites.Add(new SiteRow { Source = norm, Target = target });
        }
        _vm.HasSites = _sites.Count > 0;

        _listRows.Clear();
        foreach (var id in _scans.SelectMany(s => s.Findings).Where(f => f.Kind == FindingKind.ListId && !f.Value.StartsWith('@'))
                     .Select(f => f.Value).Distinct(StringComparer.OrdinalIgnoreCase))
            _listRows.Add(new ListMapRow { SourceListId = id, Title = Loc.T("Flows_NotResolvedYet") });

        _forms.Clear();
        var saved = FormsIdMapStore.Load();
        foreach (var id in _scans.SelectMany(s => s.Findings).Where(f => f.Kind == FindingKind.FormId).Select(f => f.Value).Distinct())
        {
            saved.TryGetValue(id, out var fromCopy);
            var row = new FormMapRow { SourceId = id, NewId = fromCopy?.NewId ?? "" };
            row.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(FormMapRow.NewId)) row.Note = FormNote(row, fromCopy); };
            row.Note = FormNote(row, fromCopy);
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
        _vm.HasValues = _values.Count > 0;

        if (_loaded is not null) ApplyLoadedMapping(_loaded);
    }

    // What to tell the user about the form ID in a row: where it came from, or why it looks wrong.
    private string FormNote(FormMapRow row, FormsMapEntry? fromCopy)
    {
        if (string.IsNullOrWhiteSpace(row.NewId)) return string.Empty;
        var id = FlowIdParser.ParseFormId(row.NewId);
        if (id is null) return Loc.T("Flows_FormNoteBadId");
        if (string.Equals(id, row.SourceId, StringComparison.Ordinal)) return Loc.T("Flows_FormNoteSameAsSource");

        // A form's ID carries its tenant. The copy must be in the target's tenant: a different one from the source when
        // copying across tenants, the same one within a tenant.
        if (FlowScanner.TryGetFormTenant(row.SourceId) is { } sourceTenant && FlowScanner.TryGetFormTenant(id) is { } copyTenant)
        {
            if (_isMultiTenant && sourceTenant == copyTenant) return Loc.T("Flows_FormNoteSameTenant");
            if (!_isMultiTenant && sourceTenant != copyTenant) return Loc.T("Flows_FormNoteOtherTenant");
        }
        if (fromCopy is not null && string.Equals(fromCopy.NewId, id, StringComparison.Ordinal))
            return fromCopy.CopiedAt is { } at ? Loc.T("Flows_FormNoteFromCopy", at.LocalDateTime.ToString("d"), fromCopy.TargetAccount ?? "") : Loc.T("Flows_FormNoteFromCopyOld");
        return string.Empty;
    }

    // What the scanned flows use: one row per service with counts and how the tool handles it.
    private void ShowResources(IReadOnlyList<FlowScan> scans)
    {
        ResourcesGrid.ItemsSource = FlowResources.Summarize(scans).Select(r => new
        {
            r.Service, r.Flows, r.Operations, r.Detail,
            HandlingText = Loc.T("Flows_Handling_" + r.Handling),
        }).ToList();
    }

    // ── Step 2: sites ─────────────────────────────────────────────────────────

    private void FillHost_Click(object sender, RoutedEventArgs e)
    {
        if (!Uri.TryCreate(HostBox.Text.Trim(), UriKind.Absolute, out var host))
        {
            _vm.StatusText = Loc.T("Flows_BadHost");
            return;
        }
        foreach (var row in _sites.Where(r => string.IsNullOrWhiteSpace(r.Target)))
            if (Uri.TryCreate(row.Source, UriKind.Absolute, out var src))
                row.Target = TargetHostFor(src, host) + src.AbsolutePath.TrimEnd('/');
        _vm.StatusText = Loc.T("Flows_HostFilled");
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
        if (_listRows.Count == 0)
        {
            _vm.MatchSummaryText = Loc.T("Flows_NoListsFound");
            _vm.StatusText = Loc.T("Flows_NoListsFound");
            return;
        }
        if (siteMap.Count == 0)
        {
            _vm.MatchSummaryText = Loc.T("Flows_NoSiteTargets");
            _vm.StatusText = Loc.T("Flows_NoSiteTargets");
            return;
        }

        await RunAsync(Loc.T("Flows_Resolving"), async ct =>
        {
            _vm.ClearNotes();
            var lookup = new SharePointFlowLookup(_sp, _prepareSite, AppendLog);
            // Targets the user typed or pasted win over name matching (this is how lists with different names are mapped).
            var overrides = _listRows.Where(r => r.UserEdited)
                .Select(r => (r.SourceListId, Target: FlowIdParser.ParseListId(r.TargetListId)))
                .Where(x => x.Target is not null)
                .ToDictionary(x => x.SourceListId, x => x.Target!, StringComparer.OrdinalIgnoreCase);
            _resolved = await FlowMappingResolver.ResolveAsync(_scans, siteMap, lookup, ct, overrides);

            foreach (var l in _resolved.Lists)
            {
                var row = _listRows.FirstOrDefault(r => string.Equals(r.SourceListId, l.SourceListId, StringComparison.OrdinalIgnoreCase));
                if (row is null) { row = new ListMapRow { SourceListId = l.SourceListId }; _listRows.Add(row); }
                row.SetResolved(l.Title ?? Loc.T("Flows_NotFound"), l.TargetListId, Loc.T("Flows_ListStatus_" + l.Status));
            }
            ShowColumnRows(_resolved);

            foreach (var n in _resolved.Notes) AppendLog(n);
            var matched = _resolved.Lists.Count(l => l.Status is ListMatchStatus.Matched or ListMatchStatus.Manual);
            _vm.MatchSummaryText = Loc.T("Flows_ResolveDone", matched, _resolved.Lists.Count,
                _resolved.Columns.Count(c => c.Status is ColumnMatchStatus.Renamed or ColumnMatchStatus.Similar),
                _resolved.Columns.Count(c => c.Status == ColumnMatchStatus.NotFoundInTarget));
            _vm.StatusText = _vm.MatchSummaryText;
        });
    }

    // One editable row per column the flows use on a matched list. The target dropdown lists that target list's
    // columns, pre-selected by the automapper; "(leave unchanged)" keeps the source name.
    private void ShowColumnRows(FlowResolveResult resolved)
    {
        _columnRows.Clear();
        static int Rank(ColumnMatchStatus s) => s switch
        {
            ColumnMatchStatus.NotFoundInTarget => 0, ColumnMatchStatus.Similar => 1, ColumnMatchStatus.Renamed => 2, _ => 3,
        };
        foreach (var c in resolved.Columns.OrderBy(c => Rank(c.Status)).ThenBy(c => c.ListTitle).ThenBy(c => c.SourceColumn, StringComparer.OrdinalIgnoreCase))
        {
            var options = new List<ColumnOption> { new(string.Empty, Loc.T("Flows_LeaveUnchanged")) };
            if (resolved.TargetColumnsBySourceList.TryGetValue(c.SourceListId, out var targets))
                options.AddRange(targets.OrderBy(t => t.DisplayName, StringComparer.OrdinalIgnoreCase)
                    .Select(t => new ColumnOption(t.InternalName, t.DisplayName == t.InternalName ? t.InternalName : $"{t.DisplayName} ({t.InternalName})")));
            _columnRows.Add(new ColumnMapRow
            {
                SourceListId = c.SourceListId,
                ListTitle = c.ListTitle ?? string.Empty,
                SourceColumn = c.SourceColumn,
                Options = options,
                // A column choice from a loaded mapping wins over the automapper, when that column exists on the target list.
                Selected = LoadedColumnChoice(c.SourceListId, c.SourceColumn, options) ?? c.TargetColumn ?? string.Empty,
                StatusText = Loc.T("Flows_ColumnStatus_" + c.Status),
            });
        }
    }

    private string? LoadedColumnChoice(string sourceListId, string sourceColumn, List<ColumnOption> options) =>
        _loadedColumns.TryGetValue(sourceListId, out var perList)
        && perList.TryGetValue(sourceColumn, out var target)
        && options.Any(o => o.Internal == target)
            ? target : null;

    // ── Load / save a mapping ─────────────────────────────────────────────────

    private void LoadMapping_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "JSON (*.json)|*.json", Title = Loc.T("Flows_LoadMapping") };
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
        foreach (var r in _listRows) if (Get(m.Lists, r.SourceListId) is { } l) r.TargetListId = l;   // typed targets win over name matching
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
        if (_vm.CurrentStep is >= 1 and <= 4) GoTo(_vm.CurrentStep - 1);
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
                GoTo(3);
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

    private void GoTo(int step)
    {
        _vm.CurrentStep = step;
        _vm.StatusText = step switch
        {
            0 => _vm.HasScan ? Loc.T("Flows_ScanDoneNext") : Loc.T("Flows_StatusStart"),
            1 => _vm.HasSites ? Loc.T("Flows_StatusSites") : Loc.T("Flows_NoSitesFound"),
            2 => _vm.MatchSummaryText.Length > 0 ? _vm.MatchSummaryText : Loc.T("Flows_StatusMatch"),
            3 => Loc.T("Flows_StatusFormsEmail"),
            4 => Loc.T("Flows_StatusWrite"),
            _ => _vm.ResultHeadline,
        };
    }

    // ── Step 4/5: build the mapping, then write ───────────────────────────────

    private FlowMapping BuildMapping()
    {
        var map = new FlowMapping();
        var siteMap = SiteMap();
        _resolved?.ApplyTo(map, siteMap);
        // The grid holds the final column decisions (automapped, then edited by the user), so it replaces the
        // automapper's own column result.
        map.Columns.Clear();
        map.ListColumns.Clear();
        foreach (var note in map.ApplyColumnChoices(_columnRows.Select(r => new ColumnChoice(r.SourceListId, r.SourceColumn, r.Selected))))
            AppendLog(note);
        foreach (var (k, v) in siteMap) map.Sites[k] = v;   // user-entered sites always apply, even if no lists resolved
        foreach (var l in _listRows.Where(l => !string.IsNullOrWhiteSpace(l.TargetListId)))
        {
            var id = FlowIdParser.ParseListId(l.TargetListId);
            if (id is null) { AppendLog(Loc.T("Flows_BadListId", l.TargetListId)); continue; }
            map.Lists[l.SourceListId] = id;
        }
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

    private void PrepareWrite()
    {
        var map = BuildMapping();
        _vm.WriteSummaryText = Loc.T("Flows_WriteSummary", _zips.Count, _vm.FlowCount, map.Sites.Count, map.Lists.Count,
            map.ListColumns.Sum(kv => kv.Value.Count), map.Forms.Count);
    }

    private async Task<bool> RewriteAsync()
    {
        var outDir = _vm.OutputPath.Trim();
        if (outDir.Length == 0) { _vm.StatusText = Loc.T("Flows_NoOutput"); return false; }

        var ok = false;
        await RunAsync(Loc.T("Flows_Rewriting"), async ct =>
        {
            var clock = Stopwatch.StartNew();
            var map = BuildMapping();
            Directory.CreateDirectory(outDir);
            // Findings of this dialog that need follow-up outside the tool, for the report's "To do after import".
            var followUps = new List<string>();
            foreach (var l in _resolved?.Lists.Where(l => l.Status is not (ListMatchStatus.Matched or ListMatchStatus.Manual)) ?? [])
                followUps.Add($"List {l.SourceListId} ({l.Title ?? "title unknown"}) on {l.SourceSite} has no target: {l.Status}. Flows that use it still point at the source list.");
            foreach (var c in _columnRows.Where(r => string.IsNullOrEmpty(r.Selected) && r.StatusText == Loc.T("Flows_ColumnStatus_NotFoundInTarget")))
                followUps.Add($"Column '{c.SourceColumn}' on list '{c.ListTitle}' has no match in the target list.");

            var report = new StringBuilder(Loc.T("Flows_ReportHeader"));
            report.Append(FlowReport.Overview(_scans, followUps));
            var rows = new List<PackageResultRow>();
            int replacements = 0, leftoverValues = 0, failed = 0;

            if (_zips.Any(z => string.Equals(Path.GetFullPath(Path.GetDirectoryName(z)!), Path.GetFullPath(outDir), StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException(Loc.T("Flows_OutputSameAsSource"));

            await Task.Run(() =>
            {
                // Packages from different folders can share a file name; each gets its own output file.
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var zip in _zips)
                {
                    ct.ThrowIfCancellationRequested();
                    var name = Path.GetFileNameWithoutExtension(zip);
                    var fileName = name + ".zip";
                    for (var n = 2; !usedNames.Add(fileName); n++) fileName = $"{name} ({n}).zip";
                    var dest = Path.Combine(outDir, fileName);

                    // One unreadable package (it was already reported as skipped by the scan) must not stop the others.
                    List<(string Flow, List<Change> Changes)> results;
                    try { results = FlowRewriter.RewritePackage(zip, dest, map); }
                    catch (Exception ex) when (ex is InvalidDataException or IOException or System.Text.Json.JsonException)
                    {
                        try { File.Delete(dest); } catch (IOException) { /* partial output: best effort */ }
                        AppendLog(Loc.T("Flows_LogBadPackage", Path.GetFileName(zip), ex.Message));
                        rows.Add(new PackageResultRow { Package = Path.GetFileName(zip), StatusText = "✗ " + Loc.T("Flows_PackageFailed") });
                        failed++;
                        continue;
                    }
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
                File.WriteAllText(Path.Combine(outDir, "rewrite-report.html"), FlowReport.ToHtml(report.ToString(), "Flow rewrite report"), new UTF8Encoding(false));
            }, ct);
            clock.Stop();

            _lastOutput = outDir;
            _lastReport = Path.Combine(outDir, "rewrite-report.html");
            AppendLog(Loc.T("Flows_LogReport", _lastReport));

            // To review: services the tool doesn't rewrite, plus lists and columns that never got a target.
            var otherIds = _scans.SelectMany(s => s.Findings).Where(f => f.Kind == FindingKind.OtherResource).Select(f => f.Value).Distinct().Count();
            var unmatchedLists = _resolved?.Lists.Count(l => l.Status is not (ListMatchStatus.Matched or ListMatchStatus.Manual)) ?? 0;
            var unmappedColumns = _columnRows.Count(r => string.IsNullOrEmpty(r.Selected) && r.StatusText == Loc.T("Flows_ColumnStatus_NotFoundInTarget"));

            _vm.PackageResults.Clear();
            foreach (var r in rows) _vm.PackageResults.Add(r);
            _vm.Finish(rows.Count - failed, replacements, leftoverValues, failed + otherIds + unmatchedLists + unmappedColumns, clock.Elapsed, outDir);
            ok = true;
        });
        return ok;
    }

    private void SaveMapping_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog { Filter = "JSON (*.json)|*.json", FileName = "flow-mapping.json" };
        if (dlg.ShowDialog(this) != true) return;
        BuildMapping().Save(dlg.FileName);
        _vm.StatusText = Loc.T("Flows_MappingSaved", dlg.FileName);
    }

    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog { Title = Loc.T("Flows_OutputFolder") };
        if (dlg.ShowDialog(this) == true) _vm.OutputPath = dlg.FolderName;
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

    private void StartOver_Click(object sender, RoutedEventArgs e) => ResetAll();

    private void ResetAll()
    {
        _zips.Clear();
        _scans = [];
        _resolved = null;
        _sites.Clear(); _forms.Clear(); _columnRows.Clear(); _listRows.Clear(); _values.Clear();
        _loaded = null;
        _loadedColumns = new(StringComparer.OrdinalIgnoreCase);
        _vm.Reset();
        ResourcesGrid.ItemsSource = null;
        _vm.StatusText = Loc.T("Flows_StatusStart");
    }

    // ── Plumbing ──────────────────────────────────────────────────────────────

    private async Task RunAsync(string busyText, Func<CancellationToken, Task> work)
    {
        _cts = new CancellationTokenSource();
        _vm.IsBusy = true;
        _vm.StatusText = busyText;
        try { await work(_cts.Token); }
        catch (OperationCanceledException) { _vm.StatusText = Loc.T("Flows_Cancelled"); }
        catch (Exception ex)
        {
            _vm.StatusText = Loc.T("Flows_Error", ex.Message);
            AppendLog(ex.Message);
        }
        finally { _vm.IsBusy = false; }
    }

    private void CancelRun_Click(object sender, RoutedEventArgs e) => _cts?.Cancel();

    // Messages worth showing (a list that couldn't be read, a column mapped two ways...) appear in a notes box on the
    // matching and write steps. Repeats are collapsed so building the mapping twice doesn't duplicate them.
    private void AppendLog(string line)
    {
        if (!Dispatcher.CheckAccess()) { Dispatcher.Invoke(() => AppendLog(line)); return; }
        _vm.AddNote(line);
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        _cts?.Cancel();
        base.OnClosing(e);
    }
}

// One line of the report's package table.
public sealed class PackageResultRow
{
    public string Package { get; init; } = string.Empty;
    public int Flows { get; init; }
    public int Replacements { get; init; }
    public int StillSource { get; init; }
    public string StatusText { get; init; } = string.Empty;
}

// State for the six steps. The dialog's code-behind does the work and sets these; the XAML binds to them.
public partial class FlowsDialogViewModel : ObservableObject
{
    public ObservableCollection<PackageResultRow> PackageResults { get; } = [];

    [ObservableProperty, NotifyPropertyChangedFor(nameof(ShowBack), nameof(ShowNext), nameof(ShowStartOver), nameof(NextButtonText), nameof(CanNext))]
    private int _currentStep;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext))] private bool _isBusy;
    [ObservableProperty] private string _statusText = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasNotes))] private string _notesText = string.Empty;
    public bool HasNotes => NotesText.Length > 0;

    public void AddNote(string line)
    {
        if (string.IsNullOrWhiteSpace(line)) return;
        var existing = NotesText.Split('\n');
        if (!existing.Contains(line)) NotesText = NotesText.Length == 0 ? line : NotesText + "\n" + line;
    }

    public void ClearNotes() => NotesText = string.Empty;

    // step 1
    [ObservableProperty] private string _packagesText = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext))] private bool _hasScan;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext))] private int _flowCount;
    [ObservableProperty] private string _scanSummaryText = string.Empty;
    // step 2
    [ObservableProperty, NotifyPropertyChangedFor(nameof(NoSites))] private bool _hasSites;
    public bool NoSites => HasScan && !HasSites;
    // step 3
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasMatchSummary))] private string _matchSummaryText = string.Empty;
    public bool HasMatchSummary => MatchSummaryText.Length > 0;
    // step 4
    [ObservableProperty, NotifyPropertyChangedFor(nameof(NoForms))] private bool _hasForms;
    public bool NoForms => HasScan && !HasForms;
    [ObservableProperty] private bool _hasValues;
    // step 5
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext))] private string _outputPath = string.Empty;
    [ObservableProperty] private string _writeSummaryText = string.Empty;
    // step 6
    [ObservableProperty] private string _outcome = "Success";
    [ObservableProperty] private string _resultHeadline = string.Empty;
    [ObservableProperty] private string _resultSubline = string.Empty;
    [ObservableProperty] private int _packagesWritten;
    [ObservableProperty] private int _replacements;
    [ObservableProperty] private int _stillSourceCount;
    [ObservableProperty] private int _reviewCount;
    [ObservableProperty] private string _durationText = string.Empty;

    public bool ShowBack => CurrentStep is >= 1 and <= 4;
    public bool ShowNext => CurrentStep <= 4;
    public bool ShowStartOver => CurrentStep == 5;
    public string NextButtonText => CurrentStep == 4 ? Loc.T("Flows_WriteButton") : Loc.T("Forms_Next");
    public bool CanNext => !IsBusy && CurrentStep switch
    {
        0 => HasScan && FlowCount > 0,
        4 => OutputPath.Trim().Length > 0,
        <= 3 => true,
        _ => false,
    };

    partial void OnCurrentStepChanged(int value) => OnPropertyChanged(nameof(CanNext));
    partial void OnHasScanChanged(bool value) { OnPropertyChanged(nameof(NoSites)); OnPropertyChanged(nameof(NoForms)); }

    public void Finish(int packages, int replacements, int stillSource, int review, TimeSpan elapsed, string outDir)
    {
        PackagesWritten = packages;
        Replacements = replacements;
        StillSourceCount = stillSource;
        ReviewCount = review;
        DurationText = elapsed.TotalMinutes >= 1 ? $"{(int)elapsed.TotalMinutes}m {elapsed.Seconds}s" : $"{Math.Max(1, (int)Math.Round(elapsed.TotalSeconds))}s";
        if (stillSource == 0 && review == 0)
        {
            Outcome = "Success";
            ResultHeadline = Loc.T("Flows_ResultSuccess", packages);
            ResultSubline = Loc.T("Flows_ResultSuccessSub", outDir);
        }
        else
        {
            Outcome = "Partial";
            ResultHeadline = Loc.T("Flows_ResultReview", packages);
            ResultSubline = Loc.T("Flows_ResultReviewSub", outDir);
        }
        CurrentStep = 5;
    }

    public void Reset()
    {
        CurrentStep = 0;
        PackagesText = string.Empty; HasScan = false; FlowCount = 0; ScanSummaryText = string.Empty;
        HasSites = false; MatchSummaryText = string.Empty; HasForms = false; HasValues = false; WriteSummaryText = string.Empty;
        PackageResults.Clear();
        ClearNotes();
    }
}

public partial class SiteRow : ObservableObject
{
    public string Source { get; init; } = string.Empty;
    [ObservableProperty] private string _target = string.Empty;
}

public partial class FormMapRow : ObservableObject
{
    public string SourceId { get; init; } = string.Empty;
    [ObservableProperty] private string _newId = string.Empty;
    // Where the ID came from, or why it looks wrong (wrong tenant, not an ID).
    [ObservableProperty] private string _note = string.Empty;
}

// A value in the flows that can be mapped by typing its replacement: an email address, the tenant ID or a library path.
public partial class ValueMapRow : ObservableObject
{
    public string Category { get; init; } = string.Empty;   // Email | Tenant | Path
    public string KindText { get; init; } = string.Empty;
    public string Source { get; init; } = string.Empty;
    [ObservableProperty] private string _target = string.Empty;
}

public sealed record ColumnOption(string Internal, string Label);

public partial class ColumnMapRow : ObservableObject
{
    public string SourceListId { get; init; } = string.Empty;
    public string ListTitle { get; init; } = string.Empty;
    public string SourceColumn { get; init; } = string.Empty;
    public string StatusText { get; init; } = string.Empty;
    public List<ColumnOption> Options { get; init; } = [];
    // Internal name of the chosen target column; empty = leave the source name unchanged.
    [ObservableProperty] private string _selected = string.Empty;
}

public partial class ListMapRow : ObservableObject
{
    private bool _suppressEdit;
    public string SourceListId { get; init; } = string.Empty;
    [ObservableProperty] private string _title = string.Empty;
    [ObservableProperty] private string _statusText = string.Empty;
    [ObservableProperty] private string _targetListId = string.Empty;
    // True once the user typed or pasted a target (as opposed to the automatic name match filling it in).
    public bool UserEdited { get; private set; }

    partial void OnTargetListIdChanged(string value)
    {
        if (!_suppressEdit) UserEdited = !string.IsNullOrWhiteSpace(value);
    }

    // Used by the matcher: fills in what it found without marking the row as edited by the user.
    public void SetResolved(string title, string? targetId, string statusText)
    {
        _suppressEdit = true;
        try
        {
            Title = title;
            StatusText = statusText;
            if (!UserEdited) TargetListId = targetId ?? string.Empty;
        }
        finally { _suppressEdit = false; }
    }
}
