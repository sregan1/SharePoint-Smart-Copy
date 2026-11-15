using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json.Nodes;
using System.Windows;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Services;
using SharePointSmartCopy.Workflows;

namespace SharePointSmartCopy.Dialogs;

// Live copy: pick flows in a source Power Automate environment and create them in a target environment, instead of
// exporting and importing packages. The flows go through the same scan, mapping steps and rewriter as package flows; only
// where they come from (the Flow API) and where they go (created stopped in the target) differ.
public partial class FlowsDialog
{
    private PowerAutomateService _srcPa = null!;
    private PowerAutomateService _dstPa = null!;
    // The flows fetched so far with their full definition, by flow id. These are the originals: every write works on a copy.
    private readonly Dictionary<string, LiveFlow> _liveFetched = new(StringComparer.OrdinalIgnoreCase);
    private bool _loadingEnvironments;   // setting the pickers during a load must not start a second flow load
    // Dataverse token for an environment's org URL: (isSource, orgUrl, ct). Null when the host does not provide one; the copies
    // are then created outside any solution.
    private Func<bool, string, CancellationToken, Task<string>>? _dataverseToken;
    private string? _solutionsLoadedFor;

    private void InitLive(Func<bool, CancellationToken, Task<string>> flowToken, Func<bool, CancellationToken, Task<string>> powerAppsToken)
    {
        _srcPa = new PowerAutomateService(ct => flowToken(true, ct), ct => powerAppsToken(true, ct));
        _dstPa = new PowerAutomateService(ct => flowToken(false, ct), ct => powerAppsToken(false, ct));
        _vm.PropertyChanged += OnLiveViewModelChanged;
    }

    private void ResetLive()
    {
        _liveFetched.Clear();
        // Start over clears the flow list but keeps the chosen environment, and choosing the same one again raises no change:
        // list its flows again so the user is not left with an empty list.
        if (_vm.IsLiveSource && _vm.SourceEnv is not null && _vm.LiveEnvironmentsLoaded)
            Dispatcher.BeginInvoke(new Action(async () => await RunAsync(Loc.T("Flows_Live_LoadingFlows"), LoadFlowsAsync)));
    }

    private async void OnLiveViewModelChanged(object? sender, PropertyChangedEventArgs e)
    {
        // Choosing another source environment lists its flows.
        if (e.PropertyName == nameof(FlowsDialogViewModel.SourceEnv) && !_loadingEnvironments && _vm.SourceEnv is not null)
            await RunAsync(Loc.T("Flows_Live_LoadingFlows"), LoadFlowsAsync);
        // Another target environment has other solutions: list them again when the Write step opens.
        if (e.PropertyName == nameof(FlowsDialogViewModel.TargetEnv))
        {
            _solutionsLoadedFor = null;
            _vm.TargetSolutions.Clear();
            _vm.TargetSolution = null;
            _vm.HasTargetSolutions = false;
            _vm.SolutionsNote = string.Empty;
        }
    }

    // Maps a service error to text the user can act on (the service's own text stays in the activity log).
    private static string FlowErrorText(FlowApiException ex) => ex.Kind switch
    {
        FlowApiError.ConsentMissing => Loc.T("Flows_Live_ErrConsent"),
        FlowApiError.DataverseRole => Loc.T("Flows_Live_ErrRole", ex.MissingPrivilege ?? ""),
        FlowApiError.Forbidden => Loc.T("Flows_Live_ErrForbidden"),
        FlowApiError.InvalidFlow => Loc.T("Flows_Live_ErrInvalid", Brief(ex.Detail)),
        FlowApiError.NotFound => Loc.T("Flows_Live_ErrNotFound"),
        _ => Loc.T("Flows_Live_ErrOther", Brief(ex.Detail)),
    };

    private static string Brief(string text) => text.Length <= 400 ? text : text[..400] + "…";

    // ── Packages step ─────────────────────────────────────────────────────────

    private void SourceModeLive_Click(object sender, RoutedEventArgs e) => _vm.SourceMode = "Live";
    private void SourceModeZip_Click(object sender, RoutedEventArgs e) => _vm.SourceMode = "Zip";

    private async void LoadEnvironments_Click(object sender, RoutedEventArgs e)
    {
        await RunAsync(Loc.T("Flows_Live_LoadingEnvs"), async ct =>
        {
            _loadingEnvironments = true;
            try
            {
                var src = await _srcPa.ListEnvironmentsAsync(ct);
                // One tenant: the target list is the same account's environments. Two tenants: the target is a different
                // sign-in, so its environments load only once it is connected ("Sign in to target"). Asking the target
                // registration for a token before that would reuse the browser's source session and be rejected
                // (the source account does not exist in the target tenant).
                var dst = !_vm.IsMultiTenant ? src
                        : _vm.TargetConn.State == "Connected" ? await _dstPa.ListEnvironmentsAsync(ct)
                        : new List<FlowEnvironment>();
                _vm.SourceEnvironments.Clear(); foreach (var env in src) _vm.SourceEnvironments.Add(env);
                _vm.TargetEnvironments.Clear(); foreach (var env in dst) _vm.TargetEnvironments.Add(env);
                _vm.LiveEnvironmentsLoaded = true;
                _vm.SourceEnv = src.FirstOrDefault(x => x.IsDefault) ?? src.FirstOrDefault();
                _vm.TargetEnv = dst.FirstOrDefault(x => x.IsDefault) ?? dst.FirstOrDefault();
                // The target account is connected but lists nothing: say so (an empty box alone looks like a fault).
                if (dst.Count == 0 && (!_vm.IsMultiTenant || _vm.TargetConn.State == "Connected")) _vm.LiveStatusText = Loc.T("Flows_Live_NoTargetEnvs");
                Log(Loc.T("Flows_Live_LogEnvs", src.Count, dst.Count));
            }
            catch (FlowApiException ex) { throw new InvalidOperationException(FlowErrorText(ex), ex); }
            finally { _loadingEnvironments = false; }
            await LoadFlowsAsync(ct);
        });
    }

    // Two tenants: signs in to the target tenant (the same sign-in as the Sites step, which also pins the chosen account), then
    // lists the target account's environments.
    private async void LiveSignInTarget_Click(object sender, RoutedEventArgs e)
    {
        await ConnectSideAsync(isSource: false);
        if (_vm.TargetConn.State != "Connected") return;
        await RunAsync(Loc.T("Flows_Live_LoadingEnvs"), async ct =>
        {
            try
            {
                var dst = await _dstPa.ListEnvironmentsAsync(ct);
                _vm.TargetEnvironments.Clear(); foreach (var env in dst) _vm.TargetEnvironments.Add(env);
                _vm.TargetEnv = dst.FirstOrDefault(x => x.IsDefault) ?? dst.FirstOrDefault();
                if (dst.Count == 0) _vm.LiveStatusText = Loc.T("Flows_Live_NoTargetEnvs");
                Log(Loc.T("Flows_Live_LogEnvs", _vm.SourceEnvironments.Count, dst.Count));
            }
            catch (FlowApiException ex) { throw new InvalidOperationException(FlowErrorText(ex), ex); }
        });
    }

    private async Task LoadFlowsAsync(CancellationToken ct)
    {
        var env = _vm.SourceEnv;
        if (env is null) return;
        _vm.LiveFlows.Clear();
        _liveFetched.Clear();
        _vm.LiveStatusText = string.Empty;
        try
        {
            var flows = await _srcPa.ListFlowsAsync(env.Name, ct);
            foreach (var f in flows.OrderBy(f => f.DisplayName, StringComparer.CurrentCultureIgnoreCase))
            {
                var row = new LiveFlowRow { Name = f.Name, DisplayName = f.DisplayName, State = f.State };
                row.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(LiveFlowRow.IsSelected)) _vm.RaiseLiveSelectionChanged(); };
                _vm.LiveFlows.Add(row);
            }
            _vm.LiveStatusText = flows.Count == 0 ? Loc.T("Flows_Live_NoFlows", env.DisplayName) : Loc.T("Flows_Live_FlowCount", flows.Count, env.DisplayName);
            Log(_vm.LiveStatusText);
        }
        catch (FlowApiException ex) { throw new InvalidOperationException(FlowErrorText(ex), ex); }
    }

    private void LiveSelectAll_Click(object sender, RoutedEventArgs e) { foreach (var f in _vm.LiveFlows) f.IsSelected = true; }
    private void LiveSelectNone_Click(object sender, RoutedEventArgs e) { foreach (var f in _vm.LiveFlows) f.IsSelected = false; }

    // Fetches the ticked flows (definition included) and scans them like package flows. Runs when Next is pressed on the
    // Packages step. Returns false when nothing could be scanned (the status line says why).
    private async Task<bool> ScanLiveAsync()
    {
        var ok = false;
        var picked = _vm.LiveFlows.Where(f => f.IsSelected).ToList();
        var source = _vm.SourceEnv;
        if (source is null || picked.Count == 0) return false;
        await RunAsync(Loc.T("Flows_Scanning"), async ct =>
        {
            var scans = new List<FlowScan>();
            Log(Loc.T("Flows_Live_LogFetching", picked.Count, source.DisplayName));
            foreach (var row in picked)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    if (!_liveFetched.TryGetValue(row.Name, out var full))
                    {
                        full = await _srcPa.GetFlowAsync(source.Name, row.Name, ct);
                        _liveFetched[row.Name] = full;
                    }
                    // Scan a copy so nothing later changes what was read.
                    var copy = JsonNode.Parse(full.Json!.ToJsonString())!.AsObject();
                    scans.Add(FlowScanner.ScanFlow(source.DisplayName, copy, false, row.DisplayName));
                }
                catch (FlowApiException ex)
                {
                    Log(Loc.T("Flows_Live_LogFlowFailed", row.DisplayName, FlowErrorText(ex)), important: true);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    Log(Loc.T("Flows_Live_LogFlowFailed", row.DisplayName, ex.Message), important: true);
                }
            }
            ApplyScan(scans, new(StringComparer.OrdinalIgnoreCase) { [source.DisplayName] = scans }, () => _vm.PackageRows.Clear());
            ok = scans.Count > 0;
            if (!ok) _vm.StatusText = Loc.T("Flows_Live_NothingScanned");
        });
        return ok;
    }

    // ── Write step ────────────────────────────────────────────────────────────

    // Creates every ticked flow in the target environment, stopped, then writes the same report as the package path.
    private async Task<bool> CreateLiveAsync()
    {
        var outDir = _vm.OutputPath.Trim();   // the report goes here
        if (outDir.Length == 0) { _vm.StatusText = Loc.T("Flows_Live_NoOutput"); return false; }
        var source = _vm.SourceEnv;
        var target = _vm.TargetEnv;
        if (source is null || target is null) { _vm.StatusText = Loc.T("Flows_Live_NoTargetEnv"); return false; }
        var picked = _vm.LiveFlows.Where(f => f.IsSelected && _liveFetched.ContainsKey(f.Name)).ToList();
        // Flows ticked after the scan were never read or mapped: they are not created, and the report says so.
        var notScanned = _vm.LiveFlows.Count(f => f.IsSelected && !_liveFetched.ContainsKey(f.Name));
        if (string.Equals(source.Name, target.Name, StringComparison.OrdinalIgnoreCase)
            && MessageBox.Show(this, Loc.T("Flows_Live_SameEnvBody", target.DisplayName), Title, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes)
            return false;

        var ok = false;
        await RunAsync(Loc.T("Flows_Live_Creating"), async ct =>
        {
            var clock = Stopwatch.StartNew();
            var map = BuildMapping();
            Directory.CreateDirectory(outDir);
            var review = BuildReviewItems();
            var followUps = review.Select(r => r.Text).ToList();

            Log(Loc.T("Flows_Live_LogCreateStart", picked.Count, target.DisplayName));
            var report = new StringBuilder(Loc.T("Flows_ReportHeader"));
            report.Append(FlowReport.Overview(_scans, followUps));
            var rows = new List<PackageResultRow>();
            var results = new List<(string Flow, List<Change> Changes)>();
            var allLeftovers = new List<Finding>();
            int replacements = 0, failed = 0, unbound = 0, globalColumnRenames = 0, omittedInQueries = 0;
            var connections = new Dictionary<string, List<FlowConnection>?>(StringComparer.OrdinalIgnoreCase);
            var bindUnavailable = false;
            var alreadyCreated = 0;
            var boundLog = new List<string>();   // "flow: connector -> connection", for the report
            var cancelled = false;
            var stoppedEarly = false;   // a missing role / permission fails every remaining flow the same way, so the loop stops

            // A cancel (or closing the window) stops the loop, but the report is still written with what was created so far.
            try
            {
            foreach (var row in picked)
            {
                ct.ThrowIfCancellationRequested();
                var original = _liveFetched[row.Name];
                // Created on an earlier press of Write (another flow failed then): never create it twice.
                if (row.CreatedId is not null && string.Equals(row.CreatedIn, target.Name, StringComparison.OrdinalIgnoreCase))
                {
                    alreadyCreated++;
                    rows.Add(new PackageResultRow { Package = original.DisplayName, Flows = 1, StatusText = "✓ " + Loc.T("Flows_Live_AlreadyCreated") });
                    continue;
                }
                var json = JsonNode.Parse(original.Json!.ToJsonString())!.AsObject();
                try
                {
                    var changes = await Task.Run(() => FlowRewriter.Apply(json, map), ct);
                    var leftovers = FlowRewriter.FindLeftovers(new[] { (original.DisplayName, json) }, map);

                    var notBound = 0;
                    string newId;
                    var solution = _vm.TargetSolution is { UniqueName.Length: > 0 } chosen && target.DataverseUrl is not null && _dataverseToken is not null ? chosen : null;
                    if (solution is not null)
                    {
                        // Inside a solution: the flow's connections go through Dataverse connection references of the target.
                        async Task<string?> ConnectionFor(string connector)
                        {
                            try { return (await _dstPa.ListConnectionsAsync(target.Name, connector, ct)).FirstOrDefault(c => c.IsConnected)?.Name; }
                            catch (FlowApiException ex) { Log(Loc.T("Flows_Live_LogBindUnavailable", FlowErrorText(ex)), important: true); return null; }
                        }
                        (newId, notBound) = await TargetDataverse(target).CreateSolutionFlowAsync(solution, json, ConnectionFor, _vm.BindConnections, ct);
                    }
                    else
                    {
                    if (_vm.BindConnections && !bindUnavailable)
                    {
                        try { notBound = await BindConnectionsAsync(json, target.Name, connections, ct, boundLog, original.DisplayName); }
                        catch (FlowApiException ex)
                        {
                            // Binding is a convenience: a problem listing connections must not cost the flow. Without the Power Apps
                            // permission no flow can be bound, so stop trying; any other error affects only this flow.
                            if (ex.Kind is FlowApiError.ConsentMissing or FlowApiError.Forbidden) bindUnavailable = true;
                            Log(Loc.T("Flows_Live_LogBindUnavailable", FlowErrorText(ex)), important: true);
                            notBound = json["properties"]?["connectionReferences"]?.AsObject().Count ?? 0;
                        }
                    }
                    else notBound = json["properties"]?["connectionReferences"]?.AsObject().Count ?? 0;

                    newId = await _dstPa.CreateFlowAsync(target.Name, json, null, ct);
                    }
                    row.CreatedIn = target.Name;
                    row.CreatedId = newId;

                    unbound += notBound;
                    var changed = changes.Sum(c => c.Count);
                    var still = leftovers.Select(l => l.Value).Distinct().Count();
                    replacements += changed;
                    globalColumnRenames += changes.Where(c => c.Category == FlowRewriter.GlobalColumnCategory).Sum(c => c.Count);
                    omittedInQueries += changes.Where(c => c.Category == FlowRewriter.OmittedQueryCategory).Sum(c => c.Count);
                    results.Add((original.DisplayName, changes));
                    allLeftovers.AddRange(leftovers);
                    rows.Add(new PackageResultRow
                    {
                        Package = original.DisplayName, Flows = 1, Replacements = changed, StillSource = still,
                        StatusText = still > 0 ? "⚠ " + Loc.T("Flows_PackageStillSource", still)
                                   : notBound > 0 ? "⚠ " + Loc.T("Flows_Live_CreatedUnbound", notBound)
                                   : solution is not null ? "✓ " + Loc.T("Flows_Live_CreatedInSolution", solution.DisplayName)
                                   : "✓ " + Loc.T("Flows_Live_Created"),
                    });
                    Log(Loc.T("Flows_Live_LogCreated", original.DisplayName));
                }
                catch (FlowApiException ex)
                {
                    failed++;
                    var why = FlowErrorText(ex);
                    AppendLog($"{original.DisplayName}: {ex.Detail}");
                    Log(Loc.T("Flows_Live_LogFlowFailed", original.DisplayName, why), important: true);
                    rows.Add(new PackageResultRow { Package = original.DisplayName, StatusText = "✗ " + why });
                    // The same missing role fails every remaining flow the same way, so stop instead of repeating the error.
                    if (ex.Kind is FlowApiError.DataverseRole or FlowApiError.ConsentMissing) { stoppedEarly = true; break; }
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    failed++;
                    AppendLog($"{original.DisplayName}: {ex.Message}");
                    rows.Add(new PackageResultRow { Package = original.DisplayName, StatusText = "✗ " + Loc.T("Flows_PackageFailed") });
                }
            }
            }
            catch (OperationCanceledException)
            {
                cancelled = true;
                Log(Loc.T("Flows_Live_LogCancelled", rows.Count - failed), important: true);
            }

            var firstName = SafeFileName(source.DisplayName);
            var reportBase = Path.Combine(outDir, Loc.T("Flows_Live_ReportName", firstName, SafeFileName(target.DisplayName), DateTime.Now.ToString("yyyyMMdd-HHmmss")));
            report.Append(FlowRewriter.ToMarkdown(source.DisplayName, results, allLeftovers));
            if (boundLog.Count > 0)
            {
                report.Append("\n**").Append(Loc.T("Flows_Live_ReportConnections")).Append("**\n\n");
                foreach (var line in boundLog) report.Append("- ").Append(line).Append('\n');
            }
            await Task.Run(() =>
            {
                File.WriteAllText(reportBase + ".md", report.ToString());
                File.WriteAllText(reportBase + ".html", FlowReport.ToHtml(report.ToString(), Loc.T("Flows_ReportTitle")), new UTF8Encoding(false));
            }, CancellationToken.None);   // written even after a cancel, so what was created is on record
            clock.Stop();

            _lastOutput = outDir;
            _lastReport = reportBase + ".html";
            AppendLog(Loc.T("Flows_LogReport", _lastReport));
            FlowsUiState.SaveOutputPath(outDir);

            foreach (var r in FlowResources.Summarize(_scans).Where(r => r.Handling == ResourceHandling.ReviewManually))
                review.Add(new ReviewItem { Text = Loc.T("Flows_Rep_TodoReview", r.Service, r.Detail) });
            var leftoverValues = allLeftovers.Select(l => l.Value).Distinct().Count();
            if (leftoverValues > 0) review.Add(new ReviewItem { Text = Loc.T("Flows_Review_StillSource", leftoverValues) });
            if (globalColumnRenames > 0 && _cards.Count > 1)
                review.Add(new ReviewItem { Text = Loc.T("Flows_Review_GlobalColumns", globalColumnRenames) });
            if (omittedInQueries > 0) review.Add(new ReviewItem { Text = Loc.T("Flows_Review_OmittedInQuery", omittedInQueries), Step = 2 });
            if (unbound > 0) review.Add(new ReviewItem { Text = Loc.T("Flows_Live_Review_Unbound", unbound) });
            if (failed > 0) review.Add(new ReviewItem { Text = Loc.T("Flows_Live_Review_Failed", failed) });
            if (alreadyCreated > 0) review.Add(new ReviewItem { Text = Loc.T("Flows_Live_Review_AlreadyCreated", alreadyCreated) });
            if (notScanned > 0) review.Add(new ReviewItem { Text = Loc.T("Flows_Live_Review_NotScanned", notScanned), Step = 0 });
            if (cancelled) review.Add(new ReviewItem { Text = Loc.T("Flows_Live_Review_Cancelled", picked.Count - rows.Count) });
            else if (stoppedEarly && picked.Count > rows.Count) review.Add(new ReviewItem { Text = Loc.T("Flows_Live_Review_NotAttempted", picked.Count - rows.Count) });

            _vm.PackageResults.Clear();
            foreach (var r in rows) _vm.PackageResults.Add(r);
            var created = rows.Count - failed - alreadyCreated;   // made in THIS run (rows skipped as already created do not count)
            if (cancelled)
            {
                // A cancelled run is not "done": stay on the Write step (the report on disk lists what was created so far, and
                // pressing Write again skips those flows). The footer says it was cancelled.
                _vm.StatusText = Loc.T("Flows_Live_LogCancelled", created);
                return;
            }
            if (created == 0 && failed > 0)
            {
                // Nothing was created: stay on the Write step so the cause can be fixed, with the reason in the footer.
                _vm.StatusText = rows.FirstOrDefault()?.StatusText.TrimStart('✗', ' ') ?? Loc.T("Flows_PackageFailed");
                return;
            }
            _vm.Finish(created, replacements, review, clock.Elapsed, target.DisplayName, live: true);
            ok = true;
        });
        return ok;
    }

    private static readonly FlowSolution NoSolution = new(string.Empty, Loc.T("Flows_Live_NoSolution"), null);

    private DataverseFlowService TargetDataverse(FlowEnvironment env) =>
        new(env.DataverseUrl!, ct => _dataverseToken!(false, env.DataverseUrl!, ct));

    // The target's unmanaged solutions, for the "Create in" choice on the Write step. Only for a target environment with
    // Dataverse; a failure (no Dataverse permission or role) leaves "My flows" as the only choice and says why.
    private async Task LoadTargetSolutionsAsync()
    {
        var env = _vm.TargetEnv;
        if (!_vm.IsLiveSource || env?.DataverseUrl is null || _dataverseToken is null)
        {
            _vm.TargetSolutions.Clear();
            _vm.TargetSolution = null;
            _vm.HasTargetSolutions = false;
            return;
        }
        if (_solutionsLoadedFor == env.Name) return;
        await RunAsync(Loc.T("Flows_Live_LoadingSolutions"), async ct =>
        {
            var list = new List<FlowSolution>();
            _vm.SolutionsNote = string.Empty;
            try { list = await TargetDataverse(env).ListSolutionsAsync(ct); }
            catch (FlowApiException ex)
            {
                _vm.SolutionsNote = Loc.T("Flows_Live_SolutionsUnavailable", FlowErrorText(ex));
                Log(_vm.SolutionsNote, important: true);
            }
            catch (Microsoft.Identity.Client.MsalException ex)
            {
                _vm.SolutionsNote = Loc.T("Flows_Live_SolutionsUnavailable", ex.Message);
                Log(_vm.SolutionsNote, important: true);
            }
            _vm.TargetSolutions.Clear();
            _vm.TargetSolutions.Add(NoSolution);
            foreach (var s in list) _vm.TargetSolutions.Add(s);
            _vm.TargetSolution = NoSolution;
            _vm.HasTargetSolutions = true;
            _solutionsLoadedFor = env.Name;
        });
    }

    // Binds each of the flow's connection references to a connected connection of the signed-in user in the target
    // environment (the first one per connector). Returns how many references stayed unbound.
    private async Task<int> BindConnectionsAsync(JsonObject flow, string targetEnv,
        Dictionary<string, List<FlowConnection>?> cache, CancellationToken ct, List<string>? boundLog = null, string flowName = "")
    {
        if (flow["properties"]?["connectionReferences"] is not JsonObject refs) return 0;
        var unbound = 0;
        foreach (var (key, value) in refs.ToList())
        {
            // A solution flow's reference may carry only api.name, not id; both say which connector it is.
            var apiId = value?["id"]?.GetValue<string>() ?? value?["api"]?["name"]?.GetValue<string>() ?? string.Empty;
            var connector = apiId[(apiId.LastIndexOf('/') + 1)..];
            if (connector.Length == 0) { unbound++; continue; }
            if (!cache.TryGetValue(connector, out var list))
                cache[connector] = list = await _dstPa.ListConnectionsAsync(targetEnv, connector, ct);
            var pick = list?.FirstOrDefault(c => c.IsConnected);
            if (pick is null) { unbound++; continue; }
            PowerAutomateService.BindConnection(refs, key, pick);
            var said = Loc.T("Flows_Live_LogBound", flowName, connector, pick.DisplayName);
            boundLog?.Add(said);
            Log(said);
        }
        return unbound;
    }
}
