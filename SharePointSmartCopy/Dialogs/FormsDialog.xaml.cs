using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Data;
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.Win32;
using SharePointSmartCopy.Forms;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Dialogs;

// Copies Microsoft Forms in four steps, like the main copy wizard:
//   1 Forms (choose what to copy)  ->  2 Target (sign in to where the copies go)  ->  3 Copy (progress)  ->  4 Results
// Single tenant: source and target are the same account. Multi-tenant: the target needs its own sign-in, which is
// done explicitly on step 2. Uses the unofficial Forms web API; see the warning banner.
public partial class FormsDialog : Window
{
    private readonly AuthService _sourceAuth;
    private readonly AuthService _targetAuth;
    private readonly FormsDialogViewModel _vm;
    private readonly ObservableCollection<FormRow> _rows = [];
    private readonly ICollectionView _view;
    private CancellationTokenSource? _cts;

    public FormsDialog(AuthService sourceAuth, AuthService targetAuth, bool isMultiTenant, string targetRegistrationName = "")
    {
        _sourceAuth = sourceAuth;
        _targetAuth = targetAuth;
        _vm = new FormsDialogViewModel { IsMultiTenant = isMultiTenant };
        _vm.TargetRegistrationText = isMultiTenant
            ? (string.IsNullOrWhiteSpace(targetRegistrationName)
                ? Loc.T("Forms_TargetRegistrationUnnamed")
                : Loc.T("Forms_TargetRegistration", targetRegistrationName))
            : string.Empty;
        DataContext = _vm;
        InitializeComponent();
        _view = CollectionViewSource.GetDefaultView(_rows);
        _view.Filter = o => o is FormRow r && r.Matches(_vm.FilterText);
        _vm.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(FormsDialogViewModel.FilterText)) _view.Refresh(); };
        FormsGrid.ItemsSource = _view;
        ResultsGrid.ItemsSource = _vm.Results;
        _vm.SourceAccountText = Loc.T("Forms_SourceAccount", sourceAuth.FormsUserName ?? sourceAuth.UserName ?? Loc.T("Forms_SignInOnDemand"));
        _vm.StatusText = Loc.T("Forms_StatusStart");
    }

    private FormsApiClient SourceClient() => new((ct, force) => _sourceAuth.GetFormsTokenAsync(ct, force));
    private FormsApiClient TargetClient() => new((ct, force) => _targetAuth.GetFormsTokenAsync(ct, force));

    private IEnumerable<FormRow> Selected => _rows.Where(r => r.IsSelected);

    private string SourceAccount => _sourceAuth.FormsUserName ?? _sourceAuth.UserName ?? string.Empty;
    private string TargetAccount => _vm.IsMultiTenant
        ? _targetAuth.FormsUserName ?? _targetAuth.UserName ?? string.Empty
        : SourceAccount;

    // ── Step 1: Forms ─────────────────────────────────────────────────────────

    private async void Load_Click(object sender, RoutedEventArgs e)
    {
        await RunAsync(Loc.T("Forms_Loading"), async ct =>
        {
            var forms = await SourceClient().ListFormsAsync(ct);
            var copied = FormsIdMapStore.Load();
            _rows.Clear();
            foreach (var f in forms.OrderByDescending(f => f.Modified))
            {
                var row = new FormRow { Id = f.Id, Title = f.Title, Type = f.Type, Modified = f.Modified };
                if (copied.TryGetValue(f.Id, out var entry)) row.CopiedText = CopiedLabel(entry);
                row.PropertyChanged += (_, args) => { if (args.PropertyName == nameof(FormRow.IsSelected)) UpdateSelection(); };
                _rows.Add(row);
            }
            _vm.HasForms = _rows.Count > 0;
            _vm.SourceAccountText = Loc.T("Forms_SourceAccount", SourceAccount);
            _vm.LoadButtonText = Loc.T("Forms_ReloadForms");
            UpdateSelection();
            _vm.StatusText = forms.Count == 0 ? Loc.T("Forms_NoForms") : Loc.T("Forms_Loaded", forms.Count);
        });
    }

    private void UpdateSelection()
    {
        _vm.SelectedCount = _rows.Count(r => r.IsSelected);
        _vm.TotalCount = _rows.Count;
    }

    // "Select all" acts on what the filter box currently shows, so you can filter to a group of forms and take just those.
    private IEnumerable<FormRow> Visible => _view.Cast<FormRow>();
    private void SelectAll_Click(object sender, RoutedEventArgs e) { foreach (var r in Visible) r.IsSelected = true; UpdateSelection(); }
    private void DeselectAll_Click(object sender, RoutedEventArgs e) { foreach (var r in Visible) r.IsSelected = false; UpdateSelection(); }

    private static string CopiedLabel(FormsMapEntry e) =>
        e.CopiedAt is { } at ? Loc.T("Forms_CopiedOn", at.LocalDateTime.ToString("d"), e.TargetAccount ?? "") : Loc.T("Forms_CopiedBefore");

    // True when the saved copy went to the account that would receive this one (or to an unknown account, from an
    // earlier version): a second copy would be a duplicate.
    private bool IsDuplicate(FormRow row, Dictionary<string, FormsMapEntry> map) =>
        map.TryGetValue(row.Id, out var e)
        && (string.IsNullOrEmpty(e.TargetAccount) || string.Equals(e.TargetAccount, TargetAccount, StringComparison.OrdinalIgnoreCase));

    // ── Step 2: Target ────────────────────────────────────────────────────────

    private async void TargetSignIn_Click(object sender, RoutedEventArgs e)
    {
        _vm.TargetHasError = false;
        _vm.TargetConnected = false;
        await RunAsync(Loc.T("Forms_SigningInTarget"), async ct =>
        {
            try
            {
                // Signs in (a browser window may open) and proves the account can use Forms by listing its forms.
                var existing = await TargetClient().ListFormsAsync(ct);
                _vm.TargetConnected = true;
                _vm.TargetStatusText = Loc.T("Forms_TargetConnected", TargetAccount, existing.Count);
                _vm.TargetButtonText = Loc.T("Forms_TargetSignInAgain");
                _vm.StatusText = Loc.T("Forms_TargetReady");
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                _vm.TargetHasError = true;
                _vm.TargetStatusText = Loc.T("Forms_TargetFailed", ex.Message) + (HelpFor(ex) is { } h ? Environment.NewLine + h : "");
                AppendLog(ex.Message);
                _vm.StatusText = Loc.T("Forms_TargetNotConnected");
            }
        });
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    private void Back_Click(object sender, RoutedEventArgs e)
    {
        if (_vm.CurrentStep == 1) GoTo(0);
    }

    private async void Next_Click(object sender, RoutedEventArgs e)
    {
        if (_vm.CurrentStep == 0)
        {
            if (!Selected.Any()) { _vm.StatusText = Loc.T("Forms_NothingSelected"); return; }
            PrepareTargetStep();
            GoTo(1);
        }
        else if (_vm.CurrentStep == 1)
        {
            await CopyAsync();
        }
    }

    private void PrepareTargetStep()
    {
        var count = Selected.Count();
        _vm.SelectedCount = count;
        if (!_vm.IsMultiTenant)
        {
            // Same account on both sides: nothing to sign in to; the source sign-in already proved access.
            _vm.TargetConnected = true;
            _vm.TargetHasError = false;
            _vm.TargetStatusText = Loc.T("Forms_TargetSingleStatus", SourceAccount);
        }
        _vm.TargetAccountName = TargetAccount;
        _vm.SourceAccountName = SourceAccount;
        _vm.StatusText = _vm.TargetConnected ? Loc.T("Forms_TargetReady") : Loc.T("Forms_SignInToContinue");
    }

    private void GoTo(int step)
    {
        _vm.CurrentStep = step;
        if (step is 0) _vm.StatusText = _vm.HasForms ? Loc.T("Forms_Loaded", _rows.Count) : Loc.T("Forms_StatusStart");
    }

    // ── Step 3: Copy ──────────────────────────────────────────────────────────

    private async Task CopyAsync()
    {
        var chosen = Selected.ToList();
        if (chosen.Count == 0) return;
        var suffix = _vm.Suffix;

        // Copying a form again creates a second copy; make sure that is intended.
        var map = FormsIdMapStore.Load();
        var dupes = chosen.Where(r => IsDuplicate(r, map)).ToList();
        if (dupes.Count > 0)
        {
            var names = string.Join(Environment.NewLine, dupes.Take(8).Select(r => "  • " + r.Title)) +
                        (dupes.Count > 8 ? Environment.NewLine + Loc.T("Forms_AndMore", dupes.Count - 8) : "");
            if (MessageBox.Show(this, Loc.T("Forms_DuplicatePrompt", dupes.Count, names), Loc.T("Forms_DuplicateTitle"),
                    MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes)
                return;
        }

        _vm.Results.Clear();
        LogBox.Clear();
        _vm.ProgressMax = chosen.Count;
        _vm.ProgressValue = 0;
        _vm.ProgressText = Loc.T("Forms_ProgressStarting");
        GoTo(2);

        var clock = Stopwatch.StartNew();
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        var cancelled = false;
        _vm.IsBusy = true;
        try
        {
            var source = SourceClient();
            var target = TargetClient();
            for (var i = 0; i < chosen.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var row = chosen[i];
                _vm.ProgressText = Loc.T("Forms_ProgressForm", i + 1, chosen.Count, row.Title);
                AppendLog(Loc.T("Forms_LogCopying", row.Title));

                var copyTitle = row.Title + suffix;
                FormCopyResult result;
                try
                {
                    result = await new FormsCopier(source, target, AppendLog).CopyAsync(row.Id, copyTitle, ct);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    AppendLog(ex.Message);
                    result = new FormCopyResult(false, null, 0, 0, [], [], ex.Message);
                }

                if (result.Success && result.NewId is not null)
                {
                    // The copy itself succeeded; a locked or unwritable map file must not turn that into a crash.
                    try
                    {
                        FormsIdMapStore.Record(row.Id, result.NewId, TargetAccount);
                        row.CopiedText = CopiedLabel(new FormsMapEntry(result.NewId, TargetAccount, DateTimeOffset.Now));
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        AppendLog(Loc.T("Forms_MapNotSaved", ex.Message));
                        result = result with { Warnings = [.. result.Warnings, Loc.T("Forms_MapNotSaved", ex.Message)] };
                    }
                }
                _vm.Results.Add(FormResultRow.From(row, copyTitle, result));
                _vm.ProgressValue = i + 1;
            }
        }
        catch (OperationCanceledException) { cancelled = true; }
        finally
        {
            clock.Stop();
            _vm.IsBusy = false;
        }

        _vm.Finish(clock.Elapsed, cancelled, chosen.Count, TargetAccount, FormsIdMapStore.DefaultPath);
        GoTo(3);
        _vm.StatusText = _vm.ResultHeadline;
    }

    private void CancelRun_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        _vm.ProgressText = Loc.T("Forms_Cancelling");
    }

    // ── Step 4: Results ───────────────────────────────────────────────────────

    private FormResultRow? SelectedResult => ResultsGrid.SelectedItem as FormResultRow;

    private void OpenSelectedCopy_Click(object sender, RoutedEventArgs e)
    {
        var row = SelectedResult ?? _vm.Results.FirstOrDefault(r => r.NewId is not null);
        if (row?.NewId is null) { _vm.StatusText = Loc.T("Forms_SelectACopy"); return; }
        OpenInBrowser($"https://forms.office.com/Pages/DesignPageV2.aspx?subpage=design&id={Uri.EscapeDataString(row.NewId)}");
    }

    private void OpenForms_Click(object sender, RoutedEventArgs e) => OpenInBrowser("https://forms.office.com");

    private static void OpenInBrowser(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { /* no default browser: nothing useful to do */ }
    }

    private void ShowMap_Click(object sender, RoutedEventArgs e)
    {
        var path = FormsIdMapStore.DefaultPath;
        try
        {
            if (File.Exists(path)) Process.Start("explorer.exe", $"/select,\"{path}\"");
            else if (Directory.Exists(Path.GetDirectoryName(path))) Process.Start("explorer.exe", $"\"{Path.GetDirectoryName(path)}\"");
        }
        catch { /* explorer unavailable */ }
    }

    private void ExportResults_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx",
            FileName = $"Forms-copy-results-{DateTime.Now:yyyyMMdd-HHmm}.xlsx",
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Results");
            string[] headers = ["Source form", "Source form ID", "Copy title", "Copy form ID", "Status", "Questions copied", "Questions in source", "Warnings", "Notes", "Error"];
            for (var c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];
            var r = 2;
            foreach (var row in _vm.Results)
            {
                ws.Cell(r, 1).Value = row.SourceTitle;
                ws.Cell(r, 2).Value = row.SourceId;
                ws.Cell(r, 3).Value = row.CopyTitle;
                ws.Cell(r, 4).Value = row.NewId ?? "";
                ws.Cell(r, 5).Value = row.StatusPlain;
                ws.Cell(r, 6).Value = row.QuestionsCopied;
                ws.Cell(r, 7).Value = row.QuestionsTotal;
                ws.Cell(r, 8).Value = string.Join("\n", row.Warnings);
                ws.Cell(r, 9).Value = string.Join("\n", row.Notes);
                ws.Cell(r, 10).Value = row.Error ?? "";
                r++;
            }
            ws.SheetView.FreezeRows(1);
            ws.Columns().AdjustToContents(1, Math.Min(r, 50));
            wb.SaveAs(dlg.FileName);
            _vm.StatusText = Loc.T("Forms_ResultsSaved", dlg.FileName);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _vm.StatusText = Loc.T("Forms_Error", ex.Message);
        }
    }

    private void CopyMore_Click(object sender, RoutedEventArgs e)
    {
        foreach (var row in _rows) row.IsSelected = false;
        UpdateSelection();
        _vm.Results.Clear();
        GoTo(0);
    }

    // ── Export responses (step 1 footer) ──────────────────────────────────────

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        var chosen = Selected.ToList();
        if (chosen.Count == 0) { _vm.StatusText = Loc.T("Forms_NothingSelected"); return; }

        if (_vm.CombineExport)
        {
            await ExportCombinedAsync(chosen);
            return;
        }

        var dlg = new OpenFolderDialog { Title = Loc.T("Forms_ChooseExportFolder") };
        if (dlg.ShowDialog(this) != true) return;
        var folder = dlg.FolderName;

        await RunAsync(Loc.T("Forms_Exporting"), async ct =>
        {
            var exporter = new FormsResponseExporter(SourceClient());
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int exported = 0, failed = 0;
            foreach (var row in chosen)
            {
                ct.ThrowIfCancellationRequested();
                // Forms can share a title; each gets its own file.
                var baseName = SafeFileName(row.Title);
                var fileName = baseName;
                for (var n = 2; !usedNames.Add(fileName); n++) fileName = $"{baseName} ({n})";
                var path = Path.Combine(folder, fileName + ".xlsx");
                try
                {
                    var count = await exporter.ExportAsync(row.Id, path, ct);
                    AppendLog(Loc.T("Forms_LogExported", row.Title, count, path));
                    exported++;
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    // One form failing (no permission, a locked file) must not stop the others.
                    AppendLog(Loc.T("Forms_LogExportFailed", row.Title, ex.Message));
                    failed++;
                }
            }
            _vm.StatusText = failed == 0
                ? Loc.T("Forms_ExportDone", exported, folder)
                : Loc.T("Forms_ExportDonePartial", exported, failed, folder);
        });
    }

    // One workbook with a sheet per form.
    private async Task ExportCombinedAsync(List<FormRow> chosen)
    {
        var dlg = new SaveFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx",
            FileName = $"Forms-responses-{DateTime.Now:yyyyMMdd-HHmm}.xlsx",
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        };
        if (dlg.ShowDialog(this) != true) return;
        var path = dlg.FileName;

        await RunAsync(Loc.T("Forms_Exporting"), async ct =>
        {
            var exporter = new FormsResponseExporter(SourceClient());
            var fetched = new List<(string Title, JsonObject Form, IReadOnlyList<JsonObject> Responses)>();
            var failed = 0;
            foreach (var row in chosen)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var (form, responses) = await exporter.FetchAsync(row.Id, ct);
                    fetched.Add((row.Title, form, responses));
                    AppendLog(Loc.T("Forms_LogExported", row.Title, responses.Count, path));
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    AppendLog(Loc.T("Forms_LogExportFailed", row.Title, ex.Message));
                    failed++;
                }
            }
            if (fetched.Count == 0) { _vm.StatusText = Loc.T("Forms_ExportNothing"); return; }
            FormsResponseExporter.WriteCombined(fetched, path);
            _vm.StatusText = failed == 0
                ? Loc.T("Forms_ExportDone", fetched.Count, path)
                : Loc.T("Forms_ExportDonePartial", fetched.Count, failed, path);
        });
    }

    // ── Plumbing ──────────────────────────────────────────────────────────────

    private async Task RunAsync(string busyText, Func<CancellationToken, Task> work)
    {
        _cts = new CancellationTokenSource();
        _vm.IsBusy = true;
        _vm.StatusText = busyText;
        try { await work(_cts.Token); }
        catch (OperationCanceledException) { _vm.StatusText = Loc.T("Forms_Cancelled"); }
        catch (Exception ex)
        {
            _vm.StatusText = Loc.T("Forms_Error", ex.Message);
            if (HelpFor(ex) is { } help) AppendLog(help);
            AppendLog(ex.Message);
        }
        finally { _vm.IsBusy = false; }
    }

    // Most first-time failures are the Microsoft Forms permission missing from (or not consented on) the app registration.
    private static string? HelpFor(Exception ex)
    {
        var m = ex.Message;
        if (m.Contains("AADSTS65001") || m.Contains("AADSTS500011") || m.Contains("AADSTS650057") || m.Contains("consent", StringComparison.OrdinalIgnoreCase))
            return Loc.T("Forms_ConsentHelp");
        if (m.Contains("AADSTS50011") || m.Contains("AADSTS500113"))
            return Loc.T("Forms_RedirectHelp");
        // Forms API errors carry the status as an HttpRequestException; matching bare digits in the message would
        // also fire on a form title or body that happens to contain "401".
        if (ex is HttpRequestException { StatusCode: System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden })
            return Loc.T("Forms_AccessHelp");
        return null;
    }

    private void AppendLog(string line)
    {
        if (!Dispatcher.CheckAccess()) { Dispatcher.Invoke(() => AppendLog(line)); return; }
        LogBox.AppendText(line + Environment.NewLine);
        LogBox.ScrollToEnd();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    protected override void OnClosing(CancelEventArgs e)
    {
        // Closing mid-copy would leave a half-built form in the target; make that a deliberate choice.
        if (_vm.IsBusy && _vm.CurrentStep == 2 &&
            MessageBox.Show(this, Loc.T("Forms_ClosePrompt"), Loc.T("Forms_CloseTitle"),
                MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) != MessageBoxResult.Yes)
        {
            e.Cancel = true;
            return;
        }
        _cts?.Cancel();
        base.OnClosing(e);
    }

    private static string SafeFileName(string name)
    {
        var sb = new StringBuilder(name.Length);
        foreach (var c in name) sb.Append(Path.GetInvalidFileNameChars().Contains(c) ? '_' : c);
        var s = sb.ToString().Trim();
        return s.Length == 0 ? "form" : s;
    }
}

public partial class FormRow : ObservableObject
{
    [ObservableProperty] private bool _isSelected;
    // Filled when the saved ID map says this form was copied before ("Copied 3/4/2026 to a@b.com").
    [ObservableProperty] private string _copiedText = string.Empty;
    public string Id { get; init; } = string.Empty;
    public string Title { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public DateTimeOffset? Modified { get; init; }
    public string ModifiedText => Modified?.LocalDateTime.ToString("g") ?? string.Empty;

    // The filter box: every word must appear in the title or type (case-insensitive).
    public bool Matches(string filter) =>
        string.IsNullOrWhiteSpace(filter)
        || filter.Split(' ', StringSplitOptions.RemoveEmptyEntries).All(w =>
            Title.Contains(w, StringComparison.OrdinalIgnoreCase) || Type.Contains(w, StringComparison.OrdinalIgnoreCase));
}

public enum FormOutcome { Copied, CopiedWithWarnings, Failed }

// One line of the results table.
public sealed class FormResultRow
{
    public string SourceTitle { get; init; } = string.Empty;
    public string SourceId { get; init; } = string.Empty;
    public string CopyTitle { get; init; } = string.Empty;
    public string? NewId { get; init; }
    public FormOutcome Outcome { get; init; }
    public int QuestionsTotal { get; init; }
    public int QuestionsCopied { get; init; }
    public IReadOnlyList<string> Warnings { get; init; } = [];
    public IReadOnlyList<string> Notes { get; init; } = [];
    public string? Error { get; init; }

    public string StatusPlain => Loc.T("Forms_Outcome_" + Outcome);
    public string StatusText => Outcome switch
    {
        FormOutcome.Copied => "✓ " + StatusPlain,
        FormOutcome.CopiedWithWarnings => "⚠ " + StatusPlain,
        _ => "✗ " + StatusPlain,
    };
    public string QuestionsText => Outcome == FormOutcome.Failed ? "–" : $"{QuestionsCopied} / {QuestionsTotal}";
    public string Details => Error ?? string.Join(" ", Warnings.Concat(Notes));

    public static FormResultRow From(FormRow source, string copyTitle, FormCopyResult r) => new()
    {
        SourceTitle = source.Title,
        SourceId = source.Id,
        CopyTitle = r.Success ? copyTitle : string.Empty,
        NewId = r.NewId,
        Outcome = !r.Success ? FormOutcome.Failed : r.IsPartial ? FormOutcome.CopiedWithWarnings : FormOutcome.Copied,
        QuestionsTotal = r.QuestionsTotal,
        QuestionsCopied = r.QuestionsCopied,
        Warnings = r.Warnings,
        Notes = r.Notes,
        Error = r.Error,
    };
}

// State for the four steps. The dialog's code-behind does the work and sets these; the XAML binds to them.
public partial class FormsDialogViewModel : ObservableObject
{
    public ObservableCollection<FormResultRow> Results { get; } = [];

    [ObservableProperty, NotifyPropertyChangedFor(nameof(ShowBack), nameof(ShowNext), nameof(ShowCancel), nameof(ShowCopyMore),
        nameof(ShowExportResponses), nameof(NextButtonText))]
    private int _currentStep;

    // step 1 filter box and export option
    [ObservableProperty] private string _filterText = string.Empty;
    [ObservableProperty] private bool _combineExport;

    [ObservableProperty, NotifyPropertyChangedFor(nameof(IsSingleTenant), nameof(NeedsTargetSignIn), nameof(TargetIntroText), nameof(CanNext))]
    private bool _isMultiTenant;
    public bool IsSingleTenant => !IsMultiTenant;

    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext), nameof(ShowCancel))] private bool _isBusy;

    // step 1
    [ObservableProperty] private string _sourceAccountText = string.Empty;
    [ObservableProperty] private string _loadButtonText = Loc.T("Forms_LoadForms");
    [ObservableProperty] private bool _hasForms;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(SelectedCountText), nameof(CanNext), nameof(NextButtonText), nameof(SummaryText))]
    private int _selectedCount;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(SelectedCountText))] private int _totalCount;
    public string SelectedCountText => Loc.T("Forms_CountText", SelectedCount, TotalCount);

    // step 2
    [ObservableProperty] private string _targetRegistrationText = string.Empty;
    [ObservableProperty] private string _targetButtonText = Loc.T("Forms_TargetSignIn");
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasTargetStatus))] private string _targetStatusText = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(NeedsTargetSignIn), nameof(CanNext), nameof(SummaryText))] private bool _targetConnected;
    [ObservableProperty] private bool _targetHasError;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(SummaryText))] private string _sourceAccountName = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(SummaryText))] private string _targetAccountName = string.Empty;
    [ObservableProperty] private string _suffix = " (copy)";
    public bool HasTargetStatus => !string.IsNullOrEmpty(TargetStatusText);
    public bool NeedsTargetSignIn => IsMultiTenant && !TargetConnected;
    public string TargetIntroText => IsMultiTenant ? Loc.T("Forms_TargetIntroMulti") : Loc.T("Forms_TargetIntroSingle");
    public string SummaryText => Loc.T("Forms_Summary", SelectedCount,
        string.IsNullOrEmpty(SourceAccountName) ? Loc.T("Forms_SourceAccountShort") : SourceAccountName,
        string.IsNullOrEmpty(TargetAccountName) ? Loc.T("Forms_TargetAccountShort") : TargetAccountName);

    // step 3
    [ObservableProperty] private double _progressMax = 1;
    [ObservableProperty] private double _progressValue;
    [ObservableProperty] private string _progressText = string.Empty;

    // step 4
    [ObservableProperty] private string _outcome = "Success";
    [ObservableProperty] private string _resultHeadline = string.Empty;
    [ObservableProperty] private string _resultSubline = string.Empty;
    [ObservableProperty] private int _copiedCount;
    [ObservableProperty] private int _partialCount;
    [ObservableProperty] private int _failedCount;
    [ObservableProperty] private string _questionsText = "0";
    [ObservableProperty] private string _durationText = string.Empty;

    [ObservableProperty] private string _statusText = string.Empty;

    // footer
    public bool ShowBack => CurrentStep == 1;
    public bool ShowNext => CurrentStep is 0 or 1;
    // Cancel is offered for any running operation (loading, signing in, exporting), not just the copy itself.
    public bool ShowCancel => CurrentStep == 2 || IsBusy;
    public bool ShowCopyMore => CurrentStep == 3;
    public bool ShowExportResponses => CurrentStep == 0;
    public string NextButtonText => CurrentStep == 1 ? Loc.T("Forms_CopyNForms", SelectedCount) : Loc.T("Forms_Next");
    public bool CanNext => !IsBusy && CurrentStep switch
    {
        0 => SelectedCount > 0,
        1 => TargetConnected,
        _ => false,
    };

    partial void OnCurrentStepChanged(int value) => OnPropertyChanged(nameof(CanNext));
    partial void OnSelectedCountChanged(int value) => OnPropertyChanged(nameof(CanNext));

    // Fills the dashboard from the per-form results.
    public void Finish(TimeSpan elapsed, bool cancelled, int planned, string targetAccount, string mapPath)
    {
        CopiedCount = Results.Count(r => r.Outcome == FormOutcome.Copied);
        PartialCount = Results.Count(r => r.Outcome == FormOutcome.CopiedWithWarnings);
        FailedCount = Results.Count(r => r.Outcome == FormOutcome.Failed);
        var qCopied = Results.Sum(r => r.QuestionsCopied);
        var qTotal = Results.Where(r => r.Outcome != FormOutcome.Failed).Sum(r => r.QuestionsTotal);
        QuestionsText = $"{qCopied} / {qTotal}";
        DurationText = elapsed.TotalMinutes >= 1 ? $"{(int)elapsed.TotalMinutes}m {elapsed.Seconds}s" : $"{Math.Max(1, (int)Math.Round(elapsed.TotalSeconds))}s";

        var done = CopiedCount + PartialCount;
        var who = string.IsNullOrWhiteSpace(targetAccount) ? Loc.T("Forms_TargetAccountShort") : targetAccount;
        if (cancelled)
        {
            Outcome = "Cancelled";
            ResultHeadline = Loc.T("Forms_ResultCancelled", Results.Count, planned);
            ResultSubline = Loc.T("Forms_ResultCancelledSub");
        }
        else if (FailedCount == 0 && PartialCount == 0)
        {
            Outcome = "Success";
            ResultHeadline = Loc.T("Forms_ResultSuccess", CopiedCount);
            ResultSubline = Loc.T("Forms_ResultSuccessSub", who);
        }
        else if (done == 0)
        {
            Outcome = "Failed";
            ResultHeadline = Loc.T("Forms_ResultAllFailed");
            ResultSubline = Loc.T("Forms_ResultAllFailedSub");
        }
        else
        {
            Outcome = "Partial";
            ResultHeadline = Loc.T("Forms_ResultPartial", done, Results.Count);
            ResultSubline = Loc.T("Forms_ResultPartialSub", who);
        }
    }
}
