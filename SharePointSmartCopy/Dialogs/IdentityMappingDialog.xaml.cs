using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.Dialogs;

public enum IdentityRowCheck { Unchecked, Checking, Found, NotFound }

// One line of the mapping table: this source user (or group) is that target user (or group).
public sealed partial class IdentityMappingRow : ObservableObject
{
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(Kind))] private string _source = string.Empty;
    [ObservableProperty] private string _target = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusText))]
    [NotifyPropertyChangedFor(nameof(IsFound))]
    [NotifyPropertyChangedFor(nameof(IsNotFound))]
    private IdentityRowCheck _check = IdentityRowCheck.Unchecked;

    [ObservableProperty] [NotifyPropertyChangedFor(nameof(HasError))] private string _error = string.Empty;

    // "author, permission · 12×" for a row that came from a run's unresolved users; empty otherwise.
    public string Note { get; init; } = string.Empty;

    public bool HasError   => Error.Length > 0;
    public bool IsFound    => Check == IdentityRowCheck.Found;
    public bool IsNotFound => Check == IdentityRowCheck.NotFound;

    // An email on the source side means a person; anything else is a group name.
    public string Kind => Source.Contains('@') ? Loc.T("Ui_IdentityKindUser") : Loc.T("Ui_IdentityKindGroup");

    public string StatusText => Check switch
    {
        IdentityRowCheck.Checking => Loc.T("Ui_IdentityStatusChecking"),
        IdentityRowCheck.Found    => Loc.T("Ui_IdentityStatusFound"),
        IdentityRowCheck.NotFound => Loc.T("Ui_IdentityStatusNotFound"),
        _                 => Loc.T("Ui_IdentityStatusUnchecked"),
    };

    // A different target means any earlier "found / not found" no longer applies.
    partial void OnTargetChanged(string value) => Check = IdentityRowCheck.Unchecked;
}

// Edits the cross-tenant identity mapping stored in AppSettings: an optional domain swap plus a table of
// explicit "this source user/group = that target user/group" pairs.
public partial class IdentityMappingDialog : Window
{
    private readonly AppSettings _settings;
    private readonly Func<string, Task<bool?>>? _checkInTarget;
    private readonly ObservableCollection<IdentityMappingRow> _rows = [];
    private CancellationTokenSource? _checkCts;
    private string _checkSummary = string.Empty;

    // prefill: people a run could not find in the target (their rows are added unless already mapped).
    // checkInTarget: looks one target email up in the target tenant; null disables the "Check in target" button.
    public IdentityMappingDialog(
        AppSettings settings,
        IEnumerable<UnresolvedUser>? prefill = null,
        Func<string, Task<bool?>>? checkInTarget = null)
    {
        InitializeComponent();
        _settings      = settings;
        _checkInTarget = checkInTarget;

        var m = settings.IdentityMapping;
        SourceDomainBox.Text = m.SourceDomain;
        TargetDomainBox.Text = m.TargetDomain;
        foreach (var pair in m.Pairs)
            AddRow(new IdentityMappingRow { Source = pair.Source, Target = pair.Target });

        var prefillList = prefill?.ToList() ?? [];
        if (prefillList.Count > 0)
        {
            PrefillBanner.Visibility = Visibility.Visible;
            NeededForColumn.Visibility = Visibility.Visible;
            var domainSwap = new IdentityMapper(m.SourceDomain, m.TargetDomain, null);
            foreach (var u in prefillList)
            {
                if (_rows.Any(r => r.Source.Trim().Equals(u.Source, StringComparison.OrdinalIgnoreCase))) continue;
                // Suggest what the domain swap would give; otherwise what the mapping already tried (if it differs).
                var suggested = domainSwap.MapEmail(u.Source) ?? string.Empty;
                if (suggested.Equals(u.Source, StringComparison.OrdinalIgnoreCase))
                    suggested = u.AttemptedTarget.Equals(u.Source, StringComparison.OrdinalIgnoreCase) ? string.Empty : u.AttemptedTarget;
                AddRow(new IdentityMappingRow
                {
                    Source = u.Source,
                    Target = suggested,
                    Note   = $"{u.Where} · {u.Count}×",
                });
            }
        }

        MappingGrid.ItemsSource = _rows;

        if (_checkInTarget == null)
        {
            CheckButton.IsEnabled = false;
            CheckButton.ToolTip   = Loc.T("Ui_IdentityCheckDisabledHint");
        }
        else
        {
            CheckButton.ToolTip = Loc.T("Ui_IdentityCheckHint");
        }

        Revalidate();
        UpdateSummary();
    }

    private void AddRow(IdentityMappingRow row)
    {
        row.PropertyChanged += Row_PropertyChanged;
        _rows.Add(row);
    }

    private void Row_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        // Editing a cell fixes (or creates) problems — re-evaluate so the highlight and count stay current.
        if (e.PropertyName is nameof(IdentityMappingRow.Source) or nameof(IdentityMappingRow.Target)) { Revalidate(); UpdateSummary(); }
    }

    private void Input_Changed(object sender, TextChangedEventArgs e) => UpdateSummary();

    // Finishes a cell that is still being edited, so Save / Check see what is on screen.
    private void CommitGrid()
    {
        MappingGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        MappingGrid.CommitEdit(DataGridEditingUnit.Row, true);
    }

    // Marks each row with its problem (if any). Returns how many rows have one.
    private int Revalidate()
    {
        var problems = IdentityMapper.ValidatePairs([.. _rows.Select(r => (r.Source, r.Target))]);
        for (int i = 0; i < _rows.Count; i++)
            _rows[i].Error = problems.TryGetValue(i, out var message) ? message : string.Empty;
        return problems.Count;
    }

    private void UpdateSummary()
    {
        if (SummaryText == null) return; // fires during InitializeComponent, before the field exists

        int usable = _rows.Count(r => r.Source.Trim().Length > 0 && r.Target.Trim().Length > 0 && !r.HasError);
        int errors = _rows.Count(r => r.HasError);
        var parts  = new List<string> { Loc.T("Ui_IdentityRowsSummary", usable) };
        bool warn  = false;

        if (errors > 0) { parts.Add(Loc.T("Ui_IdentityRowErrors", errors)); warn = true; }

        var src = SourceDomainBox.Text.Trim().TrimStart('@');
        var tgt = TargetDomainBox.Text.Trim().TrimStart('@');
        if (src.Length > 0 && src.Equals(tgt, StringComparison.OrdinalIgnoreCase))
        {
            parts.Add(Loc.T("Ui_IdentityDomainsSame"));
            warn = true;
        }
        if (_checkSummary.Length > 0) parts.Add(_checkSummary);

        SummaryText.Text = string.Join("  ", parts);
        SummaryText.Foreground = (Brush)FindResource(warn ? "WarningTextBrush" : "TextTertiaryBrush");
    }

    private void AddRow_Click(object sender, RoutedEventArgs e)
    {
        var row = new IdentityMappingRow();
        AddRow(row);
        MappingGrid.SelectedItem = row;
        MappingGrid.ScrollIntoView(row);
        MappingGrid.CurrentCell = new DataGridCellInfo(row, MappingGrid.Columns[0]);
        MappingGrid.BeginEdit();
    }

    private void RemoveRows_Click(object sender, RoutedEventArgs e)
    {
        CommitGrid();
        foreach (var row in MappingGrid.SelectedItems.OfType<IdentityMappingRow>().ToList())
        {
            row.PropertyChanged -= Row_PropertyChanged;
            _rows.Remove(row);
        }
        Revalidate();
        UpdateSummary();
    }

    // For rows with a source but no target yet: what the domain swap would produce.
    private void FillFromDomain_Click(object sender, RoutedEventArgs e)
    {
        CommitGrid();
        var swap = new IdentityMapper(SourceDomainBox.Text, TargetDomainBox.Text, null);
        foreach (var row in _rows.Where(r => r.Target.Trim().Length == 0 && r.Source.Contains('@')))
        {
            var mapped = swap.MapEmail(row.Source.Trim());
            if (!string.IsNullOrEmpty(mapped) && !mapped.Equals(row.Source.Trim(), StringComparison.OrdinalIgnoreCase))
                row.Target = mapped;
        }
        Revalidate();
        UpdateSummary();
    }

    // Looks each TARGET email up in the target tenant, a few at a time, and shows ✓ / ✗ per row.
    private async void Check_Click(object sender, RoutedEventArgs e)
    {
        // async void: nothing may escape or it would crash the app.
        if (_checkInTarget == null) return;
        try
        {
            CommitGrid();
            var rows = _rows.Where(r => r.Source.Trim().Length > 0 && r.Target.Trim().Contains('@')).ToList();
            if (rows.Count == 0)
            {
                _checkSummary = Loc.T("Ui_IdentityNothingToCheck");
                UpdateSummary();
                return;
            }

            CheckButton.IsEnabled = false;
            _checkCts = new CancellationTokenSource();
            foreach (var r in rows) r.Check = IdentityRowCheck.Checking;
            int done = 0, found = 0, missing = 0, unknown = 0;
            _checkSummary = Loc.T("Ui_IdentityChecking", 0, rows.Count);
            UpdateSummary();

            await Parallel.ForEachAsync(rows,
                new ParallelOptions { MaxDegreeOfParallelism = 4, CancellationToken = _checkCts.Token },
                async (row, ct) =>
                {
                    // The target email as it was when the check started — the user may edit the row meanwhile.
                    var email  = Dispatcher.Invoke(() => row.Target.Trim());
                    var result = await _checkInTarget(email);
                    Dispatcher.Invoke(() =>
                    {
                        // Only apply it if the row still has the target that was checked.
                        if (row.Target.Trim() == email)
                            row.Check = result == true ? IdentityRowCheck.Found : result == false ? IdentityRowCheck.NotFound : IdentityRowCheck.Unchecked;
                        if (result == true) found++; else if (result == false) missing++; else unknown++;
                        done++;
                        _checkSummary = Loc.T("Ui_IdentityChecking", done, rows.Count);
                        UpdateSummary();
                    });
                });

            _checkSummary = Loc.T("Ui_IdentityCheckResult", found, missing,
                unknown > 0 ? Loc.T("Ui_IdentityCheckCouldNot", unknown) : string.Empty);
        }
        catch (OperationCanceledException) { _checkSummary = string.Empty; }
        catch (Exception ex)
        {
            foreach (var r in _rows.Where(r => r.Check == IdentityRowCheck.Checking)) r.Check = IdentityRowCheck.Unchecked;
            _checkSummary = ex.Message;
        }
        finally
        {
            CheckButton.IsEnabled = _checkInTarget != null;
            UpdateSummary();
        }
    }

    // Appends the rows of a CSV/text file ("source,target" per line, optional header, quoted fields).
    private void ImportCsv_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "CSV / text (*.csv;*.txt)|*.csv;*.txt|All files (*.*)|*.*",
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            CommitGrid();
            foreach (var (source, target) in IdentityMapper.ParseCsvPairs(System.IO.File.ReadAllText(dlg.FileName)))
            {
                // A source that is already in the table is updated, not duplicated.
                var existing = _rows.FirstOrDefault(r => r.Source.Trim().Equals(source, StringComparison.OrdinalIgnoreCase));
                if (existing != null) existing.Target = target;
                else AddRow(new IdentityMappingRow { Source = source, Target = target });
            }
            Revalidate();
            UpdateSummary();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, Loc.T("Dlg_IdentityImportFailed", ex.Message), Title,
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void ExportCsv_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            FileName = "identity-mapping.csv",
            Filter   = "CSV (*.csv)|*.csv",
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            CommitGrid();
            var lines = new List<string> { "Source,Target" };
            lines.AddRange(_rows
                .Where(r => r.Source.Trim().Length > 0 || r.Target.Trim().Length > 0)
                .Select(r => $"{IdentityMapper.CsvQuote(r.Source.Trim())},{IdentityMapper.CsvQuote(r.Target.Trim())}"));
            System.IO.File.WriteAllLines(dlg.FileName, lines);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, Loc.T("Dlg_IdentitySaveFailed", ex.Message), Title,
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        _checkCts?.Cancel(); // stop an in-flight check when the dialog goes away
        base.OnClosing(e);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        CommitGrid();

        // Don't save a table with rows that can't be used: highlight them and stay open.
        if (Revalidate() > 0)
        {
            UpdateSummary();
            MappingGrid.SelectedItem = _rows.First(r => r.HasError);
            MappingGrid.ScrollIntoView(MappingGrid.SelectedItem);
            MessageBox.Show(this, Loc.T("Dlg_IdentityFixRows"), Title, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var previous = _settings.IdentityMapping;
        _settings.IdentityMapping = new IdentityMappingSettings
        {
            SourceDomain = SourceDomainBox.Text.Trim(),
            TargetDomain = TargetDomainBox.Text.Trim(),
            Pairs = [.. _rows
                .Where(r => r.Source.Trim().Length > 0 && r.Target.Trim().Length > 0)
                .Select(r => new IdentityPair { Source = r.Source.Trim(), Target = r.Target.Trim() })],
        };
        try
        {
            _settings.Save();
        }
        catch (Exception ex)
        {
            // Saving failed: put the old mapping back and keep the dialog open so nothing is lost.
            _settings.IdentityMapping = previous;
            MessageBox.Show(this, Loc.T("Dlg_IdentitySaveFailed", ex.Message), Title,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
