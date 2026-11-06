using SharePointSmartCopy.Localization;
using System.ComponentModel;
using System.Windows;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.ViewModels;

namespace SharePointSmartCopy.Dialogs;

public partial class ColumnMappingDialog : Window
{
    private readonly MainViewModel _vm;

    // One entry per editable mapping. Item-level copies have one per selected list (Context set); every
    // other scope has a single entry with a null Context that edits the shared VM.ColumnMappings.
    private sealed record Entry(ListMappingContext? Context, ColumnMappingViewModel Vm);
    private readonly List<Entry> _entries = [];

    public ColumnMappingDialog(MainViewModel vm)
    {
        _vm = vm;
        MappingRow.TermSetsWillBeCopied = vm.IsMultiTenant && vm.CopyTermSets;
        InitializeComponent();

        if (vm.ListMappingContexts.Count > 0)
        {
            foreach (var ctx in vm.ListMappingContexts)
                _entries.Add(new Entry(ctx, new ColumnMappingViewModel(
                    ctx.SourceColumns, ctx.TargetColumns, ctx.Mappings,
                    isLibraryScope: ctx.IsNewDestination)));
        }
        else
        {
            _entries.Add(new Entry(null, new ColumnMappingViewModel(
                vm.SourceColumns, vm.TargetColumns, vm.ColumnMappings,
                isLibraryScope: vm.IsLibraryOrSiteScope)));
        }

        foreach (var entry in _entries)
            foreach (var row in entry.Vm.Mappings)
                row.PropertyChanged += (_, e) =>
                {
                    if (e.PropertyName == nameof(MappingRow.SelectedTargetItem))
                        UpdateStatusBar();
                };

        if (_entries[0].Context != null)
        {
            ListPicker.ItemsSource = _entries.Select(en => en.Context!.Title).ToList();
            ListPickerPanel.Visibility = _entries.Count > 1 ? Visibility.Visible : Visibility.Collapsed;
            ListPicker.SelectedIndex = 0; // fires SelectionChanged, which sets the DataContext
        }
        else
        {
            DataContext = _entries[0].Vm;
        }

        UpdateStatusBar();
        if (vm.ColumnLoadError != null)
            StatusBar.Text = Loc.T("Dlg_ColLoadWarn", vm.ColumnLoadError);
    }

    private ColumnMappingViewModel DlgVM => (ColumnMappingViewModel)DataContext;

    // Columns the current list's rows can map to.
    private IReadOnlyList<ColumnDefinition> CurrentTargetColumns =>
        _entries.First(en => en.Vm == DlgVM).Context?.TargetColumns ?? (IReadOnlyList<ColumnDefinition>)_vm.TargetColumns;

    private void ListPicker_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (ListPicker.SelectedIndex < 0) return;
        DataContext = _entries[ListPicker.SelectedIndex].Vm;
        UpdateStatusBar();
    }

    private void AutoMatch_Click(object sender, RoutedEventArgs e)
    {
        if (DlgVM.IsLibraryScope)
        {
            // Library/Site scope: auto-match means "create all columns in target"
            var createOption = DlgVM.TargetColumnOptions.First(o => o.IsCreate);
            foreach (var row in DlgVM.Mappings)
            {
                row.SelectedTargetItem = createOption;
                row.Mapping.Status     = MappingStatus.WillCreate;
            }
        }
        else
        {
            var targets = CurrentTargetColumns;
            foreach (var row in DlgVM.Mappings)
            {
                if (row.Mapping.Status == MappingStatus.ManuallyMapped) continue;

                // Name matches only count when the value types are compatible — a Person
                // column must not silently match a Text column that shares its name.
                var srcType = row.Mapping.SourceColumn.FieldType;
                var exact = targets.FirstOrDefault(t =>
                    t.InternalName.Equals(row.Mapping.SourceColumn.InternalName, StringComparison.OrdinalIgnoreCase) &&
                    ColumnMapping.AreTypesCompatible(srcType, t.FieldType));
                var fuzzy = exact ?? targets.FirstOrDefault(t =>
                    t.DisplayName.Equals(row.Mapping.SourceColumn.DisplayName, StringComparison.OrdinalIgnoreCase) &&
                    ColumnMapping.AreTypesCompatible(srcType, t.FieldType));

                if (fuzzy != null)
                {
                    row.SelectedTargetItem = DlgVM.TargetColumnOptions.FirstOrDefault(
                        o => !o.IsSkip && !o.IsCreate && o.InternalName == fuzzy.InternalName);
                    row.Mapping.Status = MappingStatus.AutoMatched;
                }
                else
                {
                    row.SelectedTargetItem = DlgVM.TargetColumnOptions.First(o => o.IsSkip);
                    row.Mapping.Status     = MappingStatus.Unmatched;
                }
            }
        }
        UpdateStatusBar();
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        // Saves every list's mapping, not just the one currently shown.
        foreach (var entry in _entries)
        {
            if (entry.Context != null)
                SaveRows(entry.Vm, entry.Context.TargetColumns, entry.Context.Mappings);
            else
            {
                var shared = new List<ColumnMapping>();
                SaveRows(entry.Vm, _vm.TargetColumns, shared);
                _vm.ColumnMappings.Clear();
                foreach (var m in shared) _vm.ColumnMappings.Add(m);
            }
        }
        DialogResult = true;
    }

    private static void SaveRows(ColumnMappingViewModel dlg, IReadOnlyList<ColumnDefinition> targetColumns, ICollection<ColumnMapping> dest)
    {
        dest.Clear();
        foreach (var row in dlg.Mappings)
        {
            var sel = row.SelectedTargetItem;
            if (sel == null || sel.IsSkip)
            {
                row.Mapping.TargetColumn = null;
                row.Mapping.CreateNew    = false;
                row.Mapping.Status       = MappingStatus.Skipped;
            }
            else if (sel.IsCreate)
            {
                row.Mapping.TargetColumn = null;
                row.Mapping.CreateNew    = true;
                row.Mapping.Status       = MappingStatus.WillCreate;
            }
            else
            {
                row.Mapping.TargetColumn = targetColumns.FirstOrDefault(
                    t => t.InternalName == sel.InternalName);
                row.Mapping.CreateNew    = false;
                row.Mapping.Status       = row.Mapping.TargetColumn != null
                    ? MappingStatus.ManuallyMapped
                    : MappingStatus.Unmatched;
            }
            dest.Add(row.Mapping);
        }
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void UpdateStatusBar()
    {
        if (DlgVM.Mappings.Count == 0)
        {
            StatusBar.Text = Loc.T("Dlg_NoMappableColumns");
            return;
        }
        var skipped = DlgVM.Mappings.Count(r => r.SelectedTargetItem == null || r.SelectedTargetItem.IsSkip);
        if (DlgVM.IsLibraryScope)
        {
            var creating = DlgVM.Mappings.Count(r => r.SelectedTargetItem?.IsCreate == true);
            StatusBar.Text = Loc.T("Dlg_ColStatusCreateSkip", creating, skipped);
        }
        else
        {
            var mapped   = DlgVM.Mappings.Count(r => r.SelectedTargetItem != null && !r.SelectedTargetItem.IsSkip && !r.SelectedTargetItem.IsCreate);
            var creating = DlgVM.Mappings.Count(r => r.SelectedTargetItem?.IsCreate == true);
            StatusBar.Text = creating > 0
                ? Loc.T("Dlg_ColStatusMapCreateSkip", mapped, creating, skipped)
                : Loc.T("Dlg_ColStatusMapSkip", mapped, skipped);
        }
    }
}

// ── Dialog view model ─────────────────────────────────────────────────────────

public class ColumnMappingViewModel
{
    public List<MappingRow>          Mappings            { get; }
    public List<TargetColumnOption>  TargetColumnOptions { get; }
    public bool                      HasNoMappings       => Mappings.Count == 0;
    public bool                      IsLibraryScope      { get; }

    public string HeaderDescription  => IsLibraryScope
        ? Loc.T("Dlg_ColHeaderLibrary")
        : Loc.T("Dlg_ColHeaderFiles");
    public string TargetColumnHeader => IsLibraryScope ? Loc.T("Dlg_ColAction") : Loc.T("Dlg_ColTargetColumn");

    public ColumnMappingViewModel(
        IReadOnlyList<ColumnDefinition> sourceColumns,
        IReadOnlyList<ColumnDefinition> targetColumns,
        IEnumerable<ColumnMapping>      existingMappings,
        bool                            isLibraryScope = false)
    {
        IsLibraryScope = isLibraryScope;

        if (isLibraryScope)
        {
            // Library/Site scope: target library is being created — offer Create or Skip only.
            TargetColumnOptions =
            [
                new TargetColumnOption { DisplayName = Loc.T("Dlg_CreateInTarget"), IsCreate = true, InternalName = "__create__" },
                new TargetColumnOption { DisplayName = Loc.T("Dlg_SkipThisColumn"), IsSkip = true, InternalName = "__skip__" },
            ];
        }
        else
        {
            // Files/Pages scope: map to an existing target column, create it, or skip.
            TargetColumnOptions =
            [
                new TargetColumnOption { DisplayName = Loc.T("Dlg_SkipThisColumn"), IsSkip = true, InternalName = "__skip__" },
                new TargetColumnOption { DisplayName = Loc.T("Dlg_PlusCreateInTarget"), IsCreate = true, InternalName = "__create__" },
                .. targetColumns.Select(c => new TargetColumnOption
                {
                    DisplayName  = c.DisplayName,
                    InternalName = c.InternalName,
                    FieldType    = c.FieldType.ToString(),
                })
            ];
        }

        var skipOption   = TargetColumnOptions.First(o => o.IsSkip);
        var createOption = TargetColumnOptions.FirstOrDefault(o => o.IsCreate);

        var existingBySource = existingMappings.ToDictionary(m => m.SourceColumn.InternalName);

        Mappings = sourceColumns.Select(src =>
        {
            var mapping = existingBySource.TryGetValue(src.InternalName, out var ex)
                ? ex
                : new ColumnMapping
                {
                    SourceColumn = src,
                    Status       = isLibraryScope ? MappingStatus.WillCreate : MappingStatus.Unmatched,
                    CreateNew    = isLibraryScope,
                };

            // Per-row options: sentinels always present; real columns filtered to compatible types only.
            List<TargetColumnOption> rowOptions;
            if (isLibraryScope)
            {
                rowOptions = TargetColumnOptions; // Create + Skip only — no type filtering needed
            }
            else
            {
                rowOptions =
                [
                    skipOption,
                    .. (createOption != null ? (IEnumerable<TargetColumnOption>)[createOption] : []),
                    .. targetColumns
                        .Where(t => ColumnMapping.AreTypesCompatible(src.FieldType, t.FieldType))
                        .Select(t => TargetColumnOptions.First(o => o.InternalName == t.InternalName))
                ];
            }

            TargetColumnOption? selected;
            if (isLibraryScope)
            {
                // Default to Create; honour an explicit Skipped state from a prior save.
                selected = mapping.Status == MappingStatus.Skipped ? skipOption : createOption;
            }
            else
            {
                selected = null;
                if (mapping.TargetColumn != null)
                    selected = rowOptions.FirstOrDefault(o => !o.IsSkip && !o.IsCreate && o.InternalName == mapping.TargetColumn.InternalName);
                if (selected == null && mapping.CreateNew)
                    selected = createOption;
                if (selected == null && mapping.Status == MappingStatus.Skipped)
                    selected = skipOption;

                // First open with no saved state: auto-match by name + compatible type.
                if (selected == null && !existingBySource.ContainsKey(src.InternalName))
                {
                    var match = targetColumns.FirstOrDefault(t =>
                        t.InternalName.Equals(src.InternalName, StringComparison.OrdinalIgnoreCase) &&
                        ColumnMapping.AreTypesCompatible(src.FieldType, t.FieldType))
                        ?? targetColumns.FirstOrDefault(t =>
                        t.DisplayName.Equals(src.DisplayName, StringComparison.OrdinalIgnoreCase) &&
                        ColumnMapping.AreTypesCompatible(src.FieldType, t.FieldType));
                    if (match != null)
                    {
                        selected = rowOptions.FirstOrDefault(o => !o.IsSkip && !o.IsCreate && o.InternalName == match.InternalName);
                        mapping.TargetColumn = match;
                        mapping.Status       = MappingStatus.AutoMatched;
                    }
                }
            }

            return new MappingRow(mapping, selected, rowOptions);
        }).ToList();
    }
}

// ── Row model ─────────────────────────────────────────────────────────────────

public class MappingRow : INotifyPropertyChanged
{
    public ColumnMapping            Mapping          { get; }
    public List<TargetColumnOption> CompatibleOptions { get; }

    private TargetColumnOption? _selectedTargetItem;
    public TargetColumnOption? SelectedTargetItem
    {
        get => _selectedTargetItem;
        set
        {
            _selectedTargetItem = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedTargetItem)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StatusIcon)));
        }
    }

    public string StatusIcon
    {
        get
        {
            if (_selectedTargetItem == null || _selectedTargetItem.IsSkip) return "—";
            if (_selectedTargetItem.IsCreate) return "+";
            return "✓";
        }
    }

    // Set by the dialog: a managed-metadata column's term set will be copied to the target tenant.
    public static bool TermSetsWillBeCopied { get; set; }

    // One line under the column name that says what the column is bound to and what the copy will do about it.
    public string Info
    {
        get
        {
            var c = Mapping.SourceColumn;
            if (ColumnDefinition.IsLookupType(c.FieldType))
                return string.IsNullOrEmpty(c.LookupListTitle) ? Loc.T("Dlg_ColLooksUpOther") : Loc.T("Dlg_ColLooksUp", c.LookupListTitle);
            if (ColumnDefinition.IsTaxonomyType(c.FieldType))
                return TermSetsWillBeCopied ? Loc.T("Dlg_ColTermSetCopy") : Loc.T("Dlg_ColTermSetSame");
            return string.Empty;
        }
    }
    public bool HasInfo => Info.Length > 0;

    public MappingRow(ColumnMapping mapping, TargetColumnOption? initialSelection, List<TargetColumnOption> compatibleOptions)
    {
        Mapping             = mapping;
        _selectedTargetItem = initialSelection;
        CompatibleOptions   = compatibleOptions;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

// ── Option ────────────────────────────────────────────────────────────────────

public class TargetColumnOption
{
    public string DisplayName  { get; set; } = string.Empty;
    public string InternalName { get; set; } = string.Empty;
    public string FieldType    { get; set; } = string.Empty;
    public bool   IsSkip       { get; set; }
    public bool   IsCreate     { get; set; }

    // Shows "Column Name  (Type)" for real columns; plain name for Skip/Create sentinel items.
    public string DisplayLabel => IsSkip || IsCreate || string.IsNullOrEmpty(FieldType)
        ? DisplayName
        : $"{DisplayName}  ({FieldType})";

    public override bool Equals(object? obj) =>
        obj is TargetColumnOption other
        && IsSkip       == other.IsSkip
        && IsCreate     == other.IsCreate
        && InternalName == other.InternalName;

    public override int GetHashCode() => HashCode.Combine(IsSkip, IsCreate, InternalName);
}
