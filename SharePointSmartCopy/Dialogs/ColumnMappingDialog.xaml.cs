using System.ComponentModel;
using System.Windows;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.ViewModels;

namespace SharePointSmartCopy.Dialogs;

public partial class ColumnMappingDialog : Window
{
    private readonly MainViewModel _vm;

    public ColumnMappingDialog(MainViewModel vm)
    {
        _vm = vm;
        InitializeComponent();

        var dlgVm = new ColumnMappingViewModel(
            vm.SourceColumns, vm.TargetColumns, vm.ColumnMappings,
            isLibraryScope: vm.IsLibraryOrSiteScope);
        DataContext = dlgVm;
        UpdateStatusBar();
        if (vm.ColumnLoadError != null)
            StatusBar.Text = $"⚠ {vm.ColumnLoadError}";
    }

    private ColumnMappingViewModel DlgVM => (ColumnMappingViewModel)DataContext;

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
            foreach (var row in DlgVM.Mappings)
            {
                if (row.Mapping.Status == MappingStatus.ManuallyMapped) continue;

                var exact = _vm.TargetColumns.FirstOrDefault(t =>
                    t.InternalName.Equals(row.Mapping.SourceColumn.InternalName, StringComparison.OrdinalIgnoreCase));
                var fuzzy = exact ?? _vm.TargetColumns.FirstOrDefault(t =>
                    t.DisplayName.Equals(row.Mapping.SourceColumn.DisplayName, StringComparison.OrdinalIgnoreCase));

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
        _vm.ColumnMappings.Clear();
        foreach (var row in DlgVM.Mappings)
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
                row.Mapping.TargetColumn = _vm.TargetColumns.FirstOrDefault(
                    t => t.InternalName == sel.InternalName);
                row.Mapping.CreateNew    = false;
                row.Mapping.Status       = row.Mapping.TargetColumn != null
                    ? MappingStatus.ManuallyMapped
                    : MappingStatus.Unmatched;
            }
            _vm.ColumnMappings.Add(row.Mapping);
        }
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void UpdateStatusBar()
    {
        if (DlgVM.Mappings.Count == 0)
        {
            StatusBar.Text = "ℹ No mappable columns found. Supported types: Text, Note, Number, Boolean, Date, Choice, Multi-choice.";
            return;
        }
        var skipped = DlgVM.Mappings.Count(r => r.SelectedTargetItem == null || r.SelectedTargetItem.IsSkip);
        if (DlgVM.IsLibraryScope)
        {
            var creating = DlgVM.Mappings.Count(r => r.SelectedTargetItem?.IsCreate == true);
            StatusBar.Text = $"✅ {creating} will be created    ⚠ {skipped} skipped";
        }
        else
        {
            var mapped = DlgVM.Mappings.Count(r => r.SelectedTargetItem != null && !r.SelectedTargetItem.IsSkip && !r.SelectedTargetItem.IsCreate);
            StatusBar.Text = $"✅ {mapped} mapped    ⚠ {skipped} skipped";
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
        ? "Choose which columns to create in the new target library. Columns set to 'Skip' will not be created."
        : "Map source columns to target columns. Unmatched columns will be skipped unless you assign a target. Use Auto-match to automatically match columns by name.";
    public string TargetColumnHeader => IsLibraryScope ? "Action" : "Target Column";

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
                new TargetColumnOption { DisplayName = "Create in target", IsCreate = true, InternalName = "__create__" },
                new TargetColumnOption { DisplayName = "── Skip this column ──", IsSkip = true, InternalName = "__skip__" },
            ];
        }
        else
        {
            // Files/Pages scope: map to an existing target column, or skip.
            TargetColumnOptions =
            [
                new TargetColumnOption { DisplayName = "── Skip this column ──", IsSkip = true, InternalName = "__skip__" },
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
                    selected = TargetColumnOptions.FirstOrDefault(o => !o.IsSkip && o.InternalName == mapping.TargetColumn.InternalName);
                if (selected == null && mapping.Status == MappingStatus.Skipped)
                    selected = skipOption;
            }

            return new MappingRow(mapping, selected);
        }).ToList();
    }
}

// ── Row model ─────────────────────────────────────────────────────────────────

public class MappingRow : INotifyPropertyChanged
{
    public ColumnMapping Mapping { get; }

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

    public MappingRow(ColumnMapping mapping, TargetColumnOption? initialSelection)
    {
        Mapping             = mapping;
        _selectedTargetItem = initialSelection;
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
