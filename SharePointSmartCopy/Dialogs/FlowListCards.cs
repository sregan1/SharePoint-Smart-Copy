using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Workflows;

namespace SharePointSmartCopy.Dialogs;

public sealed record ColumnOption(string Internal, string Label);

// A list or library on the target site that a flow's list can be pointed at. An empty Id is the "(none)" choice.
public sealed record TargetListOption(string Id, string Title);

// One column of the source list and what it becomes on the target list.
public partial class ColumnMapRow : ObservableObject
{
    private bool _auto;

    public string SourceColumn { get; init; } = string.Empty;       // internal name (what flows store)
    public string SourceDisplay { get; init; } = string.Empty;
    public string SourceType { get; init; } = string.Empty;
    // True when a flow writes to, filters on or reads this column. Only used columns are changed in the flows.
    public bool IsUsed { get; init; }
    public ColumnMatchStatus Match { get; init; }
    // What the row says about itself and its chip color (Ready / Look / Info). They follow the user's choices, so a row the
    // user fixed stops asking for attention.
    [ObservableProperty] private string _statusText = string.Empty;
    [ObservableProperty] private string _statusLevel = "Info";
    public List<ColumnOption> Options { get; init; } = [];
    public string SourceLabel => SourceDisplay == SourceColumn || SourceDisplay.Length == 0 ? SourceColumn : $"{SourceDisplay} ({SourceColumn})";

    // Internal name of the chosen target column; empty = leave the source name unchanged.
    [ObservableProperty] private string _selected = string.Empty;
    // True once the user changed the dropdown (as opposed to the automatic match filling it in).
    public bool UserEdited { get; private set; }

    // The automatic match (or a loaded mapping) filling the row in, which is not a user edit.
    public void SetAuto(string selected)
    {
        _auto = true;
        try { Selected = selected; }
        finally { _auto = false; }
    }

    // Puts back a choice the user made earlier (kept across "Match again").
    public void MarkEdited() => UserEdited = true;

    // True while the row still needs the user: no target was found and none has been chosen, or a "similar name" match the
    // user has not looked at (looking = changing the dropdown, even to the same column).
    public bool NeedsLook => IsUsed && (Match == ColumnMatchStatus.NotFoundInTarget ? string.IsNullOrEmpty(Selected) : Match == ColumnMatchStatus.Similar && !UserEdited);

    // Raised when the user changes this row, so the card can recompute its status.
    public event Action? UserChanged;

    partial void OnSelectedChanged(string value)
    {
        if (_auto) return;
        UserEdited = true;
        if (IsUsed)
        {
            StatusLevel = NeedsLook ? "Look" : "Ready";
            StatusText = Loc.T(string.IsNullOrEmpty(value) ? "Flows_ColumnStatus_Unchanged" : "Flows_ColumnStatus_Chosen");
        }
        UserChanged?.Invoke();
    }
}

// Everything about ONE list the flows use, once, however many flows use it: its name and columns on the source, the target
// list it maps to (the same-named list by default, or any list of the target site chosen from the dropdown, or an ID typed
// in), and the column mapping between the two. The card is laid out as a SOURCE pane and a TARGET pane so it is always
// clear which side is which.
public partial class ListCard : ObservableObject
{
    private bool _syncing;

    public string SourceSite { get; init; } = string.Empty;          // the address stored in the flow
    public string SourceListId { get; init; } = string.Empty;
    // The flows that use this list, for the "Used by" line, and how many there are.
    public string UsedBy { get; init; } = string.Empty;
    public int UsedByCount { get; init; }
    // The target site this list's counterpart lives on ("" when the source site has no target).
    public string TargetSite { get; set; } = string.Empty;
    // Where the list is read from on the source (differs from SourceSite when the user or the tool redirected it).
    public string ReadSite { get; set; } = string.Empty;
    // Tenant hosts shown on the two panes, e.g. "contoso.sharepoint.com".
    public string SourceHost { get; set; } = string.Empty;
    public string TargetHost { get; set; } = string.Empty;
    // False when the list could not be read on the source: the column names then come from the flow itself.
    public bool SourceReadable { get; set; } = true;
    // True when the list ID is a list on the TARGET site (a package this tool already wrote).
    public bool AlreadyOnTarget { get; set; }
    public IReadOnlyList<FlowColumnInfo> SourceColumns { get; set; } = [];
    public HashSet<string> Used { get; set; } = new(StringComparer.Ordinal);
    // The "Create item" steps on this list, and the columns of the chosen target list (with which are required).
    public IReadOnlyList<CreateItemStep> CreateSteps { get; set; } = [];
    private IReadOnlyList<FlowColumnInfo> _targetCols = [];

    // What would make the flows fail on import or at run time, in words. While any is listed the wizard does not go on.
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasBlockers))] private int _blockerCount;
    [ObservableProperty] private string _blockersText = string.Empty;
    public bool HasBlockers => BlockerCount > 0;
    // Values typed for required columns the Create item steps leave out (kept across "Check again"), and the boxes for them.
    public ObservableCollection<RequiredFixRow> RequiredFixes { get; } = [];
    public bool HasFixes => RequiredFixes.Count > 0;
    private readonly Dictionary<string, string> _fixValues = new(StringComparer.Ordinal);
    public IReadOnlyDictionary<string, string> FixValues => _fixValues.Where(kv => !string.IsNullOrWhiteSpace(kv.Value)).ToDictionary(kv => kv.Key, kv => kv.Value);
    public void LoadFixValues(IReadOnlyDictionary<string, string>? values)
    {
        if (values is null) return;
        foreach (var (k, v) in values) _fixValues[k] = v;
    }
    // Column types a typed value works for; for the rest only "make it optional" applies.
    private static readonly HashSet<string> FixableTypes = new(StringComparer.OrdinalIgnoreCase) { "Text", "Note", "Number", "Currency", "Integer" };
    // Raised by "Check again": the dialog reads the target list again (after the user fixed it in SharePoint).
    public event Action<ListCard>? RecheckRequested;

    [RelayCommand]
    private void Recheck() => RecheckRequested?.Invoke(this);

    // Every list and library on the target site. Shared by all cards on that site; filled once per site.
    public ObservableCollection<TargetListOption> TargetLists { get; init; } = [];
    public ObservableCollection<ColumnMapRow> Columns { get; } = [];
    // The column grid shows the columns the flows use by default; "Show all" adds the rest.
    public ICollectionView ColumnsView { get; }

    [ObservableProperty, NotifyPropertyChangedFor(nameof(Title))] private string _sourceTitle = string.Empty;
    // The reason the source list could not be read, or that the flow already points at the target. Empty when all is well.
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasSourceNote))] private string _sourceNote = string.Empty;
    // About the target list (for example that its columns could not be read).
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasTargetNote))] private string _targetNote = string.Empty;
    // A note about the target list that belongs to the lookup (for example "no list with this name on the target"); Refresh
    // adds to it, never replaces it.
    public string TargetBaseNote { get; set; } = string.Empty;
    private int _targetColumnCount;
    [ObservableProperty] private string _summaryText = string.Empty;
    // The line under the title when the card is collapsed: the target list.
    [ObservableProperty] private string _headerStatus = string.Empty;
    // Ready / Look / Error / Info, and the word that goes in the chip.
    [ObservableProperty] private string _statusLevel = "Info";
    [ObservableProperty] private string _statusChip = string.Empty;
    [ObservableProperty] private bool _needsAttention;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(Chevron))] private bool _isExpanded;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(ColumnsToggleText))] private bool _showAllColumns;
    // Why the column grid is empty, when it is.
    [ObservableProperty] private string _emptyColumnsText = string.Empty;
    [ObservableProperty] private bool _hasVisibleColumns;
    [ObservableProperty] private bool _hasHiddenColumns;

    [ObservableProperty] private TargetListOption? _picked;
    [ObservableProperty] private string _targetListId = string.Empty;

    // True once the user chose or typed the target list (as opposed to the automatic name match filling it in). Such a
    // choice survives "Match again".
    public bool UserEdited { get; private set; }
    // Raised when the user changes the target list; the dialog then reads that list's columns and matches them again.
    public event Action<ListCard>? TargetChanged;
    // Set while the dialog fills the card in, so that is not mistaken for a user edit.
    public bool Populating { get; set; }

    public ListCard()
    {
        RequiredFixes.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasFixes));
        ColumnsView = CollectionViewSource.GetDefaultView(Columns);
        ColumnsView.Filter = o => o is ColumnMapRow r && (ShowAllColumns || r.IsUsed);
    }

    // The list's name; when it cannot be read, a plain label (the ID is shown separately, muted).
    public string Title => SourceTitle.Length > 0 ? SourceTitle : Loc.T("Flows_CardUnnamed");
    public string Chevron => IsExpanded ? "" : "";
    // The SOURCE pane's list name: the real name, or a plain label when the list could not be read there.
    public string SourceListName => SourceReadable && SourceTitle.Length > 0 ? SourceTitle : Loc.T("Flows_CardNotReadable");
    // Heading over the column grid, naming both tenants so it is clear which side is which.
    public string ColumnsHeaderText => Loc.T("Flows_CardColumnsHeader", SourceHost.Length > 0 ? SourceHost : Loc.T("Flows_PaneSource"), TargetHost.Length > 0 ? TargetHost : Loc.T("Flows_PaneTarget"));
    // True when a valid target list (an ID) is chosen. Column counts only mean something for a list that has one.
    public bool HasTarget => FlowIdParser.ParseListId(TargetListId) is not null;
    public bool HasSourceNote => SourceNote.Length > 0;
    public bool HasTargetNote => TargetNote.Length > 0;
    public bool ShowReadSite => ReadSite.Length > 0 && !string.Equals(ReadSite.TrimEnd('/'), SourceSite.TrimEnd('/'), StringComparison.OrdinalIgnoreCase);
    public string FlowSitePath => Uri.TryCreate(SourceSite, UriKind.Absolute, out var u) ? u.AbsolutePath : SourceSite;
    public string UsedByChip => Loc.T("Flows_Chip_UsedBy", UsedByCount);
    public string ColumnsToggleText => ShowAllColumns ? Loc.T("Flows_ShowUsedOnly") : Loc.T("Flows_ShowAllColumns", Columns.Count);

    [RelayCommand]
    private void Toggle() => IsExpanded = !IsExpanded;

    partial void OnShowAllColumnsChanged(bool value)
    {
        ColumnsView.Refresh();
        UpdateColumnCounts();
    }

    partial void OnPickedChanged(TargetListOption? value)
    {
        if (_syncing || value is null) return;
        _syncing = true;
        try { TargetListId = value.Id; }
        finally { _syncing = false; }
        Edited();
    }

    partial void OnTargetListIdChanged(string value)
    {
        if (_syncing) return;
        RefreshPicked();
        Edited();
    }

    private void Edited()
    {
        if (Populating) return;
        UserEdited = true;
        TargetChanged?.Invoke(this);
    }

    // Selects the listed target list whose ID equals the typed ID or link, or nothing.
    public void RefreshPicked()
    {
        var id = FlowIdParser.ParseListId(TargetListId);
        var match = id is null ? null : TargetLists.FirstOrDefault(o => o.Id.Length > 0 && string.Equals(o.Id, id, StringComparison.OrdinalIgnoreCase));
        _syncing = true;
        try { Picked = match; }
        finally { _syncing = false; }
    }

    // Sets the target list without it counting as a user edit (the name match, or a loaded mapping).
    public void SetTarget(string? id)
    {
        Populating = true;
        try { TargetListId = id ?? string.Empty; RefreshPicked(); }
        finally { Populating = false; }
    }

    // Rebuilds the column rows from the source columns, the target list's columns and the names the flows use. Every column
    // of the source list gets a row (when the source could not be read, the rows are the names the flow uses); the target
    // column is the automatic match unless `preferred` (a loaded mapping or an earlier user choice) names one that exists on
    // the target.
    public void FillColumns(IReadOnlyList<FlowColumnInfo> targetCols, Func<string, string, List<ColumnOption>, string?>? preferred)
    {
        Columns.Clear();
        _targetCols = targetCols;
        var options = new List<ColumnOption> { new(string.Empty, Loc.T("Flows_LeaveUnchanged")) };
        options.AddRange(targetCols.OrderBy(t => t.DisplayName, StringComparer.OrdinalIgnoreCase)
            .Select(t => new ColumnOption(t.InternalName, t.DisplayName == t.InternalName ? t.InternalName : $"{t.DisplayName} ({t.InternalName})")));

        static int Rank(ColumnMatchStatus s) => s switch
        {
            ColumnMatchStatus.NotFoundInTarget => 0, ColumnMatchStatus.Similar => 1, ColumnMatchStatus.Renamed => 2, _ => 3,
        };

        var rows = new List<ColumnMapRow>();
        foreach (var src in SourceColumns)
        {
            var used = Used.Contains(src.InternalName);
            var (target, match) = targetCols.Count == 0
                ? ((string?)null, ColumnMatchStatus.NotFoundInTarget)
                : FlowMappingResolver.MatchColumn(src.InternalName, src, targetCols);
            var level = !used ? "Info" : match is ColumnMatchStatus.Same or ColumnMatchStatus.Renamed ? "Ready" : "Look";
            var row = new ColumnMapRow
            {
                SourceColumn = src.InternalName,
                SourceDisplay = src.DisplayName,
                SourceType = src.TypeName,
                IsUsed = used,
                Match = match,
                Options = options,
                StatusText = used ? Loc.T("Flows_ColumnStatus_" + match) : Loc.T("Flows_ColumnNotUsed"),
                StatusLevel = level,
            };
            row.SetAuto(preferred?.Invoke(SourceListId, src.InternalName, options) ?? target ?? string.Empty);
            row.UserChanged += () => Refresh(_targetColumnCount);
            rows.Add(row);
        }
        // Used columns first, the ones that need the user at the top; then the rest alphabetically.
        foreach (var r in rows.OrderBy(r => r.IsUsed ? 0 : 1).ThenBy(r => r.IsUsed ? Rank(r.Match) : 0).ThenBy(r => r.SourceDisplay, StringComparer.OrdinalIgnoreCase))
            Columns.Add(r);
        Refresh(targetCols.Count);
    }

    private void UpdateColumnCounts()
    {
        var visible = Columns.Count(r => ShowAllColumns || r.IsUsed);
        HasVisibleColumns = visible > 0;
        HasHiddenColumns = Columns.Count > Columns.Count(r => r.IsUsed);
        OnPropertyChanged(nameof(ColumnsToggleText));
    }

    // Built-in columns that exist on every list but are not in the editable-column listing; never a reason to stop.
    private static readonly HashSet<string> SystemColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "ID", "ContentType", "Author", "Editor", "Created", "Modified", "Attachments", "FileLeafRef", "_UIVersionString",
    };

    // What would make the flows fail with this target list, judged from the columns the target list really has:
    //  - no target list at all;
    //  - a column the flows use that is not on the target list (after the mapping);
    //  - a column the target list REQUIRES that a "Create item" step does not set.
    // Nothing is judged when the columns of the target list could not be read: that is shown as a warning instead.
    private List<string> FindBlockers(int targetColumnCount, bool hasTarget)
    {
        var found = new List<string>();
        if (!hasTarget) { SyncFixRows([]); found.Add(Loc.T("Flows_Block_NoTarget")); return found; }
        if (targetColumnCount == 0 || _targetCols.Count == 0) { SyncFixRows([]); return found; }

        var targetNames = _targetCols.Select(c => c.InternalName).ToHashSet(StringComparer.Ordinal);
        string Mapped(string name)
        {
            var row = Columns.FirstOrDefault(c => c.SourceColumn == name);
            return string.IsNullOrEmpty(row?.Selected) ? name : row.Selected;
        }
        var listName = Picked?.Title is { Length: > 0 } t ? t : TargetListId;

        foreach (var row in Columns.Where(c => c.IsUsed && !SystemColumns.Contains(c.SourceColumn)))
            if (!targetNames.Contains(Mapped(row.SourceColumn)))
                found.Add(Loc.T("Flows_Block_ColumnMissing", row.SourceLabel, listName));

        // A required column of the target list that a Create item step does not set: the import fails. A typed value fixes it
        // (the rewriter adds it to those steps); otherwise the list must make the column optional.
        var fixable = new List<FlowColumnInfo>();
        foreach (var required in _targetCols.Where(c => c.Required && !SystemColumns.Contains(c.InternalName)))
        {
            var lacking = CreateSteps.Where(s => s.ColumnsKnown && !s.Columns.Select(Mapped).Contains(required.InternalName)).ToList();
            if (lacking.Count == 0) continue;
            var canFix = FixableTypes.Contains(required.TypeName);
            if (canFix) fixable.Add(required);
            if (canFix && !string.IsNullOrWhiteSpace(_fixValues.GetValueOrDefault(required.InternalName))) continue;
            foreach (var step in lacking)
                found.Add(Loc.T(canFix ? "Flows_Block_RequiredMissing" : "Flows_Block_RequiredMissingNoFix", listName, required.DisplayName, step.FlowName, step.StepName));
        }
        SyncFixRows(fixable);
        return found.Distinct().ToList();
    }

    // One box per required column that needs a value; rebuilt only when that set changes, so typing is never interrupted.
    private void SyncFixRows(List<FlowColumnInfo> columns)
    {
        if (RequiredFixes.Select(r => r.Column).SequenceEqual(columns.Select(c => c.InternalName))) return;
        RequiredFixes.Clear();
        foreach (var c in columns)
        {
            var row = new RequiredFixRow { Column = c.InternalName, Display = c.DisplayName, Value = _fixValues.GetValueOrDefault(c.InternalName) ?? string.Empty };
            row.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName != nameof(RequiredFixRow.Value)) return;
                _fixValues[row.Column] = row.Value;
                Refresh(_targetColumnCount);
            };
            RequiredFixes.Add(row);
        }
    }

    // Recomputes the summary line, the chip, the collapsed header line, and why the grid is empty if it is. The chip is
    // never "Ready" unless the list really is: a target is chosen, its columns were read, and nothing needs a look.
    public void Refresh(int targetColumnCount)
    {
        _targetColumnCount = targetColumnCount;
        UpdateColumnCounts();
        ColumnsView.Refresh();

        var usedRows = Columns.Where(c => c.IsUsed).ToList();
        int Count(ColumnMatchStatus s) => usedRows.Count(c => c.Match == s);
        var check = usedRows.Count(c => c.NeedsLook);
        var targetId = FlowIdParser.ParseListId(TargetListId);
        var hasTarget = targetId is not null;
        var onTargetNow = AlreadyOnTarget && string.Equals(targetId, SourceListId, StringComparison.OrdinalIgnoreCase);
        var targetTitle = Picked?.Title is { Length: > 0 } t ? t : TargetListId;

        var unread = hasTarget && usedRows.Count > 0 && targetColumnCount == 0 ? Loc.T("Flows_CardTargetColumnsUnread") : string.Empty;
        TargetNote = string.Join(" ", new[] { TargetBaseNote, unread }.Where(s => s.Length > 0));

        SummaryText = Columns.Count == 0 ? Loc.T("Flows_CardSummaryNoColumns")
            : !hasTarget ? Loc.T("Flows_CardSummaryNoTarget")
            : onTargetNow ? Loc.T("Flows_CardSummaryOnTarget", usedRows.Count)
            : usedRows.Count == 0 ? Loc.T("Flows_CardSummaryNoUse", Columns.Count)
            : Loc.T("Flows_CardSummary", Columns.Count, usedRows.Count, Count(ColumnMatchStatus.Same), Count(ColumnMatchStatus.Renamed),
                    Count(ColumnMatchStatus.Similar), Count(ColumnMatchStatus.NotFoundInTarget));

        EmptyColumnsText = Columns.Count > 0 && !HasVisibleColumns ? Loc.T("Flows_CardEmptyNoneUsed")
            : Columns.Count > 0 ? string.Empty
            : !SourceReadable && !onTargetNow ? Loc.T("Flows_CardEmptyUnreadable")
            : Loc.T("Flows_CardEmptyNoColumns");

        var blockers = FindBlockers(targetColumnCount, hasTarget);
        BlockersText = string.Join(Environment.NewLine, blockers);
        BlockerCount = blockers.Count;

        (StatusLevel, StatusChip) =
            !hasTarget ? ("Error", Loc.T("Flows_Chip_NoTarget"))
            : blockers.Count > 0 ? ("Error", Loc.T("Flows_Chip_WillFail", blockers.Count))
            : onTargetNow ? ("Info", Loc.T("Flows_Chip_AlreadyTarget"))
            : unread.Length > 0 ? ("Look", Loc.T("Flows_Chip_TargetColsUnread"))
            : Columns.Count == 0 && !SourceReadable ? ("Error", Loc.T("Flows_Chip_CantRead"))
            : Columns.Count == 0 ? ("Ready", Loc.T("Flows_Chip_NothingToMap"))
            : check > 0 ? ("Look", Loc.T("Flows_Chip_ToCheck", check))
            : !SourceReadable ? ("Info", Loc.T("Flows_Chip_ByFlowNames"))
            : ("Ready", Loc.T("Flows_Chip_Ready"));
        NeedsAttention = StatusLevel is "Look" or "Error";
        HeaderStatus = hasTarget ? Loc.T("Flows_CardArrow", targetTitle) : Loc.T("Flows_CardNoTarget");
    }
}
