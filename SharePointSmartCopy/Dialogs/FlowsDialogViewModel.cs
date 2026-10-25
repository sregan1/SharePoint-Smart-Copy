using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Workflows;

namespace SharePointSmartCopy.Dialogs;

// A short message with a level (Ready / Look / Error / Info), shown as a status chip.
public partial class StatusMessage : ObservableObject
{
    [ObservableProperty] private string _text = string.Empty;
    [ObservableProperty] private string _statusLevel = "Info";
}

// One numbered line of "What to do next" on the report step.
public sealed record NextStepItem(string Number, string Text);

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
//
// Where text goes (so the same message is never shown three times):
//   banners   = a condition that needs the user, at most one amber banner per step;
//   chips     = the state of one item (a list, a site, a package...);
//   activity  = the diary of what the tool did, one line, expandable;
//   footer    = what is happening right now, cleared when it finishes.
public partial class FlowsDialogViewModel : ObservableObject
{
    public FlowsDialogViewModel()
    {
        PackageRows.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasPackages));
        ReviewItems.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasReviewItems));
        TargetConn.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(SideConnection.State)) { OnPropertyChanged(nameof(ShowSignInTargetFirst)); OnPropertyChanged(nameof(ShowChangeTarget)); }
        };
        MatchSummary.PropertyChanged += (_, _) => OnPropertyChanged(nameof(HasMatchSummary));
    }

    public ObservableCollection<PackageResultRow> PackageResults { get; } = [];
    public ObservableCollection<PackageRow> PackageRows { get; } = [];
    public ObservableCollection<ChecklistItem> Checklist { get; } = [];
    public ObservableCollection<ReviewItem> ReviewItems { get; } = [];
    public ObservableCollection<NextStepItem> NextSteps { get; } = [];
    public SideConnection SourceConn { get; } = new();
    public SideConnection TargetConn { get; } = new();
    public StatusMessage MatchSummary { get; } = new();

    [ObservableProperty, NotifyPropertyChangedFor(nameof(ShowBack), nameof(ShowNext), nameof(ShowStartOver), nameof(NextButtonText), nameof(CanNext), nameof(ShowSubtitle))]
    private int _currentStep;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext))] private bool _isBusy;
    [ObservableProperty] private string _statusText = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasNotes))] private string _notesText = string.Empty;
    public bool HasNotes => NotesText.Length > 0;

    // The "What happens here" text of the current step is open. Closes whenever the step changes.
    [ObservableProperty] private bool _helpOpen;
    // The header subtitle is only needed on the first step.
    public bool ShowSubtitle => CurrentStep == 0;

    public void AddNote(string line)
    {
        if (string.IsNullOrWhiteSpace(line)) return;
        var existing = NotesText.Split('\n');
        if (!existing.Contains(line)) NotesText = NotesText.Length == 0 ? line : NotesText + "\n" + line;
    }

    public void ClearNotes() => NotesText = string.Empty;

    // ── Activity strip: the whole log, and the latest line shown while the strip is closed ──
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasActivity))] private string _activityText = string.Empty;
    [ObservableProperty] private string _activityLatest = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(ActivityChevron))] private bool _isActivityOpen;
    public bool HasActivity => ActivityText.Length > 0;
    // Up while closed (opens upward), down while open.
    public string ActivityChevron => IsActivityOpen ? "" : "";
    private readonly Queue<string> _activityLines = new();

    // `important` (a failure, a cancellation) opens the strip so it cannot be missed.
    public void AddActivity(string line, bool important = false)
    {
        if (string.IsNullOrWhiteSpace(line)) return;
        var stamped = $"{DateTime.Now:HH:mm:ss}  {line.Trim()}";
        _activityLines.Enqueue(stamped);
        while (_activityLines.Count > 300) _activityLines.Dequeue();
        ActivityText = string.Join("\n", _activityLines);
        ActivityLatest = stamped;
        if (important) IsActivityOpen = true;
    }

    public void ClearActivity()
    {
        _activityLines.Clear();
        ActivityText = string.Empty;
        ActivityLatest = string.Empty;
    }

    // ── Forms & email step ──
    // Result of loading the target account's forms, and why the email domains were prefilled.
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasTargetFormsText))] private string _targetFormsText = string.Empty;
    public bool HasTargetFormsText => TargetFormsText.Length > 0;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasDomainNote))] private string _domainNote = string.Empty;
    public bool HasDomainNote => DomainNote.Length > 0;
    // Something on this step was filled in automatically from the sign-ins (so the "clear" link shows).
    [ObservableProperty] private bool _hasAutofill;

    // ── Write step ──
    // One row per package with an editable "Save as" name.
    public ObservableCollection<OutputNameRow> OutputNames { get; } = [];
    // ── Sites step ──
    // Which side is signing in right now ("source", "target" or ""), so that side's button shows Cancel instead.
    [ObservableProperty, NotifyPropertyChangedFor(nameof(ShowSignInSource), nameof(ShowCancelSource), nameof(ShowSignInTarget),
        nameof(ShowCancelTarget), nameof(IsConnecting), nameof(ShowSignInTargetFirst), nameof(ShowChangeTarget))] private string _connectingSide = string.Empty;
    public bool IsConnecting => ConnectingSide.Length > 0;
    public bool ShowSignInSource => IsMultiTenant && ConnectingSide != "source";
    public bool ShowCancelSource => ConnectingSide == "source";
    public bool ShowSignInTarget => IsMultiTenant && ConnectingSide != "target";
    // The target button reads "Sign in to target" until connected, then "Change target" (like the source button).
    public bool ShowSignInTargetFirst => ShowSignInTarget && TargetConn.State != "Connected";
    public bool ShowChangeTarget => ShowSignInTarget && TargetConn.State == "Connected";
    public bool ShowCancelTarget => ConnectingSide == "target";
    [ObservableProperty, NotifyPropertyChangedFor(nameof(IsSingleTenant), nameof(ShowSignInSource), nameof(ShowSignInTarget), nameof(ShowSignInTargetFirst), nameof(ShowChangeTarget))] private bool _isMultiTenant;
    public bool IsSingleTenant => !IsMultiTenant;
    // The one warning shown under the connections, chosen by priority.
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasSourceWarning))] private string _sourceWarning = string.Empty;
    public bool HasSourceWarning => SourceWarning.Length > 0;
    // The warning is about packages that already point at the target, so offer to remove them.
    [ObservableProperty] private bool _showRemoveTargetPackages;
    // The manual "new tenant address" box, shown only when the address could not be read from the sign-in.
    [ObservableProperty] private bool _showManualHost;
    // The advanced "read lists from" column of the sites table.
    [ObservableProperty] private bool _showReadFrom;

    // ── Packages step ──
    public bool HasPackages => PackageRows.Count > 0;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext))] private bool _hasScan;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext))] private int _flowCount;
    [ObservableProperty] private string _scanSummaryText = string.Empty;
    [ObservableProperty] private string _resourcesHeaderText = string.Empty;
    // Shown when an added package is one this tool already wrote.
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasPackageWarning))] private string _packageWarning = string.Empty;
    public bool HasPackageWarning => PackageWarning.Length > 0;

    // ── Lists step ──
    public bool HasMatchSummary => MatchSummary.Text.Length > 0;
    // True once the lists have been matched and their cards are shown.
    [ObservableProperty] private bool _hasCards;

    // ── Sites / Forms bookkeeping ──
    [ObservableProperty, NotifyPropertyChangedFor(nameof(NoSites))] private bool _hasSites;
    public bool NoSites => HasScan && !HasSites;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(NoForms))] private bool _hasForms;
    public bool NoForms => HasScan && !HasForms;
    [ObservableProperty] private bool _hasValues;

    // ── Write / Report ──
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanNext))] private string _outputPath = string.Empty;
    [ObservableProperty] private string _outcome = "Success";
    [ObservableProperty] private string _resultHeadline = string.Empty;
    [ObservableProperty] private string _resultSubline = string.Empty;
    [ObservableProperty] private string _durationLabel = string.Empty;
    [ObservableProperty] private int _packagesWritten;
    [ObservableProperty] private int _replacements;
    [ObservableProperty] private int _reviewCount;
    public bool HasReviewItems => ReviewItems.Count > 0;

    public bool ShowBack => CurrentStep is >= 1 and <= 5;
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

    partial void OnCurrentStepChanged(int value)
    {
        HelpOpen = false;
        OnPropertyChanged(nameof(CanNext));
    }

    partial void OnHasScanChanged(bool value) { OnPropertyChanged(nameof(NoSites)); OnPropertyChanged(nameof(NoForms)); }

    // Fills the report step. The review items are what the tool could not settle by itself, each with the step that can fix it.
    public void Finish(int packages, int replacements, IReadOnlyList<ReviewItem> review, TimeSpan elapsed, string outDir)
    {
        PackagesWritten = packages;
        Replacements = replacements;
        ReviewItems.Clear();
        foreach (var r in review) ReviewItems.Add(r);
        ReviewCount = review.Count;
        DurationLabel = Loc.T("Flows_Took", elapsed.TotalMinutes >= 1
            ? $"{(int)elapsed.TotalMinutes}m {elapsed.Seconds}s"
            : $"{Math.Max(1, (int)Math.Round(elapsed.TotalSeconds))}s");

        if (review.Count == 0)
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

        NextSteps.Clear();
        for (var i = 1; i <= 4; i++) NextSteps.Add(new NextStepItem(i.ToString(), Loc.T("Flows_NextStep" + i)));
        CurrentStep = 5;
    }

    public void Reset()
    {
        CurrentStep = 0;
        PackageRows.Clear(); HasScan = false; FlowCount = 0; ScanSummaryText = string.Empty; ResourcesHeaderText = string.Empty;
        HasSites = false; HasCards = false; HasForms = false; HasValues = false; HasAutofill = false;
        MatchSummary.Text = string.Empty;
        TargetFormsText = string.Empty; DomainNote = string.Empty;
        ShowManualHost = false; ShowReadFrom = false; SourceWarning = string.Empty;
        Checklist.Clear(); ReviewItems.Clear(); NextSteps.Clear();
        PackageResults.Clear();
        ClearNotes();
        ClearActivity();
        IsActivityOpen = false;
    }
}

public partial class SiteRow : ObservableObject
{
    // The address found in the flows: what gets replaced, so it cannot be edited.
    public string Source { get; init; } = string.Empty;
    // Where the tool reads this site's lists from on the source. The same address by default; change it when the flow's
    // address is old or renamed.
    [ObservableProperty] private string _readFrom = string.Empty;
    // True once the user typed a "read from" address (so the automatic fill never overwrites it).
    public bool ReadFromEdited { get; private set; }
    private bool _autoRead;

    // The automatic fill (or the default, the flow's own address) setting the address: not a user edit.
    public void SetReadFromAuto(string value)
    {
        _autoRead = true;
        try { ReadFrom = value; }
        finally { _autoRead = false; }
    }

    partial void OnReadFromChanged(string value)
    {
        if (!_autoRead) ReadFromEdited = true;
    }

    [ObservableProperty, NotifyPropertyChangedFor(nameof(Effect), nameof(StatusLevel))] private string _target = string.Empty;

    // What the Write step will do with this site, so an empty cell is never a surprise.
    public string Effect => string.IsNullOrWhiteSpace(Target) ? Loc.T("Flows_SiteEffectUnchanged") : Loc.T("Flows_SiteEffectMapped");
    public string StatusLevel => string.IsNullOrWhiteSpace(Target) ? "Look" : "Ready";
}

// A form in the target account that a flow's form can be pointed at.
public sealed record TargetFormOption(string Id, string Title);

// One form used by the flows. The new form is chosen from the target account's forms (Picked) or typed/pasted as an ID or
// edit-page link (NewId); the two stay in step: picking a form fills NewId, and a typed ID that matches a listed form
// selects it.
public partial class FormMapRow : ObservableObject
{
    private bool _syncing;

    // Set before NewId, so the first sync has the list to look in. Shared by every row; filled when the target forms load.
    public ObservableCollection<TargetFormOption> Options { get; init; } = [];
    public string SourceId { get; init; } = string.Empty;
    // The source form's title: from the Forms copy tool's record, or read from Microsoft Forms with "Load form names". Empty
    // until known (the flow itself only holds the ID).
    [ObservableProperty, NotifyPropertyChangedFor(nameof(DisplayTitle))] private string _sourceTitle = string.Empty;
    public string DisplayTitle => SourceTitle.Length > 0 ? SourceTitle : Loc.T("Flows_FormTitleUnknown");
    // Forms IDs start with the same tenant and owner characters, so the END is what tells two forms apart.
    public string ShortId => SourceId.Length > 22 ? SourceId[..6] + "…" + SourceId[^12..] : SourceId;
    [ObservableProperty] private string _newId = string.Empty;
    [ObservableProperty] private TargetFormOption? _picked;
    // Where the ID came from, or why it looks wrong (wrong tenant, not an ID), and how serious that is (Ready / Look).
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasNote))] private string _note = string.Empty;
    [ObservableProperty] private string _statusLevel = "Info";
    public bool HasNote => Note.Length > 0;

    partial void OnPickedChanged(TargetFormOption? value)
    {
        if (_syncing || value is null) return;
        _syncing = true;
        try { NewId = value.Id; }
        finally { _syncing = false; }
    }

    partial void OnNewIdChanged(string value)
    {
        if (_syncing) return;
        RefreshPicked();
    }

    // Selects the listed form whose ID equals the typed ID or link, or nothing.
    public void RefreshPicked()
    {
        var id = FlowIdParser.ParseFormId(NewId);
        var match = id is null ? null : Options.FirstOrDefault(o => o.Id.Length > 0 && o.Id == id);
        _syncing = true;
        try { Picked = match; }
        finally { _syncing = false; }
    }
}

// A value in the flows that can be mapped by typing its replacement: an email address, the tenant ID or a library path.
public partial class ValueMapRow : ObservableObject
{
    public string Category { get; init; } = string.Empty;   // Email | Tenant | Path
    public string KindText { get; init; } = string.Empty;
    public string Source { get; init; } = string.Empty;
    [ObservableProperty] private string _target = string.Empty;
    // Set when the tool filled the replacement in for you (from your sign-ins), so it can be checked.
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasNote))] private string _note = string.Empty;
    public bool HasNote => Note.Length > 0;
    public string StatusLevel => "Ready";

    // Any later edit makes the "filled in for you" note untrue; the autofill sets the note after the target.
    partial void OnTargetChanged(string value) => Note = string.Empty;
}
