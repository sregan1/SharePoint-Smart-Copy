using CommunityToolkit.Mvvm.ComponentModel;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Dialogs;

// One side of the connection (source or target) as the Sites step shows it: the tenant address, the account, and a state
// chip. State is None (not signed in), Waiting (a sign-in window is open), Connected or Error.
public partial class SideConnection : ObservableObject
{
    [ObservableProperty] private string _host = string.Empty;
    [ObservableProperty] private string _account = string.Empty;
    // Extra line under the account: what the tool is waiting for, why a sign-in failed, or a warning.
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasNote))] private string _note = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(StatusLevel), nameof(StateText))] private string _state = "None";
    // Single tenant: both sides are the main window's sign-in; say so beside the account.
    [ObservableProperty] private string _suffix = string.Empty;

    public bool HasNote => Note.Length > 0;
    public bool HasHost => Host.Length > 0;
    public string StatusLevel => State switch { "Connected" => "Ready", "Waiting" => "Info", "Error" => "Error", _ => "None" };
    public string StateText => Loc.T("Flows_State_" + State);

    partial void OnHostChanged(string value) => OnPropertyChanged(nameof(HasHost));
}

// One package in the Packages step's list.
public sealed class PackageRow
{
    public string FileName { get; init; } = string.Empty;
    public string FullPath { get; init; } = string.Empty;
    public string FlowNames { get; init; } = string.Empty;
    public int SiteCount { get; init; }
    public int ListCount { get; init; }
    public string StatusText { get; init; } = string.Empty;
    // Ready / Look / Error, for the chip.
    public string StatusLevel { get; init; } = "Ready";
}

// One line of the "Ready to write?" checklist on the Write step. Step is the wizard step (0-based) that fixes it, or -1.
public sealed class ChecklistItem
{
    public string Text { get; init; } = string.Empty;
    public string StatusLevel { get; init; } = "Ready";   // Ready / Look / Info
    public int Step { get; init; } = -1;
    public bool HasStep => Step >= 0 && StatusLevel == "Look";
    public string Glyph => StatusLevel switch { "Ready" => "✓", "Look" => "⚠", _ => "ℹ" };
}

// One package on the Write step: what it is called now, and what the updated copy will be saved as. SaveAs is editable.
public sealed partial class OutputNameRow : CommunityToolkit.Mvvm.ComponentModel.ObservableObject
{
    public string Original { get; init; } = string.Empty;   // the file name now
    public string FullPath { get; init; } = string.Empty;
    [CommunityToolkit.Mvvm.ComponentModel.ObservableProperty] private string _saveAs = string.Empty;
}

// A value for a column the target list requires but a Create item step leaves out; the rewriter adds it to the step.
public sealed partial class RequiredFixRow : CommunityToolkit.Mvvm.ComponentModel.ObservableObject
{
    public string Column { get; init; } = string.Empty;    // internal name on the target list
    public string Display { get; init; } = string.Empty;
    [CommunityToolkit.Mvvm.ComponentModel.ObservableProperty] private string _value = string.Empty;
}

// One item to review after writing (the report step), with the step that can fix it when there is one.
public sealed class ReviewItem
{
    public string Text { get; init; } = string.Empty;
    public int Step { get; init; } = -1;
    public bool HasStep => Step >= 0;
}
