using System.ComponentModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;
using SharePointSmartCopy.Teams;

namespace SharePointSmartCopy.Dialogs;

// What the wizard's header, step indicator and footer bind to.
public sealed class TeamsWizardState : INotifyPropertyChanged
{
    public const int StepCount = 6;
    private static readonly string[] StepKeys = ["Teams_StepConnect", "Teams_StepChoose", "Teams_StepOptions", "Teams_StepReview", "Teams_StepCopy", "Teams_StepResults"];

    public event PropertyChangedEventHandler? PropertyChanged;

    public int CurrentStep { get; private set; }
    public bool IsBusy { get; private set; }
    public string StatusText { get; private set; } = string.Empty;
    public string BlockReason { get; private set; } = string.Empty;
    public bool HasBlockReason => BlockReason.Length > 0;
    public bool CanNext { get; private set; }
    public bool ShowBack => CurrentStep is 1 or 2 or 3;
    public bool ShowNext => CurrentStep is 0 or 1 or 2 or 3;
    public bool ShowCancel { get; private set; }
    public bool ShowCopyAnother => CurrentStep == 5;
    public string NextButtonText => CurrentStep == 3 ? Loc.T("Ui_StartCopy") : Loc.T("Ui_Next");
    public string StepText => Loc.T("Forms_StepOfN", CurrentStep + 1, StepCount, Loc.T(StepKeys[Math.Clamp(CurrentStep, 0, StepCount - 1)]));

    public void Apply(int step, bool busy, string status, string block, bool canNext, bool showCancel)
    {
        CurrentStep = step; IsBusy = busy; StatusText = status; BlockReason = block; CanNext = canNext; ShowCancel = showCancel;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(string.Empty));   // empty name: every property changed
    }
}

// The Teams tool, as a short wizard like Forms: choose a team or chat, set the options, review (people who have no match in the
// target are listed here, and can be mapped), copy, results. The target may be another tenant (people are then found by the
// identity mapping). Messages and chats need the target registration's client secret.
public partial class TeamsDialog : Window
{
    private TeamsApiClient _api;
    private readonly TeamsApiClient _delegatedApi;
    private bool _admin;                               // administrator view: every team, and the chats of chosen people (application sign-in)
    private TreeViewItem? _chatsRoot;
    private string? _adminFallbackNote;                // why the administrator view was dropped for the personal one, shown with the list
    private readonly HashSet<string> _seenChats = [];
    private readonly List<string> _chatPeople = [];
    private readonly AuthService _source;
    private readonly AuthService _target;
    private readonly bool _isMultiTenant;
    private readonly AzureRegistration? _targetRegistration;
    private readonly Func<IdentityMapper> _getMapper;
    private readonly Func<Window, IReadOnlyList<UnresolvedUser>, bool> _openMapping;
    private readonly SharePointService _sp;
    private readonly CopyService _copy;
    private readonly int _maxParallel;
    private readonly Func<CancellationToken, Task<string?>> _signInSource;
    private readonly AzureRegistration? _sourceRegistration;
    private readonly Func<bool, CancellationToken, Task<string?>> _getRoot;   // a side's SharePoint address, e.g. https://contoso.sharepoint.com
    private string? _srcRoot, _tgtRoot;
    private readonly Func<CancellationToken, Task<string?>> _signInTarget;
    private readonly TeamsWizardState _wizard = new();
    private readonly bool _hasSecret;

    private CancellationTokenSource? _loadCts;
    private CancellationTokenSource? _copyCts;
    private CancellationTokenSource? _peopleCts;
    private CancellationTokenSource? _srcSignInCts;
    private CancellationTokenSource? _tgtSignInCts;
    private string? _srcMessage, _tgtMessage;           // the last sign-in outcome per side (canceled, failed), shown until the next try
    private bool _treeStale = true;                     // the list of teams and chats must be loaded again (first visit, or another source account)

    private int _step;
    private bool _loading, _running, _checkingPeople;
    private string _status = string.Empty;
    private string _copyBlock = string.Empty;          // why copying is not possible (missing write permissions), when it is not
    private readonly List<TeamsApiClient.TeamInfo> _checkedTeams = [];   // the ticked teams (a chat is chosen by selecting it instead)
    private TreeViewItem? _teamsRoot;
    private List<TeamCopyPlanRow> _plans = [];          // the teams being copied and the names of their copies
    private sealed class TeamRun
    {
        public required TeamCopyPlanRow Plan { get; init; }
        public TeamsCopyService.Outcome? Outcome { get; set; }
        public string? Url { get; set; }
        public string? Error { get; set; }
        public bool Canceled { get; set; }
    }
    private List<TeamRun> _runs = [];
    private const int TeamParallelism = 2;              // teams copied at the same time (each waits a lot on Microsoft's provisioning)
    private List<TeamsCopyService.MemberPlan> _unresolved = [];
    private readonly StringBuilder _report = new();
    private string? _resultUrl;                      // where to open the copy: the team, or the chat
    private readonly List<TeamsApiClient.ChatInfo> _checkedChats = [];   // the ticked chats
    private readonly Dictionary<string, HashSet<string>> _channelsByTeam = [];   // team id -> the ticked channels (all of them until one is unticked)
    private readonly Dictionary<string, int> _channelTotals = [];
    private sealed class ChatRun
    {
        public required TeamsApiClient.ChatInfo Chat { get; init; }
        public string Label => ChatLabel(Chat);
        public TeamsCopyService.ChatOutcome? Outcome { get; set; }
        public string? Url { get; set; }
        public string? Error { get; set; }
        public bool Canceled { get; set; }
    }
    private List<ChatRun> _chatRuns = [];
    private bool _copyCanceled;
    private string? _copyError;                      // a failure that stopped the copy before it finished
    private System.Diagnostics.Stopwatch _copyWatch = new();

    public TeamsDialog(AuthService source, AuthService target, bool isMultiTenant, AzureRegistration? targetRegistration,
        Func<IdentityMapper> getMapper, Func<Window, IReadOnlyList<UnresolvedUser>, bool> openMapping,
        SharePointService sp, CopyService copy, int maxParallel,
        Func<CancellationToken, Task<string?>> signInSource, Func<CancellationToken, Task<string?>> signInTarget,
        AzureRegistration? sourceRegistration, Func<bool, CancellationToken, Task<string?>> getRoot)
    {
        InitializeComponent();
        DataContext = _wizard;
        _source = source;
        _target = target;
        _isMultiTenant = isMultiTenant;
        _targetRegistration = targetRegistration;
        _getMapper = getMapper;
        _openMapping = openMapping;
        _sp = sp;
        _copy = copy;
        _maxParallel = maxParallel;
        _signInSource = signInSource;
        _sourceRegistration = sourceRegistration;
        _getRoot = getRoot;
        _signInTarget = signInTarget;
        _api = _delegatedApi = new TeamsApiClient((ct, force) => source.GetGraphTokenAsync(TeamsApiClient.ReadScopes, ct, force));
        Closed += (_, _) => { _loadCts?.Cancel(); _copyCts?.Cancel(); _peopleCts?.Cancel(); _srcSignInCts?.Cancel(); _tgtSignInCts?.Cancel(); };

        // Messages and chats are imported with an application sign-in, which needs the target registration's client secret.
        _hasSecret = _targetRegistration?.GetClientSecret() is { Length: > 0 };
        OptMessages.IsEnabled = _hasSecret;
        OptMessages.IsChecked = false;
        OptMessages.ToolTip = _hasSecret ? Loc.T("Teams_OptMessagesTip") : Loc.T("Teams_Msg_NeedSecret");
        MessagesNote.Text = _hasSecret ? string.Empty : Loc.T("Teams_Msg_NeedSecret");
        MessagesNote.Visibility = _hasSecret ? Visibility.Collapsed : Visibility.Visible;

        // The administrator view reads with the SOURCE registration's own sign-in (its client secret), so it can see teams the signed-in person is not in.
        var adminOk = _sourceRegistration?.GetClientSecret() is { Length: > 0 };
        // It is the default whenever the source registration has a client secret; "only my own teams and chats" is the fallback (and the only view without a secret).
        _admin = adminOk;
        PersonalCheck.IsEnabled = adminOk;
        PersonalCheck.ToolTip = adminOk ? Loc.T("Teams_Personal_OptionTip") : Loc.T("Teams_Admin_NeedSecret");
        AdminNote.Text = adminOk ? string.Empty : Loc.T("Teams_Admin_NeedSecret");
        AdminNote.Visibility = adminOk ? Visibility.Collapsed : Visibility.Visible;

        Tree.SelectedItemChanged += (_, _) => RefreshWizard();
        SourceTitle.Text = Loc.T(_isMultiTenant ? "Teams_Connect_Source" : "Teams_Connect_Account");
        TargetCard.Visibility = IsTwoSignIns ? Visibility.Visible : Visibility.Collapsed;
        ConnectIntro.Text = Loc.T(IsTwoSignIns ? "Teams_Connect_IntroTwo" : "Teams_Connect_IntroOne");
        Loaded += async (_, _) =>
        {
            _ = LoadRootAsync(isSource: true);
            if (IsTwoSignIns) _ = LoadRootAsync(isSource: false);
            if (_target.IsAuthenticated) await CheckWritePermissionsAsync();
        };
        RefreshWizard();
    }

    private bool TeamMode => _checkedTeams.Count > 0;
    private bool HasChats => _checkedChats.Count > 0;
    // The target registration, named in the permission messages so it is clear which one to change in the Azure portal.
    private string TargetRegistrationLabel => _targetRegistration == null ? "" :
        (string.IsNullOrWhiteSpace(_targetRegistration.Name) ? "" : _targetRegistration.Name + " ") + $"(client ID {_targetRegistration.ClientId})";
    private string SourceRegistrationLabel => _sourceRegistration == null ? "" :
        (string.IsNullOrWhiteSpace(_sourceRegistration.Name) ? "" : _sourceRegistration.Name + " ") + $"(client ID {_sourceRegistration.ClientId})";

    // The administrator's read client: an application sign-in as the source registration, in the source tenant.
    private TeamsApiClient MakeAdminClient()
    {
        var reg = _sourceRegistration!;
        var app = new AppOnlyGraphToken(_source.TenantId ?? reg.TenantId, reg.ClientId, reg.GetClientSecret()!, "source");
        return new TeamsApiClient((t, force) => app.GetAsync(t, force), TeamsApiClient.AdminAccessMessageFor(SourceRegistrationLabel));
    }

    private bool IsTwoSignIns => _isMultiTenant && !ReferenceEquals(_source, _target);
    private bool IsCrossTenant => _isMultiTenant && !ReferenceEquals(_source, _target);

    // ── Wizard plumbing ──────────────────────────────────────────────────────

    private void RefreshWizard()
    {
        var canNext = false;
        var block = string.Empty;
        switch (_step)
        {
            case 0:
                var signingIn = _srcSignInCts != null || _tgtSignInCts != null;
                canNext = !signingIn && _source.IsAuthenticated && (!IsTwoSignIns || _target.IsAuthenticated);
                if (!canNext && !signingIn) block = Loc.T(IsTwoSignIns ? "Teams_Connect_NeedBoth" : "Teams_Connect_NeedOne");
                UpdateConnectCards();
                break;
            case 1:
                canNext = (TeamMode || HasChats) && !_loading;
                // The people filter is always shown; without the administrator view it is disabled and says what turns it on.
                AdminPanel.IsEnabled = _admin;
                var off = Loc.T("Teams_Admin_PanelOff");
                if (!_admin) PeopleChosenText.Text = off;
                else if (PeopleChosenText.Text == off) PeopleChosenText.Text = string.Empty;
                SelectionCount.Text = string.Join(", ", new[]
                {
                    TeamMode ? Loc.T("Teams_Selected", _checkedTeams.Count.ToString()) : "",
                    HasChats ? Loc.T("Teams_SelectedChats", _checkedChats.Count.ToString()) : "",
                }.Where(t => t.Length > 0));
                if (canNext && _copyBlock.Length > 0) { canNext = false; block = _copyBlock; }
                else if (canNext && HasChats && !_hasSecret) { canNext = false; block = Loc.T("Teams_Msg_NeedSecret"); }
                break;
            case 2:
                canNext = !TeamMode || NewNameBox.Text.Trim().Length > 0;
                if (!canNext) block = Loc.T("Teams_NameRequired");
                else if (_checkedTeams.Count > 1 && !NewNameBox.Text.Contains("{name}", StringComparison.Ordinal)) { canNext = false; block = Loc.T("Teams_NamePatternNeedsName"); }
                break;
            case 3:
                canNext = !_checkingPeople;
                if (canNext && TeamMode && NamesProblem() is { } problem) { canNext = false; block = problem; }
                break;
        }
        _wizard.Apply(_step, _loading || _running || _checkingPeople || _srcSignInCts != null || _tgtSignInCts != null, _status, block, canNext, showCancel: _step == 4 && _running);
    }

    private void SetStatus(string text) { _status = text; RefreshWizard(); }

    private void Log(string line) => Dispatcher.BeginInvoke(() =>
    {
        LogBox.AppendText(line + Environment.NewLine);
        LogBox.ScrollToEnd();
    });

    // A line for the Results step (and the log).
    private void Report(string line) { _report.AppendLine(line); Log(line); }

    private void NewNameBox_TextChanged(object sender, TextChangedEventArgs e) => RefreshWizard();

    private void Next_Click(object sender, RoutedEventArgs e)
    {
        switch (_step)
        {
            case 0:
                _step = 1;
                if (_treeStale) { _treeStale = false; _ = LoadAsync(); }
                break;
            case 1:
                PrepareOptions();
                _step = 2;
                break;
            case 2:
                BuildReview();
                _step = 3;
                _ = CheckPeopleAsync();
                break;
            case 3:
                _step = 4;
                _ = RunCopyAsync();
                break;
        }
        RefreshWizard();
    }

    private void Back_Click(object sender, RoutedEventArgs e)
    {
        if (_step is not (1 or 2 or 3)) return;
        if (_step == 3) _peopleCts?.Cancel();
        _step--;
        RefreshWizard();
    }

    // ── Ticking teams ────────────────────────────────────────────────────────

    private void TeamCheck_Changed(object sender, RoutedEventArgs e)
    {
        if (sender is not CheckBox { Tag: TeamsApiClient.TeamInfo team } box) return;
        if (box.IsChecked == true) { if (!_checkedTeams.Any(t => t.Id == team.Id)) _checkedTeams.Add(team); }
        else _checkedTeams.RemoveAll(t => t.Id == team.Id);
        RefreshWizard();
    }

    private void ChannelCheck_Changed(object sender, RoutedEventArgs e)
    {
        if (sender is not CheckBox { Tag: ValueTuple<string, string> key } box) return;
        if (!_channelsByTeam.TryGetValue(key.Item1, out var set)) _channelsByTeam[key.Item1] = set = [];
        if (box.IsChecked == true) set.Add(key.Item2); else set.Remove(key.Item2);
    }

    private void ChatCheck_Changed(object sender, RoutedEventArgs e)
    {
        if (sender is not CheckBox { Tag: TeamsApiClient.ChatInfo chat } box) return;
        if (box.IsChecked == true) { if (!_checkedChats.Any(t => t.Id == chat.Id)) _checkedChats.Add(chat); }
        else _checkedChats.RemoveAll(t => t.Id == chat.Id);
        RefreshWizard();
    }

    // The channels of a team that will be copied, or null when it is all of them.
    private IReadOnlySet<string>? ChannelFilter(string teamId)
        => _channelsByTeam.TryGetValue(teamId, out var set) && _channelTotals.TryGetValue(teamId, out var total) && set.Count < total ? set : null;

    private void SetAllTeams(bool on)
    {
        if (_teamsRoot == null) return;
        foreach (var item in _teamsRoot.Items.OfType<TreeViewItem>())
            if (item.Header is CheckBox box) box.IsChecked = on;
    }
    private void SelectAll_Click(object sender, RoutedEventArgs e) => SetAllTeams(true);
    private void ClearSelection_Click(object sender, RoutedEventArgs e)
    {
        SetAllTeams(false);
        foreach (var chat in _checkedChats.ToList())
            foreach (var item in ChatsRootItems()) if (item.Header is CheckBox { Tag: TeamsApiClient.ChatInfo c } box && c.Id == chat.Id) box.IsChecked = false;
    }
    private IEnumerable<TreeViewItem> ChatsRootItems() => Tree.Items.OfType<TreeViewItem>().Skip(1).SelectMany(r => r.Items.OfType<TreeViewItem>());

    private void CopyAnother_Click(object sender, RoutedEventArgs e)
    {
        _report.Clear();
        LogBox.Clear();
        _step = 1;
        RefreshWizard();
    }

    private void CancelCopy_Click(object sender, RoutedEventArgs e) => _copyCts?.Cancel();

    // ── Step 1: connect ──────────────────────────────────────────────────────

    private void UpdateConnectCards()
    {
        void Card(bool isSource)
        {
            var auth = isSource ? _source : _target;
            var busy = (isSource ? _srcSignInCts : _tgtSignInCts) != null;
            var message = isSource ? _srcMessage : _tgtMessage;
            var status = isSource ? SourceStatus : TargetStatus;
            var signIn = isSource ? SourceSignInButton : TargetSignInButton;
            var cancel = isSource ? SourceCancelButton : TargetCancelButton;
            status.Text = busy ? Loc.T("Teams_Connect_Waiting")
                : message ?? (auth.IsAuthenticated ? Loc.T("Teams_Connect_SignedInAs", auth.UserName ?? string.Empty) : Loc.T("Teams_Connect_NotSignedIn"));
            signIn.Content = Loc.T(auth.IsAuthenticated && message == null ? "Teams_Connect_SignInAgain" : message != null ? "Teams_Connect_Retry" : "Teams_Connect_SignIn");
            signIn.IsEnabled = !busy;
            signIn.Style = (Style)FindResource(auth.IsAuthenticated && message == null ? "SecondaryBtn" : "PrimaryBtn");
            cancel.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            // Light green once connected (signed in with no pending error); back to the default card look otherwise.
            var card = isSource ? SourceCard : TargetCard;
            if (auth.IsAuthenticated && message == null && !busy)
            {
                card.Background = Res("SuccessBgBrush");
                card.BorderBrush = Res("SuccessBrush");
            }
            else
            {
                card.ClearValue(Border.BackgroundProperty);
                card.ClearValue(Border.BorderBrushProperty);
            }
            var reg = isSource ? _sourceRegistration : _targetRegistration;
            var root = isSource ? _srcRoot : _tgtRoot;
            var lines = new List<string>();
            if (auth.IsAuthenticated) lines.Add(Loc.T("Teams_Connect_Tenant", root ?? Loc.T("Teams_Connect_TenantLooking")));
            if (reg != null) lines.Add(Loc.T("Teams_Connect_Registration", string.IsNullOrWhiteSpace(reg.Name) ? reg.ClientId : reg.Name, reg.ClientId));
            (isSource ? SourceDetail : TargetDetail).Text = string.Join(Environment.NewLine, lines);
        }
        Card(isSource: true);
        if (IsTwoSignIns) Card(isSource: false);
    }

    private async void SourceSignIn_Click(object sender, RoutedEventArgs e) => await SignInAsync(isSource: true);
    private async void TargetSignIn_Click(object sender, RoutedEventArgs e) => await SignInAsync(isSource: false);
    private void SourceCancel_Click(object sender, RoutedEventArgs e) => _srcSignInCts?.Cancel();
    private void TargetCancel_Click(object sender, RoutedEventArgs e) => _tgtSignInCts?.Cancel();

    // Opens the sign-in window for one side. Cancel stops waiting for it; Retry opens it again. Each side is independent, so a
    // sign-in to the target that was canceled leaves the source connected.
    private async Task SignInAsync(bool isSource)
    {
        if ((isSource ? _srcSignInCts : _tgtSignInCts) != null) return;
        var cts = new CancellationTokenSource();
        if (isSource) { _srcSignInCts = cts; _srcMessage = null; } else { _tgtSignInCts = cts; _tgtMessage = null; }
        RefreshWizard();
        try
        {
            await (isSource ? _signInSource : _signInTarget)(cts.Token);
            cts.Token.ThrowIfCancellationRequested();
            if (isSource) _srcRoot = null; else _tgtRoot = null;
            _ = LoadRootAsync(isSource);
            if (isSource || !IsTwoSignIns) _treeStale = true;   // another source account: its teams are different
            if (!isSource || !IsTwoSignIns) _copyBlock = string.Empty;
        }
        catch (OperationCanceledException) { SetSideMessage(isSource, Loc.T("Teams_Connect_Canceled")); }
        catch (Exception ex) { SetSideMessage(isSource, Loc.T("Teams_Connect_Failed", ex.Message.Split('\n')[0])); }
        finally
        {
            if (isSource) _srcSignInCts = null; else _tgtSignInCts = null;
            cts.Dispose();
        }
        if ((isSource ? _srcMessage : _tgtMessage) == null && (!isSource || !IsTwoSignIns)) await CheckWritePermissionsAsync();
        RefreshWizard();
    }

    // The side's SharePoint address, read once signed in; shown under the account so it is clear which tenant it is.
    private async Task LoadRootAsync(bool isSource)
    {
        if (!(isSource ? _source : _target).IsAuthenticated) return;
        try
        {
            var root = await _getRoot(isSource, CancellationToken.None);
            if (isSource) _srcRoot = root; else _tgtRoot = root;
        }
        catch (Exception) { /* the address is only shown; a failed lookup just leaves it out */ }
        UpdateConnectCards();
    }

    // Switching the view (the signed-in person's teams and chats, or the administrator's) makes the list load again.
    private void PersonalCheck_Changed(object sender, RoutedEventArgs e)
    {
        _admin = PersonalCheck.IsEnabled && PersonalCheck.IsChecked != true;
        _treeStale = true;
        _chatPeople.Clear();
        _seenChats.Clear();
        PersonResults.ItemsSource = null;
        PersonResults.Visibility = Visibility.Collapsed;
        PeopleChosenText.Text = string.Empty;
        RefreshWizard();
    }

    private void SetSideMessage(bool isSource, string text) { if (isSource) _srcMessage = text; else _tgtMessage = text; }

    // ── Step 3: options ──────────────────────────────────────────────────────

    private void PrepareOptions()
    {
        TeamOptions.Visibility = TeamMode ? Visibility.Visible : Visibility.Collapsed;
        ChatOptionsText.Visibility = HasChats ? Visibility.Visible : Visibility.Collapsed;
        TargetText.Text = Loc.T(IsCrossTenant ? "Teams_TargetOtherTenant" : "Teams_Target", _target.UserName ?? string.Empty);
        if (NewNameBox.Text.Length == 0) NewNameBox.Text = Loc.T("Teams_DefaultCopyPattern");
    }

    // ── Step 4: review (and the people check) ────────────────────────────────

    private void BuildReview()
    {
        var sb = new StringBuilder();
        if (TeamMode)
        {
            var pattern = NewNameBox.Text.Trim();
            _plans = _checkedTeams.OrderBy(t => t.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .Select(t => new TeamCopyPlanRow(t, pattern.Replace("{name}", t.DisplayName, StringComparison.Ordinal).Trim())).ToList();
            foreach (var plan in _plans)
            {
                var total = _channelTotals.GetValueOrDefault(plan.Team.Id);
                var picked = _channelsByTeam.TryGetValue(plan.Team.Id, out var chosen) ? chosen.Count : total;
                plan.ChannelsText = total == 0 ? string.Empty : Loc.T("Teams_Review_ChannelsOf", picked.ToString(), total.ToString());
            }
            ReviewNames.ItemsSource = _plans;
            var parts = new List<string>();
            if (OptSettings.IsChecked == true) parts.Add(Loc.T("Teams_OptSettings"));
            if (OptMembers.IsChecked == true) parts.Add(Loc.T("Teams_OptMembers"));
            if (OptFiles.IsChecked == true) parts.Add(Loc.T("Teams_OptFiles"));
            if (OptMessages.IsChecked == true) parts.Add(Loc.T("Teams_OptMessages"));
            sb.AppendLine(_plans.Count == 1 ? Loc.T("Teams_Review_Team", _plans[0].SourceName, _plans[0].NewName) : Loc.T("Teams_Review_Teams", _plans.Count.ToString()));
            sb.AppendLine(Loc.T("Teams_Review_NamesHint"));
            sb.AppendLine();
            sb.Append(Loc.T("Teams_Review_Includes", parts.Count == 0 ? Loc.T("Teams_Review_Nothing") : string.Join(", ", parts)));
        }
        if (HasChats)
        {
            if (TeamMode) sb.AppendLine().AppendLine();
            sb.Append(_checkedChats.Count == 1 ? Loc.T("Teams_Review_Chat", ChatLabel(_checkedChats[0]))
                : Loc.T("Teams_Review_Chats", _checkedChats.Count.ToString(), string.Join(", ", _checkedChats.Select(ChatLabel))));
        }
        sb.AppendLine().AppendLine();
        sb.Append(TargetText.Text);
        ReviewText.Text = sb.ToString();
        ReviewNames.Visibility = TeamMode ? Visibility.Visible : Visibility.Collapsed;
        PeopleCard.Visibility = Visibility.Collapsed;
    }

    // Why the names of the copies cannot be used yet (empty, or two the same), or null when they are fine.
    private string? NamesProblem()
    {
        if (_plans.Any(p => string.IsNullOrWhiteSpace(p.NewName))) return Loc.T("Teams_NameRequired");
        var dup = _plans.GroupBy(p => p.NewName.Trim(), StringComparer.CurrentCultureIgnoreCase).FirstOrDefault(g => g.Count() > 1);
        return dup == null ? null : Loc.T("Teams_NameDuplicate", dup.Key);
    }

    // A name edited in the grid is saved when the edit ends; the Next button then checks the names again.
    private void ReviewNames_CellEditEnding(object? sender, DataGridCellEditEndingEventArgs e)
        => Dispatcher.BeginInvoke(new Action(RefreshWizard), System.Windows.Threading.DispatcherPriority.Background);

    private static string ChatLabel(TeamsApiClient.ChatInfo chat) =>
        !string.IsNullOrWhiteSpace(chat.Topic) ? chat.Topic! : !string.IsNullOrWhiteSpace(chat.People) ? chat.People! : chat.ChatType;

    // Two tenants: who from the source team has no matching account in the target? They are left out unless they are mapped here.
    private async Task CheckPeopleAsync()
    {
        _peopleCts?.Cancel();
        if (!TeamMode || !IsCrossTenant || OptMembers.IsChecked != true)
        {
            PeopleCard.Visibility = Visibility.Collapsed;
            return;
        }
        _peopleCts = new CancellationTokenSource();
        var ct = _peopleCts.Token;
        PeopleCard.Visibility = Visibility.Visible;
        PeopleText.Text = Loc.T("Teams_Mapping_Checking");
        MapPeopleButton.Visibility = Visibility.Collapsed;
        _checkingPeople = true;
        RefreshWizard();
        try
        {
            var (src, tgt, import) = BuildClients(OptMessages.IsChecked == true, chats: false);
            // every team of the batch, a few at a time, with one shared lookup cache (a person in ten teams is looked up once)
            var service = NewService(src, tgt, import, copy: _copy, people: new TeamsPeopleCache());
            var perTeam = new List<TeamsCopyService.MemberPlan>[_plans.Count];
            await Parallel.ForEachAsync(Enumerable.Range(0, _plans.Count), new ParallelOptions { MaxDegreeOfParallelism = 3, CancellationToken = ct },
                async (i, t) => perTeam[i] = await service.PlanMembersAsync(_plans[i].Team, t));
            var plan = perTeam.SelectMany(p => p).GroupBy(p => p.Member.UserId, StringComparer.OrdinalIgnoreCase).Select(g => g.First()).ToList();
            _unresolved = plan.Where(p => !p.Resolved).ToList();
            if (_unresolved.Count == 0)
            {
                PeopleText.Text = Loc.T("Teams_Mapping_AllFound", plan.Count.ToString());
                return;
            }
            var names = string.Join(Environment.NewLine, _unresolved.Take(15).Select(p => "  " + (p.Member.Email ?? p.Member.DisplayName)))
                        + (_unresolved.Count > 15 ? Environment.NewLine + Loc.T("Teams_Mapping_More", (_unresolved.Count - 15).ToString()) : "");
            PeopleText.Text = Loc.T("Teams_Review_Unresolved", _unresolved.Count.ToString(), plan.Count.ToString(), names);
            MapPeopleButton.Visibility = Visibility.Visible;
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { PeopleText.Text = ex.Message; }   // the copy itself reports a real problem again
        finally
        {
            if (!ct.IsCancellationRequested) _checkingPeople = false;
            RefreshWizard();
        }
    }

    private async void MapPeople_Click(object sender, RoutedEventArgs e)
    {
        var users = _unresolved.Select(p => new UnresolvedUser(p.Member.Email ?? p.Member.DisplayName, p.MappedEmail ?? string.Empty, Loc.T("Teams_Title"), 1)).ToList();
        if (users.Count == 0) return;
        if (_openMapping(this, users)) await CheckPeopleAsync();   // a fresh check picks up the saved mapping
    }

    // ── Step 6: results ──────────────────────────────────────────────────────

    private Brush Res(string key) => (Brush)FindResource(key);

    // Builds the Results page like the Files report: an outcome banner, number tiles, and one status line per part of the copy.
    private void BuildResults(string headline)
    {
        var tiles = new List<TeamsResultTile>();
        var rows = new List<TeamsResultRow>();
        TeamsResultTile Tile(string value, string label, string kind) => kind switch
        {
            "Success" => new(value, label, Res("SuccessBgBrush"), Res("SuccessBrush"), Res("SuccessBrush")),
            "Danger"  => new(value, label, Res("DangerBgBrush"), Res("DangerBrush"), Res("DangerBrush")),
            "Warning" => new(value, label, Res("WarningBgBrush"), Res("WarningBrush"), Res("WarningTextBrush")),
            _         => new(value, label, Res("WindowBgBrush"), Res("BorderLightBrush"), Res("TextSecondaryBrush")),
        };
        void Row(string kind, string area, string message) => rows.Add(new TeamsResultRow(
            kind switch { "Success" => "✔", "Danger" => "✖", "Warning" => "⚠", _ => "ℹ" },
            kind switch { "Success" => Res("SuccessBrush"), "Danger" => Res("DangerBrush"), "Warning" => Res("WarningTextBrush"), _ => Res("TextSecondaryBrush") },
            area, message));
        // A warning reads "Part: text"; the part before the first colon becomes the area when it is short, else it is general.
        void WarningRows(IEnumerable<string> warnings)
        {
            foreach (var w in warnings)
            {
                var colon = w.IndexOf(':');
                if (colon is > 0 and <= 32) Row("Warning", w[..colon], w[(colon + 1)..].Trim());
                else Row("Warning", Loc.T("Teams_Res_AreaGeneral"), w);
            }
        }

        var failed = 0;
        var leftOut = 0;
        var kind = "Success";
        if (_runs.Count > 0 || _chatRuns.Count > 0)
        {
            var batch = _runs.Count + _chatRuns.Count > 1;
            var doneTeams = _runs.Where(r => r.Outcome != null).ToList();
            var doneChats = _chatRuns.Where(r => r.Outcome != null).ToList();
            var messages = OptMessages.IsChecked == true;
            failed = doneTeams.Sum(r => r.Outcome!.FilesFailed + r.Outcome.MessagesFailed) + doneChats.Sum(r => r.Outcome!.MessagesFailed)
                     + _runs.Count(r => r.Error != null) + _chatRuns.Count(r => r.Error != null);
            leftOut = doneTeams.SelectMany(r => r.Outcome!.UnresolvedPeople).Concat(doneChats.SelectMany(r => r.Outcome!.UnresolvedPeople)).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            var warnings = doneTeams.Sum(r => r.Outcome!.Warnings.Count) + doneChats.Sum(r => r.Outcome!.Warnings.Count);
            if (batch && _runs.Count > 0) tiles.Add(Tile($"{doneTeams.Count}/{_runs.Count}", Loc.T("Teams_Res_Teams"), doneTeams.Count == _runs.Count ? "Success" : "Danger"));
            if (batch && _chatRuns.Count > 0) tiles.Add(Tile($"{doneChats.Count}/{_chatRuns.Count}", Loc.T("Teams_Res_Chats"), doneChats.Count == _chatRuns.Count ? "Success" : "Danger"));
            if (_runs.Count > 0)
            {
                tiles.Add(Tile(doneTeams.Sum(r => r.Outcome!.ChannelsCreated + r.Outcome.ChannelsReused).ToString(), Loc.T("Teams_Res_Channels"), "Success"));
            }
            tiles.Add(Tile((doneTeams.Sum(r => r.Outcome!.MembersAdded) + doneChats.Sum(r => r.Outcome!.MembersMapped)).ToString(), Loc.T("Teams_Res_Members"), "Success"));
            if (_runs.Count > 0) tiles.Add(Tile(doneTeams.Sum(r => r.Outcome!.FilesCopied).ToString(), Loc.T("Teams_Res_Files"), "Success"));
            if ((messages && _runs.Count > 0) || _chatRuns.Count > 0)
                tiles.Add(Tile((doneTeams.Sum(r => r.Outcome!.MessagesImported) + doneChats.Sum(r => r.Outcome!.MessagesImported)).ToString(), Loc.T("Teams_Res_Messages"), "Success"));
            tiles.Add(Tile(failed.ToString(), Loc.T("Teams_Res_Failed"), failed > 0 ? "Danger" : "Neutral"));
            tiles.Add(Tile(leftOut.ToString(), Loc.T("Teams_Res_LeftOut"), leftOut > 0 ? "Warning" : "Neutral"));

            void Warn(Func<string, string> area, IEnumerable<string> list)
            {
                foreach (var w in list)
                {
                    var colon = w.IndexOf(':');
                    var part = colon is > 0 and <= 32 ? w[..colon] : Loc.T("Teams_Res_AreaGeneral");
                    Row("Warning", area(part), colon is > 0 and <= 32 ? w[(colon + 1)..].Trim() : w);
                }
            }
            foreach (var run in _runs)
            {
                var name = run.Plan.NewName.Trim();
                string Area(string part) => batch ? name + " · " + part : part;
                if (run.Outcome is not { } o)
                {
                    Row(run.Canceled ? "Info" : "Danger", Area(Loc.T("Teams_Res_AreaTeam")), run.Error ?? (run.Canceled ? Loc.T("Teams_CopyCancelled") : Loc.T("Teams_Res_NotStarted")));
                    continue;
                }
                Row("Success", Area(Loc.T("Teams_Res_AreaTeam")), Loc.T("Teams_Res_RowTeam", name, _target.TenantId ?? "?", _target.UserName ?? "?") + (run.Url != null && batch ? " " + run.Url : string.Empty));
                Row("Success", Area(Loc.T("Teams_Res_Channels")), Loc.T("Teams_Res_RowChannels", o.ChannelsCreated.ToString(), o.ChannelsReused.ToString()));
                Row(OptMembers.IsChecked == true ? "Success" : "Info", Area(Loc.T("Teams_Res_Members")),
                    OptMembers.IsChecked == true ? Loc.T("Teams_Res_RowMembers", o.MembersAdded.ToString()) : Loc.T("Teams_Res_NotCopied"));
                Row(OptFiles.IsChecked != true ? "Info" : o.FilesFailed > 0 ? "Danger" : "Success", Area(Loc.T("Teams_Res_Files")),
                    OptFiles.IsChecked == true ? Loc.T("Teams_Res_RowFiles", o.FilesCopied.ToString(), o.FilesFailed.ToString()) : Loc.T("Teams_Res_NotCopied"));
                if (messages)
                    Row(o.MessagesFailed > 0 ? "Danger" : "Success", Area(Loc.T("Teams_Res_Messages")),
                        Loc.T("Teams_Res_RowMessages", o.MessagesImported.ToString(), o.MessagesSkipped.ToString(), o.MessagesFailed.ToString(), o.MessagesUnmapped.ToString()));
                else Row("Info", Area(Loc.T("Teams_Res_Messages")), Loc.T("Teams_Res_NotCopied"));
                foreach (var p in o.UnresolvedPeople) Row("Warning", Area(Loc.T("Teams_Res_AreaPeople")), Loc.T("Teams_Res_RowLeftOut", p));
                if (batch) Warn(Area, o.Warnings); else WarningRows(o.Warnings);
            }
            foreach (var run in _chatRuns)
            {
                string Area(string part) => batch ? run.Label + " · " + part : part;
                if (run.Outcome is not { } c)
                {
                    Row(run.Canceled ? "Info" : "Danger", Area(Loc.T("Teams_Res_AreaChat")), run.Error ?? (run.Canceled ? Loc.T("Teams_CopyCancelled") : Loc.T("Teams_Res_NotStarted")));
                    continue;
                }
                Row("Success", Area(Loc.T("Teams_Res_AreaChat")), Loc.T("Teams_Res_RowChat", run.Label, _target.TenantId ?? "?", _target.UserName ?? "?") + (run.Url != null && batch ? " " + run.Url : string.Empty));
                Row("Success", Area(Loc.T("Teams_Res_Members")), Loc.T("Teams_Res_RowMembers", c.MembersMapped.ToString()));
                Row(c.MessagesFailed > 0 ? "Danger" : "Success", Area(Loc.T("Teams_Res_Messages")),
                    Loc.T("Teams_Res_RowMessages", c.MessagesImported.ToString(), c.MessagesSkipped.ToString(), c.MessagesFailed.ToString(), c.MessagesUnmapped.ToString()));
                foreach (var p in c.UnresolvedPeople) Row("Warning", Area(Loc.T("Teams_Res_AreaPeople")), Loc.T("Teams_Res_RowLeftOut", p));
                if (batch) Warn(Area, c.Warnings); else WarningRows(c.Warnings);
            }
            kind = _copyCanceled ? "Neutral" : doneTeams.Count + doneChats.Count == 0 ? "Danger" : failed > 0 ? "Danger" : leftOut > 0 || warnings > 0 ? "Warning" : "Success";
        }
        else
        {
            kind = _copyCanceled ? "Neutral" : "Danger";
            Row(_copyCanceled ? "Info" : "Danger", Loc.T("Teams_Res_AreaGeneral"), _copyError ?? headline);
        }
        tiles.Add(Tile(_copyWatch.Elapsed.TotalHours >= 1 ? _copyWatch.Elapsed.ToString(@"h\:mm\:ss") : _copyWatch.Elapsed.ToString(@"m\:ss"), Loc.T("Teams_Res_Duration"), "Neutral"));

        ResultHeadline.Text = headline;
        ResultSubline.Text = _runs.Any(r => r.Outcome != null) || _chatRuns.Any(r => r.Outcome != null) ? Loc.T("Teams_Result_WhereTenant", _target.TenantId ?? "?", _target.UserName ?? "?") : string.Empty;
        ResultSubline.Visibility = ResultSubline.Text.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
        ResultBanner.Background = kind switch { "Success" => Res("SuccessBgBrush"), "Danger" => Res("DangerBgBrush"), "Warning" => Res("WarningBgBrush"), _ => Res("SurfaceBrush") };
        ResultBanner.BorderBrush = kind switch { "Success" => Res("SuccessBrush"), "Danger" => Res("DangerBrush"), "Warning" => Res("WarningBrush"), _ => Res("BorderLightBrush") };
        // Open in Teams only when there is exactly one team or chat; several have their links in the list
        _resultUrl = _runs.Count + _chatRuns.Count == 1 ? _runs.FirstOrDefault()?.Url ?? _chatRuns.FirstOrDefault()?.Url : null;
        OpenTeamButton.Visibility = _resultUrl != null ? Visibility.Visible : Visibility.Collapsed;
        ResultTiles.ItemsSource = tiles;
        ResultRows.ItemsSource = rows;
    }

    private void OpenTeam_Click(object sender, RoutedEventArgs e)
    {
        if (_resultUrl == null) return;
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(_resultUrl) { UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, Title, MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    private void CopyReport_Click(object sender, RoutedEventArgs e)
    {
        try { Clipboard.SetText(_report.ToString()); } catch (Exception) { /* the clipboard can be busy: nothing to do */ }
    }

    // ── Steps 5 and 6: copy, results ─────────────────────────────────────────

    // The clients for one run. One tenant: the same scopes for both sides. Two tenants: each side asks only for its own.
    private (TeamsApiClient Source, TeamsApiClient Target, TeamsApiClient? Import) BuildClients(bool messages, bool chats)
    {
        var cross = IsCrossTenant;
        var srcScopes = TeamsApiClient.SourceScopes(messages, chats);
        var tgtScopes = TeamsApiClient.TargetScopes(cross, chats);
        if (!cross) srcScopes = tgtScopes = srcScopes.Concat(tgtScopes).Distinct().ToArray();
        var src = _admin ? MakeAdminClient() : new TeamsApiClient((t, force) => _source.GetGraphTokenAsync(srcScopes, t, force));
        var tgt = new TeamsApiClient((t, force) => _target.GetGraphTokenAsync(tgtScopes, t, force));
        TeamsApiClient? import = null;
        if ((messages || chats) && _targetRegistration?.GetClientSecret() is { Length: > 0 } secret)
        {
            var tenant = _target.TenantId ?? _targetRegistration.TenantId;
            var app = new AppOnlyGraphToken(tenant, _targetRegistration.ClientId, secret);
            import = new TeamsApiClient((t, force) => app.GetAsync(t, force), AppOnlyGraphToken.AccessMessageFor(TargetRegistrationLabel));
        }
        return (src, tgt, import);
    }

    // The identity mapper comes from the main window's view model, which only the UI thread may ask (it reads a WPF object). A team of a
    // batch is copied on a pool thread, so the mapper is fetched on the UI thread whoever asks.
    private IdentityMapper GetMapper() => Dispatcher.CheckAccess() ? _getMapper() : Dispatcher.Invoke(_getMapper);

    private TeamsCopyService NewService(TeamsApiClient src, TeamsApiClient tgt, TeamsApiClient? import, CopyService? copy = null, TeamsPeopleCache? people = null)
        => new(src, tgt, import, _sp, copy ?? _copy, _target.UserName ?? string.Empty, _maxParallel, GetMapper(), IsCrossTenant, people);

    private async Task RunCopyAsync()
    {
        _report.Clear();
        LogBox.Clear();
        _runs = []; _chatRuns = []; _resultUrl = null; _copyCanceled = false; _copyError = null;
        _copyWatch = System.Diagnostics.Stopwatch.StartNew();
        _copyCts = new CancellationTokenSource();
        var ct = _copyCts.Token;
        _running = true;
        RefreshWizard();
        var headline = Loc.T("Teams_CopyDone");
        try
        {
            if (TeamMode) await CopyTeamsAsync(ct);
            if (HasChats) await CopyChatsAsync(ct);
            headline = ComposeHeadline();
        }
        catch (OperationCanceledException) { headline = Loc.T("Teams_CopyCancelled"); _copyCanceled = true; Report(headline); }
        catch (Microsoft.Identity.Client.MsalServiceException ex)
        {
            headline = Loc.T("Teams_CopyPermissionTip", TargetRegistrationLabel);
            _copyError = headline + " (" + ex.Message.Split('\n')[0] + ")";
            Report(_copyError);
        }
        catch (TeamsApiClient.TeamsAccessException ex) { headline = ex.Message; _copyError = ex.Message; Report(ex.Message); }
        catch (Exception ex) { headline = Loc.T("Teams_CopyError", ex.Message); _copyError = headline; Report(headline); }
        finally
        {
            _copyCts?.Dispose();
            _copyCts = null;
            _running = false;
            _step = 5;
            _copyWatch.Stop();
            BuildResults(headline);
            RefreshWizard();
        }
    }

    // Copies the ticked teams, TeamParallelism at a time (a team spends much of its time waiting on Microsoft: the new team, each channel's
    // files folder), each with its own copy engine so two teams never share one's state. A team that fails does not stop the others;
    // Cancel stops after the ones running. Returns the headline for the Results page.
    private async Task CopyTeamsAsync(CancellationToken ct)
    {
        var options = new TeamsCopyService.Options(string.Empty, OptSettings.IsChecked == true, OptMembers.IsChecked == true,
            OptFiles.IsChecked == true, OptMessages.IsChecked == true);
        // The copy asks for the write permissions as well; the first time that shows one consent screen.
        var (src, tgt, import) = BuildClients(options.CopyMessages, chats: false);
        var people = new TeamsPeopleCache();
        var runs = _plans.Select(p => new TeamRun { Plan = p }).ToList();
        _runs = runs;
        var batch = runs.Count > 1;
        var finished = 0;
        var fatal = false;
        await Parallel.ForEachAsync(runs, new ParallelOptions { MaxDegreeOfParallelism = batch ? TeamParallelism : 1, CancellationToken = ct }, async (run, t) =>
        {
            if (fatal) return;   // a permission problem would hit every team the same way: stop starting new ones
            var name = run.Plan.NewName.Trim();
            var tag = batch ? $"[{run.Plan.SourceName}] " : string.Empty;
            var progress = new Progress<string>(m => Log(tag + m));
            Log(Loc.T("Teams_Batch_Starting", (runs.IndexOf(run) + 1).ToString(), runs.Count.ToString(), run.Plan.SourceName, name));
            try
            {
                // one copy engine per team when several run at once
                var copy = batch ? new CopyService(_sp, new MigrationJobService(_sp)) : _copy;
                var outcome = await NewService(src, tgt, import, copy, people).CopyAsync(run.Plan.Team, options with { NewName = name, ChannelIds = ChannelFilter(run.Plan.Team.Id) }, progress, t);
                run.Outcome = outcome;
                run.Url = await ReportWhereAsync(tgt, outcome.TargetTeamId, t, name);
                lock (_report)
                {
                    _report.AppendLine(tag + Loc.T("Teams_Summary", outcome.TargetTeamId, outcome.ChannelsCreated.ToString(), outcome.ChannelsReused.ToString(),
                        outcome.MembersAdded.ToString(), outcome.FilesCopied.ToString(), outcome.FilesFailed.ToString()));
                    if (options.CopyMessages)
                        _report.AppendLine(tag + Loc.T("Teams_SummaryMessages", outcome.MessagesImported.ToString(), outcome.MessagesSkipped.ToString(),
                            outcome.MessagesFailed.ToString(), outcome.MessagesUnmapped.ToString()));
                    foreach (var p in outcome.UnresolvedPeople) _report.AppendLine(tag + "  - " + Loc.T("Teams_Res_RowLeftOut", p));
                    foreach (var w in outcome.Warnings) _report.AppendLine(tag + "  - " + w);
                }
            }
            catch (OperationCanceledException) { run.Canceled = true; throw; }
            catch (Microsoft.Identity.Client.MsalServiceException ex)
            {
                run.Error = Loc.T("Teams_CopyPermissionTip", TargetRegistrationLabel) + " (" + ex.Message.Split('\n')[0] + ")";
                fatal = true;
            }
            catch (TeamsApiClient.TeamsAccessException ex) { run.Error = ex.Message; fatal = true; }
            catch (Exception ex) { run.Error = Loc.T("Teams_CopyError", ex.Message); }
            finally
            {
                Interlocked.Increment(ref finished);
                if (run.Error != null) lock (_report) _report.AppendLine(tag + run.Error);
            }
        });
    }

    // The headline of the Results page: one team or chat says how it went; several say how many were copied.
    private string ComposeHeadline()
    {
        var total = _runs.Count + _chatRuns.Count;
        var ok = _runs.Count(r => r.Outcome != null) + _chatRuns.Count(r => r.Outcome != null);
        if (total == 1) return _runs.FirstOrDefault()?.Error ?? _chatRuns.FirstOrDefault()?.Error ?? Loc.T("Teams_CopyDone");
        if (_runs.Count > 0 && _chatRuns.Count > 0) return ok == total ? Loc.T("Teams_Mixed_Done", ok.ToString()) : Loc.T("Teams_Mixed_Partial", ok.ToString(), total.ToString());
        if (_chatRuns.Count > 0) return ok == total ? Loc.T("Teams_Chats_Done", ok.ToString()) : Loc.T("Teams_Chats_Partial", ok.ToString(), total.ToString());
        return ok == total ? Loc.T("Teams_Batch_Done", ok.ToString()) : Loc.T("Teams_Batch_Partial", ok.ToString(), total.ToString());
    }

    // Says WHERE the new team is (tenant, owner, a link), so a team that does not show up can be traced to the right tenant.
    private async Task<string?> ReportWhereAsync(TeamsApiClient tgt, string teamId, CancellationToken ct, string name)
    {
        try
        {
            var team = await tgt.GetTeamAsync(teamId, ct);
            var webUrl = team["webUrl"]?.GetValue<string>();
            lock (_report) { _report.AppendLine(name + ": " + Loc.T("Teams_Result_Where", _target.TenantId ?? "?", _target.UserName ?? "?", webUrl ?? "(no link returned)")); }
            Log(Loc.T("Teams_Result_Where", _target.TenantId ?? "?", _target.UserName ?? "?", webUrl ?? "(no link returned)"));
            return webUrl;
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception) { return null; }
    }

    // Copies the ticked chats one after another, each into its own new group chat. A chat that fails does not stop the others.
    private async Task CopyChatsAsync(CancellationToken ct)
    {
        var (src, tgt, import) = BuildClients(messages: true, chats: true);
        var people = new TeamsPeopleCache();
        var runs = _checkedChats.Select(c => new ChatRun { Chat = c }).ToList();
        _chatRuns = runs;
        var batch = _runs.Count + runs.Count > 1;
        foreach (var run in runs)
        {
            ct.ThrowIfCancellationRequested();
            var tag = batch ? $"[{run.Label}] " : string.Empty;
            Log(Loc.T("Teams_Batch_StartingChat", (runs.IndexOf(run) + 1).ToString(), runs.Count.ToString(), run.Label));
            try
            {
                var o = await NewService(src, tgt, import, null, people).CopyChatAsync(run.Chat, new Progress<string>(m => Log(tag + m)), ct);
                run.Outcome = o;
                if (!string.IsNullOrEmpty(o.TargetChatId))
                    run.Url = "https://teams.microsoft.com/l/chat/" + Uri.EscapeDataString(o.TargetChatId) + "/0" + (string.IsNullOrEmpty(_target.TenantId) ? "" : "?tenantId=" + _target.TenantId);
                lock (_report)
                {
                    _report.AppendLine(tag + Loc.T("Teams_ChatSummary", o.TargetChatId, o.MembersMapped.ToString(), o.MessagesImported.ToString(),
                        o.MessagesSkipped.ToString(), o.MessagesFailed.ToString(), o.MessagesUnmapped.ToString()));
                    foreach (var p in o.UnresolvedPeople) _report.AppendLine(tag + "  - " + Loc.T("Teams_Res_RowLeftOut", p));
                    foreach (var w in o.Warnings) _report.AppendLine(tag + "  - " + w);
                }
            }
            catch (OperationCanceledException) { run.Canceled = true; throw; }
            catch (Microsoft.Identity.Client.MsalServiceException ex)
            {
                run.Error = Loc.T("Teams_CopyPermissionTip", TargetRegistrationLabel) + " (" + ex.Message.Split('\n')[0] + ")";
                lock (_report) _report.AppendLine(tag + run.Error);
                break;   // the same permission problem would stop every other chat
            }
            catch (TeamsApiClient.TeamsAccessException ex) { run.Error = ex.Message; lock (_report) _report.AppendLine(tag + ex.Message); break; }
            catch (Exception ex) { run.Error = Loc.T("Teams_CopyError", ex.Message); lock (_report) _report.AppendLine(tag + run.Error); }
        }
    }

    private void ReportDetails(IReadOnlyCollection<string> unresolved, IReadOnlyCollection<string> warnings)
    {
        if (unresolved.Count > 0)
        {
            Report(Loc.T("Teams_UnresolvedPeople", unresolved.Count.ToString()));
            foreach (var p in unresolved) Report("  - " + p);
        }
        if (warnings.Count > 0)
        {
            Report(Loc.T("Teams_Warnings", warnings.Count.ToString()));
            foreach (var w in warnings) Report("  - " + w);
        }
    }

    // ── Step 2: loading the teams and chats ──────────────────────────────────

    // Copying needs the write permissions on the TARGET sign-in: when the granted ones lack any, the wizard cannot go past the
    // first step and says which. A check that cannot be made (null) leaves copying possible; a real attempt then reports any problem.
    private async Task CheckWritePermissionsAsync()
    {
        _copyBlock = string.Empty;
        try
        {
            if (await _target.GetGrantedGraphScopesAsync(forceRefresh: true) is not { } granted) return;
            var missing = TeamsApiClient.WriteScopes.Where(s => !granted.Contains(s)).ToList();
            if (missing.Count == 0) return;
            _copyBlock = Loc.T("Teams_CopyPermissionMissing", TargetRegistrationLabel, string.Join(", ", missing));
            RefreshWizard();
        }
        catch (OperationCanceledException) { }
    }

    private async void Load_Click(object sender, RoutedEventArgs e) => await LoadAsync();

    private async Task LoadAsync()
    {
        _loadCts?.Cancel();
        _loadCts = new CancellationTokenSource();
        var ct = _loadCts.Token;
        _loading = true;
        var fallback = false;
        LoadButton.IsEnabled = false;
        Tree.Items.Clear();
        SetStatus(Loc.T("Teams_Loading"));
        try
        {
            var teamsRoot = new TreeViewItem { Header = Loc.T("Teams_TeamsNode"), IsExpanded = true };
            _teamsRoot = teamsRoot;
            _checkedTeams.Clear();
            _checkedChats.Clear();
            _channelsByTeam.Clear();
            _channelTotals.Clear();
            Tree.Items.Add(teamsRoot);
            _api = _admin ? MakeAdminClient() : _delegatedApi;   // built now: the source tenant is known once signed in
            _chatPeople.Clear();
            _seenChats.Clear();
            PeopleChosenText.Text = string.Empty;
            var teams = _admin ? await _api.ListAllTeamsAsync(ct) : await _api.ListJoinedTeamsAsync(ct);
            var ordered = teams.OrderBy(t => t.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();

            // Every team's channels, General channel, and each channel's members and Files folder are read at the same time (a few
            // requests in flight at once) instead of one after another; the tree is built afterward in order.
            var gate = new SemaphoreSlim(8);
            var done = 0;
            async Task<T> Limited<T>(Func<Task<T>> call)
            {
                await gate.WaitAsync(ct);
                try { return await call(); }
                finally { gate.Release(); }
            }
            var loaded = await Task.WhenAll(ordered.Select(async team =>
            {
                var channels = await Limited(() => _api.ListChannelsAsync(team.Id, ct));
                var primary = Limited(() => _api.GetPrimaryChannelIdAsync(team.Id, ct));
                var details = await Task.WhenAll(channels.Select(async ch => (
                    Members: ch.Kind == TeamsApiClient.ChannelKind.Standard
                        ? string.Empty
                        // Standard channels share the team's member list, so only private and shared ones are counted.
                        : Loc.T("Teams_Members", (await Limited(() => SafeMembersAsync(team.Id, ch.Id, ct))).ToString()),
                    Files: await Limited(() => _api.GetFilesFolderAsync(team.Id, ch.Id, ct)) is null
                        ? Loc.T("Teams_FilesNo") : Loc.T("Teams_FilesYes"))));
                var n = Interlocked.Increment(ref done);
                Dispatcher.Invoke(() => SetStatus(Loc.T("Teams_LoadingTeam", $"{team.DisplayName} ({n}/{ordered.Count})")));
                return (team, channels, primaryId: await primary, details);
            }));

            foreach (var (team, channels, primaryId, details) in loaded)
            {
                var tick = new CheckBox { Content = team.DisplayName, Tag = team, IsChecked = _checkedTeams.Any(t => t.Id == team.Id) };
                tick.Checked += TeamCheck_Changed;
                tick.Unchecked += TeamCheck_Changed;
                var teamItem = new TreeViewItem { Header = tick, Tag = team };
                teamsRoot.Items.Add(teamItem);
                _channelTotals[team.Id] = channels.Count;
                _channelsByTeam[team.Id] = channels.Select(c => c.Id).ToHashSet();
                for (var ci = 0; ci < channels.Count; ci++)
                {
                    var ch = channels[ci];
                    var (members, files) = details[ci];
                    var parts = new[] { KindText(ch.Kind), members, files }.Where(s => s.Length > 0);
                    // Each channel has a tick box (all ticked); General is always copied, so its box is fixed
                    var isGeneral = ch.Id == primaryId;
                    var channelTick = new CheckBox
                    {
                        Content = $"{ch.DisplayName}  —  {string.Join(" · ", parts)}" + (isGeneral ? "  " + Loc.T("Teams_AlwaysCopied") : ""),
                        Tag = (team.Id, ch.Id), IsChecked = true, IsEnabled = !isGeneral,
                    };
                    channelTick.Checked += ChannelCheck_Changed;
                    channelTick.Unchecked += ChannelCheck_Changed;
                    teamItem.Items.Add(new TreeViewItem { Header = channelTick });
                }
            }
            if (teams.Count == 0) teamsRoot.Items.Add(new TreeViewItem { Header = Loc.T("Teams_NoTeams") });

            var chatsRoot = new TreeViewItem { Header = Loc.T(_admin ? "Teams_Admin_ChatsNode" : "Teams_ChatsNode"), IsExpanded = _admin };
            _chatsRoot = chatsRoot;
            Tree.Items.Add(chatsRoot);
            // The administrator view has no chats of its own: they are added for the people chosen above the list.
            if (!_admin)
                foreach (var chat in await _api.ListChatsAsync(ct)) AddChatNode(chatsRoot, chat);

            _status = Loc.T("Teams_Done", teams.Count.ToString(), chatsRoot.Items.Count.ToString()) + (_adminFallbackNote == null ? string.Empty : "  " + _adminFallbackNote);
            _adminFallbackNote = null;
        }
        catch (OperationCanceledException) { _status = string.Empty; _treeStale = true; }
        catch (Exception ex) when (_admin && !ct.IsCancellationRequested && ex is InvalidOperationException or Microsoft.Identity.Client.MsalServiceException)
        {
            // The administrator sign-in failed (most often a wrong or expired client secret): do not stop there, list the person's own teams and chats.
            var first = ex.Message.Split((char)10)[0];
            _adminFallbackNote = Loc.T("Teams_Admin_Fallback", first.Contains("AADSTS7000215", StringComparison.Ordinal)
                ? Loc.T("Teams_Admin_FallbackSecret") : first.Length > 160 ? first[..160] + "…" : first);
            fallback = true;
        }
        catch (TeamsApiClient.TeamsAccessException ex) { _status = ex.Message; _treeStale = true; }
        catch (Microsoft.Identity.Client.MsalServiceException)
        {
            // The registration lacks a Teams permission (or consent was refused): say what to add.
            _status = Loc.T("Teams_PermissionMissingTip");
        }
        catch (Exception ex) { _status = Loc.T("Teams_Error", ex.Message); }
        finally
        {
            if (!ct.IsCancellationRequested) _loading = false;
            LoadButton.IsEnabled = true;
            RefreshWizard();
        }
        if (fallback)
        {
            PersonalCheck.IsChecked = true;   // the personal view from now on (until the secret is fixed and this is unticked)
            await LoadAsync();
            _treeStale = false;
        }
    }

    private void AddChatNode(TreeViewItem root, TeamsApiClient.ChatInfo chat)
    {
        if (!_seenChats.Add(chat.Id)) return;   // a chat between two chosen people is listed once
        var isMeeting = string.Equals(chat.ChatType, "meeting", StringComparison.OrdinalIgnoreCase);   // Microsoft cannot import into these
        var who = !string.IsNullOrWhiteSpace(chat.Topic) ? chat.Topic : chat.People;
        var chatText = string.IsNullOrWhiteSpace(who) ? chat.ChatType : $"{who}  —  {chat.ChatType}";
        if (isMeeting) { root.Items.Add(new TreeViewItem { Header = chatText }); return; }
        var chatTick = new CheckBox { Content = chatText, Tag = chat, IsChecked = _checkedChats.Any(t => t.Id == chat.Id) };
        chatTick.Checked += ChatCheck_Changed;
        chatTick.Unchecked += ChatCheck_Changed;
        root.Items.Add(new TreeViewItem { Header = chatTick });
    }

    // ── Administrator view: choosing people whose chats are listed ───────────

    private void PersonBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Enter) { e.Handled = true; _ = FindPeopleAsync(); }
    }
    private async void FindPerson_Click(object sender, RoutedEventArgs e) => await FindPeopleAsync();
    private void PersonResults_SelectionChanged(object sender, SelectionChangedEventArgs e) => AddChatsButton.IsEnabled = PersonResults.SelectedItems.Count > 0;

    private async Task FindPeopleAsync()
    {
        var text = PersonBox.Text.Trim();
        if (text.Length == 0 || !_admin) return;
        FindPersonButton.IsEnabled = false;
        try
        {
            var found = await _api.SearchUsersAsync(text);
            PersonResults.ItemsSource = found.Select(u => new PersonRow(u)).ToList();
            PersonResults.Visibility = found.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            PeopleChosenText.Text = found.Count == 0 ? Loc.T("Teams_Admin_NoPeople", text) : Loc.T("Teams_Admin_PickPeople");
            if (found.Count == 1) PersonResults.SelectedIndex = 0;
        }
        catch (TeamsApiClient.TeamsAccessException ex) { PeopleChosenText.Text = ex.Message; }
        catch (Exception ex) { PeopleChosenText.Text = Loc.T("Teams_Error", ex.Message); }
        finally { FindPersonButton.IsEnabled = true; }
    }

    // Lists the chats of the people picked in the results and adds them to the Chats node (no chat twice).
    private async void AddChats_Click(object sender, RoutedEventArgs e)
    {
        var picked = PersonResults.SelectedItems.OfType<PersonRow>().ToList();
        if (picked.Count == 0 || _chatsRoot == null) return;
        AddChatsButton.IsEnabled = false;
        FindPersonButton.IsEnabled = false;
        var added = 0;
        try
        {
            foreach (var person in picked)
            {
                SetStatus(Loc.T("Teams_Admin_LoadingChats", person.User.DisplayName));
                var before = _chatsRoot.Items.Count;
                foreach (var chat in await _api.ListUserChatsAsync(person.User.Id)) AddChatNode(_chatsRoot, chat);
                added += _chatsRoot.Items.Count - before;
                if (!_chatPeople.Contains(person.User.DisplayName, StringComparer.OrdinalIgnoreCase)) _chatPeople.Add(person.User.DisplayName);
            }
            _chatsRoot.IsExpanded = true;
            PeopleChosenText.Text = Loc.T("Teams_Admin_ChatsLoaded", string.Join(", ", _chatPeople), _chatsRoot.Items.Count.ToString());
            SetStatus(Loc.T("Teams_Admin_ChatsAdded", added.ToString()));
        }
        catch (TeamsApiClient.TeamsAccessException ex) { PeopleChosenText.Text = ex.Message; SetStatus(string.Empty); }
        catch (Exception ex) { PeopleChosenText.Text = Loc.T("Teams_Error", ex.Message); SetStatus(string.Empty); }
        finally
        {
            FindPersonButton.IsEnabled = true;
            AddChatsButton.IsEnabled = PersonResults.SelectedItems.Count > 0;
        }
    }

    // A channel whose members cannot be read still gets listed; its count shows as 0.
    private async Task<int> SafeMembersAsync(string teamId, string channelId, CancellationToken ct)
    {
        try { return (await _api.ListChannelMembersAsync(teamId, channelId, ct)).Count; }
        catch (OperationCanceledException) { throw; }
        catch { return 0; }
    }

    private static string KindText(TeamsApiClient.ChannelKind kind) => kind switch
    {
        TeamsApiClient.ChannelKind.Private => Loc.T("Teams_KindPrivate"),
        TeamsApiClient.ChannelKind.Shared  => Loc.T("Teams_KindShared"),
        _                                  => Loc.T("Teams_KindStandard"),
    };
}
