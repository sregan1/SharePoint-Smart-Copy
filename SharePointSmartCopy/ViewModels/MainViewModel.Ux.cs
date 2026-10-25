// Partial class MainViewModel: user-experience helpers — confirmations, "why is Next disabled",
// friendly connection errors, registration pickers, header account text and taskbar progress.
using System.Windows.Shell;
using CommunityToolkit.Mvvm.ComponentModel;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;

namespace SharePointSmartCopy.ViewModels;

public partial class MainViewModel
{
    // ── Confirmations ─────────────────────────────────────────────────────────
    // The view model never shows UI itself: the window installs this. When nothing is installed
    // (demo mode, headless use) every action simply proceeds.
    public Func<string, string, bool>? ConfirmAction { get; set; }

    private bool Confirm(string titleKey, string bodyKey) =>
        ConfirmAction == null || ConfirmAction(Loc.T(titleKey), Loc.T(bodyKey));

    // True while a run is in flight and closing the window / cancelling would throw work away.
    public bool IsRunInProgress => IsCopying || IsUpdatingMetadata;

    // Asked before the window closes during a run. Returns false to veto the close.
    public bool ConfirmCloseDuringRun() =>
        !IsRunInProgress || Confirm("Dlg_StopCopyTitle", "Dlg_CloseDuringCopyBody");

    // ── Why is Next disabled? ─────────────────────────────────────────────────
    // The Next button used to just grey out. This names the missing step next to it. Refreshed from
    // NextCommand.CanExecuteChanged (hooked in the constructor), which every input that can change
    // CanGoNext already raises.
    public string NextBlockedReason
    {
        get
        {
            if (IsDemoMode || CanGoNext()) return string.Empty;
            return CurrentStep switch
            {
                0 => Loc.T("Why_ConnectSource"),
                1 => Loc.T("Why_SelectItems"),
                2 => !TargetConnected                                                      ? Loc.T("Why_ConnectTarget")
                   : NeedsTargetFolder && SelectedTargetFolder == null                     ? Loc.T("Why_ChooseFolder")
                   : IsItemSelectionActive && ItemDestinationNodes.Any(n => n.DestinationList == null) ? Loc.T("Why_ChooseList")
                   : ItemDestinationNodes.Any(n => n.IsCreatingNewDestination && string.IsNullOrWhiteSpace(n.NewDestinationName)) ? Loc.T("Why_NameList")
                   : string.Empty,
                3 => Loc.T("Why_NothingToCopy"),
                4 => IsRunInProgress ? Loc.T("Why_CopyRunning") : string.Empty,
                _ => string.Empty,
            };
        }
    }

    // ── Connection errors ─────────────────────────────────────────────────────
    [ObservableProperty] private bool _sourceHasError;
    [ObservableProperty] private bool _sourceNeedsSettings;
    [ObservableProperty] private bool _targetHasError;
    [ObservableProperty] private bool _targetNeedsSettings;

    // Any status change clears the error flags; SetSourceError/SetTargetError set them AFTER the text.
    partial void OnSourceStatusChanged(string value) { SourceHasError = false; SourceNeedsSettings = false; }
    partial void OnTargetStatusChanged(string value) { TargetHasError = false; TargetNeedsSettings = false; }

    private void SetSourceError(Exception ex)
    {
        var (message, needsSettings) = ConnectionErrors.Translate(ex, AuthService.IsConfigured);
        SourceStatus        = message;
        SourceHasError      = true;
        SourceNeedsSettings = needsSettings;
    }

    private void SetTargetError(Exception ex)
    {
        var (message, needsSettings) = ConnectionErrors.Translate(ex, TargetAuthService.IsConfigured);
        TargetStatus        = message;
        TargetHasError      = true;
        TargetNeedsSettings = needsSettings;
    }

    // ── First run / settings changes ──────────────────────────────────────────
    // No app registration yet → the first thing a new user must do is open Settings.
    public bool IsNotConfigured => !Settings.IsConfigured;

    // A persistent message at the top of the Source step (the 3-second toast was easy to miss).
    [ObservableProperty] private string _connectionNotice = string.Empty;

    // ── Registration pickers (Multi-tenant) ───────────────────────────────────
    public IEnumerable<string> RegistrationNames => Settings.Registrations.Select(r => r.DisplayName);

    // Item 0 = "same as source"; item N+1 = registration N.
    public IEnumerable<string> TargetRegistrationChoices =>
        new[] { Loc.T("Ui_SameAsSource") }.Concat(RegistrationNames);

    public int SourceRegistrationIndex
    {
        get => Settings.ActiveRegistrationIndex;
        set
        {
            if (value < 0 || value >= Settings.Registrations.Count || value == Settings.ActiveRegistrationIndex) return;
            Settings.ActiveRegistrationIndex = value;
            ApplyRegistrationChoice();
        }
    }

    public int TargetRegistrationChoiceIndex
    {
        get => Settings.TargetRegistrationIndex is { } i && i >= 0 && i < Settings.Registrations.Count ? i + 1 : 0;
        set
        {
            if (value < 0 || value > Settings.Registrations.Count) return;
            int? chosen = value == 0 ? null : value - 1;
            if (chosen == Settings.TargetRegistrationIndex) return;
            Settings.TargetRegistrationIndex = chosen;
            ApplyRegistrationChoice();
        }
    }

    // The registration the second (target) sign-in uses — shown on the Target step.
    public string TargetRegistrationName => Settings.TargetRegistration?.DisplayName ?? string.Empty;

    // A different registration means a different sign-in: drop both connections so nothing keeps using
    // a token from the old one, persist the choice and rebuild the routing.
    private void ApplyRegistrationChoice()
    {
        try { Settings.Save(); } catch { /* the choice still applies for this session */ }
        AuthService.Configure(Settings);
        DisconnectSource();
        DisconnectTarget();
        SignedInUser       = string.Empty;
        TargetSignedInUser = string.Empty;
        ConfigureTenancy();
        SpService.Initialize();
        RefreshRegistrationProperties();
    }

    private void RefreshRegistrationProperties()
    {
        OnPropertyChanged(nameof(RegistrationNames));
        OnPropertyChanged(nameof(TargetRegistrationChoices));
        OnPropertyChanged(nameof(SourceRegistrationIndex));
        OnPropertyChanged(nameof(TargetRegistrationChoiceIndex));
        OnPropertyChanged(nameof(TargetRegistrationName));
        OnPropertyChanged(nameof(IsNotConfigured));
    }

    private static bool SameRegistrations(AppSettings a, AppSettings b)
    {
        if (a.ActiveRegistrationIndex != b.ActiveRegistrationIndex) return false;
        if (a.TargetRegistrationIndex != b.TargetRegistrationIndex) return false;
        if (a.Registrations.Count != b.Registrations.Count) return false;
        for (int i = 0; i < a.Registrations.Count; i++)
        {
            var x = a.Registrations[i];
            var y = b.Registrations[i];
            if (x.ClientId != y.ClientId || x.TenantId != y.TenantId || x.Name != y.Name) return false;
        }
        return true;
    }

    // ── Header: who is signed in ──────────────────────────────────────────────
    // Single tenant: just the account. Multi-tenant: labelled per side, so "→ b@y" can't be misread.
    public string HeaderAccountsText
    {
        get
        {
            if (!IsMultiTenant) return SignedInUser;
            var parts = new List<string>();
            if (!string.IsNullOrEmpty(SignedInUser))       parts.Add(Loc.T("Hdr_SourceAccount", SignedInUser));
            if (!string.IsNullOrEmpty(TargetSignedInUser)) parts.Add(Loc.T("Hdr_TargetAccount", TargetSignedInUser));
            return string.Join("   ·   ", parts);
        }
    }

    // ── Access for tools outside the wizard (Compare, History → Verify) ───────
    // In multi-tenant mode each site needs its own side's sign-in. The wizard registers that when it
    // connects, but Compare and History work from their own site URLs (or a saved report from an earlier
    // session), so they ask here first: sign in with the right registration and register the host so
    // requests use that token. A no-op in single-tenant mode.
    public async Task PrepareSiteAccessAsync(string siteUrl, bool isSource, CancellationToken ct = default)
    {
        if (!IsMultiTenant || string.IsNullOrWhiteSpace(siteUrl)) return;
        var auth = isSource ? AuthService : TargetAuthService;

        // The other side already owns this host: routing is by host, so one host can't serve both.
        if (Router.AuthForHostOrNull(siteUrl) is { } owner && !ReferenceEquals(owner, auth))
            throw new InvalidOperationException(Loc.T("VM_SameTenantError"));

        await auth.GetAccessTokenAsync(forceInteractive: !auth.IsAuthenticated, cancellationToken: ct);
        Router.RegisterHost(siteUrl.Trim(), auth);
        if (isSource) SignedInUser = auth.UserName ?? SignedInUser;
        else TargetSignedInUser = auth.UserName ?? TargetSignedInUser;
    }

    // The SharePoint address of the tenant an ALREADY signed-in account belongs to (for example
    // https://contoso.sharepoint.com), read from Microsoft Graph with the existing token. Never opens a sign-in window:
    // returns null when that side is not signed in or the lookup fails. The Flows tool shows this as each side's tenant,
    // instead of guessing from the site URL typed on the main window.
    public async Task<string?> GetConnectedRootAsync(bool isSource, CancellationToken ct = default)
    {
        var auth = isSource ? AuthService : TargetAuthService;
        if (!auth.IsAuthenticated) return null;
        try { return await auth.GetRootSiteUrlAsync(ct); }
        catch (OperationCanceledException) { throw; }
        catch { return null; }
    }
    // The Flows tool's "Change source…": the same for the source side. Always opens the sign-in window, so the source
    // account (and with it the source tenant) can be changed, then reads that tenant's SharePoint root address and
    // registers the host with the source sign-in. Returns the address, e.g. https://contoso.sharepoint.com.
    public async Task<string?> SignInSourceForToolsAsync(CancellationToken ct = default)
    {
        if (!IsMultiTenant) return null;
        await AuthService.GetAccessTokenAsync(forceInteractive: true, cancellationToken: ct);
        SignedInUser = AuthService.UserName ?? SignedInUser;
        var root = await AuthService.GetRootSiteUrlAsync(ct);
        await PrepareSiteAccessAsync(root, isSource: true, ct);
        return root;
    }
    // The Flows tool's "Sign in to target": always opens the sign-in window (so it also works to switch accounts or
    // browsers), then reads the address of the target tenant's SharePoint root site and registers that host with the
    // target sign-in, so every later lookup on it uses the right account. Returns the address, e.g.
    // https://fabrikam.sharepoint.com. Multi-tenant only (single tenant has one sign-in and nothing to infer).
    public async Task<string?> SignInTargetForToolsAsync(CancellationToken ct = default)
    {
        if (!IsMultiTenant) return null;
        await TargetAuthService.GetAccessTokenAsync(forceInteractive: true, cancellationToken: ct);
        TargetSignedInUser = TargetAuthService.UserName ?? TargetSignedInUser;
        var root = await TargetAuthService.GetRootSiteUrlAsync(ct);
        await PrepareSiteAccessAsync(root, isSource: false, ct);
        return root;
    }
    // Same, for a saved run: both sites, plus the drives its roots refer to (so a verify started in a
    // later session — when nothing has been listed yet — can still route every request).
    public async Task PrepareReportAccessAsync(
        string sourceUrl, string targetUrl, IEnumerable<SavedReportRoot> roots, CancellationToken ct = default)
    {
        if (!IsMultiTenant) return;
        await PrepareSiteAccessAsync(sourceUrl, isSource: true, ct);
        await PrepareSiteAccessAsync(targetUrl, isSource: false, ct);
        foreach (var root in roots)
        {
            Router.RegisterDrive(root.SourceDriveId, AuthService);
            Router.RegisterDrive(root.TargetDriveId, TargetAuthService);
        }
    }

    // ── Mapping help: check a target user, and the people a run could not find ──

    // Can the mapping dialog look users up in the target right now? Needs multi-tenant mode and a target URL.
    public bool CanCheckTargetUsers => IsMultiTenant && !string.IsNullOrWhiteSpace(TargetUrl);

    // Is this person (email) present in the TARGET tenant? true / false, or null if it could not be checked
    // (no target URL, sign-in failed, ...). Looking a user up with ensureuser also adds them to the target
    // site's user list — which the copy itself does before it can stamp or grant them.
    public async Task<bool?> CheckTargetUserAsync(string email, CancellationToken ct = default)
    {
        if (!CanCheckTargetUsers || string.IsNullOrWhiteSpace(email)) return null;
        try
        {
            await PrepareSiteAccessAsync(TargetUrl, isSource: false, ct);
            return await SpService.EnsureUserAsync(TargetUrl.Trim(), email.Trim(), ct) != null;
        }
        catch (OperationCanceledException) { throw; }
        catch { return null; }
    }

    // People the last cross-tenant run could not find in the target (most-affecting first).
    public IReadOnlyList<UnresolvedUser> UnresolvedUsers => SpService.Unresolved.Snapshot();

    // The report banner: only after a cross-tenant run that left someone unresolved.
    public bool ShowUnresolvedBanner => IsMultiTenant && IsReadyForReport && SpService.Unresolved.DistinctCount > 0;
    public int UnresolvedUserCount => SpService.Unresolved.DistinctCount;

    // ── Taskbar progress ──────────────────────────────────────────────────────
    // Long runs happen behind other windows: show progress on the taskbar icon, and leave it red after
    // a run that had failures until the user opens the report.
    public TaskbarItemProgressState TaskbarProgressState =>
        IsRunInProgress
            ? (TotalCount > 0 ? TaskbarItemProgressState.Normal : TaskbarItemProgressState.Indeterminate)
            : IsReadyForReport && CurrentStep != 5 && FailedCount > 0
                ? TaskbarItemProgressState.Error
                : TaskbarItemProgressState.None;

    public double TaskbarProgressValue => Math.Clamp(TotalProgress / 100.0, 0, 1);

    private void RefreshTaskbar()
    {
        OnPropertyChanged(nameof(TaskbarProgressState));
        OnPropertyChanged(nameof(TaskbarProgressValue));
        OnPropertyChanged(nameof(IsRunInProgress));
        // Same trigger points (run started / finished / progress tick) also decide the report banner.
        OnPropertyChanged(nameof(ShowUnresolvedBanner));
        OnPropertyChanged(nameof(UnresolvedUserCount));
        OnPropertyChanged(nameof(CanCheckTargetUsers));
    }

    partial void OnCurrentStepChanged(int value) => RefreshTaskbar();
    partial void OnIsCopyCompleteChanged(bool value) => RefreshTaskbar();
}
