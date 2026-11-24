// Partial class MainViewModel: which "Other tools" can work with the current app registration. A tool whose API the
// registration does not list at all is grayed out with a tooltip naming what to add. This never shows UI. For Forms
// and Flows a permission that is on the registration but not yet consented to still works on first click (the tool's own
// sign-in shows the consent screen). Teams: Entra gives no "not on the registration" answer for a Graph permission, and
// the token's scp claim cannot tell "not on the registration" from "not consented yet", so the tile is never disabled;
// it only shows a warning tooltip (copyable for an admin) when a required Graph permission is not granted yet.
using CommunityToolkit.Mvvm.ComponentModel;
using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Services;
using SharePointSmartCopy.Teams;

namespace SharePointSmartCopy.ViewModels;

public partial class MainViewModel
{
    // Enabled until a check proves otherwise, so a failed or skipped check can never lock a tool.
    [ObservableProperty] private bool _formsAvailable = true;
    [ObservableProperty] private string _formsBlockedTip = string.Empty;
    [ObservableProperty] private bool _teamsPermissionWarning;
    [ObservableProperty] private string _teamsBlockedTip = string.Empty;

    private int _toolAccessGeneration;   // a check that finishes after a reset or a newer check must not overwrite it

    private void ResetToolAccess()
    {
        _toolAccessGeneration++;
        FormsAvailable = true;
        FormsBlockedTip = string.Empty;
        TeamsPermissionWarning = false;
        TeamsBlockedTip = string.Empty;
    }

    // True when any sign-in the tool would use is told by Entra that the registration lacks the API.
    private async Task<bool> AnySideMissingAsync(Func<AuthService, Task<AuthService.ResourceAccess>> probe)
    {
        var auths = IsMultiTenant ? new[] { AuthService, TargetAuthService } : new[] { AuthService };
        var missing = false;
        foreach (var auth in auths)
            if (await probe(auth) == AuthService.ResourceAccess.NotInRegistration)
                missing = true;
        return missing;
    }

    // The listed Graph permissions that no sign-in has been granted (empty when all are granted or it cannot be told).
    private async Task<List<string>> MissingGraphScopesAsync(string[] required)
    {
        var auths = IsMultiTenant ? new[] { AuthService, TargetAuthService } : new[] { AuthService };
        var missing = new List<string>();
        foreach (var auth in auths)
            if (await auth.GetGrantedGraphScopesAsync(forceRefresh: true) is { } granted)
                foreach (var s in required.Where(s => !granted.Contains(s) && !missing.Contains(s)))
                    missing.Add(s);
        return missing;
    }

    // Safe to call repeatedly (after each connect).
    private async Task RefreshToolAccessAsync()
    {
        var generation = ++_toolAccessGeneration;
        try
        {
            var forms = await AnySideMissingAsync(a => a.ProbeFormsAccessAsync());
            var missing = await MissingGraphScopesAsync(TeamsApiClient.ReadScopes);
            if (generation != _toolAccessGeneration) return;
            FormsAvailable  = !forms;
            FormsBlockedTip = forms ? Loc.T("Forms_PermissionMissingTip") : string.Empty;
            TeamsPermissionWarning = missing.Count > 0;
            TeamsBlockedTip = missing.Count > 0 ? Loc.T("Teams_PermissionMissingTip") + " " + string.Join(", ", missing) : string.Empty;
        }
        catch { /* leave the tools enabled: a failed check says nothing about the permission */ }
    }

    // The Flows tool stays open (packages need nothing extra); only its live copy needs the Power Automate permissions.
    // Returns the tooltip to show on the disabled live option, or null when live copy may be tried.
    public async Task<string?> GetFlowsLiveBlockedTipAsync()
    {
        try
        {
            return await AnySideMissingAsync(a => a.ProbeFlowAccessAsync()) ? Loc.T("Flows_Live_PermissionMissingTip") : null;
        }
        catch { return null; }
    }
}
