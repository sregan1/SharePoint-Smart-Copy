using SharePointSmartCopy.Localization;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class PermissionCopyService(SharePointService spService)
{
    // name→ID cache pre-loaded from the target site once per copy session
    private Dictionary<string, int> _targetRoleDefs = [];

    // principal resolution caches for the duration of the copy session.
    // Concurrent — CopyObjectPermissionsAsync is called from parallel file-copy tasks.
    private readonly ResolutionCache _userCache  = new();
    private readonly ResolutionCache _groupCache = new();

    // Successful resolutions are kept for the whole session. A null (unresolved) result is kept only
    // briefly: EnsureUserAsync / GetSiteGroupIdAsync return null both for a principal that genuinely
    // isn't on the target AND for any transient failure (throttle, network). A permanent negative entry
    // turned one hiccup into a dropped grant — and a broken-inheritance item — for every later item.
    // The short TTL still spares repeat lookups for a genuinely missing principal.
    private sealed class ResolutionCache
    {
        private static readonly TimeSpan NegativeTtl = TimeSpan.FromSeconds(90);
        private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (int? Id, DateTimeOffset Expires)> _map
            = new(StringComparer.OrdinalIgnoreCase);

        public bool TryGet(string key, out int? id)
        {
            if (_map.TryGetValue(key, out var e) && (e.Id != null || e.Expires > DateTimeOffset.UtcNow))
            {
                id = e.Id;
                return true;
            }
            id = null;
            return false;
        }

        public void Set(string key, int? id)
            => _map[key] = (id, id != null ? DateTimeOffset.MaxValue : DateTimeOffset.UtcNow + NegativeTtl);

        public void Clear() => _map.Clear();
    }

    // Call once before copy loops start. Pre-loads target role definitions in one request.
    public async Task InitializeAsync(string targetSiteUrl, CancellationToken ct = default)
    {
        _targetRoleDefs = await spService.GetAllRoleDefinitionsAsync(targetSiteUrl, ct);
        _userCache.Clear();
        _groupCache.Clear();
    }

    // Copies role assignments from sourceApiPath to targetApiPath.
    // hasUniquePermissions should come from the already-fetched item metadata — zero extra calls.
    // sourceApiPath / targetApiPath: e.g. "web/lists('guid')/items(3)"
    public async Task<PermissionCopyResult> CopyObjectPermissionsAsync(
        string sourceSiteUrl,
        string targetSiteUrl,
        string sourceApiPath,
        string targetApiPath,
        bool? hasUniquePermissions,
        string itemDisplayName,
        CancellationToken ct = default)
    {
        // null = the source could not be asked (see TryGetHasUniqueRoleAssignmentsAsync): say so instead of
        // silently leaving the item with inherited permissions.
        if (hasUniquePermissions is null)
            return new PermissionCopyResult(itemDisplayName, 0, [], Loc.T("Svc_Perm_UniqueUnknown"));
        if (hasUniquePermissions == false)
            return new PermissionCopyResult(itemDisplayName, 0, []);

        // Guard BEFORE breaking inheritance: with an empty role-definition cache (failed/skipped
        // InitializeAsync) every role would fail "no such role on target" AFTER inheritance was
        // already broken — stripping the item's permissions down to the migrating account. Not
        // copying is recoverable; breaking without applying is destructive.
        if (_targetRoleDefs.Count == 0)
            return new PermissionCopyResult(itemDisplayName, 0, [],
                Loc.T("Svc_Perm_NoRoleDefs"));

        var assignments = await spService.GetRoleAssignmentsAsync(sourceSiteUrl, sourceApiPath, ct);
        if (assignments == null)
            return new PermissionCopyResult(itemDisplayName, 0, [],
                Loc.T("Svc_Perm_NoSource"));

        // "Limited Access" bindings are hierarchy plumbing SharePoint maintains itself and rejects
        // when granted directly — copying them only produced failed-role noise, and an object whose
        // unique assignments were ONLY Limited Access broke inheritance and applied nothing.
        static bool IsLimitedAccess(string roleName) =>
            roleName.Equals("Limited Access", StringComparison.OrdinalIgnoreCase) ||
            roleName.Equals("Web-Only Limited Access", StringComparison.OrdinalIgnoreCase);
        assignments = assignments
            .Select(a => a with { RoleNames = a.RoleNames.Where(r => !IsLimitedAccess(r)).ToList() })
            .Where(a => a.RoleNames.Count > 0)
            .ToList();

        if (assignments.Count == 0)
            return new PermissionCopyResult(itemDisplayName, 0, []);

        // The root web is already the top of the permission hierarchy — BreakRoleInheritance is
        // not valid on it and SharePoint returns 500 "Cannot change permissions of root web".
        bool isRootWeb = string.Equals(targetApiPath, "web", StringComparison.OrdinalIgnoreCase);
        if (!isRootWeb)
        {
            try
            {
                await spService.BreakPermissionInheritanceAsync(targetSiteUrl, targetApiPath, ct);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                return new PermissionCopyResult(itemDisplayName, 0, [], Loc.T("Svc_Perm_BreakFailed", ex.Message));
            }
        }

        int applied = 0;
        var skipped = new List<string>();
        var failed  = new List<string>();

        foreach (var assignment in assignments)
        {
            int? principalId = await ResolvePrincipalAsync(targetSiteUrl, assignment, ct);
            if (principalId == null)
            {
                skipped.Add(string.IsNullOrEmpty(assignment.LoginName) ? assignment.Title : assignment.LoginName);
                // Cross-tenant: remember WHO could not be found so the report can offer to map them.
                if (spService.IsCrossTenant)
                    spService.Unresolved.Record(
                        IdentityMapper.DisplayIdentity(assignment.LoginName, assignment.Title),
                        IdentityMapper.DisplayIdentity(spService.Identity.MapPrincipal(assignment.LoginName, assignment.Title), assignment.Title),
                        "permission");
                continue;
            }

            foreach (var roleName in assignment.RoleNames)
            {
                if (!_targetRoleDefs.TryGetValue(roleName, out var roleDefId))
                {
                    // Role definition doesn't exist on target — record so the user sees it.
                    failed.Add(Loc.T("Svc_Perm_NoRoleSuffix", assignment.Title, roleName));
                    continue;
                }

                try
                {
                    await spService.AddRoleAssignmentAsync(targetSiteUrl, targetApiPath, principalId.Value, roleDefId, ct);
                    applied++;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    failed.Add(Loc.T("Svc_Perm_RoleError", assignment.Title, roleName, ex.Message));
                }
            }
        }

        // Inheritance was broken but nothing could be applied — the target object is now
        // accessible only to the migrating account. This includes the all-principals-unresolvable
        // case (cross-tenant users, missing groups), which used to report Success.
        string? error = null;
        if (!isRootWeb && applied == 0 && (failed.Count > 0 || skipped.Count > 0))
            error = failed.Count > 0
                ? Loc.T("Svc_Perm_NoneApplied", string.Join("; ", failed.Take(3)))
                : Loc.T("Svc_Perm_NoPrincipals", string.Join("; ", skipped.Take(3)));

        return new PermissionCopyResult(itemDisplayName, applied, skipped, error, failed);
    }

    private async Task<int?> ResolvePrincipalAsync(
        string targetSiteUrl, RoleAssignmentInfo assignment, CancellationToken ct)
    {
        // PrincipalType 8 = SharePoint group → resolve by title
        if (assignment.PrincipalType == 8)
        {
            // Cross-tenant: a group's target name can differ — an override keyed by its title remaps it.
            var groupTitle = spService.Identity.MapPrincipal(string.Empty, assignment.Title);
            if (_groupCache.TryGet(groupTitle, out var cached))
                return cached;
            var id = await spService.GetSiteGroupIdAsync(targetSiteUrl, groupTitle, ct);
            _groupCache.Set(groupTitle, id);
            return id;
        }

        // PrincipalType 1 (user) or 4 (AAD security group) → resolve via ensureuser.
        // Cross-tenant: map the source login (domain swap / explicit override) to its target identity first.
        var loginKey = spService.Identity.MapPrincipal(assignment.LoginName, assignment.Title);
        if (_userCache.TryGet(loginKey, out var cachedUser))
            return cachedUser;
        var userId = await spService.EnsureUserAsync(targetSiteUrl, loginKey, ct);
        _userCache.Set(loginKey, userId);
        return userId;
    }
}

public record PermissionCopyResult(
    string ItemName,
    int Applied,
    List<string> SkippedPrincipals,
    string? Error = null,
    List<string>? FailedRoles = null)
{
    // True when there is anything worth showing in the copy log.
    public bool HasActivity => Applied > 0 || SkippedPrincipals.Count > 0
                            || Error != null || (FailedRoles?.Count ?? 0) > 0;
}
