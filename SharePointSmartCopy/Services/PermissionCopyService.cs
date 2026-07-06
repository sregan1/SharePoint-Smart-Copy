using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class PermissionCopyService(SharePointService spService)
{
    // name→ID cache pre-loaded from the target site once per copy session
    private Dictionary<string, int> _targetRoleDefs = [];

    // principal resolution caches for the duration of the copy session.
    // Concurrent — CopyObjectPermissionsAsync is called from parallel file-copy tasks.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int?> _userCache  = new(StringComparer.OrdinalIgnoreCase);
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int?> _groupCache = new(StringComparer.OrdinalIgnoreCase);

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
        bool hasUniquePermissions,
        string itemDisplayName,
        CancellationToken ct = default)
    {
        if (!hasUniquePermissions)
            return new PermissionCopyResult(itemDisplayName, 0, []);

        // Guard BEFORE breaking inheritance: with an empty role-definition cache (failed/skipped
        // InitializeAsync) every role would fail "no such role on target" AFTER inheritance was
        // already broken — stripping the item's permissions down to the migrating account. Not
        // copying is recoverable; breaking without applying is destructive.
        if (_targetRoleDefs.Count == 0)
            return new PermissionCopyResult(itemDisplayName, 0, [],
                "Target role definitions unavailable — permissions not copied (inheritance left unchanged); re-run to retry");

        var assignments = await spService.GetRoleAssignmentsAsync(sourceSiteUrl, sourceApiPath, ct);

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
                return new PermissionCopyResult(itemDisplayName, 0, [], $"Break inheritance failed: {ex.Message}");
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
                continue;
            }

            foreach (var roleName in assignment.RoleNames)
            {
                if (!_targetRoleDefs.TryGetValue(roleName, out var roleDefId))
                {
                    // Role definition doesn't exist on target — record so the user sees it.
                    failed.Add($"{assignment.Title} ({roleName}: no such role on target)");
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
                    failed.Add($"{assignment.Title} ({roleName}: {ex.Message})");
                }
            }
        }

        // Inheritance was broken but nothing could be applied — the target object is now
        // accessible only to the migrating account. This includes the all-principals-unresolvable
        // case (cross-tenant users, missing groups), which used to report Success.
        string? error = null;
        if (!isRootWeb && applied == 0 && (failed.Count > 0 || skipped.Count > 0))
            error = failed.Count > 0
                ? $"Inheritance broken but no role assignments applied: {string.Join("; ", failed.Take(3))}"
                : $"Inheritance broken but no principals could be resolved on target: {string.Join("; ", skipped.Take(3))}";

        return new PermissionCopyResult(itemDisplayName, applied, skipped, error, failed);
    }

    private async Task<int?> ResolvePrincipalAsync(
        string targetSiteUrl, RoleAssignmentInfo assignment, CancellationToken ct)
    {
        // PrincipalType 8 = SharePoint group → resolve by title
        if (assignment.PrincipalType == 8)
        {
            if (_groupCache.TryGetValue(assignment.Title, out var cached))
                return cached;
            var id = await spService.GetSiteGroupIdAsync(targetSiteUrl, assignment.Title, ct);
            _groupCache[assignment.Title] = id;
            return id;
        }

        // PrincipalType 1 (user) or 4 (AAD security group) → resolve via ensureuser
        var loginKey = string.IsNullOrEmpty(assignment.LoginName) ? assignment.Title : assignment.LoginName;
        if (_userCache.TryGetValue(loginKey, out var cachedUser))
            return cachedUser;
        var userId = await spService.EnsureUserAsync(targetSiteUrl, loginKey, ct);
        _userCache[loginKey] = userId;
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
