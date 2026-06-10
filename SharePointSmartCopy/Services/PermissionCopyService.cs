using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

public class PermissionCopyService(SharePointService spService)
{
    // name→ID cache pre-loaded from the target site once per copy session
    private Dictionary<string, int> _targetRoleDefs = [];

    // principal resolution caches for the duration of the copy session
    private readonly Dictionary<string, int?> _userCache  = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, int?> _groupCache = new(StringComparer.OrdinalIgnoreCase);

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

        var assignments = await spService.GetRoleAssignmentsAsync(sourceSiteUrl, sourceApiPath, ct);
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
                    continue; // role definition doesn't exist on target — skip

                try
                {
                    await spService.AddRoleAssignmentAsync(targetSiteUrl, targetApiPath, principalId.Value, roleDefId, ct);
                    applied++;
                }
                catch (OperationCanceledException) { throw; }
                catch { /* non-fatal — continue with remaining assignments */ }
            }
        }

        return new PermissionCopyResult(itemDisplayName, applied, skipped);
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
    string? Error = null);
