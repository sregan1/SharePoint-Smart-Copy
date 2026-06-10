namespace SharePointSmartCopy.Models;

public record RoleAssignmentInfo(
    int PrincipalId,
    int PrincipalType,
    string LoginName,
    string Title,
    List<string> RoleNames);
