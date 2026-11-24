using System.Windows.Media;

namespace SharePointSmartCopy.Dialogs;

// One number on the Teams Results page, like the Succeeded / Failed tiles of the Files report.
public sealed record TeamsResultTile(string Value, string Label, Brush Back, Brush Border, Brush Fore);

// One line of the status list: what happened to one part of the copy.
public sealed record TeamsResultRow(string Glyph, Brush GlyphBrush, string Area, string Message);

// One person in the administrator's people search.
public sealed class PersonRow(SharePointSmartCopy.Teams.TeamsApiClient.UserInfo user)
{
    public SharePointSmartCopy.Teams.TeamsApiClient.UserInfo User { get; } = user;
    public string Label => string.IsNullOrWhiteSpace(User.Email) ? User.DisplayName : $"{User.DisplayName}  ({User.Email})";
}

// One team of a batch on the Review step: the source team and the name its copy will get (editable in the grid).
public sealed class TeamCopyPlanRow(SharePointSmartCopy.Teams.TeamsApiClient.TeamInfo team, string newName)
{
    public SharePointSmartCopy.Teams.TeamsApiClient.TeamInfo Team { get; } = team;
    public string SourceName => Team.DisplayName;
    public string NewName { get; set; } = newName;
    public string ChannelsText { get; set; } = string.Empty;   // "5 of 7": the channels that will be copied
}
