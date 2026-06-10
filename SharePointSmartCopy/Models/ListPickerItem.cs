namespace SharePointSmartCopy.Models;

public record ListPickerItem(string Id, string Title)
{
    public override string ToString() => Title;
}
