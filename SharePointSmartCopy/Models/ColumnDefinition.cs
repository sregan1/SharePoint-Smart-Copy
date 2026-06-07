namespace SharePointSmartCopy.Models;

public enum SupportedFieldType { Text, Note, Number, Boolean, DateTime, Choice, MultiChoice }

public class ColumnDefinition
{
    public string InternalName { get; set; } = string.Empty;
    public string DisplayName  { get; set; } = string.Empty;
    public SupportedFieldType FieldType { get; set; }
    public string[]? Choices  { get; set; }
}
