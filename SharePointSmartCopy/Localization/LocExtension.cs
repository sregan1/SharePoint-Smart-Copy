using System.Windows.Markup;

namespace SharePointSmartCopy.Localization;

// XAML usage: Text="{loc:Loc Ui_Key}"
[MarkupExtensionReturnType(typeof(string))]
public sealed class LocExtension : MarkupExtension
{
    public string Key { get; set; } = string.Empty;

    public LocExtension() { }
    public LocExtension(string key) => Key = key;

    public override object ProvideValue(IServiceProvider serviceProvider) => Loc.T(Key);
}
