using System.Windows;
using Microsoft.Win32;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Swaps the color ResourceDictionary (Themes/Colors.Light.xaml or Colors.Dark.xaml)
// in Application.Current.Resources at runtime. All styles reference palette brushes
// via DynamicResource, so the running UI restyles immediately.
public static class ThemeManager
{
    public static void Apply(AppTheme theme)
    {
        bool dark = theme == AppTheme.Dark || (theme == AppTheme.System && IsSystemDark());
        var uri = new Uri($"Themes/Colors.{(dark ? "Dark" : "Light")}.xaml", UriKind.Relative);

        var dicts = Application.Current.Resources.MergedDictionaries;
        var existing = dicts.FirstOrDefault(d =>
            d.Source?.OriginalString.Contains("Themes/Colors.", StringComparison.OrdinalIgnoreCase) == true);

        var replacement = new ResourceDictionary { Source = uri };
        if (existing != null)
            dicts[dicts.IndexOf(existing)] = replacement;
        else
            dicts.Insert(0, replacement);
    }

    private static bool IsSystemDark()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            return key?.GetValue("AppsUseLightTheme") is int v && v == 0;
        }
        catch { return false; }
    }
}
