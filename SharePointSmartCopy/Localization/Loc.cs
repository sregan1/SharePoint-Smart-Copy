using System.Globalization;
using System.Resources;
using System.Windows;

namespace SharePointSmartCopy.Localization;

// Runtime string lookup backed by Resources/Strings*.resx (English is the neutral resource).
// The language is applied once at startup (before any window is created); changing it in
// Settings takes effect on the next launch.
public static class Loc
{
    public sealed record LanguageInfo(string Code, string NativeName);

    // The same 30 languages as the SharePoint Smart Permissions web part.
    public static readonly IReadOnlyList<LanguageInfo> Languages =
    [
        new("ar-SA", "العربية"),           new("cs-CZ", "Čeština"),
        new("da-DK", "Dansk"),             new("de-DE", "Deutsch"),
        new("el-GR", "Ελληνικά"),          new("en-US", "English"),
        new("es-ES", "Español"),           new("fi-FI", "Suomi"),
        new("fr-FR", "Français"),          new("he-IL", "עברית"),
        new("hi-IN", "हिन्दी"),             new("hu-HU", "Magyar"),
        new("id-ID", "Bahasa Indonesia"),  new("it-IT", "Italiano"),
        new("ja-JP", "日本語"),             new("ko-KR", "한국어"),
        new("nb-NO", "Norsk bokmål"),      new("nl-NL", "Nederlands"),
        new("pl-PL", "Polski"),            new("pt-BR", "Português (Brasil)"),
        new("pt-PT", "Português (Portugal)"), new("ro-RO", "Română"),
        new("ru-RU", "Русский"),           new("sv-SE", "Svenska"),
        new("th-TH", "ไทย"),               new("tr-TR", "Türkçe"),
        new("uk-UA", "Українська"),        new("vi-VN", "Tiếng Việt"),
        new("zh-CN", "简体中文"),           new("zh-TW", "繁體中文"),
    ];

    private static readonly ResourceManager Rm =
        new("SharePointSmartCopy.Resources.Strings", typeof(Loc).Assembly);

    public static CultureInfo Culture { get; private set; } = CultureInfo.CurrentUICulture;

    public static bool IsRtl => Culture.TextInfo.IsRightToLeft;

    public static FlowDirection FlowDirection =>
        IsRtl ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;

    // code: "" / null = follow the Windows display language.
    public static void Apply(string? code)
    {
        CultureInfo culture;
        try
        {
            culture = string.IsNullOrWhiteSpace(code)
                ? ClosestSupported(CultureInfo.InstalledUICulture)
                : CultureInfo.GetCultureInfo(code);
        }
        catch (CultureNotFoundException) { culture = CultureInfo.InvariantCulture; }

        Culture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
        CultureInfo.CurrentUICulture = culture;
    }

    // Maps e.g. de-AT -> de-DE, nn-NO/no -> nb-NO, zh-HK -> zh-TW. Unsupported languages fall back to English.
    private static CultureInfo ClosestSupported(CultureInfo c)
    {
        var exact = Languages.FirstOrDefault(l => l.Code.Equals(c.Name, StringComparison.OrdinalIgnoreCase));
        if (exact != null) return CultureInfo.GetCultureInfo(exact.Code);

        var lang = c.TwoLetterISOLanguageName;
        if (lang is "nn" or "no") lang = "nb";
        if (lang == "zh")
            return CultureInfo.GetCultureInfo(
                c.Name.Contains("Hant", StringComparison.OrdinalIgnoreCase) ||
                c.Name.EndsWith("-TW", StringComparison.OrdinalIgnoreCase) ||
                c.Name.EndsWith("-HK", StringComparison.OrdinalIgnoreCase) ||
                c.Name.EndsWith("-MO", StringComparison.OrdinalIgnoreCase) ? "zh-TW" : "zh-CN");

        var byLang = Languages.FirstOrDefault(l => l.Code.StartsWith(lang + "-", StringComparison.OrdinalIgnoreCase));
        return byLang != null ? CultureInfo.GetCultureInfo(byLang.Code) : CultureInfo.GetCultureInfo("en-US");
    }

    public static string T(string key)
    {
        try { return Rm.GetString(key, Culture) ?? key; }
        catch (MissingManifestResourceException) { return key; }
    }

    public static string T(string key, params object?[] args)
    {
        var fmt = T(key);
        try { return string.Format(Culture, fmt, args); }
        catch (FormatException) { return fmt; }
    }
}
