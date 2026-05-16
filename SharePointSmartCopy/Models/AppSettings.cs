using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointSmartCopy.Models;

public partial class AzureRegistration : ObservableObject
{
    [ObservableProperty] private string _name = string.Empty;
    [ObservableProperty] private string _clientId = string.Empty;
    [ObservableProperty] private string _tenantId = string.Empty;

    [JsonIgnore]
    public string DisplayName => string.IsNullOrWhiteSpace(Name) ? "(Unnamed)" : Name;
}

public enum CopyMode { MigrationApi, EnhancedRest }

public class AppSettings
{
    public List<AzureRegistration> Registrations { get; set; } = [];
    public int ActiveRegistrationIndex { get; set; } = 0;
    public string SourceUrl { get; set; } = string.Empty;
    public string TargetUrl { get; set; } = string.Empty;
    public CopyMode PreferredCopyMode { get; set; } = CopyMode.MigrationApi;

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() }
    };

    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SharePointSmartCopy",
        "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<AppSettings>(json, _jsonOptions) ?? new AppSettings();
            }
        }
        catch { /* fall through to default */ }
        return new AppSettings();
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
        File.WriteAllText(SettingsPath, JsonSerializer.Serialize(this, _jsonOptions));
    }

    public AzureRegistration? ActiveRegistration =>
        Registrations.Count > 0
        && ActiveRegistrationIndex >= 0
        && ActiveRegistrationIndex < Registrations.Count
            ? Registrations[ActiveRegistrationIndex]
            : null;

    public bool IsConfigured => !string.IsNullOrWhiteSpace(ActiveRegistration?.ClientId);
}
