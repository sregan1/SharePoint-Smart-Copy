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

public class AppSettings
{
    public List<AzureRegistration> Registrations { get; set; } = [];
    public int ActiveRegistrationIndex { get; set; } = 0;
    public string SourceUrl { get; set; } = string.Empty;
    public string TargetUrl { get; set; } = string.Empty;

    // Legacy single-registration fields — kept only for migration from older settings files
    public string? ClientId { get; set; }
    public string? TenantId { get; set; }

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
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
                var s = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();

                // Migrate from old single-registration format
                if (s.Registrations.Count == 0 && !string.IsNullOrWhiteSpace(s.ClientId))
                {
                    s.Registrations.Add(new AzureRegistration
                    {
                        Name     = "Default",
                        ClientId = s.ClientId,
                        TenantId = string.IsNullOrEmpty(s.TenantId) ? "common" : s.TenantId
                    });
                    s.ClientId = null;
                    s.TenantId = null;
                    s.Save();
                }

                return s;
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
