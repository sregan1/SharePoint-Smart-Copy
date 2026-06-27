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

// What to do when a file already exists at the target:
// Skip — leave it; Overwrite — always replace; IfNewer — replace only when the
// source was modified more recently (incremental re-copy for staged cutovers).
public enum OverwriteMode { Skip, Overwrite, IfNewer }

public enum AppTheme { System, Light, Dark }

public class AppSettings
{
    public List<AzureRegistration> Registrations { get; set; } = [];
    public int ActiveRegistrationIndex { get; set; } = 0;
    public string SourceUrl { get; set; } = string.Empty;
    public string TargetUrl { get; set; } = string.Empty;
    public CopyMode PreferredCopyMode    { get; set; } = CopyMode.MigrationApi;
    // Legacy two-state setting; superseded by OverwriteMode but kept so old
    // settings.json files migrate cleanly (see Load).
    public bool     OverwriteFiles       { get; set; } = false;
    public OverwriteMode? OverwriteMode  { get; set; }
    public bool     CopyVersions         { get; set; } = true;
    public bool     CopyAllVersions      { get; set; } = true;
    public int      MaxVersions          { get; set; } = 10;
    public int      MaxParallelCopies    { get; set; } = 8;
    public CopyScope Scope                { get; set; } = CopyScope.Files;
    public bool      CopyCustomColumns    { get; set; } = false;
    public bool      CopyLibraryContent   { get; set; } = true;
    public bool      RemapPageWebPartUrls { get; set; } = true;
    public bool      PreserveMetadata     { get; set; } = true;
    public bool      CopyNavigation       { get; set; } = true;
    public bool      CopyPermissions      { get; set; } = false;
    public AppTheme  Theme                { get; set; } = AppTheme.System;

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
                var settings = JsonSerializer.Deserialize<AppSettings>(json, _jsonOptions) ?? new AppSettings();
                // Migrate pre-OverwriteMode settings from the legacy bool.
                settings.OverwriteMode ??= settings.OverwriteFiles
                    ? Models.OverwriteMode.Overwrite
                    : Models.OverwriteMode.Skip;
                return settings;
            }
        }
        catch { /* fall through to default */ }
        return new AppSettings();
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
        // Write-then-move so a crash mid-write can't corrupt the settings file
        // (Load silently resets to defaults on corrupt JSON, losing all registrations).
        var tempPath = SettingsPath + ".tmp";
        File.WriteAllText(tempPath, JsonSerializer.Serialize(this, _jsonOptions));
        File.Move(tempPath, SettingsPath, overwrite: true);
    }

    public AzureRegistration? ActiveRegistration =>
        Registrations.Count > 0
        && ActiveRegistrationIndex >= 0
        && ActiveRegistrationIndex < Registrations.Count
            ? Registrations[ActiveRegistrationIndex]
            : null;

    public bool IsConfigured => !string.IsNullOrWhiteSpace(ActiveRegistration?.ClientId);
}
