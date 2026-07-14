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

    // Deep verify: opt-in, off by default — see SharePointSmartCopy/Docs/DEEP-VERIFY-PLAN.md.
    // The checkbox in each Verify UI is seeded from this value and persists the user's last choice,
    // but the LIVE checkbox state (not a re-read of this setting) is what gates any given run.
    // No candidate-count cap or time budget — a run always deep-verifies every candidate (removed
    // 2026-07-10 at the user's request; a per-file size cap still exists in
    // VerificationReportService to avoid downloading pathologically large single files).
    public bool DeepVerifyOfficeFiles    { get; set; } = false;

    // Prefixes default report filenames with "{SourceSite}-{TargetSite}-", e.g.
    // "Marketing-Archive-CopyReport_Files_....csv" — see SiteUrlHelper.
    public bool PrefixReportFilenamesWithSiteNames { get; set; } = true;

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
        catch
        {
            // Preserve the corrupt file before falling back to defaults: the next Save() would
            // otherwise overwrite the evidence and the user's registrations would be lost with
            // no way to recover them by hand.
            try { File.Copy(SettingsPath, SettingsPath + ".corrupt", overwrite: true); } catch { }
        }
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
