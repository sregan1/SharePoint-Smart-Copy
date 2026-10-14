using SharePointSmartCopy.Localization;
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
    public string DisplayName => string.IsNullOrWhiteSpace(Name) ? Loc.T("VM_Unnamed") : Name;
}

public enum CopyMode { MigrationApi, EnhancedRest }

// What to do when a file already exists at the target:
// Skip — leave it; Overwrite — always replace; IfNewer — replace only when the
// source was modified more recently (incremental re-copy for staged cutovers).
public enum OverwriteMode { Skip, Overwrite, IfNewer }

public enum AppTheme { System, Light, Dark }

// SingleTenant: source and target share one sign-in (the original behavior).
// MultiTenant: source and target each sign in with their own app registration and account.
public enum ConnectionMode { SingleTenant, MultiTenant }

// One "this source identity is that target identity" mapping. Source is a user's email (or a group's
// name); Target is the matching email or group name in the other tenant.
public class IdentityPair
{
    public string Source { get; set; } = string.Empty;
    public string Target { get; set; } = string.Empty;
}

// Cross-tenant identity mapping (multi-tenant mode only). See IdentityMapper.
public class IdentityMappingSettings
{
    // user@SourceDomain is written to the target as user@TargetDomain unless a pair says otherwise.
    public string SourceDomain { get; set; } = string.Empty;
    public string TargetDomain { get; set; } = string.Empty;

    // Explicit pairs. They win over the domain swap and are the only way to remap groups (e.g. an Entra
    // group's name) and users whose names differ between tenants.
    public List<IdentityPair> Pairs { get; set; } = [];

    // LEGACY: an earlier build kept the pairs as "source,target" text lines. It is only ever READ — once,
    // by MigrateLegacyOverrides — and then cleared, so it is no longer written back (null is omitted).
    public string? Overrides { get; set; }

    // Converts the legacy text into Pairs (if there are no Pairs yet) and clears it. Idempotent.
    public void MigrateLegacyOverrides()
    {
        if (string.IsNullOrWhiteSpace(Overrides)) { Overrides = null; return; }
        if (Pairs.Count == 0)
        {
            foreach (var raw in Overrides.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith('#')) continue;
                var parts = line.Split([',', ';', '\t'], 2, StringSplitOptions.TrimEntries);
                if (parts.Length == 2 && parts[0].Length > 0 && parts[1].Length > 0)
                    Pairs.Add(new IdentityPair { Source = parts[0], Target = parts[1] });
            }
        }
        Overrides = null;
    }
}

public class AppSettings
{
    public List<AzureRegistration> Registrations { get; set; } = [];
    public int ActiveRegistrationIndex { get; set; } = 0;
    public ConnectionMode ConnectionMode { get; set; } = ConnectionMode.SingleTenant;
    // Multi-tenant only. The source side uses ActiveRegistrationIndex; the target side uses this.
    // Null (or out of range) means the target falls back to the active registration.
    public int? TargetRegistrationIndex { get; set; }
    public IdentityMappingSettings IdentityMapping { get; set; } = new();
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
    // When true (default), folder Created/Modified dates and Author/Editor are re-applied for EVERY
    // selected folder on every run — needed to repair folder metadata on an already-copied target
    // without re-transferring files, but it costs a per-folder Graph fetch that does not shrink with
    // the skip rate. When false, only folders that actually receive a copied file this run are dated,
    // so repeated "mostly skipped" incremental runs skip that whole phase. Only has any effect while
    // PreserveMetadata is on. See MigrationJobService's folder-metadata fetch.
    public bool      ReapplyFolderMetadataEveryRun { get; set; } = true;
    public bool      CopyNavigation       { get; set; } = true;
    public bool      CopyPermissions      { get; set; } = false;
    public AppTheme  Theme                { get; set; } = AppTheme.System;
    // UI language code (e.g. "de-DE"); empty = follow the Windows display language. Applied at next launch.
    public string    Language             { get; set; } = string.Empty;

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
                // Older builds stored identity pairs as text — convert them to the structured form.
                settings.IdentityMapping ??= new IdentityMappingSettings();
                settings.IdentityMapping.MigrateLegacyOverrides();
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

    public AzureRegistration? TargetRegistration =>
        TargetRegistrationIndex is { } i && i >= 0 && i < Registrations.Count
            ? Registrations[i]
            : ActiveRegistration;

    public bool IsConfigured => !string.IsNullOrWhiteSpace(ActiveRegistration?.ClientId);
}
