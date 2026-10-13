using SharePointSmartCopy.Localization;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Text;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SharePointSmartCopy.Models;
using SharePointSmartCopy.Services;
using SpColumnDef = SharePointSmartCopy.Models.ColumnDefinition;

namespace SharePointSmartCopy.ViewModels;

// Partial class MainViewModel: Options. The shared fields, constructor and core setup live in MainViewModel.cs.
public partial class MainViewModel : ObservableObject
{
    // ── Copy Scope ────────────────────────────────────────────────────────────

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsFilesScope))]
    [NotifyPropertyChangedFor(nameof(IsLibraryScope))]
    [NotifyPropertyChangedFor(nameof(IsSiteScope))]
    [NotifyPropertyChangedFor(nameof(IsPagesScope))]
    [NotifyPropertyChangedFor(nameof(IsLibraryOrSiteScope))]
    [NotifyPropertyChangedFor(nameof(IsFilesOrPagesScope))]
    [NotifyPropertyChangedFor(nameof(NeedsTargetFolder))]
    [NotifyPropertyChangedFor(nameof(IsItemSelectionActive))]
    [NotifyPropertyChangedFor(nameof(ShowLibrarySummaryLine))]
    [NotifyPropertyChangedFor(nameof(ShowLibraryFullCopyMapping))]
    [NotifyPropertyChangedFor(nameof(LibrarySummaryCount))]
    [NotifyPropertyChangedFor(nameof(LibraryPreviewItems))]
    [NotifyPropertyChangedFor(nameof(EffectiveCopyCustomColumns))]
    [NotifyPropertyChangedFor(nameof(CustomFieldColumnWidth))]
    [NotifyPropertyChangedFor(nameof(CustomFieldDetailsColumnWidth))]
    [NotifyPropertyChangedFor(nameof(ShowFileCopyOptions))]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    private CopyScope _copyScope = CopyScope.Files;

    public bool IsFilesScope         => CopyScope == CopyScope.Files;
    public bool IsLibraryScope       => CopyScope == CopyScope.Library;
    public bool IsSiteScope          => CopyScope == CopyScope.Site;
    public bool IsPagesScope         => CopyScope == CopyScope.Pages;
    public bool IsLibraryOrSiteScope => CopyScope is CopyScope.Library or CopyScope.Site;
    public bool IsFilesOrPagesScope  => CopyScope is CopyScope.Files or CopyScope.Pages;
    public bool NeedsTargetFolder      => CopyScope is CopyScope.Files or CopyScope.Pages;
    // True when item-level selection is active: either individual items are checked (partial),
    // or a list is in items-only mode (IsChecked == null).
    public bool IsItemSelectionActive  =>
        CopyScope == CopyScope.Library && (
            SourceLibraries.Any(lib => lib.IsCustomList && lib.IsChecked == null) ||
            SourceLibraries.Any(lib => lib.Children.Any(c => c.Type == NodeType.ListItem && c.IsChecked == true)));

    // True when the Site scope summary text should be shown (whole-site copy, not individual items).
    public bool ShowLibrarySummaryLine => IsSiteScope && !IsItemSelectionActive;
    // True when the Library/List full-copy mapping button should be shown.
    public bool ShowLibraryFullCopyMapping => IsLibraryScope && !IsItemSelectionActive;

    // For Library scope: count of checked library nodes. For Site scope: count of all loaded libraries.
    // Count of whole libraries/lists selected (not individual items).
    public int LibrarySummaryCount => IsSiteScope
        ? SourceLibraries.Count
        : SourceLibraries.SelectMany(l => l.GetCheckedNodes()).Count(n => n.Type == NodeType.Library);

    // Count of individually selected list items (partial or items-only selection).
    public int SelectedItemCount => SourceLibraries
        .SelectMany(l => l.Children)
        .Count(c => c.Type == NodeType.ListItem && c.IsChecked == true);

    // Best-effort total size for the Files/Pages-scope preview header. Graph's children listing
    // already returns an aggregate Size for folders (not just files), so this is accurate without
    // a recursive walk — but CopyJob.SourceSize itself stays null until the copy-time enumeration
    // walk (see CopyJob.cs), so this reads the originally-checked SharePointNodes instead. Blank
    // (not a misleading partial sum) when any checked node's size isn't known yet, or when nothing
    // sized is selected — e.g. Library/Site scope, where sizes aren't known until copy time.
    public string CopyPreviewTotalSizeDisplay
    {
        get
        {
            var nodes = SourceLibraries.SelectMany(l => l.GetCheckedNodes())
                .Where(n => n.Type is NodeType.File or NodeType.Folder)
                .ToList();
            if (nodes.Count == 0 || nodes.Any(n => n.Size == null)) return string.Empty;
            return $" ({SizeFormatter.FormatBytes(nodes.Sum(n => n.Size ?? 0))})";
        }
    }

    // Libraries shown in the Options step preview for Library/Site scope.
    // The null (items-only) state only has meaning for custom lists — a stray null on
    // any other node type must not surface a phantom entry here.
    public IEnumerable<SharePointNode> LibraryPreviewItems => IsSiteScope
        ? SourceLibraries
        : SourceLibraries.Where(n => n.IsChecked == true
              || (n.IsCustomList && n.IsChecked == null)
              || (n.IsCustomList && n.Children.Any(c => c.Type == NodeType.ListItem && c.IsChecked == true)));

    // Called from code-behind when any source node check state changes.
    public void NotifySelectionChanged()
    {
        OnPropertyChanged(nameof(IsItemSelectionActive));
        OnPropertyChanged(nameof(ShowLibrarySummaryLine));
        OnPropertyChanged(nameof(ShowLibraryFullCopyMapping));
        OnPropertyChanged(nameof(LibrarySummaryCount));
        OnPropertyChanged(nameof(SelectedItemCount));
        OnPropertyChanged(nameof(CopyPreviewTotalSizeDisplay));
        OnPropertyChanged(nameof(LibraryPreviewItems));
        NextCommand.NotifyCanExecuteChanged();
    }

    // ── Step 3: Options ───────────────────────────────────────────────────────

    [ObservableProperty] private OverwriteMode _overwriteMode = OverwriteMode.Skip;
    [ObservableProperty] private bool _copyVersions = true;
    [ObservableProperty] private bool _copyAllVersions = true;
    [ObservableProperty] private int _maxVersions = 10;
    [ObservableProperty] private int _maxParallelCopies = 4;
    [ObservableProperty] private CopyMode _copyMode = CopyMode.MigrationApi;
    [ObservableProperty] private ObservableCollection<CopyJob> _copyJobs = [];

    // New options
    [ObservableProperty] private bool _preserveMetadata = true;
    // When false, folder date/author repair is limited to folders receiving new files this run (faster
    // repeated incremental runs). Only meaningful while PreserveMetadata is on. See AppSettings.
    [ObservableProperty] private bool _reapplyFolderMetadataEveryRun = true;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(EffectiveCopyCustomColumns))]
    [NotifyPropertyChangedFor(nameof(CanConfigureMappings))]
    [NotifyPropertyChangedFor(nameof(CustomFieldColumnWidth))]
    [NotifyPropertyChangedFor(nameof(CustomFieldDetailsColumnWidth))]
    private bool _copyCustomColumns = false;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowFileCopyOptions))]
    private bool _copyLibraryContent = true;
    [ObservableProperty] private bool _remapPageWebPartUrls = true;
    [ObservableProperty] private bool _copyNavigation = true;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PermColumnWidth))]
    [NotifyPropertyChangedFor(nameof(PermDetailsColumnWidth))]
    private bool _copyPermissions = false;
    public ObservableCollection<ColumnMapping> ColumnMappings { get; } = [];

    // Library, Site, and Pages scopes always copy all custom column values; Files scope uses the checkbox.
    public bool EffectiveCopyCustomColumns => IsLibraryOrSiteScope || IsPagesScope || CopyCustomColumns;

    // File-level options (overwrite, versions, permissions) apply in Files scope, and in
    // Library/Site scope only when library content is being copied.
    public bool ShowFileCopyOptions => IsFilesScope || (IsLibraryOrSiteScope && CopyLibraryContent);

    // Mapping configuration requires columns loaded and (in Files scope) the opt-in checkbox.
    public bool CanConfigureMappings => IsColumnsReady && CopyCustomColumns;
    public IReadOnlyList<SpColumnDef> SourceColumns => _sourceColumns;
    public IReadOnlyList<SpColumnDef> TargetColumns => _targetColumns;

    public string MappingButtonLabel
    {
        get
        {
            var total = _sourceColumns.Count;
            if (ColumnMappings.Count == 0)
                return total > 0 ? Loc.T("VM_ConfigureMappingsCount", total) : Loc.T("VM_ConfigureMappings");

            var unmatched = ColumnMappings.Count(m => m.Status == MappingStatus.Unmatched);
            var mapped    = ColumnMappings.Count - unmatched;
            return unmatched > 0
                ? Loc.T("VM_ConfigureMappingsUnmatched", unmatched, total)
                : Loc.T("VM_ConfigureMappingsMapped", mapped, total);
        }
    }

    public void BuildCopyJobs()
    {
        CopyJobs.Clear();
        if (SelectedTargetFolder == null) return;

        // Find the library ancestor to get its server-relative URL for Migration API mode
        var libraryNode = SelectedTargetFolder;
        while (libraryNode.Parent != null)
            libraryNode = libraryNode.Parent;

        // Compute the subfolder path relative to the library root by walking the parent chain.
        // ServerRelativePath is only populated on library root nodes, so we use node names instead.
        var subFolderRelPath = BuildRelativePath(SelectedTargetFolder, libraryNode);

        var sourceSiteUrl = SourceUrl.TrimEnd('/');
        var targetSiteUrl = SelectedTargetFolder.SiteUrl.TrimEnd('/');

        foreach (var lib in SourceLibraries)
        {
            foreach (var node in lib.GetCheckedNodes())
            {
                var isLibrary = node.Type == NodeType.Library;

                // When the source is a whole library, copy its contents directly into the
                // target folder — no wrapper folder with the library's name.
                var targetDisplayPath = isLibrary
                    ? (string.IsNullOrEmpty(subFolderRelPath)
                        ? $"{targetSiteUrl}/{libraryNode.Name}"
                        : $"{targetSiteUrl}/{libraryNode.Name}/{subFolderRelPath}")
                    : (string.IsNullOrEmpty(subFolderRelPath)
                        ? $"{targetSiteUrl}/{libraryNode.Name}/{node.Name}"
                        : $"{targetSiteUrl}/{libraryNode.Name}/{subFolderRelPath}/{node.Name}");

                var job = new CopyJob
                {
                    SourceDriveId                  = node.DriveId,
                    SourceItemId                   = node.Id,
                    SourceName                     = node.Name,
                    SourceSiteUrl                  = sourceSiteUrl,
                    SourceDisplayPath              = $"{sourceSiteUrl}/{BuildPath(node)}",
                    TargetDriveId                  = SelectedTargetFolder.DriveId,
                    TargetParentItemId             = libraryNode.Id,
                    TargetSubFolderPath            = subFolderRelPath,
                    TargetSiteId                   = SelectedTargetFolder.SiteId,
                    TargetSiteUrl                  = SelectedTargetFolder.SiteUrl,
                    TargetDisplayPath              = targetDisplayPath,
                    TargetLibraryServerRelativeUrl = libraryNode.ServerRelativePath ?? string.Empty,
                    IsFolder                       = node.Type != NodeType.File,
                    IsLibrary                      = isLibrary,
                    IsPage                         = node.IsPage,
                };
                CopyJobs.Add(job);
            }
        }
    }

    // Resumes a file-selection run from History (typically one that was interrupted): reconnects both
    // sites, rebuilds the copy jobs from the run's saved roots, switches to "Copy if newer", and lands
    // on the Options step so the user can review and press Next. Files the earlier run already copied
    // are skipped by the Copy-if-newer check, so only what is missing or changed is transferred.
    // Returns false (after setting a status message) if a site could not be connected.
    public async Task<bool> ResumeFromReportAsync(SavedReportSummary report)
    {
        if (IsCopying || report.Roots.Count == 0) return false;

        // Start from a clean wizard so nothing from an earlier run in this window leaks in.
        IsCopyComplete = false;
        ColumnMappings.Clear();
        CopyScope = CopyScope.Files;
        SourceUrl = report.SourceUrl;
        TargetUrl = report.TargetUrl;

        await ConnectSourceCommand.ExecuteAsync(null);
        if (!SourceConnected) { CurrentStep = 0; return false; }
        await ConnectTargetCommand.ExecuteAsync(null);
        if (!TargetConnected) { CurrentStep = 2; return false; }

        var sourceSite = SourceUrl.TrimEnd('/');
        var targetSite = TargetUrl.TrimEnd('/');
        var jobs = new List<CopyJob>();
        foreach (var root in report.Roots)
        {
            var libraryRelUrl = await SpService.GetLibraryServerRelativeUrlAsync(root.TargetDriveId) ?? string.Empty;
            var libraryName   = libraryRelUrl.TrimEnd('/').Split('/').LastOrDefault() ?? string.Empty;
            var sub           = root.TargetSubFolderPath;
            // Same display-path shapes BuildCopyJobs produces.
            var targetDisplay = root.IsLibrary
                ? (string.IsNullOrEmpty(sub) ? $"{targetSite}/{libraryName}" : $"{targetSite}/{libraryName}/{sub}")
                : (string.IsNullOrEmpty(sub) ? $"{targetSite}/{libraryName}/{root.SourceName}" : $"{targetSite}/{libraryName}/{sub}/{root.SourceName}");
            jobs.Add(new CopyJob
            {
                SourceDriveId                  = root.SourceDriveId,
                SourceItemId                   = root.SourceItemId,
                SourceName                     = root.SourceName,
                SourceSiteUrl                  = sourceSite,
                SourceDisplayPath              = $"{sourceSite}/{root.SourceName}",
                TargetDriveId                  = root.TargetDriveId,
                TargetParentItemId             = root.TargetParentItemId,
                TargetSubFolderPath            = sub,
                TargetSiteId                   = TargetSiteId,
                TargetSiteUrl                  = targetSite,
                TargetDisplayPath              = targetDisplay,
                TargetLibraryServerRelativeUrl = libraryRelUrl,
                IsFolder                       = root.IsFolder,
                IsLibrary                      = root.IsLibrary,
            });
        }

        CopyJobs.Clear();
        foreach (var job in jobs) CopyJobs.Add(job);
        OverwriteMode = Models.OverwriteMode.IfNewer;
        StatusMessage = Loc.T("VM_ResumeReady", jobs.Count);
        CurrentStep   = 3;
        NextCommand.NotifyCanExecuteChanged();
        return true;
    }

}
