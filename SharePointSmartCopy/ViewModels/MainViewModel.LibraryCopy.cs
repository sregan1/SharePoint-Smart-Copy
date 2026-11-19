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

// Partial class MainViewModel: Copy. The shared fields, constructor and core setup live in MainViewModel.cs.
// Partial class MainViewModel: LibraryCopy.
public partial class MainViewModel : ObservableObject
{
    // A custom list read from the source site; Def is null (and ReadError set) when reading it failed.
    private async Task ApplyDeferredStructureSettingsAsync()
    {
        try
        {
            foreach (var note in await _libraryCopyService.ApplyPendingSettingsAsync())
                PushActivity(note);
        }
        catch (Exception ex) when (ex is not OperationCanceledException) { PushActivity(ex.Message); }

        // Managed metadata: term sets copied to the target, and how many values were remapped or left out.
        var (taxRemapped, taxSkipped) = SpService.TaxonomyStats;
        if (_libraryCopyService.TermStoreNotes.Count > 0 || taxRemapped + taxSkipped > 0)
        {
            var parts = new List<string>(_libraryCopyService.TermStoreNotes.Take(6));
            if (taxRemapped + taxSkipped > 0) parts.Add(Loc.T("VM_TaxonomySummary", taxRemapped, taxSkipped));
            System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(new CopyResult
            {
                FileName          = Loc.T("VM_TaxonomySummaryTitle"),
                SourcePath        = string.Empty,
                TargetPath        = TargetUrl.TrimEnd('/'),
                Status            = taxSkipped > 0 ? CopyStatus.Failed : CopyStatus.Success,
                ErrorMessage      = string.Join(" ", parts),
                IsLibraryCreation = true,
            }));
        }

        // One summary line for lookup values (matched by item ID where the item was copied in this run).
        var (byId, byText, unresolved) = SpService.LookupStats;
        if (byId + byText + unresolved > 0)
        {
            var text = Loc.T("VM_LookupSummary", byId, byText, unresolved);
            System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(new CopyResult
            {
                FileName          = Loc.T("VM_LookupSummaryTitle"),
                SourcePath        = string.Empty,
                TargetPath        = TargetUrl.TrimEnd('/'),
                Status            = unresolved > 0 ? CopyStatus.Failed : CopyStatus.Success,
                ErrorMessage      = text,
                IsLibraryCreation = true,
            }));
        }
    }

    // Adds what could not be copied of the list's/library's views (filled in when the list was created) to its result line.
    // A list or library that already existed is normally reported as skipped (only what it lacked was added). With "Update existing
    // lists" on, its structure was brought in line with the source, so it is reported as a success that says so.
    private void MarkExistingUpdated(CopyResult result, string existsMessage, bool structureApplied)
    {
        if (!UpdateExistingStructure || !structureApplied || result.Status == CopyStatus.Failed) return;
        result.Status = CopyStatus.Success;
        var rest = result.ErrorMessage ?? string.Empty;
        if (rest.StartsWith(existsMessage, StringComparison.Ordinal)) rest = rest[existsMessage.Length..].TrimStart();
        result.ErrorMessage = string.IsNullOrEmpty(rest) ? Loc.T("VM_ExistingUpdated") : Loc.T("VM_ExistingUpdated") + " " + rest;
    }

    private void AppendViewNotes(CopyResult result, string title)
    {
        var notes = _libraryCopyService.ViewFailuresFor(title);
        if (notes.Count == 0) return;
        var text = Loc.T("VM_ViewsNotes", notes.Count, string.Join("; ", notes.Take(3)));
        result.ErrorMessage = string.IsNullOrEmpty(result.ErrorMessage) ? text : result.ErrorMessage + " " + text;
    }

    private sealed record SiteListInfo(string Id, string Title, int BaseTemplate, LibraryDefinition? Def, string? ReadError);

    // Orders lists so each one comes after every list its Lookup columns point at (depth-first
    // over source list GUIDs, taken from each definition's SourceListId). A cycle just stops
    // recursing — those lists keep their given order. Entries with no definition go last.
    private static List<T> OrderByLookupDependency<T>(List<T> items, Func<T, LibraryDefinition?> defOf)
    {
        static string Norm(string? id) => (id ?? string.Empty).Trim('{', '}').ToLowerInvariant();

        var byId = new Dictionary<string, T>();
        foreach (var l in items.Where(l => defOf(l) != null))
            byId.TryAdd(Norm(defOf(l)!.SourceListId), l);

        var ordered = new List<T>();
        var visited = new HashSet<string>();
        void Visit(T l)
        {
            if (!visited.Add(Norm(defOf(l)!.SourceListId))) return;
            foreach (var col in defOf(l)!.Columns.Where(c => SpColumnDef.IsLookupType(c.FieldType) && !string.IsNullOrEmpty(c.LookupListId)))
                if (byId.TryGetValue(Norm(col.LookupListId), out var dep))
                    Visit(dep);
            ordered.Add(l);
        }

        foreach (var l in items.Where(l => defOf(l) != null)) Visit(l);
        ordered.AddRange(items.Where(l => defOf(l) == null));
        return ordered;
    }

    // Site scope: copy custom lists (Tasks, Calendars, Announcements, etc.). Shells were already
    // created by the structure pass (preLists); a list missing from it is created here, which also
    // surfaces the real error / already-exists case for it.
    private async Task CopySiteCustomListsAsync(
        List<SiteListInfo> siteLists,
        Dictionary<string, string> preLists,
        Dictionary<string, List<string>> preColumnFailures)
    {
        _copyCts!.Token.ThrowIfCancellationRequested();
        try
        {
            foreach (var info in siteLists)
            {
                _copyCts.Token.ThrowIfCancellationRequested();
                var (srcListId, listTitle, baseTemplate, definition, readError) = info;

                var listResult = new CopyResult
                {
                    FileName          = listTitle,
                    SourcePath        = $"{SourceUrl.TrimEnd('/')}/{listTitle}",
                    TargetPath        = $"{TargetUrl.TrimEnd('/')}/{listTitle}",
                    Status            = CopyStatus.Copying,
                    IsLibraryCreation = true,
                };
                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(listResult));

                if (definition == null)
                {
                    listResult.Status       = CopyStatus.Failed;
                    listResult.ErrorMessage = readError;
                    continue;
                }

                try
                {
                    StatusMessage = Loc.T("VM_CopyingList", listTitle);
                    var hasShell     = preLists.TryGetValue(listTitle, out var shellListId);
                    var targetListId = hasShell
                        ? shellListId!
                        : await _libraryCopyService.CreateCustomListAsync(
                            TargetUrl.TrimEnd('/'), TargetSiteId, definition, baseTemplate);
                    var colFailures = hasShell ? preColumnFailures[listTitle] : _libraryCopyService.LastColumnCreationFailures;
                    if (colFailures.Count > 0)
                        listResult.ErrorMessage =
                            Loc.T("VM_ColumnsCouldNotBeCreated", colFailures.Count, string.Join("; ", colFailures.Take(3)));
                    AppendViewNotes(listResult, listTitle);

                    if (CopyPermissions)
                    {
                        try
                        {
                            var srcHasUnique = await SpService.TryGetHasUniqueRoleAssignmentsAsync(
                                SourceUrl.TrimEnd('/'), $"web/lists('{srcListId}')", _copyCts.Token);
                            var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                                SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                                $"web/lists('{srcListId}')",
                                $"web/lists('{targetListId}')",
                                hasUniquePermissions: srcHasUnique,
                                listTitle, _copyCts.Token);
                            if (perm.HasActivity)
                                AddPermissionResult(perm, $"{TargetUrl.TrimEnd('/')}/{listTitle}");
                        }
                        catch (OperationCanceledException) { throw; }
                        catch { }
                    }

                    if (CopyLibraryContent)
                        await CopyListItemsAsync(targetListId, definition, [], false, listResult);

                    if (listResult.Status != CopyStatus.Failed)
                        listResult.Status = CopyStatus.Success;
                }
                catch (OperationCanceledException) { throw; }
                catch (LibraryAlreadyExistsException alreadyEx)
                {
                    // Stays "Copying" while the existing list is brought up to date; the final status is set below.
                    listResult.ErrorMessage = alreadyEx.Message;

                    if (!string.IsNullOrEmpty(alreadyEx.ListId) && definition.Columns.Count > 0)
                    {
                        try
                        {
                            var createdCols = await _libraryCopyService.AddMissingColumnsAsync(
                                TargetUrl.TrimEnd('/'), alreadyEx.ListId, definition.Columns, definition.SourceSiteUrl);
                            AddColumnCreationResult(listTitle, $"{TargetUrl.TrimEnd('/')}/{listTitle}", createdCols);
                        }
                        catch { }
                    }

                    // Existing list: add the content types and views it lacks (nothing it already has is changed).
                    if (!string.IsNullOrEmpty(alreadyEx.ListId) && EffectiveCopyCustomColumns)
                    {
                        try
                        {
                            await _libraryCopyService.AddMissingStructureAsync(definition, TargetUrl.TrimEnd('/'), alreadyEx.ListId, isLibrary: false);
                            AppendViewNotes(listResult, definition.Title);
                        }
                        catch (Exception ex) when (ex is not OperationCanceledException) { }
                    }

                    if (CopyLibraryContent && !string.IsNullOrEmpty(alreadyEx.ListId))
                        await CopyListItemsAsync(alreadyEx.ListId, definition, [], false, listResult);

                    if (listResult.Status != CopyStatus.Failed)
                        listResult.Status = CopyStatus.Skipped;
                    MarkExistingUpdated(listResult, alreadyEx.Message, !string.IsNullOrEmpty(alreadyEx.ListId) && EffectiveCopyCustomColumns);
                }
                catch (Exception ex)
                {
                    listResult.Status       = CopyStatus.Failed;
                    listResult.ErrorMessage = ex.Message;
                }
            }
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) { StatusMessage = Loc.T("VM_ListCopyWarning", ex.Message); }
        StatusMessage = string.Empty;
    }

    private async Task StartLibraryCopyAsync()
    {
        if (!ConfirmUserMapping()) return;
        IsCopying      = true;
        IsCopyComplete = false;
        CopyDuration   = string.Empty;
        CopyResults.Clear();
        ResultFilter     = ResultFilterKind.All;
        FileResultFilter = ResultFilterKind.All;
        _copyStartTime = DateTimeOffset.Now;
        _copyEndTime   = null;
        // Library / site runs get the same crash journal as a file copy (SaveReport completes it), and the
        // same fresh "people the target can't find" log.
        BeginRunJournal();
        SpService.Unresolved.Clear();
        SpService.ClearLookupCache();   // lookup values resolve against the target as it is NOW
        _libraryCopyService.ResetTermStoreRun();
        _libraryCopyService.CopyTermSets = CopyTermSets;
        _libraryCopyService.Structure = new LibraryCopyService.StructureOptions
        {
            Views = CopyViews, ListSettings = CopyListSettings, ContentTypes = CopyContentTypes, ColumnRules = CopyColumnRules,
            UpdateExisting = UpdateExistingStructure, Templates = CopyTemplates, SiteDefinitions = CopySiteDefinitions, Navigation = CopyNavigation, SiteSettings = CopySiteSettings,
        };
        _libraryCopyService.TermStore = new TermStoreCopyService(AuthService, IsMultiTenant ? TargetAuthService : AuthService);

        _copyCts?.Dispose();
        _copyCts = new CancellationTokenSource();

        var progressTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(400)
        };
        progressTimer.Tick += (_, _) => UpdateProgress();
        progressTimer.Start();

        // Library/Site scope can copy content into MULTIPLE libraries in sequence below, each via its
        // own _copyService.ExecuteAsync call — and each of those launches its own detached (fire-and-
        // forget) folder-metadata pass that outlives the ExecuteAsync call itself. This method used to
        // pass no onMetadataDone at all, so IsUpdatingMetadata never went true here: the run was
        // declared complete (IsReadyForReport/CanGoNext gate open) and the app could be closed while
        // folder dates/authors/permissions were still being written in the background, and a Back→Next
        // re-entry could dispose the still-running pass's CancellationTokenSource out from under it
        // (B6). pendingMetadataPasses starts at 1 as a scheduling sentinel — released in the outer
        // `finally` below — so IsUpdatingMetadata can't flip back to false mid-loop just because one
        // library's pass happened to finish before the next library's copy even starts.
        _metadataStartTime      = DateTimeOffset.Now;
        IsMetadataCancelled     = false;
        IsUpdatingMetadata      = false;
        int  pendingMetadataPasses = 1;
        bool anyMetadataCancelled  = false;
        void OnOneMetadataPassDone(bool completed)
        {
            if (!completed) anyMetadataCancelled = true;
            if (Interlocked.Decrement(ref pendingMetadataPasses) == 0)
            {
                if (anyMetadataCancelled) IsMetadataCancelled = true;
                else MetadataElapsedTime = FormatDuration(DateTimeOffset.Now - _metadataStartTime);
                MetadataEta        = string.Empty;
                IsUpdatingMetadata = false;
                // Content approval and require check-out are switched on only now, after every file and the metadata passes
                // are done: on the empty new list they would have changed how the copy's own writes behave.
                if (!anyMetadataCancelled) _ = ApplyDeferredStructureSettingsAsync();
            }
        }
        IProgress<bool>? MakeMetadataProgress()
        {
            if (!CopyLibraryContent) return null;
            IsUpdatingMetadata = true;
            Interlocked.Increment(ref pendingMetadataPasses);
            return new Progress<bool>(OnOneMetadataPassDone);
        }

        try
        {
            // Pre-warm permission cache once before any copy loops.
            // For Site scope, also copy the web-level permissions.
            if (CopyPermissions)
            {
                try { await _permissionCopyService.InitializeAsync(TargetUrl.TrimEnd('/'), _copyCts.Token); }
                catch { /* non-fatal — proceed without permission copy */ }

                if (IsSiteScope)
                {
                    try
                    {
                        var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                            SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                            "web", "web",
                            hasUniquePermissions: true,
                            "Site", _copyCts.Token);
                        if (perm.HasActivity)
                            AddPermissionResult(perm, TargetUrl.TrimEnd('/'));
                    }
                    catch (OperationCanceledException) { throw; }
                    catch { /* non-fatal */ }
                }
            }

            // Site scope: the site-wide definitions first (custom site columns and content types, navigation, site settings),
            // so the lists created below can use them.
            if (IsSiteScope)
            {
                var siteResult = new CopyResult
                {
                    FileName   = Loc.T("VM_SiteStructure"),
                    SourcePath = SourceUrl.TrimEnd('/'),
                    TargetPath = TargetUrl.TrimEnd('/'),
                    Status     = CopyStatus.Copying,
                };
                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(siteResult));
                try
                {
                    StatusMessage = Loc.T("VM_CopyingSiteStructure");
                    var siteNotes = await _libraryCopyService.CopySiteStructureAsync(
                        SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'), _copyCts.Token);
                    siteResult.Status       = CopyStatus.Success;
                    siteResult.ErrorMessage = siteNotes.Count > 0 ? string.Join("; ", siteNotes.Take(5)) : null;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    siteResult.Status       = CopyStatus.Failed;
                    siteResult.ErrorMessage = ex.Message;
                }
            }

            List<LibraryDefinition> definitions;
            var listDefinitions = new List<(LibraryDefinition Def, int BaseTemplate, SharePointNode? SourceNode)>();
            if (IsSiteScope)
            {
                StatusMessage = Loc.T("VM_ReadingSiteStructure");
                definitions   = await _libraryCopyService.ReadAllLibraryDefinitionsAsync(
                    SourceSiteId, SourceUrl.TrimEnd('/'));
            }
            else
            {
                StatusMessage = Loc.T("VM_ReadingLibraryStructure");
                definitions   = [];
                foreach (var lib in SourceLibraries)
                {
                    foreach (var node in lib.GetCheckedNodes().Where(n => n.Type == NodeType.Library))
                    {
                        if (node.IsCustomList)
                        {
                            var def = await _libraryCopyService.ReadListDefinitionAsync(
                                SourceUrl.TrimEnd('/'), node.Id, node.Name);
                            if (!string.IsNullOrWhiteSpace(node.OverrideName))
                                def.Title = node.OverrideName.Trim();
                            listDefinitions.Add((def, node.ListBaseTemplate, node));
                        }
                        else
                        {
                            var def = await _libraryCopyService.ReadLibraryDefinitionAsync(
                                SourceUrl.TrimEnd('/'), node.DriveId);
                            if (!string.IsNullOrWhiteSpace(node.OverrideName))
                                def.Title = node.OverrideName.Trim();
                            definitions.Add(def);
                        }
                    }

                    // Items-only mode (IsChecked == null) or partial item selection (IsChecked == false with
                    // individual children checked): list node itself is not fully checked.
                    if (lib.Type == NodeType.Library && lib.IsCustomList && lib.IsChecked != true)
                    {
                        var isItemsOnly   = lib.IsChecked == null;
                        var hasSelected   = isItemsOnly || lib.Children.Any(c => c.Type == NodeType.ListItem && c.IsChecked == true);
                        if (hasSelected)
                        {
                            var def = await _libraryCopyService.ReadListDefinitionAsync(
                                SourceUrl.TrimEnd('/'), lib.Id, lib.Name);
                            if (!string.IsNullOrWhiteSpace(lib.OverrideName))
                                def.Title = lib.OverrideName.Trim();
                            listDefinitions.Add((def, lib.ListBaseTemplate, lib));
                        }
                    }
                }
            }

            // Library scope: lists that the selected lists look up are copied too (so lookup columns can
            // bind and values resolve), repeating until no new dependency appears. They get a synthetic
            // checked node so they are copied as whole lists, not as items-only.
            if (!IsSiteScope && CopyLookupDependencies && EffectiveCopyCustomColumns)
            {
                try
                {
                    static string Norm(string? id) => (id ?? string.Empty).Trim('{', '}').ToLowerInvariant();
                    var customLists = (await SpService.GetCustomListsAsync(SourceUrl.TrimEnd('/')))
                        .ToDictionary(l => Norm(l.Item1), l => l, StringComparer.OrdinalIgnoreCase);
                    var have = new HashSet<string>(
                        listDefinitions.Select(l => Norm(l.Def.SourceListId))
                            .Concat(definitions.Select(d => Norm(d.SourceListId))),
                        StringComparer.OrdinalIgnoreCase);
                    for (var pass = 0; pass < 10; pass++)
                    {
                        var wanted = listDefinitions.Select(l => l.Def).Concat(definitions)
                            .SelectMany(d => d.Columns)
                            .Where(c => SpColumnDef.IsLookupType(c.FieldType) && !string.IsNullOrEmpty(c.LookupListId))
                            .Select(c => Norm(c.LookupListId))
                            .Where(id => !have.Contains(id) && customLists.ContainsKey(id))
                            .Distinct().ToList();
                        if (wanted.Count == 0) break;
                        foreach (var id in wanted)
                        {
                            have.Add(id);
                            var (listId, listTitle, listBase) = customLists[id];
                            var dep = await _libraryCopyService.ReadListDefinitionAsync(
                                SourceUrl.TrimEnd('/'), listId, listTitle);
                            var node = new SharePointNode
                            {
                                Name = listTitle, Id = listId, Type = NodeType.Library,
                                IsCustomList = true, ListBaseTemplate = listBase,
                                SiteId = SourceSiteId, SiteUrl = SourceUrl.TrimEnd('/'),
                                IsChecked = true,
                            };
                            listDefinitions.Add((dep, listBase, node));
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[LibraryCopy] lookup dependency expansion failed: {ex.Message}");
                }
            }

            StatusMessage = string.Empty;

            // Site scope, in dependency order so a Lookup column can bind and its values resolve:
            //   1. structure: create EVERY library and custom list (no content yet)
            //   2. fix-up: re-add columns that couldn't be created in step 1 (a Lookup column needs
            //      its target list to exist, and lists are created in source order)
            //   3. custom-list items, ordered so a list is copied after the lists it looks up
            //   4. library files (below) — their lookup values resolve against the copied items
            // Values are matched by display text against the target list's items, so the referenced
            // items must already be there when each value is written.
            var preLibs           = new Dictionary<string, (string newDriveId, string newServerRelativeUrl)>(StringComparer.OrdinalIgnoreCase);
            var preLists          = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var preColumnFailures = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var siteLists         = new List<SiteListInfo>();
            if (IsSiteScope)
            {
                _copyCts.Token.ThrowIfCancellationRequested();
                try
                {
                    StatusMessage = Loc.T("VM_ReadingCustomLists");
                    foreach (var (srcListId, listTitle, baseTemplate) in await SpService.GetCustomListsAsync(SourceUrl.TrimEnd('/')))
                    {
                        _copyCts.Token.ThrowIfCancellationRequested();
                        try
                        {
                            var def = await _libraryCopyService.ReadListDefinitionAsync(
                                SourceUrl.TrimEnd('/'), srcListId, listTitle);
                            siteLists.Add(new SiteListInfo(srcListId, listTitle, baseTemplate, def, null));
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex) { siteLists.Add(new SiteListInfo(srcListId, listTitle, baseTemplate, null, ex.Message)); }
                    }
                    siteLists = OrderByLookupDependency(siteLists, l => l.Def);

                    StatusMessage = Loc.T("VM_CreatingTargetColumns");
                    foreach (var def in definitions)
                    {
                        _copyCts.Token.ThrowIfCancellationRequested();
                        try
                        {
                            preLibs[def.Title] = await _libraryCopyService.CreateLibraryAsync(
                                TargetUrl.TrimEnd('/'), TargetSiteId, def, ColumnMappings);
                            preColumnFailures[def.Title] = [.. _libraryCopyService.LastColumnCreationFailures];
                        }
                        catch (OperationCanceledException) { throw; }
                        catch { /* the library's own pass below retries and reports (incl. already-exists) */ }
                    }
                    foreach (var sl in siteLists.Where(l => l.Def != null))
                    {
                        _copyCts.Token.ThrowIfCancellationRequested();
                        try
                        {
                            preLists[sl.Title] = await _libraryCopyService.CreateCustomListAsync(
                                TargetUrl.TrimEnd('/'), TargetSiteId, sl.Def!, sl.BaseTemplate);
                            preColumnFailures[sl.Title] = [.. _libraryCopyService.LastColumnCreationFailures];
                        }
                        catch (OperationCanceledException) { throw; }
                        catch { /* the list's own pass below retries and reports (incl. already-exists) */ }
                    }

                    // Fix-up: any column still missing (typically a Lookup whose list didn't exist yet).
                    // AddMissingColumnsAsync skips columns that exist, so this is a no-op otherwise.
                    var shellDefs = definitions.Where(d => preLibs.ContainsKey(d.Title))
                        .Select(d => (d.Title, Def: d, ListId: (string?)null, RelUrl: (string?)preLibs[d.Title].newServerRelativeUrl))
                        .Concat(siteLists.Where(l => l.Def != null && preLists.ContainsKey(l.Title))
                            .Select(l => (l.Title, Def: l.Def!, ListId: (string?)preLists[l.Title], RelUrl: (string?)null)));
                    foreach (var (title, d, knownListId, relUrl) in shellDefs.ToList())
                    {
                        if (!preColumnFailures.TryGetValue(title, out var failed) || failed.Count == 0) continue;
                        _copyCts.Token.ThrowIfCancellationRequested();
                        try
                        {
                            var listId = knownListId
                                ?? await SpService.GetListIdByServerRelativeUrlAsync(TargetUrl.TrimEnd('/'), relUrl!);
                            var fix = await _libraryCopyService.AddMissingColumnsAsync(
                                TargetUrl.TrimEnd('/'), listId, d.Columns, d.SourceSiteUrl);
                            preColumnFailures[title] = fix.Failed;
                        }
                        catch (OperationCanceledException) { throw; }
                        catch { /* keep the original failure list */ }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) { StatusMessage = Loc.T("VM_ListCopyWarning", ex.Message); }
                StatusMessage = string.Empty;

                await CopySiteCustomListsAsync(siteLists, preLists, preColumnFailures);
            }

            foreach (var def in definitions)
            {
                _copyCts.Token.ThrowIfCancellationRequested();

                // Emit a "library created" result row
                var libResult = new CopyResult
                {
                    FileName          = def.Title,
                    SourcePath        = $"{SourceUrl.TrimEnd('/')}/{def.Title}",
                    TargetPath        = $"{TargetUrl.TrimEnd('/')}/{def.Title}",
                    Status            = CopyStatus.Copying,
                    IsLibraryCreation = true,
                };
                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(libResult));

                try
                {
                    // Site scope already created every library/list shell up front (see the structure
                    // pass above) so lookup columns could bind to lists that exist.
                    var hasShell = preLibs.TryGetValue(def.Title, out var shell);
                    var (newDriveId, newServerRelUrl) = hasShell
                        ? shell
                        : await _libraryCopyService.CreateLibraryAsync(
                            TargetUrl.TrimEnd('/'), TargetSiteId, def, ColumnMappings);
                    libResult.Status = CopyStatus.Success;
                    var colFailures = hasShell ? preColumnFailures[def.Title] : _libraryCopyService.LastColumnCreationFailures;
                    if (colFailures.Count > 0)
                        libResult.ErrorMessage =
                            Loc.T("VM_ColumnsCouldNotBeCreated", colFailures.Count, string.Join("; ", colFailures.Take(3)));
                    AppendViewNotes(libResult, def.Title);

                    if (CopyPermissions)
                    {
                        try
                        {
                            var newListId = await SpService.GetListIdByServerRelativeUrlAsync(TargetUrl.TrimEnd('/'), newServerRelUrl);
                            var srcHasUnique = await SpService.TryGetHasUniqueRoleAssignmentsAsync(
                                SourceUrl.TrimEnd('/'), $"web/lists('{def.SourceListId}')", _copyCts.Token);
                            var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                                SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                                $"web/lists('{def.SourceListId}')",
                                $"web/lists('{newListId}')",
                                hasUniquePermissions: srcHasUnique,
                                def.Title, _copyCts.Token);
                            if (perm.HasActivity)
                                AddPermissionResult(perm, $"{TargetUrl.TrimEnd('/')}/{def.Title}");
                        }
                        catch (OperationCanceledException) { throw; }
                        catch { /* non-fatal */ }
                    }

                    if (CopyLibraryContent)
                    {
                        StatusMessage = Loc.T("VM_CopyingFilesInto", def.Title);

                        // Build file jobs targeting the new library
                        var newLibRoot = await SpService.GetLibraryRootItemIdAsync(newDriveId);
                        if (newLibRoot == null)
                        {
                            libResult.ErrorMessage = Loc.T("VM_LibraryRootIdMissing");
                        }
                        else
                        {
                            var fileJobs = new List<CopyJob>
                            {
                                new CopyJob
                                {
                                    SourceDriveId                  = def.SourceDriveId,
                                    SourceItemId                   = "root",
                                    SourceName                     = def.Title,
                                    SourceSiteUrl                  = def.SourceSiteUrl,
                                    SourceDisplayPath              = $"{def.SourceSiteUrl}/{def.Title}",
                                    TargetDriveId                  = newDriveId,
                                    TargetParentItemId             = newLibRoot,
                                    TargetSiteId                   = TargetSiteId,
                                    TargetSiteUrl                  = TargetUrl.TrimEnd('/'),
                                    TargetDisplayPath              = $"{TargetUrl.TrimEnd('/')}/{def.Title}",
                                    TargetLibraryServerRelativeUrl = newServerRelUrl,
                                    IsFolder                       = true,
                                    IsLibrary                      = true,
                                    ColumnMappings                 = [.. ColumnMappings],
                                }
                            };

                            // Build bulk field cache for this library; also permission flags if enabled
                            _bulkFieldCache = [];
                            _sourceMetaCache = [];
                            var libPermFlags = new Dictionary<string, bool>();
                            try
                            {
                                if (EffectiveCopyCustomColumns)
                                {
                                    var wantsSystemMeta = CopyMode == CopyMode.MigrationApi && PreserveMetadata;
                                    var systemMeta = wantsSystemMeta ? new Dictionary<string, FileMetadata>() : null;
                                    var raw = await SpService.BulkReadCustomFieldsAsync(
                                        def.SourceSiteUrl, def.SourceListId, def.Columns,
                                        systemMetadataOut: systemMeta, ct: _copyCts.Token);
                                    _bulkFieldCache = raw.ToDictionary(
                                        kvp => $"{def.SourceListId}:{kvp.Key}", kvp => kvp.Value);
                                    if (systemMeta != null)
                                        _sourceMetaCache = systemMeta.ToDictionary(
                                            kvp => $"{def.SourceListId}:{kvp.Key}", kvp => kvp.Value);
                                }
                                if (CopyPermissions)
                                    libPermFlags = await SpService.BulkReadPermissionFlagsAsync(
                                        def.SourceSiteUrl, def.SourceListId, _copyCts.Token);
                            }
                            catch (Exception ex) { ReportBulkReadFailure(def.SourceSiteUrl, ex); }

                            int versionLimit = CopyVersions && !CopyAllVersions ? MaxVersions : 0;
                            await _copyService.ExecuteAsync(
                                fileJobs, CopyResults,
                                new CopyRunOptions
                                {
                                    OverwriteMode         = OverwriteMode,
                                    CopyVersions          = CopyVersions,
                                    MaxParallel           = MaxParallelCopies,
                                    MaxVersions           = versionLimit,
                                    CopyMode              = CopyMode,
                                    OnMetadataDone        = MakeMetadataProgress(),
                                    CopyCustomColumns     = EffectiveCopyCustomColumns,
                                    ColumnMappings        = [.. ColumnMappings],
                                    BulkFieldCache        = _bulkFieldCache,
                                    PreserveMetadata      = PreserveMetadata,
                                    ReapplyFolderMetadata = ReapplyFolderMetadataEveryRun,
                                    CopyPermissions       = CopyPermissions,
                                    PermissionService     = CopyPermissions ? _permissionCopyService : null,
                                    PermissionFlags       = libPermFlags,
                                    SourceMetaCache       = _sourceMetaCache,
                                },
                                _copyCts.Token);
                        }
                        StatusMessage = string.Empty;
                    }
                }
                catch (LibraryAlreadyExistsException alreadyEx)
                {
                    // Stays "Copying" while the existing library is brought up to date; the final status is set below.
                    libResult.ErrorMessage = alreadyEx.Message;

                    // Sync schema: add any source columns missing from the existing target library
                    if (!string.IsNullOrEmpty(alreadyEx.ServerRelativeUrl) && def.Columns.Count > 0)
                    {
                        try
                        {
                            var existingListId = await SpService.GetListIdByServerRelativeUrlAsync(
                                TargetUrl.TrimEnd('/'), alreadyEx.ServerRelativeUrl);
                            var createdCols = await _libraryCopyService.AddMissingColumnsAsync(
                                TargetUrl.TrimEnd('/'), existingListId, def.Columns, def.SourceSiteUrl);
                            AddColumnCreationResult(def.Title, $"{TargetUrl.TrimEnd('/')}/{def.Title}", createdCols);
                        }
                        catch { }
                    }

                    // Existing library: add the content types and views it lacks (nothing it already has is changed).
                    if (!string.IsNullOrEmpty(alreadyEx.ServerRelativeUrl) && EffectiveCopyCustomColumns)
                    {
                        try
                        {
                            var structListId = await SpService.GetListIdByServerRelativeUrlAsync(
                                TargetUrl.TrimEnd('/'), alreadyEx.ServerRelativeUrl);
                            await _libraryCopyService.AddMissingStructureAsync(def, TargetUrl.TrimEnd('/'), structListId, isLibrary: true);
                            AppendViewNotes(libResult, def.Title);
                        }
                        catch (Exception ex) when (ex is not OperationCanceledException) { }
                    }

                    if (CopyLibraryContent && !string.IsNullOrEmpty(alreadyEx.DriveId))
                    {
                        StatusMessage = Loc.T("VM_CopyingFilesIntoExisting", def.Title);

                        var newLibRoot = await SpService.GetLibraryRootItemIdAsync(alreadyEx.DriveId);
                        if (newLibRoot == null)
                        {
                            libResult.ErrorMessage += Loc.T("VM_FileCopySkippedRootId");
                        }
                        else
                        {
                            var serverRelUrl = alreadyEx.ServerRelativeUrl ?? string.Empty;
                            var fileJobs = new List<CopyJob>
                            {
                                new CopyJob
                                {
                                    SourceDriveId                  = def.SourceDriveId,
                                    SourceItemId                   = "root",
                                    SourceName                     = def.Title,
                                    SourceSiteUrl                  = def.SourceSiteUrl,
                                    SourceDisplayPath              = $"{def.SourceSiteUrl}/{def.Title}",
                                    TargetDriveId                  = alreadyEx.DriveId,
                                    TargetParentItemId             = newLibRoot,
                                    TargetSiteId                   = TargetSiteId,
                                    TargetSiteUrl                  = TargetUrl.TrimEnd('/'),
                                    TargetDisplayPath              = $"{TargetUrl.TrimEnd('/')}/{def.Title}",
                                    TargetLibraryServerRelativeUrl = serverRelUrl,
                                    IsFolder                       = true,
                                    IsLibrary                      = true,
                                    ColumnMappings                 = [.. ColumnMappings],
                                }
                            };

                            _bulkFieldCache = [];
                            _sourceMetaCache = [];
                            var existLibPermFlags = new Dictionary<string, bool>();
                            try
                            {
                                if (EffectiveCopyCustomColumns)
                                {
                                    var wantsSystemMeta = CopyMode == CopyMode.MigrationApi && PreserveMetadata;
                                    var systemMeta = wantsSystemMeta ? new Dictionary<string, FileMetadata>() : null;
                                    var raw = await SpService.BulkReadCustomFieldsAsync(
                                        def.SourceSiteUrl, def.SourceListId, def.Columns,
                                        systemMetadataOut: systemMeta, ct: _copyCts.Token);
                                    _bulkFieldCache = raw.ToDictionary(
                                        kvp => $"{def.SourceListId}:{kvp.Key}", kvp => kvp.Value);
                                    if (systemMeta != null)
                                        _sourceMetaCache = systemMeta.ToDictionary(
                                            kvp => $"{def.SourceListId}:{kvp.Key}", kvp => kvp.Value);
                                }
                                if (CopyPermissions)
                                    existLibPermFlags = await SpService.BulkReadPermissionFlagsAsync(
                                        def.SourceSiteUrl, def.SourceListId, _copyCts.Token);
                            }
                            catch (Exception ex) { ReportBulkReadFailure(def.SourceSiteUrl, ex); }

                            int versionLimit = CopyVersions && !CopyAllVersions ? MaxVersions : 0;
                            await _copyService.ExecuteAsync(
                                fileJobs, CopyResults,
                                new CopyRunOptions
                                {
                                    OverwriteMode         = OverwriteMode,
                                    CopyVersions          = CopyVersions,
                                    MaxParallel           = MaxParallelCopies,
                                    MaxVersions           = versionLimit,
                                    CopyMode              = CopyMode,
                                    OnMetadataDone        = MakeMetadataProgress(),
                                    CopyCustomColumns     = EffectiveCopyCustomColumns,
                                    ColumnMappings        = [.. ColumnMappings],
                                    BulkFieldCache        = _bulkFieldCache,
                                    PreserveMetadata      = PreserveMetadata,
                                    ReapplyFolderMetadata = ReapplyFolderMetadataEveryRun,
                                    CopyPermissions       = CopyPermissions,
                                    PermissionService     = CopyPermissions ? _permissionCopyService : null,
                                    PermissionFlags       = existLibPermFlags,
                                    SourceMetaCache       = _sourceMetaCache,
                                },
                                _copyCts.Token);
                        }
                        StatusMessage = string.Empty;
                    }
                    if (libResult.Status != CopyStatus.Failed)
                        libResult.Status = CopyStatus.Skipped;
                    MarkExistingUpdated(libResult, alreadyEx.Message, !string.IsNullOrEmpty(alreadyEx.ServerRelativeUrl) && EffectiveCopyCustomColumns);
                }
                catch (Exception ex)
                {
                    libResult.Status       = CopyStatus.Failed;
                    libResult.ErrorMessage = ex.Message;
                }
            }

            // For Library scope: copy any selected custom lists (Tasks, Calendars, etc.).
            // Lists that others look up are copied first, so the referencing lists' lookup values resolve.
            foreach (var (def, baseTemplate, sourceNode) in OrderByLookupDependency(listDefinitions, x => x.Def))
            {
                _copyCts.Token.ThrowIfCancellationRequested();

                // Determine copy mode for this list node.
                // isItemsOnly: IsChecked == null — copy all items, skip structure creation, use SelectedTargetList.
                // isPartialSelection: IsChecked == false with individual children checked — copy those items only.
                var isItemsOnly       = sourceNode?.IsChecked == null;
                var loadedItems       = sourceNode?.Children.Where(c => c.Type == NodeType.ListItem).ToList() ?? [];
                var selectedItemIds   = loadedItems.Where(c => c.IsChecked == true).Select(c => c.Id).ToHashSet();
                // Items-only ticks every loaded item (the check cascades down), so a user who then
                // UNCHECKS some of them is narrowing the selection — honoring only that subset. Before,
                // items-only ignored the children entirely and copied every item in the list.
                var narrowedItemsOnly  = isItemsOnly && selectedItemIds.Count > 0 && selectedItemIds.Count < loadedItems.Count;
                var isPartialSelection = selectedItemIds.Count > 0 && (!isItemsOnly || narrowedItemsOnly);
                var needsItemCopy      = CopyLibraryContent || isPartialSelection || isItemsOnly;

                var listResult = new CopyResult
                {
                    FileName          = def.Title,
                    SourcePath        = $"{SourceUrl.TrimEnd('/')}/{def.Title}",
                    TargetPath        = $"{TargetUrl.TrimEnd('/')}/{def.Title}",
                    Status            = CopyStatus.Copying,
                    IsLibraryCreation = true,
                };
                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(listResult));

                // This list's own destination on the Target step (null for whole-list copies).
                var itemDest = sourceNode?.DestinationList;
                var itemDestTitle = (isPartialSelection || isItemsOnly) && itemDest != null
                    ? sourceNode!.EffectiveDestinationName : null;

                try
                {
                    StatusMessage = Loc.T("VM_CopyingList", def.Title);
                    string targetListId;
                    if ((isPartialSelection || isItemsOnly) && itemDest != null)
                    {
                        // Items-only or partial selection: create a new list or use this list's chosen destination.
                        if (sourceNode!.IsCreatingNewDestination)
                        {
                            // The Target step is the single authority for each list's destination
                            // name; the preview's Destination Name column mirrors it read-only.
                            def.Title    = sourceNode.EffectiveDestinationName;
                            targetListId = await _libraryCopyService.CreateCustomListAsync(
                                TargetUrl.TrimEnd('/'), TargetSiteId, def, baseTemplate);
                            AppendViewNotes(listResult, def.Title);
                        }
                        else
                        {
                            targetListId = itemDest.Id;
                        }
                        if (CopyCustomColumns && def.Columns.Count > 0)
                        {
                            var colTitle = sourceNode.IsCreatingNewDestination ? def.Title : itemDest.Title;
                            try
                            {
                                var createdCols = await _libraryCopyService.AddMissingColumnsAsync(
                                    TargetUrl.TrimEnd('/'), targetListId, def.Columns, def.SourceSiteUrl);
                                AddColumnCreationResult(colTitle, $"{TargetUrl.TrimEnd('/')}/{colTitle}", createdCols);
                            }
                            catch { /* non-fatal — column creation best-effort */ }
                        }
                    }
                    else
                    {
                        targetListId = await _libraryCopyService.CreateCustomListAsync(
                            TargetUrl.TrimEnd('/'), TargetSiteId, def, baseTemplate);
                        AppendViewNotes(listResult, def.Title);
                    }

                    if (CopyPermissions)
                    {
                        try
                        {
                            var srcHasUnique = await SpService.TryGetHasUniqueRoleAssignmentsAsync(
                                SourceUrl.TrimEnd('/'), $"web/lists('{def.SourceListId}')", _copyCts.Token);
                            var perm = await _permissionCopyService.CopyObjectPermissionsAsync(
                                SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'),
                                $"web/lists('{def.SourceListId}')",
                                $"web/lists('{targetListId}')",
                                hasUniquePermissions: srcHasUnique,
                                def.Title, _copyCts.Token);
                            if (perm.HasActivity)
                                AddPermissionResult(perm, $"{TargetUrl.TrimEnd('/')}/{def.Title}");
                        }
                        catch (OperationCanceledException) { throw; }
                        catch { /* non-fatal */ }
                    }

                    if (needsItemCopy)
                        await CopyListItemsAsync(targetListId, def, selectedItemIds, isPartialSelection, listResult, itemDestTitle);

                    if (listResult.Status != CopyStatus.Failed)
                        listResult.Status = CopyStatus.Success;
                }
                catch (OperationCanceledException) { throw; }
                catch (LibraryAlreadyExistsException alreadyEx)
                {
                    // Stays "Copying" while the existing list is brought up to date; the final status is set below.
                    listResult.ErrorMessage = alreadyEx.Message;

                    if (!string.IsNullOrEmpty(alreadyEx.ListId) && def.Columns.Count > 0)
                    {
                        try
                        {
                            var createdCols = await _libraryCopyService.AddMissingColumnsAsync(
                                TargetUrl.TrimEnd('/'), alreadyEx.ListId, def.Columns, def.SourceSiteUrl);
                            AddColumnCreationResult(def.Title, $"{TargetUrl.TrimEnd('/')}/{def.Title}", createdCols);
                        }
                        catch { }
                    }

                    // Existing list: add the content types and views it lacks (nothing it already has is changed).
                    if (!string.IsNullOrEmpty(alreadyEx.ListId) && EffectiveCopyCustomColumns)
                    {
                        try
                        {
                            await _libraryCopyService.AddMissingStructureAsync(def, TargetUrl.TrimEnd('/'), alreadyEx.ListId, isLibrary: false);
                            AppendViewNotes(listResult, def.Title);
                        }
                        catch (Exception ex) when (ex is not OperationCanceledException) { }
                    }

                    if (needsItemCopy && !string.IsNullOrEmpty(alreadyEx.ListId))
                        await CopyListItemsAsync(alreadyEx.ListId, def, selectedItemIds, isPartialSelection, listResult, itemDestTitle);

                    if (listResult.Status != CopyStatus.Failed)
                        listResult.Status = CopyStatus.Skipped;
                    MarkExistingUpdated(listResult, alreadyEx.Message, !string.IsNullOrEmpty(alreadyEx.ListId) && EffectiveCopyCustomColumns);
                }
                catch (Exception ex)
                {
                    listResult.Status       = CopyStatus.Failed;
                    listResult.ErrorMessage = ex.Message;
                }
            }
            StatusMessage = string.Empty;

            // For Site scope: also copy Site Pages (excluded from the Drives API, so not
            // captured by ReadAllLibraryDefinitionsAsync). The target Site Pages library
            // always exists — no creation step needed, just copy pages into it.
            if (IsSiteScope && CopyLibraryContent)
            {
                _copyCts.Token.ThrowIfCancellationRequested();
                StatusMessage = Loc.T("VM_CopyingSitePages");
                try
                {
                    var srcSitePages = await SpService.GetSitePagesLibraryAsync(SourceSiteId, SourceUrl.TrimEnd('/'));
                    var tgtSitePages = await SpService.GetSitePagesLibraryAsync(TargetSiteId, TargetUrl.TrimEnd('/'));

                    if (srcSitePages != null && tgtSitePages != null)
                    {
                        var pages = await SpService.GetChildrenAsync(
                            srcSitePages.DriveId, srcSitePages.Id,
                            srcSitePages.SiteId, srcSitePages.SiteUrl);

                        var pageJobs = pages
                            .Where(p => p.Type == NodeType.File)
                            .Select(p => new CopyJob
                            {
                                SourceDriveId                  = srcSitePages.DriveId,
                                SourceItemId                   = p.Id,
                                SourceName                     = p.Name,
                                SourceSiteUrl                  = SourceUrl.TrimEnd('/'),
                                SourceDisplayPath              = $"{SourceUrl.TrimEnd('/')}/SitePages/{p.Name}",
                                TargetDriveId                  = tgtSitePages.DriveId,
                                TargetParentItemId             = tgtSitePages.Id,
                                TargetSubFolderPath            = string.Empty,
                                TargetSiteId                   = TargetSiteId,
                                TargetSiteUrl                  = TargetUrl.TrimEnd('/'),
                                TargetDisplayPath              = $"{TargetUrl.TrimEnd('/')}/SitePages/{p.Name}",
                                TargetLibraryServerRelativeUrl = tgtSitePages.ServerRelativePath ?? string.Empty,
                                IsFolder                       = false,
                                IsPage                         = true,
                            })
                            .ToList();

                        if (pageJobs.Count > 0)
                        {
                            // Pre-add result rows — non-folder jobs are not dynamically added during execution
                            foreach (var pj in pageJobs)
                            {
                                var pr = new CopyResult
                                {
                                    FileName   = pj.SourceName,
                                    SourcePath = pj.SourceDisplayPath,
                                    TargetPath = pj.TargetDisplayPath,
                                };
                                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(pr));
                            }

                            // Build field cache for any custom columns on the Site Pages library
                            var pageBulkCache   = new Dictionary<string, Dictionary<string, object?>>();
                            var pagePermFlags   = new Dictionary<string, bool>();
                            if (srcSitePages.ServerRelativePath != null)
                            {
                                try
                                {
                                    var pagesListId = await SpService.GetListIdByServerRelativeUrlAsync(
                                        SourceUrl.TrimEnd('/'), srcSitePages.ServerRelativePath);
                                    if (EffectiveCopyCustomColumns)
                                    {
                                        var pageCols = await SpService.GetLibraryColumnsAsync(
                                            SourceUrl.TrimEnd('/'), pagesListId);
                                        if (pageCols.Count > 0)
                                        {
                                            // BulkReadCustomFieldsAsync keys its result by bare item ID;
                                            // CopyService's lookup (ApplyCustomColumnsIfAny) uses the
                                            // composite "{listId}:{itemId}" key the library-scope call
                                            // sites already re-key to (see the def.SourceListId path
                                            // above) — without this, every lookup here missed and page
                                            // custom columns silently never applied, with no error.
                                            var rawPageFields = await SpService.BulkReadCustomFieldsAsync(
                                                SourceUrl.TrimEnd('/'), pagesListId, pageCols,
                                                ct: _copyCts.Token);
                                            pageBulkCache = rawPageFields.ToDictionary(
                                                kvp => $"{pagesListId}:{kvp.Key}", kvp => kvp.Value);
                                        }
                                    }
                                    if (CopyPermissions)
                                        pagePermFlags = await SpService.BulkReadPermissionFlagsAsync(
                                            SourceUrl.TrimEnd('/'), pagesListId, _copyCts.Token);
                                }
                                catch (Exception ex) { ReportBulkReadFailure(SourceUrl, ex); }
                            }

                            await _copyService.ExecuteAsync(
                                pageJobs, CopyResults,
                                new CopyRunOptions
                                {
                                    OverwriteMode         = OverwriteMode,
                                    CopyVersions          = false,
                                    MaxParallel           = MaxParallelCopies,
                                    MaxVersions           = 0,
                                    CopyMode              = CopyMode.EnhancedRest,
                                    OnMetadataDone        = MakeMetadataProgress(),
                                    CopyPages             = true,
                                    RemapPageWebPartUrls  = RemapPageWebPartUrls,
                                    PreserveMetadata      = PreserveMetadata,
                                    ReapplyFolderMetadata = ReapplyFolderMetadataEveryRun,
                                    CopyCustomColumns     = EffectiveCopyCustomColumns,
                                    ColumnMappings        = [.. ColumnMappings],
                                    BulkFieldCache        = pageBulkCache,
                                    CopyPermissions       = CopyPermissions,
                                    PermissionService     = CopyPermissions ? _permissionCopyService : null,
                                    PermissionFlags       = pagePermFlags,
                                },
                                _copyCts.Token);
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    StatusMessage = Loc.T("VM_SitePagesWarning", ex.Message);
                }
                StatusMessage = string.Empty;
            }

            // Site scope: lookup site columns, now that the lists they point at exist on the target.
            if (IsSiteScope && CopySiteDefinitions)
            {
                _copyCts.Token.ThrowIfCancellationRequested();
                try
                {
                    var lookupNotes = await _libraryCopyService.CopySiteLookupColumnsAsync(
                        SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'), _copyCts.Token);
                    if (lookupNotes.Count > 0)
                    {
                        var lookupResult = new CopyResult
                        {
                            FileName     = Loc.T("VM_SiteLookupColumns"),
                            SourcePath   = SourceUrl.TrimEnd('/'),
                            TargetPath   = TargetUrl.TrimEnd('/'),
                            Status       = CopyStatus.Success,
                            ErrorMessage = string.Join("; ", lookupNotes.Take(5)),
                        };
                        System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(lookupResult));
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch { /* best effort */ }
            }

            // For Site scope: copy navigation (Quick Launch + Top Navigation Bar).
            if (IsSiteScope && CopyNavigation)
            {
                _copyCts.Token.ThrowIfCancellationRequested();
                StatusMessage = Loc.T("VM_CopyingNavigation");
                var navResult = new CopyResult
                {
                    FileName          = Loc.T("VM_Navigation"),
                    SourcePath        = $"{SourceUrl.TrimEnd('/')}/navigation",
                    TargetPath        = $"{TargetUrl.TrimEnd('/')}/navigation",
                    Status            = CopyStatus.Copying,
                    IsLibraryCreation = true,
                };
                System.Windows.Application.Current.Dispatcher.Invoke(() => CopyResults.Add(navResult));
                try
                {
                    await SpService.CopyNavigationAsync(SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'), quickLaunch: true);
                    await SpService.CopyNavigationAsync(SourceUrl.TrimEnd('/'), TargetUrl.TrimEnd('/'), quickLaunch: false);
                    navResult.Status = CopyStatus.Success;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    navResult.Status       = CopyStatus.Failed;
                    navResult.ErrorMessage = ex.Message;
                }
                StatusMessage = string.Empty;
            }
        }
        catch (OperationCanceledException) { StatusMessage = Loc.T("VM_CopyCancelled"); }
        catch (Exception ex)              { StatusMessage = Loc.T("VM_LibraryCopyError", ex.Message); }
        finally
        {
            // Release the scheduling sentinel (see its declaration above) now that every
            // MakeMetadataProgress() call this run will ever make has already happened — on any exit
            // path (success, cancellation, or exception). If no library/page copy ever ran, this is
            // the ONLY decrement and IsUpdatingMetadata (never set true) stays false, unchanged from
            // before. If at least one detached folder-metadata pass is still running, this drops the
            // count without zeroing it, so IsUpdatingMetadata correctly stays true until that pass's
            // own onMetadataDone callback brings the count the rest of the way to zero.
            OnOneMetadataPassDone(true);

            _copyEndTime   = DateTimeOffset.Now;
            progressTimer.Stop();
            IsCopying      = false;
            IsCopyComplete = true;
            TotalCount     = CopyResults.Count;
            CopyDuration   = FormatDuration(_copyEndTime.Value - _copyStartTime);
            UpdateProgress();
            OnPropertyChanged(nameof(SuccessCount));
            OnPropertyChanged(nameof(FailedCount));
            OnPropertyChanged(nameof(SkippedCount));
            OnPropertyChanged(nameof(CancelledCount));
            OnPropertyChanged(nameof(FileTotalCount));
            OnPropertyChanged(nameof(FileSuccessCount));
            OnPropertyChanged(nameof(FileFailedCount));
            OnPropertyChanged(nameof(FileSkippedCount));
            OnPropertyChanged(nameof(FileCancelledCount));
            OnPropertyChanged(nameof(CustomFieldWarningCount));
            SaveReport();
        }
    }
}
