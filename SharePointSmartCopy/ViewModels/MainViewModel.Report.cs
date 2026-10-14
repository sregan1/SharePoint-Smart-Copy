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

// Partial class MainViewModel: Report. The shared fields, constructor and core setup live in MainViewModel.cs.
public partial class MainViewModel : ObservableObject
{
    // ── Step 5: Report ────────────────────────────────────────────────────────

    [RelayCommand]
    private void ExportFilesReport()
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = Loc.T("VM_FilterCsvTxt"),
            FileName = $"{SiteUrlHelper.ReportFilenamePrefix(SourceUrl, TargetUrl, Settings.PrefixReportFilenamesWithSiteNames)}CopyReport_Files_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true) return;

        static string Csv(string? s) => $"\"{(s ?? "").Replace("\"", "\"\"")}\"";
        var sb = new StringBuilder();
        sb.AppendLine(Loc.T("VM_CsvHeader"));
        foreach (var r in CopyResults.Where(r => !r.IsPermissionResult))
        {
            sb.AppendLine($"{Csv(r.FileName)},{Csv(r.SourcePath)},{Csv(r.TargetPath)}," +
                          $"{r.Status},{r.VersionsCopied},{Csv(r.ErrorMessage)}," +
                          $"{r.PermissionStatus?.ToString() ?? ""},{Csv(r.PermissionDetails)}");
        }
        System.IO.File.WriteAllText(dlg.FileName, sb.ToString());
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
    }

    private bool CanExportVerificationReport() => !IsCopying && !IsVerifying && CopyJobs.Count > 0;

    [RelayCommand(CanExecute = nameof(CanExportVerificationReport))]
    private async Task ExportVerificationReportAsync()
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter   = Loc.T("VM_FilterXlsx"),
            // "DeepVerificationReport_" when the deep Office-file pass is on, so the file is
            // distinguishable from a standard verification at a glance. Live checkbox state.
            FileName = $"{SiteUrlHelper.ReportFilenamePrefix(SourceUrl, TargetUrl, Settings.PrefixReportFilenamesWithSiteNames)}{(DeepVerifyOfficeFiles ? "Deep" : "")}VerificationReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
        };
        if (dlg.ShowDialog() != true) return;

        _verifyCts?.Dispose();
        _verifyCts   = new CancellationTokenSource();
        IsVerifying  = true;
        // Verification previously only appended to whatever log file a prior copy in this session had
        // already opened — running Verify on its own (no copy this session, or after reopening the app)
        // left zero on-disk record of the run, so a real failure had nothing to inspect afterward.
        // Give it its own log file, same pattern as StartCopyAsync.
        StartActivityLogFile();
        if (ActivityLogPath != null)
            PushActivity(Loc.T("VM_ActivityLogPath", ActivityLogPath));
        // (No crash journal here: a verification is not a copy run — it has no results of its own to recover.)
        // StatusMessage previously only ever got updated by scan progress below — once the scan
        // finished and the deep-verify pass started, its own progress messages (which DO arrive here
        // via onActivity/PushActivity, appended to the scrolling log) never replaced the single
        // status line, which stayed frozen on the last "found X source, Y target" text for the
        // entire deep-verify phase, making an actively-running verification look stalled. "⚠"-
        // prefixed messages (throttle waits, scan errors) are excluded — those are transient blips
        // that belong in the log, not overwriting the main status line on every occurrence.
        // Live checkbox state, not a settings re-read — see the field's own doc comment. Read up front
        // so every status/activity line below can label the run as "Deep verification" when the deep
        // Office-file pass is enabled, matching the checkbox the user ticked.
        bool deepVerify   = DeepVerifyOfficeFiles;
        string verifyName = deepVerify ? Loc.T("VM_DeepVerification") : Loc.T("VM_Verification");  // sentence-initial
        string verifyLow  = deepVerify ? Loc.T("VM_DeepVerificationLow") : Loc.T("VM_VerificationLow");  // mid-sentence
        string verifyIng  = deepVerify ? Loc.T("VM_DeepVerifying") : Loc.T("VM_Verifying");

        var onActivity = new Progress<string>(msg =>
        {
            PushActivity(msg);
            if (!msg.StartsWith('⚠'))
                StatusMessage = msg;
        });
        var onScanned  = new Progress<VerificationReportService.ScanProgress>(p =>
            StatusMessage = Loc.T("VM_VerifyScanProgress", verifyIng, p.SourceFilesFound, p.TargetFilesFound));

        try
        {
            StatusMessage = Loc.T("VM_VerifyRescanning", verifyLow);
            var roots  = VerificationRoot.FromCopyJobs(CopyJobs);
            var result = await _verificationReportService.RunAsync(
                roots, MaxParallelCopies, onActivity, onScanned, _verifyCts.Token, deepVerify);
            StatusMessage = Loc.T("VM_VerifyWriting", verifyLow);
            // Off the UI thread: ClosedXML SaveAs is CPU-heavy on 100k+ row runs.
            await Task.Run(() => ExcelReportWriter.Write(dlg.FileName, result));

            int matched         = result.Comparison.Count(r => r.Status == ComparisonStatus.Match);
            int contentMismatch = result.Comparison.Count(r => r.Status == ComparisonStatus.ContentMismatch);
            int dateMismatch    = result.Comparison.Count(r => r.Status == ComparisonStatus.DateMismatch);
            int onlyInSource    = result.Comparison.Count(r => r.Status == ComparisonStatus.OnlyInSource);
            int onlyInTarget    = result.Comparison.Count(r => r.Status == ComparisonStatus.OnlyInTarget);
            int unverified      = result.Comparison.Count(r => r.Status == ComparisonStatus.Unverified);
            PushActivity(Loc.T("VM_VerifyComplete", verifyName, matched, contentMismatch, dateMismatch, onlyInSource, onlyInTarget)
                + (unverified > 0 ? Loc.T("VM_VerifyUnverified", unverified) : ""));
            PushActivity(Loc.T("VM_VerifyReportWritten", verifyName, dlg.FileName));
            if (result.ScanErrors.Count > 0)
                PushActivity(Loc.T("VM_VerifyScanErrors", result.ScanErrors.Count));
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
            StatusMessage = string.Empty;
        }
        catch (OperationCanceledException) { StatusMessage = Loc.T("VM_VerifyCancelled", verifyName); PushActivity(Loc.T("VM_VerifyCancelled", verifyName)); }
        catch (Exception ex)
        {
            StatusMessage = Loc.T("VM_VerifyError", verifyName, ex.Message);
            PushActivity(Loc.T("VM_VerifyErrorWarn", verifyName, ex.Message));
        }
        finally { IsVerifying = false; }
    }

    [RelayCommand]
    private void CancelVerification() => _verifyCts?.Cancel();

    // ── Navigation ────────────────────────────────────────────────────────────

    // Set to true by MainViewModel.Demo.cs LoadDemoData(); always false in production.
    private bool _isDemoMode = false;
    public bool IsDemoMode => _isDemoMode;
    partial void AdvanceDemoToReport();  // implemented in MainViewModel.Demo.cs; no-op when absent

    [RelayCommand(CanExecute = nameof(CanGoBack))]
    private void Back()
    {
        if (CurrentStep == 3)
            ColumnMappings.Clear();
        // Leaving the progress/report steps invalidates the completed run: without this, Back →
        // change options → Next silently jumped to the OLD results ("case 3: if (IsCopyComplete)")
        // and the new options never ran, while the UI presented them as applied. Next now re-runs
        // the copy with whatever is currently selected.
        if (CurrentStep == 4 && IsCopyComplete && !IsUpdatingMetadata)
            IsCopyComplete = false;
        CurrentStep--;
    }

    private bool CanGoBack() => CurrentStep > 0 && !IsCopying;

    [RelayCommand(CanExecute = nameof(CanGoNext))]
    private void Next()
    {
        switch (CurrentStep)
        {
            case 0 when SourceConnected || IsDemoMode:
                CurrentStep = 1;
                if (!IsDemoMode) _ = LoadLibrariesAsync();
                break;
            case 1:
                OnPropertyChanged(nameof(IsItemSelectionActive));
                NextCommand.NotifyCanExecuteChanged();
                CurrentStep = 2;
                break;
            case 2 when (TargetConnected && (NeedsTargetFolder ? SelectedTargetFolder != null : true)) || IsDemoMode:
                if (!IsDemoMode)
                {
                    if (NeedsTargetFolder) BuildCopyJobs();
                    // Item-level copies write into the chosen destination list (or a new one) —
                    // reflect that name in the preview's Destination Name column so what the
                    // user typed at the Target step is what the preview shows.
                    if (IsItemSelectionActive && !string.IsNullOrWhiteSpace(EffectiveDestinationListName))
                    {
                        foreach (var lib in SourceLibraries.Where(l => l.IsCustomList &&
                                     (l.IsChecked == null ||
                                      l.Children.Any(c => c.Type == NodeType.ListItem && c.IsChecked == true))))
                            lib.OverrideName = EffectiveDestinationListName;
                        OnPropertyChanged(nameof(LibraryPreviewItems));
                    }
                    // Eagerly load target columns for mapping dialog; store task so dialog can await it
                    _columnLoadTask = LoadTargetColumnsAsync();
                }
                CurrentStep = 3;
                break;
            case 3 when IsLibraryOrSiteScope || CopyJobs.Count > 0 || IsDemoMode:
                if (IsCopyComplete)
                {
                    // Copy already ran — navigate forward without re-running.
                    CurrentStep = 4;
                    break;
                }
                if (!IsDemoMode)
                {
                    Settings.PreferredCopyMode    = CopyMode;
                    Settings.OverwriteMode        = OverwriteMode;
                    // Keep the legacy bool in sync for older builds reading the same file.
                    Settings.OverwriteFiles       = OverwriteMode == Models.OverwriteMode.Overwrite;
                    Settings.CopyVersions         = CopyVersions;
                    Settings.CopyAllVersions      = CopyAllVersions;
                    Settings.MaxVersions          = MaxVersions;
                    Settings.MaxParallelCopies    = MaxParallelCopies;
                    Settings.CopyCustomColumns    = CopyCustomColumns;
                    Settings.CopyLibraryContent   = CopyLibraryContent;
                    Settings.RemapPageWebPartUrls = RemapPageWebPartUrls;
                    Settings.PreserveMetadata     = PreserveMetadata;
                    Settings.ReapplyFolderMetadataEveryRun = ReapplyFolderMetadataEveryRun;
                    Settings.CopyNavigation       = CopyNavigation;
                    Settings.Scope                = CopyScope;
                    Settings.Save();
                    _ = IsLibraryOrSiteScope ? StartLibraryCopyAsync() : StartCopyAsync();
                }
                CurrentStep = 4;
                break;
            case 4 when (IsCopyComplete && !IsUpdatingMetadata) || IsDemoMode:
                if (IsDemoMode) AdvanceDemoToReport();
                CurrentStep = 5;
                break;
        }
    }

    private bool CanGoNext() => IsDemoMode || CurrentStep switch
    {
        0 => SourceConnected,
        1 => IsSiteScope
               ? SourceConnected
               : SourceLibraries.Any(l => l.GetCheckedNodes().Any()),
        2 => TargetConnected &&
             (NeedsTargetFolder ? SelectedTargetFolder != null : true) &&
             (!IsItemSelectionActive || (SelectedTargetList != null &&
                 (!IsCreatingNewList || !string.IsNullOrWhiteSpace(NewListName)))),
        3 => IsLibraryOrSiteScope || CopyJobs.Count > 0,
        4 => IsCopyComplete && !IsUpdatingMetadata,
        _ => false
    };

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string FormatDuration(TimeSpan ts)
    {
        if (ts.TotalHours >= 1)   return Loc.T("VM_DurationHMS", (int)ts.TotalHours, ts.Minutes, ts.Seconds);
        if (ts.TotalMinutes >= 1) return Loc.T("VM_DurationMS", (int)ts.TotalMinutes, ts.Seconds);
        return Loc.T("VM_DurationS", ts.Seconds);
    }

    // ── Crash journal (see RunJournal) ────────────────────────────────────────

    private RunJournal? _journal;
    private readonly HashSet<CopyResult> _journaledResults = [];
    private System.Windows.Threading.DispatcherTimer? _journalTimer;

    private static SavedReportItem ToSavedItem(CopyResult r) => new()
    {
        FileName           = r.FileName,
        SourcePath         = r.SourcePath,
        TargetPath         = r.TargetPath,
        Status             = r.Status,
        VersionsCopied     = r.VersionsCopied,
        VersionsTotal      = r.VersionsTotal,
        ErrorMessage       = r.ErrorMessage,
        IsPermissionResult = r.IsPermissionResult,
        PermissionStatus   = r.PermissionStatus,
        PermissionDetails  = r.PermissionDetails,
        CustomFieldStatus  = r.CustomFieldStatus,
        CustomFieldDetails = r.CustomFieldDetails,
    };

    private List<SavedReportRoot> BuildSavedRoots() => CopyJobs.Select(j => new SavedReportRoot
    {
        SourceDriveId       = j.SourceDriveId,
        SourceItemId        = j.SourceItemId,
        SourceName          = j.SourceName,
        IsFolder            = j.IsFolder,
        IsLibrary           = j.IsLibrary,
        TargetDriveId       = j.TargetDriveId,
        TargetParentItemId  = j.TargetParentItemId,
        TargetSubFolderPath = j.TargetSubFolderPath,
    }).ToList();

    // Starts journaling this run: a header now, then every finished row every few seconds. If the app
    // dies mid-run, the next launch turns what was journaled into an "Interrupted" History entry.
    private void BeginRunJournal()
    {
        EndRunJournal(complete: false);
        _journaledResults.Clear();
        _journal = RunJournal.Start(new SavedReport
        {
            Id          = _copyStartTime.ToString("yyyyMMdd_HHmmss"),
            Timestamp   = _copyStartTime,
            SourceUrl   = SourceUrl,
            TargetUrl   = TargetUrl,
            CopyMode    = CopyMode,
            Roots       = BuildSavedRoots(),
            Interrupted = true,
        });
        if (_journal == null) return;

        _journalTimer ??= new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(15) };
        _journalTimer.Tick -= OnJournalTick;
        _journalTimer.Tick += OnJournalTick;
        _journalTimer.Start();
    }

    private void OnJournalTick(object? sender, EventArgs e) => FlushRunJournal();

    // Writes the rows that reached a final state since the last tick. Runs on the UI thread, where
    // CopyResults is only ever modified, and is cheap (a pass over the rows plus a small append).
    private void FlushRunJournal()
    {
        if (_journal == null) return;
        var finished = new List<SavedReportItem>();
        foreach (var r in CopyResults)
        {
            if (r.Status == CopyStatus.Copying) continue;
            if (_journaledResults.Add(r)) finished.Add(ToSavedItem(r));
        }
        if (finished.Count > 0) _journal.Append(finished);
    }

    // complete=true: the run's normal report was saved, so the journal is no longer needed.
    private void EndRunJournal(bool complete)
    {
        _journalTimer?.Stop();
        var journal = _journal;
        _journal = null;
        if (journal == null) return;
        if (complete) journal.Complete(); else journal.Dispose();
    }

    private void SaveReport()
    {
        try
        {
            var report = new SavedReport
            {
                Id           = _copyStartTime.ToString("yyyyMMdd_HHmmss"),
                Timestamp    = _copyStartTime,
                Duration     = DateTimeOffset.Now - _copyStartTime,
                SourceUrl    = SourceUrl,
                TargetUrl    = TargetUrl,
                SuccessCount   = SuccessCount,
                FailedCount    = FailedCount,
                SkippedCount   = SkippedCount,
                CancelledCount = CancelledCount,
                TotalCount     = TotalCount,
                TotalSize      = HasKnownTotalSize ? _bytesFinalTally : 0,
                CopyMode     = CopyMode,
                Items        = CopyResults.Select(ToSavedItem).ToList(),
                Roots        = BuildSavedRoots(),
                // Kept with the run so History can offer "Map unresolved users" later, not just this screen.
                UnresolvedUsers = [.. SpService.Unresolved.Snapshot()],
            };
            // The full report supersedes the journal — but only once it is really on disk.
            if (ReportHistoryService.Save(report)) EndRunJournal(complete: true);
        }
        catch { /* non-critical */ }
    }

    private static string BuildPath(SharePointNode node)
    {
        var parts   = new List<string>();
        var current = (SharePointNode?)node;
        while (current != null)
        {
            parts.Insert(0, current.Name);
            current = current.Parent;
        }
        return string.Join("/", parts);
    }

    // Returns the path of `node` relative to `libraryRoot` by walking the parent chain.
    // Returns empty string when node IS the library root.
    private static string BuildRelativePath(SharePointNode node, SharePointNode libraryRoot)
    {
        var parts   = new List<string>();
        var current = node;
        while (current != null && current != libraryRoot)
        {
            parts.Insert(0, current.Name);
            current = current.Parent;
        }
        return string.Join("/", parts);
    }

}
