using System.Collections.Concurrent;
using System.IO;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Independently re-scans a source location and a target location via fresh Graph calls (never
// reusing in-memory CopyResults) and compares them by relative path, for post-copy verification.
public sealed class VerificationReportService(SharePointService spService)
{
    // Mirrors CopyService/MigrationJobService's throttle-driven parallelism reporting.
    public event Action<int>? ParallelismChanged;

    public sealed record Result(
        List<ScannedFile> SourceFiles,
        List<ScannedFile> TargetFiles,
        List<ComparisonRow> Comparison,
        List<string> ScanErrors);

    // Reported as files are discovered on each side — kept separate (rather than one combined
    // count) so the displayed text can say exactly what it means: files found so far in the
    // source scan and files found so far in the target scan, never folders or file versions.
    public readonly record struct ScanProgress(int SourceFilesFound, int TargetFilesFound);

    public async Task<Result> RunAsync(
        IReadOnlyList<VerificationRoot> roots,
        int maxParallel,
        IProgress<string>? activityLog,
        IProgress<ScanProgress>? progress,
        CancellationToken ct)
    {
        int softStart = Math.Min(maxParallel, 8);
        using var controller = new AdaptiveParallelismController(maxParallel, softStart);
        controller.LimitChanged += n => ParallelismChanged?.Invoke(n);
        void onThrottle(TimeSpan delay, int _, int __, string? ___) => controller.StepDown(delay);
        spService.Throttled += onThrottle;

        // Without this, a throttle wait (up to 120s, observed recurring for over an hour on a
        // busy tenant) is completely silent — the scan-progress text only updates when a new file
        // is found, so a long Retry-After window is indistinguishable from a hang. Mirrors the
        // same "⚠ Graph throttled — waiting Ns" message and 5s dedup CopyService.ExecuteAsync
        // already uses for copy runs.
        if (activityLog != null)
        {
            var throttleLogLock = new object();
            var lastThrottleLog = DateTimeOffset.MinValue;
            spService.Throttled += (delay, attempt, max, reason) =>
            {
                lock (throttleLogLock)
                {
                    var now = DateTimeOffset.UtcNow;
                    if (now - lastThrottleLog < TimeSpan.FromSeconds(5)) return;
                    lastThrottleLog = now;
                }
                activityLog.Report($"⚠ Graph throttled — waiting {delay.TotalSeconds:0}s"
                    + (string.IsNullOrEmpty(reason) ? "" : $" [{reason}]"));
            };
        }

        try
        {
            var scanErrors = new ConcurrentBag<string>();
            int sourceCount = 0, targetCount = 0;

            // Reporting on every single file (up to ~240k dispatcher round-trips for a 120k-per-side
            // scan, now more concurrent than ever with the folder walk parallelized above) is exactly
            // the UI-freeze pattern CopyService already hit and fixed once for CopyResults — throttle
            // to a fixed cadence instead, plus one guaranteed final report so the displayed count is
            // never stale after a throttled-out update.
            var progressLock = new object();
            var lastReport   = DateTime.MinValue;
            var reportInterval = TimeSpan.FromMilliseconds(250);
            void ReportProgress(int s, int t)
            {
                if (progress == null) return;
                lock (progressLock)
                {
                    var now = DateTime.UtcNow;
                    if (now - lastReport < reportInterval) return;
                    lastReport = now;
                }
                progress.Report(new ScanProgress(s, t));
            }
            void BumpSource() => ReportProgress(Interlocked.Increment(ref sourceCount), targetCount);
            void BumpTarget() => ReportProgress(sourceCount, Interlocked.Increment(ref targetCount));

            // Source roots: de-duplicate by (drive, item) — the actual scan starting point.
            var sourceRoots = roots
                .Select(r => (r.SourceDriveId, r.SourceItemId, r.SourceName, r.IsFolder, r.IsLibrary))
                .Distinct()
                .ToList();

            // Target roots: TargetParentItemId is the library root, and TargetSubFolderPath is the
            // *destination container* the user picked in Step 2 — NOT the copied item's own path.
            // The copied item itself lands at TargetSubFolderPath/SourceName (File and Folder jobs
            // both get their own name appended there); a Library job's contents land directly under
            // TargetSubFolderPath with no wrapper name. NavigatePath is what we resolve-then-scan
            // from; BasePath is what the results get labeled with — matching the source side's
            // SourceName/"" labeling so the two scans' relative paths actually align for the join.
            var targetRoots = roots
                .Select(r => new TargetRoot(
                    r.TargetDriveId,
                    r.TargetParentItemId,
                    r.IsLibrary ? r.TargetSubFolderPath : CombinePath(r.TargetSubFolderPath, r.SourceName),
                    r.IsLibrary ? "" : r.SourceName,
                    r.IsFolder,
                    r.IsLibrary))
                .Distinct()
                .ToList();

            var sourceTask = ScanSourceRootsAsync(sourceRoots, controller, scanErrors, activityLog, BumpSource, ct);
            var targetTask = ScanTargetRootsAsync(targetRoots, controller, scanErrors, activityLog, BumpTarget, ct);
            await Task.WhenAll(sourceTask, targetTask);

            var sourceFiles = await sourceTask;
            var targetFiles = await targetTask;
            progress?.Report(new ScanProgress(sourceCount, targetCount)); // final tally, bypassing the throttle
            var comparison  = Merge(sourceFiles, targetFiles);

            return new Result(sourceFiles, targetFiles, comparison, scanErrors.ToList());
        }
        finally
        {
            spService.Throttled -= onThrottle;
        }
    }

    private async Task<List<ScannedFile>> ScanSourceRootsAsync(
        List<(string SourceDriveId, string SourceItemId, string SourceName, bool IsFolder, bool IsLibrary)> roots,
        AdaptiveParallelismController controller, ConcurrentBag<string> scanErrors,
        IProgress<string>? activityLog, Action bump, CancellationToken ct)
    {
        var results = new ConcurrentBag<ScannedFile>();
        await Task.WhenAll(roots.Select(async root =>
        {
            try
            {
                if (!root.IsFolder && !root.IsLibrary)
                {
                    // Plain single-file root — fetch directly, no recursion.
                    var single = await spService.GetFileForVerificationAsync(
                        root.SourceDriveId, root.SourceItemId, root.SourceName);
                    if (single != null) { results.Add(single); bump(); }
                    return;
                }

                var basePath = root.IsLibrary ? "" : root.SourceName;
                await foreach (var f in spService.EnumerateFilesWithMetadataAsync(
                    root.SourceDriveId, root.SourceItemId, basePath, controller, ct))
                {
                    results.Add(f);
                    bump();
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                scanErrors.Add($"Source '{root.SourceName}': {ex.Message}");
                activityLog?.Report($"⚠ Could not scan source root '{root.SourceName}': {ex.Message}");
            }
        }));
        return results.ToList();
    }

    private sealed record TargetRoot(
        string DriveId, string ParentItemId, string NavigatePath, string BasePath, bool IsFolder, bool IsLibrary);

    private static string CombinePath(string basePath, string name) =>
        string.IsNullOrEmpty(basePath) ? name : $"{basePath}/{name}";

    private async Task<List<ScannedFile>> ScanTargetRootsAsync(
        List<TargetRoot> roots,
        AdaptiveParallelismController controller, ConcurrentBag<string> scanErrors,
        IProgress<string>? activityLog, Action bump, CancellationToken ct)
    {
        var results = new ConcurrentBag<ScannedFile>();
        await Task.WhenAll(roots.Select(async root =>
        {
            var label = string.IsNullOrEmpty(root.NavigatePath) ? "(library root)" : root.NavigatePath;
            try
            {
                // Navigate from the library root down to the actual copied item — TargetParentItemId
                // is never the item itself, so this resolution is required before scanning anything.
                string? scanRootId = string.IsNullOrEmpty(root.NavigatePath)
                    ? root.ParentItemId
                    : await spService.ResolveItemIdByPathAsync(root.DriveId, root.ParentItemId, root.NavigatePath);
                if (scanRootId == null)
                {
                    scanErrors.Add($"Target '{label}': not found (it may have been deleted or renamed since the copy)");
                    activityLog?.Report($"⚠ Target '{label}' no longer exists");
                    return;
                }

                if (!root.IsFolder && !root.IsLibrary)
                {
                    var single = await spService.GetFileForVerificationAsync(root.DriveId, scanRootId, root.BasePath);
                    if (single != null) { results.Add(single); bump(); }
                    return;
                }

                await foreach (var f in spService.EnumerateFilesWithMetadataAsync(
                    root.DriveId, scanRootId, root.BasePath, controller, ct))
                {
                    results.Add(f);
                    bump();
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                scanErrors.Add($"Target '{label}': {ex.Message}");
                activityLog?.Report($"⚠ Could not scan target root '{label}': {ex.Message}");
            }
        }));
        return results.ToList();
    }

    // Office/OLE compound-document formats: SharePoint's backend re-serializes these independently
    // of content changes (indexing, thumbnails, co-authoring), so size/hash are unreliable — modified
    // date is used instead (see ComparisonStatus for the full rationale). Internal (not private) so
    // ExcelReportWriter can reuse the same list when deciding which raw value to display per row.
    //
    // Covers both container families: modern OOXML (.docx/.xlsx/.pptx, ZIP-based) AND legacy
    // binary Office formats (.doc/.xls/.ppt, OLE Compound File Binary Format) plus other formats
    // sharing that same OLE container (.msg). Confirmed 2026-07-03: a real verification run showed
    // 100% of ContentMismatch rows were legacy .xls and .msg files with completely different hashes
    // on both sides — the OLE container's internal metadata streams (summary info, calculation
    // chains, etc.) get touched by background processing the same way OOXML's ZIP does, so the
    // original list (which only covered modern OOXML) was producing false positives for these.
    internal static readonly HashSet<string> OfficeReserializedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        // Modern OOXML (ZIP-based)
        ".docx", ".xlsx", ".pptx", ".docm", ".xlsm", ".pptm",
        ".dotx", ".xltx", ".potx", ".ppsx", ".ppsm", ".dotm",
        // Legacy binary Office formats (OLE Compound File Binary Format)
        ".doc", ".dot", ".xls", ".xlt", ".xla", ".ppt", ".pot", ".pps", ".ppa",
        // Other OLE-compound formats subject to the same background metadata churn
        ".msg"
    };

    // Absorbs clock/rounding differences (e.g. Migration API manifest timestamps) without masking
    // a genuine metadata-preservation failure.
    private static readonly TimeSpan DateMismatchTolerance = TimeSpan.FromSeconds(5);

    private static List<ComparisonRow> Merge(List<ScannedFile> sourceFiles, List<ScannedFile> targetFiles)
    {
        var targetByPath = new Dictionary<string, ScannedFile>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in targetFiles) targetByPath.TryAdd(t.RelativePath, t);

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var rows = new List<ComparisonRow>();

        foreach (var s in sourceFiles)
        {
            seen.Add(s.RelativePath);
            targetByPath.TryGetValue(s.RelativePath, out var t);
            rows.Add(new ComparisonRow
            {
                RelativePath   = s.RelativePath,
                Status         = ClassifyMatch(s, t),
                SourceHash     = s.QuickXorHash,
                TargetHash     = t?.QuickXorHash,
                SourceModified = s.LastModified,
                TargetModified = t?.LastModified
            });
        }

        foreach (var t in targetFiles)
        {
            if (seen.Contains(t.RelativePath)) continue;
            rows.Add(new ComparisonRow
            {
                RelativePath   = t.RelativePath,
                Status         = ComparisonStatus.OnlyInTarget,
                TargetHash     = t.QuickXorHash,
                TargetModified = t.LastModified
            });
        }

        return rows;
    }

    private static ComparisonStatus ClassifyMatch(ScannedFile s, ScannedFile? t)
    {
        if (t == null) return ComparisonStatus.OnlyInSource;

        if (OfficeReserializedExtensions.Contains(Path.GetExtension(s.RelativePath)))
        {
            // Hash/size are unreliable here — modified date is the signal instead, since the app is
            // already responsible for preserving it onto the target.
            if (s.LastModified is not { } srcDate || t.LastModified is not { } tgtDate)
                return ComparisonStatus.Match; // missing signal on either side → don't manufacture a false mismatch
            return (srcDate - tgtDate).Duration() <= DateMismatchTolerance
                ? ComparisonStatus.Match
                : ComparisonStatus.DateMismatch;
        }

        if (s.QuickXorHash == null || t.QuickXorHash == null)
            return ComparisonStatus.Match; // missing signal on either side → don't manufacture a false mismatch

        return string.Equals(s.QuickXorHash, t.QuickXorHash, StringComparison.Ordinal)
            ? ComparisonStatus.Match
            : ComparisonStatus.ContentMismatch;
    }
}
