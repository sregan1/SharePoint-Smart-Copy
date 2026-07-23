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
        CancellationToken ct,
        // Opt-in "Deep verify Office files" (see Docs/DEEP-VERIFY-PLAN.md). Callers must pass the
        // LIVE checkbox state here, not a settings re-read — this is the only gate for the whole
        // feature. Every candidate is deep-verified when this is on — no count cap or time budget.
        bool deepVerifyOffice = false)
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
        // Named so it can be unsubscribed in the finally — an anonymous handler leaked once per
        // verification run onto the app-lifetime service (duplicate throttle lines on later runs,
        // plus closed HistoryDialog UI kept alive via the captured IProgress).
        Action<TimeSpan, int, int, string?>? onThrottleLog = null;
        if (activityLog != null)
        {
            var throttleLogLock = new object();
            var lastThrottleLog = DateTimeOffset.MinValue;
            onThrottleLog = (delay, attempt, max, reason) =>
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
            spService.Throttled += onThrottleLog;
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
            var (comparison, deepCandidates) = Merge(sourceFiles, targetFiles);

            if (deepVerifyOffice && deepCandidates.Count > 0)
                await RunDeepVerifyPassAsync(deepCandidates, controller, maxParallel, activityLog, ct);

            return new Result(sourceFiles, targetFiles, comparison, scanErrors.ToList());
        }
        finally
        {
            spService.Throttled -= onThrottle;
            if (onThrottleLog != null) spService.Throttled -= onThrottleLog;
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
    // Despite the name, the combined set below also covers OneNote (see OneNoteExtensions) for a
    // related but distinct reason — different mechanism, same size/hash-unreliable symptom.
    //
    // Covers both container families used by Office applications, not just Word/Excel/PowerPoint's
    // primary document types:
    //   - Modern OOXML/ZIP-based: Word, Excel (including the binary-sheet .xlsb variant, which is
    //     still a ZIP with the same docProps/customXml parts), PowerPoint, and Visio, plus their
    //     templates and add-ins.
    //   - Legacy OLE Compound File Binary Format: the pre-2007 Word/Excel/PowerPoint/Visio formats,
    //     Outlook .msg, Publisher, and Project — all share the same OLE metadata-stream container
    //     that gets touched by SharePoint's background processing the same way OOXML's ZIP does.
    // Confirmed 2026-07-03: a real verification run showed 100% of ContentMismatch rows were legacy
    // .xls and .msg files with completely different hashes on both sides despite identical content —
    // the original list (modern OOXML only) was producing false positives for every legacy format.
    // Confirmed 2026-07-06: .xlsb (missing from the original OOXML list) showed the same false
    // ContentMismatch pattern — any Office container format not in this set will.
    //
    // Split into OOXML (ZIP-based) and OLE (compound-binary) so the opt-in deep-verify pass can
    // target OOXML only — OpcDeepComparer opens the file as a ZIP, which OLE formats aren't.
    internal static readonly HashSet<string> OoxmlExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        // Word
        ".docx", ".docm", ".dotx", ".dotm",
        // Excel, including the binary-sheet .xlsb variant (still a ZIP container with the same
        // docProps/customXml parts SharePoint rewrites)
        ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam",
        // PowerPoint
        ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".ppam", ".sldx", ".sldm",
        // Visio
        ".vsdx", ".vsdm", ".vssx", ".vssm", ".vstx", ".vstm",
    };

    internal static readonly HashSet<string> OleExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        // Legacy binary Office formats (OLE Compound File Binary Format) — Word/Excel/PowerPoint
        ".doc", ".dot", ".xls", ".xlt", ".xla", ".ppt", ".pot", ".pps", ".ppa",
        // Legacy binary Office formats (OLE Compound File Binary Format) — Visio/Publisher/Project
        ".vsd", ".vst", ".vss", ".pub", ".mpp",
        // Other OLE-compound formats subject to the same background metadata churn
        ".msg"
    };

    // OneNote section/table-of-contents files: unlike the OOXML/OLE formats above, these aren't
    // touched by SharePoint's backend re-serialization — the churn comes from OneNote's own
    // MS-ONESTORE storage format and its client/server sync-and-consolidation behavior, which can
    // rewrite a section file's bytes with no logical content change. Different mechanism, same
    // false-positive symptom, so it gets the same date-based fallback rather than its own bespoke
    // path. Confirmed 2026-07-21: .one files were showing ContentMismatch on size alone despite the
    // user confirming identical content — .one/.onetoc2 were previously in neither set here.
    internal static readonly HashSet<string> OneNoteExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".one", ".onetoc2",
    };

    internal static readonly HashSet<string> OfficeReserializedExtensions =
        new(OoxmlExtensions.Concat(OleExtensions).Concat(OneNoteExtensions), StringComparer.OrdinalIgnoreCase);

    // Absorbs clock/rounding differences (e.g. Migration API manifest timestamps) without masking
    // a genuine metadata-preservation failure.
    private static readonly TimeSpan DateMismatchTolerance = TimeSpan.FromSeconds(5);

    // A row whose cheap-tier signals disagreed (or were missing) on an OOXML file — a candidate
    // for the opt-in deep-verify pass, which needs both sides' DriveId/ItemId to download them.
    private sealed record DeepVerifyCandidate(ComparisonRow Row, ScannedFile Source, ScannedFile Target);

    private static (List<ComparisonRow> Rows, List<DeepVerifyCandidate> DeepVerifyCandidates) Merge(
        List<ScannedFile> sourceFiles, List<ScannedFile> targetFiles)
    {
        var targetByPath = new Dictionary<string, ScannedFile>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in targetFiles) targetByPath.TryAdd(t.RelativePath, t);

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var rows = new List<ComparisonRow>();
        var candidates = new List<DeepVerifyCandidate>();

        foreach (var s in sourceFiles)
        {
            seen.Add(s.RelativePath);
            targetByPath.TryGetValue(s.RelativePath, out var t);
            var row = new ComparisonRow
            {
                RelativePath   = s.RelativePath,
                Status         = ClassifyMatch(s, t),
                SourceHash     = s.QuickXorHash,
                TargetHash     = t?.QuickXorHash,
                SourceModified = s.LastModified,
                TargetModified = t?.LastModified,
                SourceSize     = s.Size,
                TargetSize     = t?.Size
            };
            rows.Add(row);

            // Candidates need both sides present (to download both) and an OOXML extension (the
            // deep comparer opens the file as a ZIP, which OLE formats aren't). Excludes the
            // trusted equal-hash fast path — those are never candidates, deep verify would just
            // re-confirm what the hash already proved.
            if (t != null && OoxmlExtensions.Contains(Path.GetExtension(s.RelativePath)) &&
                !(s.QuickXorHash != null && t.QuickXorHash != null &&
                  string.Equals(s.QuickXorHash, t.QuickXorHash, StringComparison.Ordinal)))
            {
                candidates.Add(new DeepVerifyCandidate(row, s, t));
            }
        }

        foreach (var t in targetFiles)
        {
            if (seen.Contains(t.RelativePath)) continue;
            rows.Add(new ComparisonRow
            {
                RelativePath   = t.RelativePath,
                Status         = ComparisonStatus.OnlyInTarget,
                TargetHash     = t.QuickXorHash,
                TargetModified = t.LastModified,
                TargetSize     = t.Size
            });
        }

        return (rows, candidates);
    }

    // Downloads both copies of each OOXML candidate and compares their content parts, ignoring the
    // parts SharePoint rewrites independently of content. See Docs/DEEP-VERIFY-PLAN.md §5.5.
    private async Task RunDeepVerifyPassAsync(
        List<DeepVerifyCandidate> candidates,
        AdaptiveParallelismController controller,
        int maxParallel,
        IProgress<string>? activityLog,
        CancellationToken ct)
    {
        // Per-file download size cap — a handful of pathologically huge mismatched files shouldn't
        // be able to exhaust memory. This doesn't shrink the run — it just marks that one file
        // NotComparable and moves on.
        const long SizeCapBytes = 250L * 1024 * 1024;

        activityLog?.Report(
            $"{candidates.Count:N0} Office file(s) need deep verification (hash/date signals disagreed)...");

        int attempted = 0, resolvedToMatch = 0, mismatched = 0, inconclusive = 0;

        var reportLock = new object();
        var lastReport = DateTimeOffset.MinValue;
        void ReportDeepProgress(int done, int total)
        {
            if (activityLog == null) return;
            lock (reportLock)
            {
                var now = DateTimeOffset.UtcNow;
                if (now - lastReport < TimeSpan.FromSeconds(3) && done < total) return;
                lastReport = now;
            }
            activityLog.Report($"Deep-verifying Office files: {done:N0} / {total:N0}");
        }

        // The outer degree just needs to be high enough that AdaptiveParallelismController (not
        // this loop) is the real gate — same shape as the migration/copy pipelines elsewhere in
        // this app, which pair a fixed-width Parallel.ForEachAsync with an adaptive inner gate.
        await Parallel.ForEachAsync(candidates,
            new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, maxParallel), CancellationToken = ct },
            async (candidate, itemCt) =>
            {
                await controller.WaitAsync(itemCt);
                try
                {
                    var (outcome, skipReason) = await DeepVerifyOneAsync(candidate, SizeCapBytes, itemCt);
                    ApplyDeepVerifyOutcome(candidate.Row, outcome, skipReason);
                    Interlocked.Increment(ref attempted);
                    switch (outcome.Result)
                    {
                        case OpcCompareResult.VolatileOnlyDifferences: Interlocked.Increment(ref resolvedToMatch); break;
                        case OpcCompareResult.ContentMismatch:         Interlocked.Increment(ref mismatched);      break;
                        case OpcCompareResult.NotComparable:           Interlocked.Increment(ref inconclusive);    break;
                    }
                }
                finally { controller.Release(); }

                ReportDeepProgress(attempted, candidates.Count);
            });

        activityLog?.Report(
            $"Deep verify complete: {attempted:N0} compared" +
            (resolvedToMatch > 0 ? $", {resolvedToMatch:N0} confirmed matched" : "") +
            (mismatched      > 0 ? $", {mismatched:N0} content mismatch" : "") +
            (inconclusive    > 0 ? $", {inconclusive:N0} could not be compared" : ""));
    }

    private async Task<(OpcCompareOutcome Outcome, string? SkipReason)> DeepVerifyOneAsync(
        DeepVerifyCandidate candidate, long sizeCapBytes, CancellationToken ct)
    {
        if ((candidate.Source.Size ?? 0) > sizeCapBytes || (candidate.Target.Size ?? 0) > sizeCapBytes)
            return (new OpcCompareOutcome(OpcCompareResult.NotComparable, []),
                $"file exceeds the {sizeCapBytes / (1024 * 1024):N0} MB deep-verify size cap");

        try
        {
            using var sourceMs = new MemoryStream();
            using var targetMs = new MemoryStream();
            await Task.WhenAll(
                DownloadIntoAsync(candidate.Source.DriveId, candidate.Source.ItemId, sourceMs, ct),
                DownloadIntoAsync(candidate.Target.DriveId, candidate.Target.ItemId, targetMs, ct));
            sourceMs.Position = 0;
            targetMs.Position = 0;

            var outcome = OpcDeepComparer.Compare(sourceMs, targetMs);
            return (outcome, outcome.Result == OpcCompareResult.NotComparable
                ? "not a valid OOXML package (label-encrypted or corrupt)"
                : null);
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            return (new OpcCompareOutcome(OpcCompareResult.NotComparable, []), $"download failed: {ex.Message}");
        }
    }

    private async Task DownloadIntoAsync(string driveId, string itemId, Stream destination, CancellationToken ct)
    {
        using var stream = await spService.DownloadFileAsync(driveId, itemId);
        await stream.CopyToAsync(destination, ct);
    }

    // Never downgrades the cheap-tier status just because deep verify couldn't reach a conclusion
    // (NotComparable) — that would destroy real information (e.g. a genuine DateMismatch) in favor
    // of a less informative result. Only a definite content verdict changes Status.
    private static void ApplyDeepVerifyOutcome(ComparisonRow row, OpcCompareOutcome outcome, string? skipReason)
    {
        switch (outcome.Result)
        {
            case OpcCompareResult.VolatileOnlyDifferences:
                row.Status = ComparisonStatus.Match;
                row.Note   = "Deep verify: content parts identical — only SharePoint-rewritten metadata (Document ID, custom properties, etc.) differs.";
                break;
            case OpcCompareResult.ContentMismatch:
                row.Status = ComparisonStatus.ContentMismatch;
                row.Note   = outcome.DifferingParts.Count > 0
                    ? $"Deep verify: content differs in {string.Join(", ", outcome.DifferingParts)}"
                    : "Deep verify: content differs.";
                break;
            case OpcCompareResult.NotComparable:
                row.Note = $"Deep verify could not run — {skipReason ?? "unknown reason"}. Status reflects the date/hash check only.";
                break;
        }
    }

    private static ComparisonStatus ClassifyMatch(ScannedFile s, ScannedFile? t)
    {
        if (t == null) return ComparisonStatus.OnlyInSource;

        bool hashesPresent = s.QuickXorHash != null && t.QuickXorHash != null;
        bool hashesEqual   = hashesPresent && string.Equals(s.QuickXorHash, t.QuickXorHash, StringComparison.Ordinal);

        if (OfficeReserializedExtensions.Contains(Path.GetExtension(s.RelativePath)))
        {
            // Equal hashes are always trustworthy — SharePoint didn't rewrite this one (common
            // for Migration API imports, which bypass the upload pipeline). Only a hash
            // DIFFERENCE is meaningless for these formats.
            if (hashesEqual) return ComparisonStatus.Match;

            // Otherwise modified date is the signal, since the app is already responsible for
            // preserving it onto the target. A missing date on either side is Unverified — the
            // old "call it Match" behavior let a 0-byte or corrupt Office file pass green.
            if (s.LastModified is not { } srcDate || t.LastModified is not { } tgtDate)
                return ComparisonStatus.Unverified;
            return (srcDate - tgtDate).Duration() <= DateMismatchTolerance
                ? ComparisonStatus.Match
                : ComparisonStatus.DateMismatch;
        }

        if (!hashesPresent)
        {
            // Graph returns null quickXorHash for a nontrivial share of items in listings.
            // Fall back to size before giving up; never report an unverifiable row as Match.
            if (s.Size is { } srcSize && t.Size is { } tgtSize)
                return srcSize == tgtSize ? ComparisonStatus.Match : ComparisonStatus.ContentMismatch;
            return ComparisonStatus.Unverified;
        }

        return hashesEqual
            ? ComparisonStatus.Match
            : ComparisonStatus.ContentMismatch;
    }
}
