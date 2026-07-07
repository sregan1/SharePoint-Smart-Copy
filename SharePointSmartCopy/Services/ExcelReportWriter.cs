using ClosedXML.Excel;
using System.IO;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Writes a VerificationReportService.Result to an .xlsx workbook: Overview, Source, Target, and
// Comparison sheets, plus a Scan Errors sheet when any root couldn't be scanned.
public static class ExcelReportWriter
{
    private const string DateFormat = "yyyy-mm-dd hh:mm:ss";
    private static readonly XLColor MismatchFill   = XLColor.FromHtml("#FDE7E9");
    private static readonly XLColor MatchFill      = XLColor.FromHtml("#DFF6DD");
    private static readonly XLColor UnverifiedFill = XLColor.FromHtml("#FFF4CE");

    public static void Write(string path, VerificationReportService.Result result)
    {
        using var wb = new XLWorkbook();
        WriteOverviewSheet(wb.Worksheets.Add("Overview"), result);
        WriteFileSheet(wb.Worksheets.Add("Source"), result.SourceFiles);
        WriteFileSheet(wb.Worksheets.Add("Target"), result.TargetFiles);
        WriteComparisonSheet(wb.Worksheets.Add("Comparison"), result.Comparison);
        if (result.ScanErrors.Count > 0)
            WriteScanErrorsSheet(wb.Worksheets.Add("Scan Errors"), result.ScanErrors);
        wb.SaveAs(path);
    }

    private static void WriteOverviewSheet(IXLWorksheet ws, VerificationReportService.Result result)
    {
        int onlyInSource    = result.Comparison.Count(r => r.Status == ComparisonStatus.OnlyInSource);
        int onlyInTarget    = result.Comparison.Count(r => r.Status == ComparisonStatus.OnlyInTarget);
        int matched         = result.Comparison.Count(r => r.Status == ComparisonStatus.Match);
        int contentMismatch = result.Comparison.Count(r => r.Status == ComparisonStatus.ContentMismatch);
        int dateMismatch    = result.Comparison.Count(r => r.Status == ComparisonStatus.DateMismatch);
        int unverified      = result.Comparison.Count(r => r.Status == ComparisonStatus.Unverified);

        ws.Cell(1, 1).Value = "Verification Summary";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;

        // Headline banner comes FIRST, ahead of the breakdown table — a user should be able to open
        // this sheet and know in one glance whether everything matched, not read eight numbers and
        // do the arithmetic themselves. Not merged (Excel doesn't reliably auto-size row height for
        // wrapped text in a merged cell, a well-known Excel limitation); the fill spans the row for
        // the banner look, AdjustToContents below widens column 1 to fit each line.
        bool noDifferences = onlyInSource == 0 && onlyInTarget == 0 && contentMismatch == 0
                          && dateMismatch == 0 && result.ScanErrors.Count == 0;
        var headline = ws.Cell(3, 1);
        var detail   = ws.Cell(4, 1);
        headline.Style.Font.Bold = true;
        headline.Style.Font.FontSize = 16;
        XLColor fill;
        if (noDifferences && unverified == 0)
        {
            headline.Value = "✓ ALL FILES MATCH";
            detail.Value   = $"{matched:N0} file(s) matched — every file in source was found in target, with no extras.";
            fill = MatchFill;
        }
        else if (noDifferences)
        {
            headline.Value = "✓ NO DIFFERENCES FOUND";
            detail.Value   = $"{matched:N0} file(s) matched, but {unverified:N0} file(s) had no comparable signal (hash and size unavailable) and could not be verified — see the Comparison tab.";
            fill = UnverifiedFill;
        }
        else
        {
            headline.Value = "⚠ CONTENT DOES NOT MATCH";
            var parts = new List<string> { $"{matched:N0} file(s) matched." };
            if (contentMismatch > 0) parts.Add($"{contentMismatch:N0} content mismatch.");
            if (dateMismatch    > 0) parts.Add($"{dateMismatch:N0} date mismatch.");
            if (onlyInSource    > 0) parts.Add($"{onlyInSource:N0} only in source.");
            if (onlyInTarget    > 0) parts.Add($"{onlyInTarget:N0} only in target.");
            if (unverified      > 0) parts.Add($"{unverified:N0} could not be verified (no comparable signal).");
            if (result.ScanErrors.Count > 0) parts.Add($"{result.ScanErrors.Count} location(s) could not be scanned — see the Scan Errors tab.");
            parts.Add("See the Comparison tab for details.");
            detail.Value = string.Join(" ", parts);
            fill = MismatchFill;
        }
        ws.Range(3, 1, 3, 5).Style.Fill.BackgroundColor = fill;
        ws.Range(4, 1, 4, 5).Style.Fill.BackgroundColor = fill;

        ws.Cell(6, 1).Value = "Files in Source";
        ws.Cell(6, 2).Value = result.SourceFiles.Count;
        ws.Cell(7, 1).Value = "Files in Target";
        ws.Cell(7, 2).Value = result.TargetFiles.Count;
        ws.Cell(8, 1).Value = "Matched (in both)";
        ws.Cell(8, 2).Value = matched;
        ws.Cell(9, 1).Value = "Content Mismatch";
        ws.Cell(9, 2).Value = contentMismatch;
        ws.Cell(10, 1).Value = "Date Mismatch";
        ws.Cell(10, 2).Value = dateMismatch;
        ws.Cell(11, 1).Value = "Only in Source";
        ws.Cell(11, 2).Value = onlyInSource;
        ws.Cell(12, 1).Value = "Only in Target";
        ws.Cell(12, 2).Value = onlyInTarget;
        ws.Cell(13, 1).Value = "Unverified (no comparable signal)";
        ws.Cell(13, 2).Value = unverified;
        ws.Range(6, 1, 13, 1).Style.Font.Bold = true;

        ws.Columns(1, 2).AdjustToContents();
    }

    private static void WriteFileSheet(IXLWorksheet ws, List<ScannedFile> files)
    {
        ws.Cell(1, 1).Value = "Relative Path";
        ws.Cell(1, 2).Value = "Name";
        ws.Cell(1, 3).Value = "Last Modified (UTC)";
        FormatHeader(ws, 3);

        int row = 2;
        foreach (var f in files.OrderBy(f => f.RelativePath, StringComparer.OrdinalIgnoreCase))
        {
            ws.Cell(row, 1).Value = f.RelativePath;
            ws.Cell(row, 2).Value = f.Name;
            SetModified(ws.Cell(row, 3), f.LastModified);
            row++;
        }

        FinishSheet(ws, row - 1, 3);
    }

    private static void WriteComparisonSheet(IXLWorksheet ws, List<ComparisonRow> rows)
    {
        ws.Cell(1, 1).Value = "Relative Path";
        ws.Cell(1, 2).Value = "Status";
        ws.Cell(1, 3).Value = "Source Value";
        ws.Cell(1, 4).Value = "Target Value";
        ws.Cell(1, 5).Value = "Note";
        FormatHeader(ws, 5);

        int row = 2;
        foreach (var r in rows.OrderBy(r => r.RelativePath, StringComparer.OrdinalIgnoreCase))
        {
            ws.Cell(row, 1).Value = r.RelativePath;
            ws.Cell(row, 2).Value = r.Status.ToString();
            // Decided ONCE per row, not independently per cell: ClassifyMatch only trusts a hash
            // comparison when BOTH sides have one — if either is missing, the verdict was decided
            // by size on both sides. Picking hash-if-available per cell independently could show a
            // hash for the side that has one right next to a size for the side that doesn't, which
            // looks like an apples-to-oranges comparison even though the actual decision was
            // size-vs-size throughout.
            bool bothHashesPresent = r.SourceHash != null && r.TargetHash != null;
            SetComparisonValue(ws.Cell(row, 3), r, r.SourceHash, r.SourceModified, r.SourceSize, bothHashesPresent);
            SetComparisonValue(ws.Cell(row, 4), r, r.TargetHash, r.TargetModified, r.TargetSize, bothHashesPresent);
            ws.Cell(row, 5).Value = r.Note ?? "";

            var fill = r.Status switch
            {
                ComparisonStatus.Match       => MatchFill,
                ComparisonStatus.Unverified  => UnverifiedFill,
                _                            => MismatchFill,
            };
            ws.Range(row, 1, row, 5).Style.Fill.BackgroundColor = fill;
            row++;
        }

        FinishSheet(ws, row - 1, 5);
    }

    // Shows whichever signal was actually used to compare this file (see ComparisonStatus):
    // modified date for Office/OLE compound-document formats, content hash when BOTH sides have
    // one, otherwise file size for BOTH sides — matching ClassifyMatch's own logic, which only
    // trusts a hash comparison when both sides have a hash and falls back to size for the pair
    // otherwise (Graph returns a null quickXorHash for a nontrivial share of items in listings).
    // Left genuinely blank only when that side has no file at all (Only in Source/Target) or no
    // signal of any kind was available.
    private static void SetComparisonValue(IXLCell cell, ComparisonRow r, string? hash, DateTimeOffset? modified, long? size, bool bothHashesPresent)
    {
        if (VerificationReportService.OfficeReserializedExtensions.Contains(Path.GetExtension(r.RelativePath)))
            SetModified(cell, modified);
        else if (bothHashesPresent)
            cell.Value = hash;
        else if (size.HasValue)
            cell.Value = $"{size.Value:N0} bytes (by size — hash unavailable on at least one side)";
    }

    private static void WriteScanErrorsSheet(IXLWorksheet ws, List<string> errors)
    {
        ws.Cell(1, 1).Value = "Root that could not be scanned";
        FormatHeader(ws, 1);

        int row = 2;
        foreach (var e in errors)
        {
            ws.Cell(row, 1).Value = e;
            row++;
        }

        FinishSheet(ws, row - 1, 1);
    }

    private static void FormatHeader(IXLWorksheet ws, int lastCol)
    {
        var header = ws.Range(1, 1, 1, lastCol);
        header.Style.Font.Bold = true;
        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#F3F2F1");
    }

    // Column auto-fit measures every cell in the sampled range — on a 100k+ row sheet that's a
    // full scan just to size columns. A few thousand rows is a representative sample for
    // consistent-format data (paths/dates), so cap it rather than scanning the whole sheet.
    private const int AutoFitSampleRows = 2000;

    private static void FinishSheet(IXLWorksheet ws, int lastRow, int lastCol)
    {
        ws.SheetView.FreezeRows(1);
        if (lastRow >= 1)
            ws.Range(1, 1, lastRow, lastCol).SetAutoFilter();
        int sampleEnd = Math.Min(lastRow, AutoFitSampleRows);
        ws.Columns(1, lastCol).AdjustToContents(1, sampleEnd);
    }

    private static void SetModified(IXLCell cell, DateTimeOffset? modified)
    {
        if (!modified.HasValue) return;
        cell.Value = modified.Value.UtcDateTime;
        cell.Style.DateFormat.Format = DateFormat;
    }
}
