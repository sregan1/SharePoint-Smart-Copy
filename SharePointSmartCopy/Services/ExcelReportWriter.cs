using ClosedXML.Excel;
using System.IO;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Writes a VerificationReportService.Result to an .xlsx workbook: Overview, Source, Target, and
// Comparison sheets, plus a Scan Errors sheet when any root couldn't be scanned.
public static class ExcelReportWriter
{
    private const string DateFormat = "yyyy-mm-dd hh:mm:ss";
    private static readonly XLColor MismatchFill = XLColor.FromHtml("#FDE7E9");
    private static readonly XLColor MatchFill    = XLColor.FromHtml("#DFF6DD");

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
        int onlyInSource   = result.Comparison.Count(r => r.Status == ComparisonStatus.OnlyInSource);
        int onlyInTarget   = result.Comparison.Count(r => r.Status == ComparisonStatus.OnlyInTarget);
        int matched        = result.Comparison.Count(r => r.Status == ComparisonStatus.Match);
        int contentMismatch = result.Comparison.Count(r => r.Status == ComparisonStatus.ContentMismatch);
        int dateMismatch    = result.Comparison.Count(r => r.Status == ComparisonStatus.DateMismatch);

        ws.Cell(1, 1).Value = "Verification Summary";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;

        ws.Cell(3, 1).Value = "Files in Source";
        ws.Cell(3, 2).Value = result.SourceFiles.Count;
        ws.Cell(4, 1).Value = "Files in Target";
        ws.Cell(4, 2).Value = result.TargetFiles.Count;
        ws.Cell(5, 1).Value = "Matched (in both)";
        ws.Cell(5, 2).Value = matched;
        ws.Cell(6, 1).Value = "Content Mismatch";
        ws.Cell(6, 2).Value = contentMismatch;
        ws.Cell(7, 1).Value = "Date Mismatch";
        ws.Cell(7, 2).Value = dateMismatch;
        ws.Cell(8, 1).Value = "Only in Source";
        ws.Cell(8, 2).Value = onlyInSource;
        ws.Cell(9, 1).Value = "Only in Target";
        ws.Cell(9, 2).Value = onlyInTarget;
        ws.Range(3, 1, 9, 1).Style.Font.Bold = true;

        // Not merged: Excel does not reliably auto-size row height for wrapped text in a merged
        // cell (a well-known Excel limitation, not a ClosedXML bug), which clipped this message.
        // Left in a single unmerged cell, AdjustToContents below widens column 1 to fit it on one
        // line instead — the fill is still applied across the row for the banner look.
        var messageCell = ws.Cell(11, 1);
        XLColor fill;
        if (onlyInSource == 0 && onlyInTarget == 0 && contentMismatch == 0 && dateMismatch == 0 && result.ScanErrors.Count == 0)
        {
            messageCell.Value = "✓ Exact match — every file in source was found in target, with no extras.";
            fill = MatchFill;
        }
        else
        {
            var parts = new List<string> { "⚠ Differences found — see the Comparison tab for details." };
            if (result.ScanErrors.Count > 0)
                parts.Add($"{result.ScanErrors.Count} location(s) could not be scanned — see the Scan Errors tab.");
            messageCell.Value = string.Join(" ", parts);
            fill = MismatchFill;
        }
        ws.Range(11, 1, 11, 4).Style.Fill.BackgroundColor = fill;

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
        FormatHeader(ws, 4);

        int row = 2;
        foreach (var r in rows.OrderBy(r => r.RelativePath, StringComparer.OrdinalIgnoreCase))
        {
            ws.Cell(row, 1).Value = r.RelativePath;
            ws.Cell(row, 2).Value = r.Status.ToString();
            SetComparisonValue(ws.Cell(row, 3), r, r.SourceHash, r.SourceModified);
            SetComparisonValue(ws.Cell(row, 4), r, r.TargetHash, r.TargetModified);

            var fill = r.Status == ComparisonStatus.Match ? MatchFill : MismatchFill;
            ws.Range(row, 1, row, 4).Style.Fill.BackgroundColor = fill;
            row++;
        }

        FinishSheet(ws, row - 1, 4);
    }

    // Shows whichever signal was actually used to compare this file (see ComparisonStatus):
    // modified date for Office/OLE compound-document formats, content hash for everything else.
    // Left blank when that side has no file at all (Only in Source/Target) or the signal wasn't available.
    private static void SetComparisonValue(IXLCell cell, ComparisonRow r, string? hash, DateTimeOffset? modified)
    {
        if (VerificationReportService.OfficeReserializedExtensions.Contains(Path.GetExtension(r.RelativePath)))
            SetModified(cell, modified);
        else
            cell.Value = hash;
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
