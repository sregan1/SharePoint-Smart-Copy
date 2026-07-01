using ClosedXML.Excel;
using SharePointSmartCopy.Models;

namespace SharePointSmartCopy.Services;

// Writes a VerificationReportService.Result to an .xlsx workbook: Source / Target / Comparison
// sheets, plus a Scan Errors sheet when any root couldn't be scanned.
public static class ExcelReportWriter
{
    private const string DateFormat = "yyyy-mm-dd hh:mm:ss";
    private static readonly XLColor MismatchFill = XLColor.FromHtml("#FDE7E9");
    private static readonly XLColor MatchFill    = XLColor.FromHtml("#DFF6DD");

    public static void Write(string path, VerificationReportService.Result result)
    {
        using var wb = new XLWorkbook();
        WriteFileSheet(wb.Worksheets.Add("Source"), result.SourceFiles);
        WriteFileSheet(wb.Worksheets.Add("Target"), result.TargetFiles);
        WriteComparisonSheet(wb.Worksheets.Add("Comparison"), result.Comparison);
        if (result.ScanErrors.Count > 0)
            WriteScanErrorsSheet(wb.Worksheets.Add("Scan Errors"), result.ScanErrors);
        wb.SaveAs(path);
    }

    private static void WriteFileSheet(IXLWorksheet ws, List<ScannedFile> files)
    {
        ws.Cell(1, 1).Value = "Relative Path";
        ws.Cell(1, 2).Value = "Name";
        ws.Cell(1, 3).Value = "Size (bytes)";
        ws.Cell(1, 4).Value = "Last Modified (UTC)";
        FormatHeader(ws, 4);

        int row = 2;
        foreach (var f in files.OrderBy(f => f.RelativePath, StringComparer.OrdinalIgnoreCase))
        {
            ws.Cell(row, 1).Value = f.RelativePath;
            ws.Cell(row, 2).Value = f.Name;
            SetSize(ws.Cell(row, 3), f.Size);
            SetModified(ws.Cell(row, 4), f.LastModified);
            row++;
        }

        FinishSheet(ws, row - 1, 4);
    }

    private static void WriteComparisonSheet(IXLWorksheet ws, List<ComparisonRow> rows)
    {
        ws.Cell(1, 1).Value = "Relative Path";
        ws.Cell(1, 2).Value = "Source Size";
        ws.Cell(1, 3).Value = "Target Size";
        ws.Cell(1, 4).Value = "Source Modified (UTC)";
        ws.Cell(1, 5).Value = "Target Modified (UTC)";
        ws.Cell(1, 6).Value = "Status";
        FormatHeader(ws, 6);

        int row = 2;
        foreach (var r in rows.OrderBy(r => r.RelativePath, StringComparer.OrdinalIgnoreCase))
        {
            ws.Cell(row, 1).Value = r.RelativePath;
            SetSize(ws.Cell(row, 2), r.SourceSize);
            SetSize(ws.Cell(row, 3), r.TargetSize);
            SetModified(ws.Cell(row, 4), r.SourceModified);
            SetModified(ws.Cell(row, 5), r.TargetModified);
            ws.Cell(row, 6).Value = r.Status.ToString();

            var fill = r.Status == ComparisonStatus.Match ? MatchFill : MismatchFill;
            ws.Range(row, 1, row, 6).Style.Fill.BackgroundColor = fill;
            row++;
        }

        FinishSheet(ws, row - 1, 6);
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

    private static void FinishSheet(IXLWorksheet ws, int lastRow, int lastCol)
    {
        ws.SheetView.FreezeRows(1);
        if (lastRow >= 1)
            ws.Range(1, 1, lastRow, lastCol).SetAutoFilter();
        ws.Columns(1, lastCol).AdjustToContents();
    }

    private static void SetSize(IXLCell cell, long? size)
    {
        if (size.HasValue) cell.Value = size.Value;
    }

    private static void SetModified(IXLCell cell, DateTimeOffset? modified)
    {
        if (!modified.HasValue) return;
        cell.Value = modified.Value.UtcDateTime;
        cell.Style.DateFormat.Format = DateFormat;
    }
}
