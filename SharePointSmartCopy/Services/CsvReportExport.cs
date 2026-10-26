using System.IO;
using System.Text;

namespace SharePointSmartCopy.Services;

// The one place the copy report and the saved-run report are written as CSV. Written as UTF-8 WITH a byte-order mark:
// without it Excel assumes the ANSI code page and garbles accented, Cyrillic, CJK, Arabic and Hebrew file names.
public static class CsvReportExport
{
    public static string Field(string? value) => $"\"{(value ?? "").Replace("\"", "\"\"")}\"";

    // One row of the file-level report, in the column order of the (localized) header line.
    public static string Row(string? fileName, string? sourcePath, string? targetPath, object? status, int versionsCopied,
        string? error, object? permissionStatus, string? permissionDetails) =>
        $"{Field(fileName)},{Field(sourcePath)},{Field(targetPath)},{status},{versionsCopied},{Field(error)}," +
        $"{permissionStatus?.ToString() ?? ""},{Field(permissionDetails)}";

    public static void Save(string path, string headerLine, IEnumerable<string> rows)
    {
        var sb = new StringBuilder();
        sb.AppendLine(headerLine);
        foreach (var row in rows) sb.AppendLine(row);
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
    }
}
