using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Forms;

// Exports a form's responses to an .xlsx (one row per response, one column per question), either one workbook per
// form or one workbook with a sheet per form.
// Responses can't be re-submitted into the copy with their original timestamps/respondents, so this export is
// the migration path for response data.
//
// UNVERIFIED against live data: the account used for the spike has no responses. The shape assumed for
// ResponseEntity.answers is a JSON string holding an array of objects with "questionId" and one or more
// answer values (answer1, answer2...). Anything unexpected is written to the cell as raw JSON, never dropped.
// The quiz score is read from whichever top-level response property has "score" in its name, if any.
public sealed class FormsResponseExporter
{
    private const string DateFormat = "yyyy-mm-dd hh:mm:ss";
    private readonly FormsApiClient _client;

    public FormsResponseExporter(FormsApiClient client) => _client = client;

    public async Task<int> ExportAsync(string formId, string xlsxPath, CancellationToken ct = default)
    {
        var (form, responses) = await FetchAsync(formId, ct);
        Write(form, responses, xlsxPath);
        return responses.Count;
    }

    public async Task<(JsonObject Form, List<JsonObject> Responses)> FetchAsync(string formId, CancellationToken ct = default)
    {
        var (st, body) = await _client.SendRawAsync(HttpMethod.Get, $"formapi/api/forms('{formId}')?$expand=questions", null, ct);
        if (st != 200) throw new HttpRequestException($"read form failed: {st} {body}", null, (System.Net.HttpStatusCode)st);
        var form = JsonNode.Parse(body)!.AsObject();
        var responses = await _client.GetAllPagesAsync($"formapi/api/forms('{formId}')/responses", "read responses", ct);
        return (form, responses);
    }

    // Pure: builds the workbook from already-fetched JSON (no HTTP).
    public static void Write(JsonObject form, IReadOnlyList<JsonObject> responses, string xlsxPath)
    {
        using var wb = new XLWorkbook();
        AddSheet(wb, "Responses", form, responses);
        wb.SaveAs(xlsxPath);
    }

    // One workbook, one sheet per form. Sheet names are the form titles made valid for Excel (31 characters, no
    // []:*?/\) and unique.
    public static void WriteCombined(IReadOnlyList<(string Title, JsonObject Form, IReadOnlyList<JsonObject> Responses)> forms, string xlsxPath)
    {
        using var wb = new XLWorkbook();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (title, form, responses) in forms)
        {
            var baseName = new string(title.Select(c => "[]:*?/\\".Contains(c) ? '_' : c).ToArray()).Trim();
            if (baseName.Length == 0) baseName = "Form";
            var name = Clip(baseName, 31);
            for (var n = 2; !used.Add(name); n++)
            {
                var suffix = $" ({n})";
                name = Clip(baseName, 31 - suffix.Length) + suffix;
            }
            AddSheet(wb, name, form, responses);
        }
        wb.SaveAs(xlsxPath);
    }

    private static string Clip(string s, int max) => s.Length > max ? s[..max] : s;

    private static void AddSheet(XLWorkbook wb, string sheetName, JsonObject form, IReadOnlyList<JsonObject> responses)
    {
        // Matrix rows are separate questions (MatrixChoice); the group itself carries no answer.
        var questions = form["questions"]!.AsArray().Select(q => q!.AsObject())
            .Where(q => q["type"]?.GetValue<string>() != "Question.MatrixChoiceGroup" && q["id"] is not null)
            .GroupBy(q => q["id"]!.GetValue<string>()).Select(g => g.First())   // a repeated question id must not abort the export
            .OrderBy(q => q["order"]?.GetValue<double>() ?? double.MaxValue).ToList();
        var index = questions.Select((q, i) => (id: q["id"]!.GetValue<string>(), col: i + 6)).ToDictionary(x => x.id, x => x.col);

        // Quiz forms: a score column when the responses carry one.
        var scoreKey = responses.SelectMany(r => r.Select(kv => kv.Key))
            .FirstOrDefault(k => k.Contains("score", StringComparison.OrdinalIgnoreCase));
        var scoreCol = 6 + questions.Count;
        var extra = scoreCol + (scoreKey is null ? 0 : 1);

        var ws = wb.AddWorksheet(sheetName);
        string[] fixedHeaders = [Loc.T("Forms_Exp_RId"), Loc.T("Forms_Exp_RStart"), Loc.T("Forms_Exp_REnd"), Loc.T("Forms_Exp_RResponder"), Loc.T("Forms_Exp_RName")];
        for (var i = 0; i < fixedHeaders.Length; i++) ws.Cell(1, i + 1).Value = fixedHeaders[i];
        foreach (var q in questions) ws.Cell(1, index[q["id"]!.GetValue<string>()]).Value = q["title"]?.GetValue<string>() ?? "";
        if (scoreKey is not null) ws.Cell(1, scoreCol).Value = Loc.T("Forms_Exp_RScore");
        ws.Cell(1, extra).Value = Loc.T("Forms_Exp_RUnmapped");

        var row = 2;
        foreach (var r in responses)
        {
            ws.Cell(row, 1).Value = r["id"]?.ToString();
            SetDate(ws.Cell(row, 2), r["startDate"]);
            SetDate(ws.Cell(row, 3), r["submitDate"]);
            ws.Cell(row, 4).Value = r["responder"]?.ToString();
            ws.Cell(row, 5).Value = r["responderName"]?.ToString();
            if (scoreKey is not null && r[scoreKey] is JsonValue sv)
            {
                if (sv.TryGetValue<double>(out var score)) ws.Cell(row, scoreCol).Value = score;
                else ws.Cell(row, scoreCol).Value = sv.ToString();
            }

            var unmapped = new List<string>();
            foreach (var a in ParseAnswers(r["answers"]))
            {
                var qid = a["questionId"]?.GetValue<string>();
                var text = string.Join("; ", a.Where(kv => kv.Key != "questionId" && kv.Value is not null)
                    .Select(kv => Render(kv.Value!)).Where(t => t.Length > 0));
                if (qid is not null && index.TryGetValue(qid, out var col)) ws.Cell(row, col).Value = Clip(text);
                else if (text.Length > 0 || a.Any(kv => kv.Key != "questionId" && kv.Value is not null && Render(kv.Value).Length > 0))
                    unmapped.Add(a.ToJsonString());   // empty answers for non-question rows (e.g. matrix groups) are noise
            }
            if (unmapped.Count > 0) ws.Cell(row, extra).Value = Clip(string.Join("\n", unmapped));
            row++;
        }
        ws.SheetView.FreezeRows(1);
    }

    // Excel rejects a cell over 32,767 characters, and one over-long answer used to fail the whole export.
    private const int MaxCellChars = 32_000;
    private static string Clip(string text) =>
        text.Length <= MaxCellChars ? text : text[..MaxCellChars] + Loc.T("Forms_Exp_Cut");

    // A real Excel date (local time) when the value parses, so the column sorts and filters; otherwise the text as is.
    private static void SetDate(IXLCell cell, JsonNode? node)
    {
        var text = node?.ToString();
        if (string.IsNullOrWhiteSpace(text)) return;
        if (DateTimeOffset.TryParse(text, out var when))
        {
            cell.Value = when.LocalDateTime;
            cell.Style.DateFormat.Format = DateFormat;
        }
        else cell.Value = text;
    }

    // Verified shapes: answer1 is a string, an array (ranking, multi-select), or a choice object
    // ({"id":2,"displayText":"Option 2",...}). Choice objects render as their text.
    private static string Render(JsonNode node) => node switch
    {
        JsonValue v when v.TryGetValue<string>(out var s) => s,
        JsonValue v => v.ToJsonString(),
        JsonArray arr => string.Join("; ", arr.Where(i => i is not null).Select(i => Render(i!)).Where(t => t.Length > 0)),
        JsonObject o when o["displayText"] is { } d => Render(d),
        JsonObject o when o["description"] is { } d => Render(d),
        _ => node.ToJsonString(),
    };

    private static IEnumerable<JsonObject> ParseAnswers(JsonNode? node)
    {
        if (node is null) yield break;
        JsonNode? parsed = node;
        if (node is JsonValue v && v.TryGetValue<string>(out var s))
        {
            var failed = false;
            try { parsed = JsonNode.Parse(s); }
            catch (JsonException) { failed = true; }
            if (failed) { yield return new JsonObject { ["raw"] = s }; yield break; }
        }
        if (parsed is JsonArray arr)
            foreach (var item in arr) { if (item is JsonObject o) yield return o; else if (item is not null) yield return new JsonObject { ["raw"] = item.DeepClone() }; }
        else if (parsed is JsonObject single) yield return single;
    }
}
