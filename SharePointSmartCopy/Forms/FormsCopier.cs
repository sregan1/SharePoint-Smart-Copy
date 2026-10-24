using System.Net.Http;
using System.Text.Json.Nodes;

namespace SharePointSmartCopy.Forms;

// The outcome of copying one form, for the results dashboard.
//   Warnings = something that did not copy or may be wrong (the copy is "partial").
//   Notes    = known limits worth checking (images are not copied yet, quiz scoring is unverified).
public sealed record FormCopyResult(
    bool Success, string? NewId, int QuestionsTotal, int QuestionsCopied,
    IReadOnlyList<string> Warnings, IReadOnlyList<string> Notes, string? Error)
{
    public bool IsPartial => Success && Warnings.Count > 0;
}

// Copies one form: create shell, patch form-level properties one at a time (so unsupported ones are
// reported instead of failing the whole copy), add questions with their ORIGINAL ids, activate, then verify
// the copy has as many questions as the source.
public sealed class FormsCopier
{
    // Form-level properties we try to carry over. Each is patched individually.
    private static readonly string[] FormProps =
    [
        "description", "settings", "themeV2", "thankYouMessage", "defaultLanguage", "emailReceiptEnabled",
        "progressBarEnabled", "fillOutTimeLimit", "header", "background", "logo", "mfpBranchingData",
        "responseThresholdCount", "inviteExpiryDays", "PrivacyUrl", "otherInfo",
    ];

    // Question properties we try to carry over (read-only/server-owned ones such as modifiedDate are skipped).
    private static readonly string[] QuestionProps =
    [
        "id", "title", "subtitle", "order", "type", "required", "questionInfo", "isQuiz", "justification",
        "groupId", "defaultValue", "allowMultipleValues", "allowCustomChoice", "image",
    ];

    private readonly FormsApiClient _source;
    private readonly FormsApiClient _target;
    private readonly Action<string> _log;

    public FormsCopier(FormsApiClient source, FormsApiClient target, Action<string> log)
    {
        _source = source;
        _target = target;
        _log = log;
    }

    public async Task<FormCopyResult> CopyAsync(string sourceFormId, string newTitle, CancellationToken ct = default)
    {
        var warnings = new List<string>();
        var notes = new List<string>();

        var (st, body) = await _source.SendRawAsync(HttpMethod.Get, $"formapi/api/forms('{sourceFormId}')?$expand=questions", null, ct);
        if (st != 200)
        {
            _log($"FAIL read source: {st} {Trim(body)}");
            return Failed($"Could not read the source form ({st}).");
        }
        var src = JsonNode.Parse(body)!.AsObject();
        var questions = src["questions"]?.AsArray().OrderBy(q => q!["order"]!.GetValue<double>()).ToList() ?? [];

        // The form type ("form" or "quiz") can only be set at creation; PATCHing type/flags later is rejected.
        var create = new JsonObject { ["title"] = newTitle, ["description"] = src["description"]?.DeepClone() };
        if (src["type"] is { } type) create["type"] = type.DeepClone();
        (st, body) = await _target.SendRawAsync(HttpMethod.Post, "formapi/api/forms", create.ToJsonString(), ct);
        if (st is not (200 or 201))
        {
            _log($"FAIL create: {st} {Trim(body)}");
            return Failed($"Could not create the form in the target ({st}). {ShortReason(body)}", questions.Count);
        }
        var newId = JsonNode.Parse(body)!["id"]!.GetValue<string>();
        _log($"created {newId}");

        var copied = 0;
        try
        {
            await CopyContentAsync(src, questions, newId, warnings, ct, n => copied = n);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            // The shell exists in the target: report it (with its ID) so it isn't lost, rather than only the error.
            _log($"FAIL after create: {ex.Message}");
            return new FormCopyResult(false, newId, questions.Count, copied, warnings, notes, $"The copy stopped part-way ({ex.Message}). A partial form was left in the target; delete it in Microsoft Forms.");
        }

        // Known limits, worth a look in the copy
        foreach (var (prop, label) in new[] { ("background", "background image"), ("logo", "logo"), ("header", "header image") })
            if (src[prop] is JsonObject img && img["resourceId"] is not null)
                notes.Add($"The {label} was not copied (images are not supported yet).");
        if (src["type"]?.GetValue<string>() == "quiz")
            notes.Add("Quiz scoring (correct answers and points) is not verified; please check the copy.");

        return new FormCopyResult(true, newId, questions.Count, copied, warnings, notes, null);
    }

    private async Task CopyContentAsync(JsonObject src, List<JsonNode?> questions, string newId, List<string> warnings,
        CancellationToken ct, Action<int> setCopied)
    {
        int st;
        string body;

        // One PATCH with every form-level property first (one round trip); only if the service rejects the batch
        // fall back to one property at a time, so an unsupported property is reported instead of failing the copy.
        var all = new JsonObject();
        foreach (var prop in FormProps)
            if (src[prop] is { } pv) all[prop] = pv.DeepClone();
        var patched = false;
        if (all.Count > 0)
        {
            (st, body) = await _target.SendRawAsync(HttpMethod.Patch, $"formapi/api/forms('{newId}')", all.ToJsonString(), ct);
            _log($"  form properties (batch): {st}{(st >= 400 ? " " + Trim(body) : "")}");
            patched = st < 400;
        }
        if (!patched)
            foreach (var prop in FormProps)
            {
                if (src[prop] is null) { _log($"  skip form.{prop} (null in source)"); continue; }
                var patch = new JsonObject { [prop] = src[prop]!.DeepClone() };
                (st, body) = await _target.SendRawAsync(HttpMethod.Patch, $"formapi/api/forms('{newId}')", patch.ToJsonString(), ct);
                _log($"  form.{prop}: {st}{(st >= 400 ? " " + Trim(body) : "")}");
                if (st >= 400) warnings.Add($"Setting '{prop}' could not be copied ({st}).");
            }

        var copied = 0;
        foreach (var q in questions)
        {
            var qb = new JsonObject();
            foreach (var p in QuestionProps)
                if (q![p] is { } v) qb[p] = v.DeepClone();
            (st, body) = await _target.SendRawAsync(HttpMethod.Post, $"formapi/api/forms('{newId}')/questions", qb.ToJsonString(), ct);
            _log($"  question {q!["id"]} ({q["type"]}): {st}{(st >= 400 ? " " + Trim(body) : "")}");
            if (st is 200 or 201) { copied++; setCopied(copied); }
            else warnings.Add($"Question '{Clip(q["title"]?.GetValue<string>() ?? "(untitled)", 60)}' was not copied ({st}). {ShortReason(body)}".TrimEnd());
        }

        (st, body) = await _target.SendRawAsync(HttpMethod.Patch, $"formapi/api/forms('{newId}')", "{\"softDeleted\":0}", ct);
        _log($"  activate: {st}");
        if (st >= 400) warnings.Add($"The copy was created but could not be made visible in Microsoft Forms ({st}); it may be in the recycle bin.");

        // Verify: the copy should hold as many questions as the source.
        (st, body) = await _target.SendRawAsync(HttpMethod.Get, $"formapi/api/forms('{newId}')/questions", null, ct);
        if (st == 200 && JsonNode.Parse(body)?["value"] is JsonArray found && found.Count != questions.Count)
            warnings.Add($"The copy has {found.Count} questions but the source has {questions.Count}.");
    }

    private static FormCopyResult Failed(string error, int questions = 0) =>
        new(false, null, questions, 0, [], [], error);

    private static string Trim(string s) => s.Length > 300 ? s[..300] : s;
    private static string Clip(string s, int n) => s.Length > n ? s[..n] + "…" : s;

    // The service's error JSON carries a short "message"; surface it when present.
    private static string ShortReason(string body)
    {
        try
        {
            var msg = JsonNode.Parse(body)?["error"]?["message"]?.GetValue<string>();
            return msg is null ? "" : Clip(msg, 140);
        }
        catch { return ""; }
    }
}
