using System.Net.Http;
using System.Text.Json.Nodes;
using SharePointSmartCopy.Localization;

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
        "progressBarEnabled", "fillOutTimeLimit", "header", "background", "logo",
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
            return Failed(Loc.T("Forms_Cp_ReadFail", st));
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
            return Failed(Loc.T("Forms_Cp_CreateFail", st, ShortReason(body)), questions.Count);
        }
        var newId = JsonNode.Parse(body)!["id"]!.GetValue<string>();
        _log($"created {newId}");

        var copied = 0;
        try
        {
            await CopyContentAsync(src, questions, newId, warnings, ct, n => copied = n);
        }
        catch (OperationCanceledException)
        {
            // Return (instead of throwing) so the partial form's ID reaches the results table; the caller's next loop pass
            // sees the cancellation and stops.
            _log($"cancelled after create: {newId}");
            return new FormCopyResult(false, newId, questions.Count, copied, warnings, notes, Loc.T("Forms_Cp_Cancelled"));
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            // The shell exists in the target: report it (with its ID) so it isn't lost, rather than only the error.
            _log($"FAIL after create: {ex.Message}");
            return new FormCopyResult(false, newId, questions.Count, copied, warnings, notes, Loc.T("Forms_Cp_Stopped", ex.Message));
        }

        // Known limits, worth a look in the copy
        foreach (var (prop, label) in new[] { ("background", Loc.T("Forms_Cp_LabelBackground")), ("logo", Loc.T("Forms_Cp_LabelLogo")), ("header", Loc.T("Forms_Cp_LabelHeader")) })
            if (src[prop] is JsonObject img && img["resourceId"] is not null)
                notes.Add(Loc.T("Forms_Cp_ImageNote", label));
        if (src["type"]?.GetValue<string>() == "quiz")
            notes.Add(Loc.T("Forms_Cp_QuizNote"));

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
                if (st >= 400) warnings.Add(Loc.T("Forms_Cp_SettingFail", prop, st));
            }

        var copied = 0;
        foreach (var q in questions)
        {
            var qb = new JsonObject();
            foreach (var p in QuestionProps)
                if (q![p] is { } v) qb[p] = v.DeepClone();
            (st, body) = await _target.SendRawAsync(HttpMethod.Post, $"formapi/api/forms('{newId}')/questions", qb.ToJsonString(), ct);
            _log($"  question {q!["id"]} ({q["type"]}): {st}{(st >= 400 ? " " + Trim(body) : "")}");
            if (st is 200 or 201)
            {
                copied++; setCopied(copied);
                await RepairQuestionInfoAsync(q, newId, body, warnings, ct);
            }
            else warnings.Add(Loc.T("Forms_Cp_QuestionFail", Clip(q["title"]?.GetValue<string>() ?? Loc.T("Forms_Cp_Untitled"), 60), st, ShortReason(body)).TrimEnd());
        }

        // Branching rules point at question IDs, so they are set once the questions exist (they used to be part of the
        // first property batch, before any question was created).
        if (src["mfpBranchingData"] is { } branching)
        {
            var bp = new JsonObject { ["mfpBranchingData"] = branching.DeepClone() };
            (st, body) = await _target.SendRawAsync(HttpMethod.Patch, $"formapi/api/forms('{newId}')", bp.ToJsonString(), ct);
            _log($"  branching rules: {st}{(st >= 400 ? " " + Trim(body) : "")}");
            if (st >= 400) warnings.Add(Loc.T("Forms_Cp_BranchFail", st));
        }

        (st, body) = await _target.SendRawAsync(HttpMethod.Patch, $"formapi/api/forms('{newId}')", "{\"softDeleted\":0}", ct);
        _log($"  activate: {st}");
        if (st >= 400) warnings.Add(Loc.T("Forms_Cp_NotVisible", st));

        // Verify: the copy should hold as many questions as the source.
        (st, body) = await _target.SendRawAsync(HttpMethod.Get, $"formapi/api/forms('{newId}')/questions", null, ct);
        if (st == 200 && JsonNode.Parse(body)?["value"] is JsonArray found && found.Count != questions.Count)
            warnings.Add(Loc.T("Forms_Cp_CountMismatch", found.Count, questions.Count));
    }

    // Ranking and Likert questions come back from the service with default content (one "Option 1", no rating scale)
    // when created with their questionInfo, so the questionInfo of the created question is compared with the source's
    // and, if it differs, PATCHed onto the question. Both shapes are logged so the real payload can be inspected.
    private async Task RepairQuestionInfoAsync(JsonNode? q, string newId, string createResponse, List<string> warnings, CancellationToken ct)
    {
        var type = q?["type"]?.GetValue<string>() ?? "";
        if (!type.Contains("rank", StringComparison.OrdinalIgnoreCase) && !type.Contains("likert", StringComparison.OrdinalIgnoreCase)) return;
        var info = q!["questionInfo"];
        var qid = q["id"]?.GetValue<string>();
        if (info is null || qid is null) return;

        _log($"    {type} source questionInfo: {Trim(info.ToJsonString())}");
        string? createdInfo = null;
        try { createdInfo = JsonNode.Parse(createResponse)?["questionInfo"]?.ToJsonString(); } catch { }
        _log($"    {type} created questionInfo: {Trim(createdInfo ?? "(none)")}");
        if (createdInfo == info.ToJsonString()) return;

        var patch = new JsonObject { ["questionInfo"] = info.DeepClone() };
        var (st, body) = await _target.SendRawAsync(HttpMethod.Patch, $"formapi/api/forms('{newId}')/questions('{qid}')", patch.ToJsonString(), ct);
        _log($"    {type} repair questionInfo: {st}{(st >= 400 ? " " + Trim(body) : "")}");
        if (st >= 400)
            warnings.Add(Loc.T("Forms_Cp_OptionsFail", Clip(q["title"]?.GetValue<string>() ?? Loc.T("Forms_Cp_Untitled"), 60), type, st));
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
