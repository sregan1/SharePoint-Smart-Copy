using System.IO;
using System.Net.Http;
using System.Text.Json;
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
            await CopyContentAsync(src, sourceFormId, questions, newId, warnings, ct, n => copied = n);
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
        // Only the images that really did not arrive: the source having an image says nothing about the copy (the background is
        // carried over by the property patch, and the old note claimed otherwise).
        var missingImages = await MissingImagesAsync(src, newId, ct);
        foreach (var (prop, label) in new[] { ("background", Loc.T("Forms_Cp_LabelBackground")), ("logo", Loc.T("Forms_Cp_LabelLogo")), ("header", Loc.T("Forms_Cp_LabelHeader")) })
            if (missingImages.TryGetValue(prop, out var savedAt))
                notes.Add(savedAt is null ? Loc.T("Forms_Cp_ImageNote", label) : Loc.T("Forms_Cp_ImageSaved", label, savedAt));
        if (src["type"]?.GetValue<string>() == "quiz")
            notes.Add(Loc.T("Forms_Cp_QuizNote"));

        return new FormCopyResult(true, newId, questions.Count, copied, warnings, notes, null);
    }

    private async Task CopyContentAsync(JsonObject src, string sourceFormId, List<JsonNode?> questions, string newId, List<string> warnings,
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

        // The form-level values (theme, images, ...) without the questions, for diagnosing what a copy lost.
        if ((questions.Any(q => FormsCopyTrace.IsTrickyType(q?["type"]?.GetValue<string>())) || ImageProps.Any(p => HasImage(src, p)))
            && src.DeepClone() is JsonObject formOnly)
        {
            formOnly.Remove("questions");
            FormsCopyTrace.Write($"SOURCE FORM (without questions): {formOnly.ToJsonString()}");
        }

        var copied = 0;
        foreach (var q in questions)
        {
            var qb = new JsonObject();
            foreach (var p in QuestionProps)
                if (q![p] is { } v) qb[p] = v.DeepClone();
            (st, body) = await _target.SendRawAsync(HttpMethod.Post, $"formapi/api/forms('{newId}')/questions", qb.ToJsonString(), ct);
            _log($"  question {q!["id"]} ({q["type"]}): {st}{(st >= 400 ? " " + Trim(body) : "")}");
            if (FormsCopyTrace.IsTrickyType(q["type"]?.GetValue<string>()))
                FormsCopyTrace.Write($"POST {q["type"]} id={q["id"]} groupId={q["groupId"]} -> {st}{Environment.NewLine}  source: {q.ToJsonString()}{Environment.NewLine}  sent:   {qb.ToJsonString()}{Environment.NewLine}  answer: {body}");
            if (st is 200 or 201)
            {
                copied++; setCopied(copied);
                await RepairQuestionInfoAsync(q, newId, body, warnings, ct);
                await CopyChoicesAsync(q, sourceFormId, newId, body, warnings, ct);
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

    // Ranking and Likert questions can come back from the service with default content (one "Option 1", no rating scale) when
    // created with their questionInfo. So the created question is compared with the source's (as JSON, not as text: the
    // service sends questionInfo as an escaped JSON string whose spacing and key order can differ), PATCHed when it differs,
    // read back, and a warning is shown if it STILL differs, instead of a silently wrong question. Every step is written in
    // full to FormsCopyTrace so a failing case can be diagnosed from the real payloads.
    private async Task RepairQuestionInfoAsync(JsonNode? q, string newId, string createResponse, List<string> warnings, CancellationToken ct)
    {
        var type = q?["type"]?.GetValue<string>() ?? "";
        if (!type.Contains("rank", StringComparison.OrdinalIgnoreCase) && !type.Contains("likert", StringComparison.OrdinalIgnoreCase)) return;
        var info = q!["questionInfo"];
        JsonNode? created = null;
        try { created = JsonNode.Parse(createResponse); } catch (JsonException) { }
        // The target may have given the question an ID of its own; the PATCH and the read-back must use that one.
        var qid = created?["id"]?.GetValue<string>() ?? q["id"]?.GetValue<string>();
        if (info is null || qid is null) return;
        var title = Clip(q["title"]?.GetValue<string>() ?? Loc.T("Forms_Cp_Untitled"), 60);

        _log($"    {type} source questionInfo: {Trim(info.ToJsonString())}");
        _log($"    {type} created questionInfo: {Trim(created?["questionInfo"]?.ToJsonString() ?? "(none)")}");
        if (SameInfo(created?["questionInfo"], info)) return;

        var patch = new JsonObject { ["questionInfo"] = info.DeepClone() };
        var questionUrl = $"formapi/api/forms('{newId}')/questions('{qid}')";
        var (st, body) = await _target.SendRawAsync(HttpMethod.Patch, questionUrl, patch.ToJsonString(), ct);
        _log($"    {type} repair questionInfo: {st}{(st >= 400 ? " " + Trim(body) : "")}");
        FormsCopyTrace.Write($"PATCH {type} {qid} -> {st}{Environment.NewLine}  sent:   {patch.ToJsonString()}{Environment.NewLine}  answer: {body}");
        if (st >= 400)
        {
            warnings.Add(Loc.T("Forms_Cp_OptionsFail", title, type, st));
            return;
        }

        // The PATCH can be accepted and still change nothing; read the question back to be sure.
        var (gst, gbody) = await _target.SendRawAsync(HttpMethod.Get, questionUrl, null, ct);
        JsonNode? after = null;
        try { after = JsonNode.Parse(gbody); } catch (JsonException) { }
        FormsCopyTrace.Write($"GET {type} {qid} after repair -> {gst}{Environment.NewLine}  answer: {gbody}");
        if (gst != 200 || !SameInfo(after?["questionInfo"], info))
        {
            _log($"    {type} still differs from the source after the repair");
            warnings.Add(Loc.T("Forms_Cp_OptionsFail", title, type, "options differ"));
        }
    }

    // questionInfo normally arrives as a string holding JSON; compare what it holds, so spacing and escaping do not matter.
    private static bool SameInfo(JsonNode? a, JsonNode? b) => Canonical(a) == Canonical(b);

    private static string Canonical(JsonNode? node)
    {
        if (node is null) return "";
        if (node is JsonValue v && v.TryGetValue<string>(out var text))
        {
            try { return JsonNode.Parse(text)?.ToJsonString() ?? ""; }
            catch (JsonException) { return text; }
        }
        return node.ToJsonString();
    }

    // Ranking options and the Likert answer scale are NOT part of the question's questionInfo (which only holds the shuffle / label
    // flags): they are a separate "choices" collection of the question, so creating the question copied none of them. The target
    // starts with a default choice or two, so those are updated in place, further choices are added, and any extra defaults
    // are removed. Requests and answers are written in full to FormsCopyTrace (the exact shape the service accepts is undocumented).
    private static bool HasChoiceCollection(string? type) =>
        type is not null && (type.Contains("rank", StringComparison.OrdinalIgnoreCase)
                             || type.Contains("MatrixChoiceGroup", StringComparison.OrdinalIgnoreCase));

    private static readonly string[] ChoiceProps = ["displayText", "description", "order", "isCorrect", "image"];

    private static string ChoiceKey(JsonNode? id) =>
        id is JsonValue v && v.TryGetValue<string>(out var text) ? $"'{text}'" : id?.ToJsonString() ?? "0";

    private static List<JsonObject> ChoicesOf(string body)
    {
        try
        {
            return (JsonNode.Parse(body)?["value"] as JsonArray)?.OfType<JsonObject>()
                .OrderBy(c => c["order"]?.GetValue<double>() ?? 0).ToList() ?? [];
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException) { return []; }
    }

    private async Task CopyChoicesAsync(JsonNode? q, string sourceFormId, string newId, string createResponse, List<string> warnings, CancellationToken ct)
    {
        var type = q?["type"]?.GetValue<string>();
        if (!HasChoiceCollection(type)) return;
        var srcQid = q!["id"]?.GetValue<string>();
        string? newQid = null;
        try { newQid = JsonNode.Parse(createResponse)?["id"]?.GetValue<string>(); } catch (JsonException) { }
        newQid ??= srcQid;
        if (srcQid is null || newQid is null) return;
        var title = Clip(q["title"]?.GetValue<string>() ?? Loc.T("Forms_Cp_Untitled"), 60);

        var (st, body) = await _source.SendRawAsync(HttpMethod.Get, $"formapi/api/forms('{sourceFormId}')/questions('{srcQid}')/choices", null, ct);
        FormsCopyTrace.Write($"CHOICES source {type} {srcQid} -> {st}{Environment.NewLine}  answer: {body}");
        var wanted = st == 200 ? ChoicesOf(body) : [];
        if (wanted.Count == 0) return;

        var choicesUrl = $"formapi/api/forms('{newId}')/questions('{newQid}')/choices";
        (st, body) = await _target.SendRawAsync(HttpMethod.Get, choicesUrl, null, ct);
        FormsCopyTrace.Write($"CHOICES target before {type} {newQid} -> {st}{Environment.NewLine}  answer: {body}");
        var existing = st == 200 ? ChoicesOf(body) : [];

        var failed = 0;
        for (var i = 0; i < wanted.Count; i++)
        {
            var full = new JsonObject();
            foreach (var p in ChoiceProps)
                if (wanted[i][p] is { } v && !(p == "image" && !HasImage(wanted[i], "image"))) full[p] = v.DeepClone();
            var minimal = new JsonObject { ["displayText"] = wanted[i]["displayText"]?.DeepClone(), ["order"] = wanted[i]["order"]?.DeepClone() ?? (i + 1) };

            // The first choices take the place of the defaults, in order; the rest are new.
            var method = i < existing.Count ? HttpMethod.Patch : HttpMethod.Post;
            var url = i < existing.Count ? $"{choicesUrl}({ChoiceKey(existing[i]["id"])})" : choicesUrl;
            (st, body) = await _target.SendRawAsync(method, url, full.ToJsonString(), ct);
            FormsCopyTrace.Write($"CHOICE {i + 1}/{wanted.Count} {method} {url} -> {st}{Environment.NewLine}  sent: {full.ToJsonString()}{Environment.NewLine}  answer: {body}");
            if (st >= 400)
            {
                // an unexpected property can make the service refuse the whole request: try with just the text and order
                (st, body) = await _target.SendRawAsync(method, url, minimal.ToJsonString(), ct);
                FormsCopyTrace.Write($"CHOICE {i + 1}/{wanted.Count} retry minimal {method} {url} -> {st}{Environment.NewLine}  sent: {minimal.ToJsonString()}{Environment.NewLine}  answer: {body}");
            }
            _log($"    {type} choice {i + 1}/{wanted.Count}: {st}{(st >= 400 ? " " + Trim(body) : "")}");
            if (st >= 400) failed++;
        }
        for (var i = wanted.Count; i < existing.Count; i++)
        {
            var url = $"{choicesUrl}({ChoiceKey(existing[i]["id"])})";
            (st, body) = await _target.SendRawAsync(HttpMethod.Delete, url, null, ct);
            FormsCopyTrace.Write($"CHOICE extra default DELETE {url} -> {st}{Environment.NewLine}  answer: {body}");
        }
        if (failed > 0) warnings.Add(Loc.T("Forms_Cp_OptionsFail", title, type, $"{failed} of {wanted.Count} options"));
    }

    // The image property of a form is read-only in practice: patching the form returns 204 and changes nothing, because the file
    // belongs to the SOURCE form (it lives under that form's address in the image store). What the Forms web app does (seen in the
    // browser's network log) is PUT the raw file to forms('id')/logo (or /background, /header) with the file's type as Content-Type
    // and a Content-Disposition header "attachment; name=<guid>; filename=<guid>"; the answer is the new image record. The file is
    // downloaded from the source image's address (readable without signing in). After the PUT the target is read back, so success
    // is only reported when the image really is on the form. The answer is written to FormsCopyTrace.
    private async Task<bool> TryUploadImageAsync(string prop, JsonNode image, string newId, Func<Task<JsonNode?>> readTarget, CancellationToken ct)
    {
        var url = image["resourceUrl"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(url)) return false;
        var (dst, bytes, mediaType) = await _source.DownloadAnonymousAsync(url, ct);
        FormsCopyTrace.Write($"IMAGE {prop} download {url} -> {dst}, {bytes?.Length ?? 0} bytes, {mediaType}");
        if (bytes is null or { Length: 0 }) return false;

        var contentType = image["contentType"]?.GetValue<string>() is { Length: > 0 } ct0 ? ct0 : mediaType ?? "image/png";
        var disposition = $"attachment; name={Guid.NewGuid()}; filename={Guid.NewGuid()}";
        var target = $"formapi/api/forms('{newId}')/{prop}";
        var (st, answer) = await _target.SendBytesAsync(HttpMethod.Put, target, bytes, contentType, disposition, ct);
        var landed = st is >= 200 and < 300 && HasImage(await readTarget(), prop);
        FormsCopyTrace.Write($"IMAGE {prop} PUT {target} ({contentType}, {bytes.Length} bytes) -> {st}, now present: {landed}{Environment.NewLine}  answer: {answer}");
        _log($"  form.{prop} image upload: {st}{(landed ? " (ok)" : "")}");
        return landed;
    }

    // Last resort when the service will not take an image: keep the file, so the user can add it to the new form by hand in a few
    // clicks (Microsoft Forms: the form header / theme > Upload). Saved under Documents\SharePointSmartCopy\Forms images; the path is
    // returned for the note. Null when the image could not be downloaded or saved.
    private async Task<string?> SaveImageFileAsync(string prop, JsonNode image, string? formTitle, CancellationToken ct)
    {
        var url = image["resourceUrl"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(url)) return null;
        var (st, bytes, mediaType) = await _source.DownloadAnonymousAsync(url, ct);
        FormsCopyTrace.Write($"IMAGE {prop} saved-file fallback: download -> {st}, {bytes?.Length ?? 0} bytes, {mediaType}");
        if (bytes is null or { Length: 0 }) return null;
        var type = image["contentType"]?.GetValue<string>() ?? mediaType ?? "";
        var ext = type.ToLowerInvariant() switch
        {
            "image/png" => ".png", "image/jpeg" or "image/jpg" => ".jpg", "image/gif" => ".gif", "image/webp" => ".webp",
            "image/bmp" => ".bmp", "image/svg+xml" => ".svg", _ => ".img",
        };
        try
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SharePointSmartCopy", "Forms images");
            Directory.CreateDirectory(dir);
            var safeTitle = string.Concat((formTitle ?? "form").Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c)).Trim();
            if (safeTitle.Length == 0) safeTitle = "form";
            var path = Path.Combine(dir, $"{safeTitle} - {prop}{ext}");
            await File.WriteAllBytesAsync(path, bytes, ct);
            FormsCopyTrace.Write($"IMAGE {prop} saved to {path}");
            return path;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            FormsCopyTrace.Write($"IMAGE {prop} could not be saved: {ex.Message}");
            return null;
        }
    }

    // The copied theme (themeV2, otherInfo) still points at the SOURCE form's background image, so the copy looks right only while the
    // source form exists. Once the new form has a background of its own, swap the source's address, resource ID and file ID for the
    // new ones in those two values.
    private async Task RepointThemeAsync(JsonObject src, JsonNode? target, string newId, CancellationToken ct)
    {
        var from = src["background"] as JsonObject;
        var to = target?["background"] as JsonObject;
        if (from is null || to is null) return;
        var swaps = new List<(string Old, string New)>();
        foreach (var key in new[] { "resourceUrl", "resourceId", "fileIdentifier" })
            if (from[key]?.GetValue<string>() is { Length: > 0 } o && to[key]?.GetValue<string>() is { Length: > 0 } n && o != n)
                swaps.Add((o, n));
        if (swaps.Count == 0) return;

        var patch = new JsonObject();
        foreach (var prop in new[] { "themeV2", "otherInfo" })
        {
            if (target?[prop]?.GetValue<string>() is not { Length: > 0 } text) continue;
            var changed = text;
            foreach (var (o, n) in swaps) changed = changed.Replace(o, n, StringComparison.OrdinalIgnoreCase);
            if (changed != text) patch[prop] = changed;
        }
        if (patch.Count == 0) return;
        var (st, body) = await _target.SendRawAsync(HttpMethod.Patch, $"formapi/api/forms('{newId}')", patch.ToJsonString(), ct);
        FormsCopyTrace.Write($"THEME repointed to the new background ({string.Join(", ", patch.Select(kv => kv.Key))}) -> {st}{Environment.NewLine}  answer: {body}");
        _log($"  theme now uses the new background: {st}");
    }

    private static readonly string[] ImageProps = ["background", "logo", "header"];

    private static bool HasImage(JsonNode? form, string prop) => form?[prop] is JsonObject o && o["resourceId"] is not null;

    // Which of the source's images (background, logo, header) are NOT on the copy. The target form is read back; an image that is
    // missing is patched once more on its own (the batch patch can be accepted and still drop a property), and read again. What was
    // sent and what came back goes to FormsCopyTrace, since the exact shape the service wants for an image is not documented.
    private async Task<Dictionary<string, string?>> MissingImagesAsync(JsonObject src, string newId, CancellationToken ct)
    {
        var missing = new Dictionary<string, string?>();
        var wanted = ImageProps.Where(p => HasImage(src, p)).ToList();
        if (wanted.Count == 0) return missing;

        async Task<JsonNode?> ReadTargetAsync()
        {
            var (st, body) = await _target.SendRawAsync(HttpMethod.Get, $"formapi/api/forms('{newId}')", null, ct);
            if (st != 200) return null;
            try { return JsonNode.Parse(body); } catch (JsonException) { return null; }
        }

        var target = await ReadTargetAsync();
        // What the source sent and what the target now holds, for every image, whether or not it "landed": a logo can be present
        // on the target form as a reference and still not show (the file can belong to the source form).
        foreach (var prop in wanted)
            FormsCopyTrace.Write($"IMAGE {prop}{Environment.NewLine}  source: {src[prop]!.ToJsonString()}{Environment.NewLine}  target: {target?[prop]?.ToJsonString() ?? "(none)"}");
        if (target is JsonObject targetOnly && targetOnly.DeepClone() is JsonObject t2)
        {
            t2.Remove("questions");
            FormsCopyTrace.Write($"TARGET FORM (without questions): {t2.ToJsonString()}");
        }
        foreach (var prop in wanted)
        {
            if (HasImage(target, prop)) continue;
            var patch = new JsonObject { [prop] = src[prop]!.DeepClone() };
            var (st, body) = await _target.SendRawAsync(HttpMethod.Patch, $"formapi/api/forms('{newId}')", patch.ToJsonString(), ct);
            _log($"  form.{prop} was missing after the copy; patched on its own: {st}{(st >= 400 ? " " + Trim(body) : "")}");
            target = await ReadTargetAsync();
            var landed = HasImage(target, prop);
            FormsCopyTrace.Write($"IMAGE {prop} missing after batch patch -> single patch {st}, now present: {landed}{Environment.NewLine}  source: {src[prop]!.ToJsonString()}{Environment.NewLine}  target: {target?[prop]?.ToJsonString() ?? "(none)"}{Environment.NewLine}  answer: {body}");
            if (!landed) landed = await TryUploadImageAsync(prop, src[prop]!, newId, ReadTargetAsync, ct);
            if (landed && prop == "background")
            {
                target = await ReadTargetAsync();
                await RepointThemeAsync(src, target, newId, ct);
            }
            if (!landed) missing[prop] = await SaveImageFileAsync(prop, src[prop]!, src["title"]?.GetValue<string>(), ct);
        }
        return missing;
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
