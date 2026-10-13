using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions;
using SharePointSmartCopy.Localization;

namespace SharePointSmartCopy.Services;

// Turns the raw exception from a failed "Connect" into something a person can act on.
//
// The raw text (an MSAL/Entra "AADSTS65001…" blob, a Graph OData error, an HTTP status) was shown
// verbatim, clipped, and with no hint of what to do — and the most common failures are all
// configuration problems with a specific fix. The original message is kept, shortened, after the
// advice so a bug report still has the technical detail.
public static class ConnectionErrors
{
    // needsSettings: the fix is in Settings (so the UI offers an "Open Settings" button).
    public static (string Message, bool NeedsSettings) Translate(Exception ex, bool isConfigured)
    {
        if (!isConfigured)
            return (Loc.T("Err_NoRegistration"), true);

        var raw = Describe(ex);

        if (ex is MsalClientException { ErrorCode: "authentication_canceled" } ||
            raw.Contains("authentication_canceled", StringComparison.OrdinalIgnoreCase))
            return (Loc.T("Err_SignInCancelled"), false);

        // Entra "AADSTS" codes — the stable part of an otherwise long, wordy message.
        if (HasCode(raw, "AADSTS65001") || raw.Contains("consent_required", StringComparison.OrdinalIgnoreCase))
            return (Loc.T("Err_ConsentMissing") + Details(raw), false);
        if (HasCode(raw, "AADSTS700016"))
            return (Loc.T("Err_AppNotInTenant") + Details(raw), true);
        if (HasCode(raw, "AADSTS90002") || HasCode(raw, "AADSTS900144") || HasCode(raw, "AADSTS50049"))
            return (Loc.T("Err_TenantNotFound") + Details(raw), true);
        if (HasCode(raw, "AADSTS50011"))
            return (Loc.T("Err_RedirectMismatch") + Details(raw), true);
        if (HasCode(raw, "AADSTS7000218") || HasCode(raw, "AADSTS70002"))
            return (Loc.T("Err_PublicClient") + Details(raw), true);
        if (HasCode(raw, "AADSTS50020") || HasCode(raw, "AADSTS50105") || HasCode(raw, "AADSTS50034"))
            return (Loc.T("Err_AccountNotAllowed") + Details(raw), true);

        int? status = ex switch
        {
            ApiException api                    => api.ResponseStatusCode,
            System.Net.Http.HttpRequestException h => (int?)h.StatusCode,
            _                                   => null,
        };
        if (status == null)
        {
            if (raw.Contains("accessDenied", StringComparison.OrdinalIgnoreCase) || raw.Contains("Forbidden", StringComparison.OrdinalIgnoreCase)) status = 403;
            else if (raw.Contains("itemNotFound", StringComparison.OrdinalIgnoreCase) || raw.Contains("Not Found", StringComparison.OrdinalIgnoreCase)) status = 404;
            else if (raw.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase)) status = 401;
        }

        return status switch
        {
            401 => (Loc.T("Err_Unauthorized") + Details(raw), false),
            403 => (Loc.T("Err_Forbidden") + Details(raw), false),
            404 => (Loc.T("Err_SiteNotFound") + Details(raw), false),
            _   => (Loc.T("VM_ErrorPrefix", raw), false),
        };
    }

    private static bool HasCode(string text, string code) =>
        text.Contains(code, StringComparison.OrdinalIgnoreCase);

    // The exception's own message, plus its first inner message when that adds something.
    private static string Describe(Exception ex)
    {
        var msg = ex.Message ?? string.Empty;
        if (ex.InnerException?.Message is { Length: > 0 } inner && !msg.Contains(inner, StringComparison.Ordinal))
            msg += " " + inner;
        return msg;
    }

    // Keep the technical text for bug reports, but never let it dominate the message.
    private static string Details(string raw)
    {
        var oneLine = raw.Replace('\r', ' ').Replace('\n', ' ').Trim();
        if (oneLine.Length > 220) oneLine = oneLine[..220] + "…";
        return oneLine.Length == 0 ? "" : " " + Loc.T("Err_Details", oneLine);
    }
}
