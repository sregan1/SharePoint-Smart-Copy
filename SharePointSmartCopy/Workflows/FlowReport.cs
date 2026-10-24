using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace SharePointSmartCopy.Workflows;

// The parts of the rewrite report that are about the flows as a whole rather than one package: what the flows use,
// and what must still be done by hand after import. Also renders the report's Markdown as a standalone HTML page.
public static partial class FlowReport
{
    // "What these flows use" plus "To do after import". extraFollowUps are the dialog's own findings (lists with no
    // target, columns with no match); everything else is derived from the scan.
    public static string Overview(IReadOnlyList<FlowScan> scans, IEnumerable<string> extraFollowUps)
    {
        var sb = new StringBuilder();
        var resources = FlowResources.Summarize(scans);

        sb.Append("## What these flows use\n\n");
        foreach (var r in resources)
            sb.Append($"- **{r.Service}** — {r.Flows} flow(s), {r.Operations} step(s): {r.Detail} ({HandlingLabel(r.Handling)})\n");
        sb.Append('\n');

        var todo = new List<string>
        {
            "Import each patched package in the target environment and recreate the connections (they are never copied).",
        };
        foreach (var r in resources.Where(r => r.Handling == ResourceHandling.ReviewManually))
            todo.Add($"{r.Service}: {r.Detail}.");
        if (scans.Any(s => s.IsSolution))
            todo.Add("Solution packages: only the workflow definitions were rewritten. Environment variable values and " +
                     "connection references in the solution (customizations.xml) still hold source values; set them after import.");
        if (scans.SelectMany(s => s.Findings).Any(f => f.Kind == FindingKind.UserId))
            todo.Add("Flow owners and run-only users are not part of the package; share the flows with the right people after import.");
        todo.AddRange(extraFollowUps);

        sb.Append("## To do after import\n\n");
        foreach (var t in todo.Distinct()) sb.Append($"- {t}\n");
        sb.Append('\n');
        return sb.ToString();
    }

    private static string HandlingLabel(ResourceHandling h) => h switch
    {
        ResourceHandling.Rewritten => "rewritten by this tool",
        ResourceHandling.ReviewManually => "review manually",
        _ => "connection only",
    };

    // Renders the Markdown subset the report uses (headings, bullets, paragraphs, **bold**, `code`, _italic_ lines).
    public static string ToHtml(string markdown, string title)
    {
        var body = new StringBuilder();
        var inList = false;
        void CloseList() { if (inList) { body.Append("</ul>\n"); inList = false; } }

        foreach (var raw in markdown.Replace("\r\n", "\n").Split('\n'))
        {
            var line = raw.TrimEnd();
            if (line.Length == 0) { CloseList(); continue; }
            if (line.StartsWith("### ")) { CloseList(); body.Append($"<h3>{Inline(line[4..])}</h3>\n"); }
            else if (line.StartsWith("## ")) { CloseList(); body.Append($"<h2>{Inline(line[3..])}</h2>\n"); }
            else if (line.StartsWith("# ")) { CloseList(); body.Append($"<h1>{Inline(line[2..])}</h1>\n"); }
            else if (line.StartsWith("- "))
            {
                if (!inList) { body.Append("<ul>\n"); inList = true; }
                body.Append($"<li>{Inline(line[2..])}</li>\n");
            }
            else { CloseList(); body.Append($"<p>{Inline(line)}</p>\n"); }
        }
        CloseList();

        return "<!doctype html>\n<html lang=\"en\"><head><meta charset=\"utf-8\">" +
               "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">" +
               $"<title>{WebUtility.HtmlEncode(title)}</title>\n" +
               "<style>:root{color-scheme:light dark}body{font:14px/1.5 Segoe UI,system-ui,sans-serif;max-width:900px;margin:24px auto;padding:0 16px}" +
               "code{background:rgba(127,127,127,.18);padding:1px 4px;border-radius:3px;font-family:Consolas,monospace;font-size:13px;word-break:break-all}" +
               "h1{font-size:22px}h2{font-size:18px;margin-top:28px;border-bottom:1px solid rgba(127,127,127,.35);padding-bottom:4px}h3{font-size:15px;margin-top:20px}" +
               "li{margin:2px 0}</style></head><body>\n" + body + "</body></html>\n";
    }

    private static string Inline(string text)
    {
        var s = WebUtility.HtmlEncode(text);
        s = Code().Replace(s, "<code>$1</code>");
        s = Bold().Replace(s, "<strong>$1</strong>");
        s = ItalicLine().Replace(s, "<em>$1</em>");
        return s;
    }

    [GeneratedRegex("`([^`]+)`")] private static partial Regex Code();
    [GeneratedRegex(@"\*\*(.+?)\*\*")] private static partial Regex Bold();
    [GeneratedRegex("^_(.+)_$")] private static partial Regex ItalicLine();
}
