using System.Net.Http;

namespace SharePointSmartCopy.Services;

// Sits inside Kiota's RetryHandler in the Graph HTTP pipeline so raw 429/503 responses
// are visible before the retry handler absorbs them. Fires the callback so the adaptive
// parallelism controller can step down — Graph throttles would otherwise be invisible to it.
internal sealed class GraphThrottleNotifyHandler(Action<TimeSpan, int, int, string?> onThrottled) : DelegatingHandler
{
    protected override async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var response = await base.SendAsync(request, cancellationToken);
        // 504 included: Kiota's RetryHandler retries gateway timeouts too, so without this those
        // storms were absorbed invisibly with no step-down.
        if (response.StatusCode is System.Net.HttpStatusCode.TooManyRequests
                                or System.Net.HttpStatusCode.ServiceUnavailable
                                or System.Net.HttpStatusCode.GatewayTimeout)
        {
            // No Retry-After header: real SharePoint throttles nearly always carry one; a bare
            // 503/504 is more often a transient hiccup. The old 60s default froze all new slot
            // acquisition for a minute on a single blip — use a modest default instead.
            var delay = response.Headers.RetryAfter?.Delta
                ?? (response.Headers.RetryAfter?.Date is { } when
                        ? when - DateTimeOffset.UtcNow
                        : TimeSpan.FromSeconds(10));
            if (delay < TimeSpan.Zero) delay = TimeSpan.FromSeconds(1);
            if (delay > TimeSpan.FromSeconds(120)) delay = TimeSpan.FromSeconds(120);

            response.Headers.TryGetValues("x-ms-throttle-reason", out var vals);
            var reason = vals?.FirstOrDefault();

            // attempt/max 0,0: this handler only observes (Kiota retries); passing a fake "1/3"
            // made the status bar permanently show "attempt 1/3" for Graph throttles.
            onThrottled(delay, 0, 0, reason);
        }
        return response;
    }
}
