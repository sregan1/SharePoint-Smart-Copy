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
        if (response.StatusCode is System.Net.HttpStatusCode.TooManyRequests
                                or System.Net.HttpStatusCode.ServiceUnavailable)
        {
            var delay = response.Headers.RetryAfter?.Delta
                ?? (response.Headers.RetryAfter?.Date is { } when
                        ? when - DateTimeOffset.UtcNow
                        : TimeSpan.FromSeconds(60));
            if (delay < TimeSpan.Zero) delay = TimeSpan.FromSeconds(1);
            if (delay > TimeSpan.FromSeconds(120)) delay = TimeSpan.FromSeconds(120);

            response.Headers.TryGetValues("x-ms-throttle-reason", out var vals);
            var reason = vals?.FirstOrDefault();

            onThrottled(delay, 1, 3, reason);
        }
        return response;
    }
}
