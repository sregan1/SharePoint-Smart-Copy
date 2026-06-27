
namespace SharePointSmartCopy.Services;

/// Scales down the effective copy-parallelism when SharePoint throttles (HTTP 429/503),
/// then steps it back up once the throttle window clears.  Thread-safe.
///
/// How it works:
///   StepDown(retryAfter) — called on each throttle event; reserves one slot to be withheld on
///                 the next Release() call, shrinking the live pool by 1 (floor: 1).
///                 Rate-limited to 1 step-down per 2s so a burst of 429s doesn't crater
///                 the limit to 1 in one shot.
///                 Also extends _nextRestoreTime = now + retryAfter + RestoreBuffer so that
///                 both the restore heartbeat and WaitAsync respect the active throttle window.
///   WaitAsync()  — blocks callers not only on the semaphore but also on _nextRestoreTime.
///                 This prevents a completing non-throttled download from immediately
///                 handing its slot to a fresh download that would hit the same throttle
///                 window and trigger further escalation (the cascade-to-1 failure mode).
///   Release()    — called after each file completes; either absorbs the slot (if a step-down
///                 is pending) or returns it to the semaphore normally.
///   TryRestore() — fires on a fixed 5-second heartbeat.  Returns one absorbed slot per tick,
///                 but ONLY after _nextRestoreTime has passed.
internal sealed class AdaptiveParallelismController : IDisposable
{
    private readonly SemaphoreSlim _sem;
    private readonly int           _max;
    private readonly object        _lock = new();
    private          int           _limit;           // effective concurrency limit (1.._max)
    private          int           _pendingWithhold; // step-downs waiting for a Release() to absorb
    private          int           _withheld;        // slots fully absorbed (not in semaphore pool)
    private          bool          _disposed;
    private readonly System.Timers.Timer _restoreTimer;
    private          DateTimeOffset _lastStepDown    = DateTimeOffset.MinValue;
    private          DateTimeOffset _nextRestoreTime = DateTimeOffset.MinValue;
    private static readonly TimeSpan StepDownCooldown = TimeSpan.FromSeconds(2);
    // Extra cushion added on top of the Retry-After value before any slot is reused or restored.
    private static readonly TimeSpan RestoreBuffer    = TimeSpan.FromSeconds(10);

    // Fired on the thread-pool when the effective limit changes.
    public event Action<int>? LimitChanged;

    // softStart: optional lower initial slot count.  When provided the controller begins at
    // softStart slots and ramps up to max via the 5-second restore heartbeat, only if no
    // throttle fires.  Avoids the initial full-burst → throttle → cascade pattern.
    // Omit (or pass -1) to start at max (default behaviour for direct-copy mode).
    public AdaptiveParallelismController(int max, int softStart = -1)
    {
        _max   = Math.Max(1, max);
        int start = (softStart > 0 && softStart < _max) ? softStart : _max;
        _limit    = start;
        _withheld = _max - start;   // pre-withheld slots; TryRestore releases one per 5s tick
        _sem      = new SemaphoreSlim(start, _max);

        _restoreTimer = new System.Timers.Timer(5_000) { AutoReset = true };
        _restoreTimer.Elapsed += (_, _) => TryRestore();
        _restoreTimer.Start();
    }

    public int EffectiveLimit => _limit;

    // Acquires a slot, but also waits out any active throttle window before doing so.
    // Downloads already holding a slot continue in Kiota's retry wait; only the acquisition
    // of NEW slots is blocked here.  This prevents the cascade-to-1 pattern where each
    // completing non-throttled download immediately starts a fresh download that hits the
    // same throttle window, triggering another step-down.
    public async Task WaitAsync(CancellationToken ct)
    {
        while (true)
        {
            TimeSpan remaining;
            lock (_lock) { remaining = _nextRestoreTime - DateTimeOffset.UtcNow; }
            if (remaining > TimeSpan.Zero)
                await Task.Delay(remaining, ct);

            await _sem.WaitAsync(ct);

            // Re-check: a new throttle may have extended _nextRestoreTime while we were
            // waiting for a semaphore slot.  If so, return the unused slot and loop.
            lock (_lock) { remaining = _nextRestoreTime - DateTimeOffset.UtcNow; }
            if (remaining <= TimeSpan.Zero) return;

            _sem.Release(); // direct release — this slot was never used, bypasses withhold logic
        }
    }

    // Must be called once per successful WaitAsync when the work unit finishes.
    public void Release()
    {
        bool absorb;
        lock (_lock)
        {
            absorb = _pendingWithhold > 0;
            if (absorb) { _pendingWithhold--; _withheld++; }
        }
        if (!absorb) _sem.Release();
    }

    // Called on each HTTP 429/503 throttle event.
    // retryAfter: the Retry-After duration from the HTTP response header.  When provided,
    // both WaitAsync and TryRestore are blocked until retryAfter + RestoreBuffer elapses.
    // Every call (even those suppressed by the step-down cooldown) extends the quiet period
    // so that repeated throttles within the same window don't allow premature slot reuse.
    public void StepDown(TimeSpan retryAfter = default)
    {
        lock (_lock)
        {
            var now = DateTimeOffset.UtcNow;

            // Extend the quiet period on every 429 call, not just on successful step-downs.
            if (retryAfter > TimeSpan.Zero)
            {
                var candidate = now + retryAfter + RestoreBuffer;
                if (candidate > _nextRestoreTime)
                    _nextRestoreTime = candidate;
            }

            if (_limit <= 1) return;
            if (now - _lastStepDown < StepDownCooldown) return;
            _lastStepDown = now;
            _limit--;
            _pendingWithhold++;
            LimitChanged?.Invoke(_limit);
        }
    }

    private void TryRestore()
    {
        if (_disposed) return;
        bool released;
        int  newLimit;
        lock (_lock)
        {
            // Honour the quiet period — do not restore during an active throttle window.
            if (DateTimeOffset.UtcNow < _nextRestoreTime) return;
            if (_limit >= _max) return;

            // Prefer cancelling a pending withhold (slot never actually left pool) over
            // releasing a fully absorbed slot — either way the effective limit goes up by 1.
            if (_pendingWithhold > 0)
            {
                _pendingWithhold--;
                _limit++;
                newLimit = _limit;
                released = false;
            }
            else if (_withheld > 0)
            {
                _withheld--;
                _limit++;
                newLimit = _limit;
                released = true;
            }
            else
            {
                return;
            }
        }

        if (released) _sem.Release();
        LimitChanged?.Invoke(newLimit);
    }

    public void Dispose()
    {
        _disposed = true;
        _restoreTimer.Dispose();
        _sem.Dispose();
    }
}
