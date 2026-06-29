
namespace SharePointSmartCopy.Services;

/// Scales down the effective copy-parallelism when SharePoint throttles (HTTP 429/503),
/// then steps it back up once the throttle window clears.  Thread-safe.
///
/// How it works:
///   StepDown(retryAfter) — called on each throttle event; HALVES the live limit (multiplicative
///                 decrease) and lowers the AIMD ceiling to match, so the pool settles below the
///                 tenant's throttle threshold instead of climbing straight back into it.
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
///   TryRestore() — fires on a fixed 5-second heartbeat.  Returns absorbed slots (climbing only up
///                 to the AIMD ceiling, not _max), but ONLY after _nextRestoreTime has passed. After
///                 a stable stretch with no throttle it nudges the ceiling up by 1 (additive
///                 increase) to re-probe for more capacity — so it converges on the safe rate.
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
    // The server's Retry-After already includes its own margin; a large extra cushion just leaves
    // concurrency stuck low across a long run, so keep it small.
    private static readonly TimeSpan RestoreBuffer    = TimeSpan.FromSeconds(3);

    // AIMD ceiling: the highest limit we'll currently climb to. A throttle HALVES it (multiplicative
    // decrease) so the pool doesn't immediately climb back into the same throttle; it then re-probes
    // upward by 1 every ReprobeInterval of clean running (additive increase). This makes the pool
    // settle just below the tenant's actual throttle threshold instead of oscillating at _max.
    private          int            _ceiling;
    private          DateTimeOffset _lastCeilingRaise = DateTimeOffset.MinValue;
    private static readonly TimeSpan ReprobeInterval  = TimeSpan.FromSeconds(45);

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
        _ceiling  = start;          // start conservative at the soft-start; probe UPWARD (additive
                                    // increase) only after clean running, so a high slider can't cause
                                    // an opening burst → throttle cascade. The slider is just the cap.
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

            // Multiplicative decrease: halve the live limit and remember it as the new ceiling, so
            // the pool settles below the throttle instead of climbing straight back into it. Withhold
            // the dropped slots (absorbed as in-flight work releases them).
            int target = Math.Max(1, _limit / 2);
            _pendingWithhold += _limit - target;
            _limit            = target;
            _ceiling          = Math.Min(_ceiling, target);
            LimitChanged?.Invoke(_limit);
        }
    }

    // Number of slots restored per heartbeat tick. Restoring more than one lets the pool climb
    // back to full concurrency in a few seconds after a transient throttle instead of taking one
    // tick (5 s) per slot — important on long runs where a brief 429 burst shouldn't cost minutes.
    private const int RestorePerTick = 2;

    private void TryRestore()
    {
        if (_disposed) return;
        for (int i = 0; i < RestorePerTick; i++)
        {
            bool released;
            int  newLimit;
            lock (_lock)
            {
                var now = DateTimeOffset.UtcNow;
                // Honour the quiet period — do not restore during an active throttle window.
                if (now < _nextRestoreTime) return;

                // Additive increase: after a stable stretch with no throttle, cautiously raise the
                // ceiling by 1 to probe whether more capacity is now available. Capped at _max.
                if (_ceiling < _max
                    && now - _lastStepDown     > ReprobeInterval
                    && now - _lastCeilingRaise > ReprobeInterval)
                {
                    _ceiling++;
                    _lastCeilingRaise = now;
                }

                // Climb only up to the discovered ceiling, not all the way back to _max.
                if (_limit >= _ceiling) return;

                // Only grow UNDER LOAD. If slots are sitting free, downloads aren't saturating the
                // current limit, so there's no demand for more — and ramping up while idle (e.g. during
                // the pre-download provisioning/analyze phase) just makes the next burst start at full
                // concurrency, which is exactly what triggers the opening throttle. When the pool is
                // exhausted (callers waiting on slots), CurrentCount is 0 and we climb normally.
                if (_sem.CurrentCount > 0) return;

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
    }

    public void Dispose()
    {
        _disposed = true;
        _restoreTimer.Dispose();
        _sem.Dispose();
    }
}
