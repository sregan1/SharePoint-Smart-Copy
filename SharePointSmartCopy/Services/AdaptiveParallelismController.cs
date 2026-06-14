
namespace SharePointSmartCopy.Services;

/// Scales down the effective copy-parallelism when SharePoint throttles (HTTP 429/503),
/// then steps it back up after a quiet period.  Thread-safe.
///
/// How it works:
///   StepDown()  — called on each throttle event; reserves one slot to be withheld on the
///                 next Release() call, shrinking the live pool by 1 (floor: 1).
///   Release()   — called after each file completes; either absorbs the slot (if a step-down
///                 is pending) or returns it to the semaphore normally.
///   TryRestore()— fires on a 30-second quiet timer; returns one absorbed slot to the pool
///                 and re-arms until the pool is back to its original size.
internal sealed class AdaptiveParallelismController : IDisposable
{
    private readonly SemaphoreSlim _sem;
    private readonly int           _max;
    private readonly object        _lock = new();
    private          int           _limit;          // effective concurrency limit (1.._max)
    private          int           _pendingWithhold; // step-downs waiting for a Release() to absorb
    private          int           _withheld;        // slots fully absorbed (not in semaphore pool)
    private readonly System.Timers.Timer _restoreTimer;

    // Fired on the thread-pool when the effective limit changes.
    public event Action<int>? LimitChanged;

    public AdaptiveParallelismController(int max)
    {
        _max   = Math.Max(1, max);
        _limit = _max;
        _sem   = new SemaphoreSlim(_max, _max);

        _restoreTimer = new System.Timers.Timer(30_000) { AutoReset = false };
        _restoreTimer.Elapsed += (_, _) => TryRestore();
    }

    public int EffectiveLimit => _limit;

    public Task WaitAsync(CancellationToken ct) => _sem.WaitAsync(ct);

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
    public void StepDown()
    {
        lock (_lock)
        {
            if (_limit <= 1) return;
            _limit--;
            _pendingWithhold++;
            LimitChanged?.Invoke(_limit);
        }
        // Reset the restore timer so quiet-period is measured from the last throttle.
        _restoreTimer.Stop();
        _restoreTimer.Start();
    }

    private void TryRestore()
    {
        bool released;
        int  newLimit;
        lock (_lock)
        {
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

        // Schedule another restore tick if still below max.
        if (newLimit < _max)
        {
            _restoreTimer.Stop();
            _restoreTimer.Start();
        }
    }

    public void Dispose()
    {
        _restoreTimer.Dispose();
        _sem.Dispose();
    }
}
