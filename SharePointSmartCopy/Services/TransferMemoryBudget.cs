namespace SharePointSmartCopy.Services;

/// Bounds the total BYTES of file payloads in flight across the whole Migration API pipeline
/// (raw download buffers + encrypted copies, from download start until the file's blobs finish
/// uploading). Thread-safe, FIFO.
///
/// Why bytes and not counts: every other gate in the pipeline is count-based (download slots,
/// per-batch pipe capacity, per-batch upload concurrency, the 2-slot large-file gate) — and all
/// but the download gate are PER-BATCH, so 3 concurrent batches multiply them. A library of
/// ~300-490 MB files slides under the 500 MB large-file threshold entirely, and the count gates
/// happily admit 20+ such files at once ≈ 15+ GB of live buffers (observed 2026-07-18: 19 GB
/// heap on a 32 GB machine with Parallel Copies at just 5, plus connection-reset storms from the
/// resulting GC pauses). A single byte-denominated budget shared across all batches bounds the
/// real quantity — memory — no matter how file sizes are distributed.
///
/// FIFO matters: waiters are granted strictly in arrival order, so a large request parked at the
/// head cannot be starved by a stream of small ones slipping past it.
internal sealed class TransferMemoryBudget(long capacityBytes)
{
    private sealed class Waiter(long bytes)
    {
        public long Bytes { get; } = bytes;
        public TaskCompletionSource Tcs { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        public volatile bool Cancelled;
        public CancellationTokenRegistration Registration;
    }

    private readonly object _lock = new();
    private readonly Queue<Waiter> _waiters = new();
    private long _available = Math.Max(1, capacityBytes);

    public long Capacity { get; } = Math.Max(1, capacityBytes);

    // A single file larger than the whole budget must still be able to proceed (by itself, once
    // everything else drains) — clamp its charge to the full capacity rather than deadlocking.
    public long ClampCharge(long bytes) => Math.Clamp(bytes, 1, Capacity);

    public Task WaitAsync(long bytes, CancellationToken ct)
    {
        if (ct.IsCancellationRequested) return Task.FromCanceled(ct);
        bytes = ClampCharge(bytes);
        Waiter waiter;
        lock (_lock)
        {
            // Grant immediately only when no one is queued ahead — preserving FIFO.
            if (_waiters.Count == 0 && _available >= bytes)
            {
                _available -= bytes;
                return Task.CompletedTask;
            }
            waiter = new Waiter(bytes);
            _waiters.Enqueue(waiter);
        }
        if (ct.CanBeCanceled)
        {
            // Callback deliberately takes no lock (it can run synchronously inside Register):
            // it only flags the entry; Release() skips and disposes cancelled entries lazily —
            // and if cancellation races a grant, Release's TrySetResult failure re-credits the
            // already-deducted bytes.
            waiter.Registration = ct.Register(() =>
            {
                waiter.Cancelled = true;
                waiter.Tcs.TrySetCanceled(ct);
            });
        }
        return waiter.Tcs.Task;
    }

    public void Release(long bytes)
    {
        List<Waiter>? granted = null;
        lock (_lock)
        {
            _available = Math.Min(Capacity, _available + ClampCharge(bytes));
            while (_waiters.Count > 0)
            {
                var head = _waiters.Peek();
                if (head.Cancelled)
                {
                    _waiters.Dequeue();
                    head.Registration.Dispose();
                    continue;
                }
                if (_available < head.Bytes) break; // FIFO: never skip the head
                _waiters.Dequeue();
                _available -= head.Bytes;
                (granted ??= []).Add(head);
            }
        }
        if (granted == null) return;
        foreach (var w in granted)
        {
            w.Registration.Dispose();
            // Lost the race against cancellation after we already deducted its bytes —
            // the waiter will never run, so put its reservation back.
            if (!w.Tcs.TrySetResult())
                Release(w.Bytes);
        }
    }
}
