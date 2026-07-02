using System.Runtime.InteropServices;

namespace SharePointSmartCopy.Services;

// Prevents Windows from sleeping while a long-running operation (copy, metadata update,
// verification) is active. A multi-hour migration that the OS suspends mid-run loses all
// in-flight work — observed on a 114k-file run (2026-07-01) where a sleep cost 95 minutes of
// wall-clock and stalled every pending Graph retry. The display is still allowed to turn off;
// only system suspend is held back.
//
// SetThreadExecutionState with ES_CONTINUOUS is per-thread state, so Begin/End must be called
// from the same thread — in practice both are invoked from the WPF UI thread via the
// MainViewModel property-changed callbacks.
internal static class SleepPreventionService
{
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint SetThreadExecutionState(uint esFlags);

    private const uint ES_CONTINUOUS      = 0x80000000;
    private const uint ES_SYSTEM_REQUIRED = 0x00000001;

    public static void Begin() => SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
    public static void End()   => SetThreadExecutionState(ES_CONTINUOUS);
}
