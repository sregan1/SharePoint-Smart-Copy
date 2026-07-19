using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SharePointSmartCopy.Services;

// Periodic resource snapshot written to the activity log during long Migration API runs. A
// UCEERR_RENDERTHREADFAILURE crash is a native WPF composition-engine failure that bypasses every
// managed exception handler (see CopyResult.OnPropertyChanged) and every other diagnostic in the
// app, so the activity log — flushed to disk one line at a time — is the only record left after
// the process dies. GDI/USER object counts matter because Windows caps each process at 10,000 of
// either; exhausting one is a distinct, unrelated way to trigger the same WPF failure. The queued
// UI-dispatch count matters because CopyResult.OnPropertyChanged marshals via BeginInvoke: if
// background threads enqueue faster than the UI thread drains, that backlog grows unbounded with
// no other visible symptom until the crash.
internal static class ProcessDiagnostics
{
    [DllImport("user32.dll")]
    private static extern uint GetGuiResources(IntPtr hProcess, uint uiFlags);

    private const uint GR_GDIOBJECTS  = 0;
    private const uint GR_USEROBJECTS = 1;

    public static string Snapshot()
    {
        using var proc = Process.GetCurrentProcess();
        var workingSetMb = proc.WorkingSet64 / 1024 / 1024;
        var gcMb         = GC.GetTotalMemory(false) / 1024 / 1024;
        var gdiCount     = GetGuiResources(proc.Handle, GR_GDIOBJECTS);
        var userCount    = GetGuiResources(proc.Handle, GR_USEROBJECTS);
        return $"♥ Heartbeat: {workingSetMb:N0} MB working set, {gcMb:N0} MB managed heap, "
             + $"{proc.Threads.Count} threads, {gdiCount} GDI / {userCount} USER objects, "
             + $"{Models.CopyResult.PendingUiDispatches} UI update(s) queued";
    }
}
