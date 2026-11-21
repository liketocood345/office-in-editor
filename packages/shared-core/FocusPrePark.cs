namespace OfficeInEditor.Core;

/// <summary>
/// One watch per Office application while focus sessions are opening. Every frame of that application that appears
/// after the first opening started is moved out of view and off the taskbar at once, whichever process shows it: a
/// cold start, a second document in an Office that is already running, or several tabs restored together. Frames that
/// were there before stay untouched. Frames no session took go back where they were when the last opening ends.
/// </summary>
internal static class FocusPrePark
{
    const int SweepMs = 15;

    sealed class Watch(string application)
    {
        public readonly string Application = application;
        public readonly CancellationTokenSource Stop = new();
        /// <summary>Frames already seen (present at the start, or pre-parked since). Touched by the sweep only.</summary>
        public readonly HashSet<long> Known = [];
        public readonly List<IntPtr> Parked = [];
        public int Openings;
    }

    static readonly object gate = new();
    static readonly Dictionary<string, Watch> watches = new(StringComparer.Ordinal);

    /// <summary>Joins (or starts) the watch for <paramref name="application"/>; dispose once the session owns its window.</summary>
    internal static IDisposable Enter(string application, FocusRecovery recovery)
    {
        lock (gate)
        {
            if (!watches.TryGetValue(application, out var watch))
            {
                watch = new Watch(application);
                foreach (var frame in Win32FocusWindowOps.OfficeFrames(application, includeMinimized: true)) watch.Known.Add(frame.ToInt64());
                watches[application] = watch;
                // Its own thread: while Office starts, Core's pool threads wait on add-in calls, and a pool-based sweep
                // was seen to run a second late, after the frame had shown on screen.
                new Thread(() => Sweep(watch, recovery)) { IsBackground = true, Name = "OieFocusPrePark" }.Start();
            }
            watch.Openings++;
            return new Lease(watch);
        }
    }

    static void Sweep(Watch watch, FocusRecovery recovery)
    {
        var stop = watch.Stop.Token;
        do
        {
            try
            {
                foreach (var frame in Win32FocusWindowOps.OfficeFrames(watch.Application, includeMinimized: false))
                {
                    // Already out of view (the add-in opened it there): nothing to hide, and its placement now is not
                    // where it belongs. Looked at again later in case it comes into view.
                    if (watch.Known.Contains(frame.ToInt64()) || Win32FocusWindowOps.IsOutOfView(frame)) continue;
                    watch.Known.Add(frame.ToInt64());
                    if (Win32FocusWindowOps.PreParkFrame(frame, recovery)) watch.Parked.Add(frame);
                }
            }
            catch (Exception) { }
        }
        while (!stop.WaitHandle.WaitOne(SweepMs));
        // Released here, after the last sweep, so no frame is pre-parked after its release.
        var released = 0;
        foreach (var frame in watch.Parked)
        {
            if (!Win32FocusWindowOps.IsPreParked(frame)) continue;
            try { Win32FocusWindowOps.ReleasePrePark(frame, recovery); released++; } catch (Exception) { }
        }
        if (released > 0) DebugNdjson.Write("F", "FocusPrePark.End", "office_frames_released", new { count = released });
        watch.Stop.Dispose();
    }

    sealed class Lease(Watch watch) : IDisposable
    {
        int disposed;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) != 0) return;
            lock (gate)
            {
                if (--watch.Openings > 0) return;
                watches.Remove(watch.Application);
                watch.Stop.Cancel();
            }
        }
    }
}
