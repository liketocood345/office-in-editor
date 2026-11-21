using System.Collections.Concurrent;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeInEditor.Core;

/// <summary>Work that lives on the focus dispatcher thread.</summary>
internal interface IFocusTicker
{
    void Tick(long now);
    void ForegroundChanged(long now);
}

/// <summary>
/// One STA thread for every focus session: owns the EVENT_SYSTEM_FOREGROUND hook, the 100 ms tick and all window
/// calls, so the state machines never race. UIA and COM may pump messages re-entrantly, so the hook only sets a flag
/// and work is drained from a queue between pumps, never from inside a callback.
/// </summary>
internal sealed class FocusDispatcher : IDisposable
{
    public const int TickMs = 100;
    const uint EventSystemForeground = 3;
    const uint WinEventOutOfContext = 0;
    const uint QsAllInput = 0x04FF;
    const uint MwmoInputAvailable = 0x0004;
    readonly ConcurrentQueue<(string Label, Action Action)> work = new();
    readonly AutoResetEvent wake = new(false);
    readonly List<IFocusTicker> tickers = [];
    readonly Thread thread;
    readonly WinEventProc hook;
    volatile bool foregroundDirty;
    volatile bool stopping;
    volatile string? currentWork;
    long currentStart;

    public FocusDispatcher()
    {
        hook = (_, _, _, _, _, _, _) => { foregroundDirty = true; };
        thread = new Thread(Run) { IsBackground = true, Name = "focus-dispatcher" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public bool IsDispatcherThread => Thread.CurrentThread == thread;

    /// <summary>What the dispatcher thread is running now (a code label) and since when; read by the stall watchdog.</summary>
    public (string? Work, long StartedAt) CurrentWork => (currentWork, Interlocked.Read(ref currentStart));

    /// <summary>
    /// Upper bound for one call to wait on the dispatcher. A window call that hangs on the dispatcher thread must not
    /// leave broker requests waiting forever: piled-up requests exhausted the 16 broker connections once on the lab.
    /// </summary>
    public static readonly TimeSpan CallTimeout = TimeSpan.FromSeconds(10);

    public async Task<T> InvokeAsync<T>(Func<T> action, [CallerMemberName] string caller = "", [CallerLineNumber] int line = 0)
    {
        var done = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        if (stopping) throw new CoreFault("focus_unavailable", "The focus dispatcher is stopping.");
        work.Enqueue((caller + ":" + line, () =>
        {
            try { done.SetResult(action()); }
            catch (Exception error) { done.SetException(error); }
        }));
        wake.Set();
        try { return await done.Task.WaitAsync(CallTimeout).ConfigureAwait(false); }
        catch (TimeoutException)
        {
            DebugNdjson.Write("F", "FocusDispatcher", "call_timeout", null);
            CoreTrace.Mark("dispatcher", "call-timeout:" + caller + ":" + line);
            throw new CoreFault("focus_busy", "The focus window service did not answer in time.");
        }
    }

    public Task InvokeAsync(Action action, [CallerMemberName] string caller = "", [CallerLineNumber] int line = 0) =>
        InvokeAsync(() => { action(); return true; }, caller, line);

    public void Add(IFocusTicker ticker) => Post("add", () => tickers.Add(ticker));
    public void Remove(IFocusTicker ticker) => Post("remove", () => tickers.Remove(ticker));
    void Post(string label, Action action) { work.Enqueue((label, action)); wake.Set(); }

    void Run()
    {
        NativeWindow.TryEnableDpi();
        var events = SetWinEventHook(EventSystemForeground, EventSystemForeground, IntPtr.Zero, hook, 0, 0, WinEventOutOfContext);
        if (events == IntPtr.Zero) DebugNdjson.Write("F", "FocusDispatcher", "foreground_hook_failed", null);
        var handles = new[] { wake.SafeWaitHandle.DangerousGetHandle() };
        var nextTick = Environment.TickCount64 + TickMs;
        try
        {
            while (!stopping)
            {
                var wait = (uint)Math.Clamp(nextTick - Environment.TickCount64, 0, TickMs);
                MsgWaitForMultipleObjectsEx(1, handles, wait, QsAllInput, MwmoInputAvailable);
                while (PeekMessage(out var message, IntPtr.Zero, 0, 0, 1))
                {
                    TranslateMessage(ref message);
                    DispatchMessage(ref message);
                }
                while (work.TryDequeue(out var item)) Guard(item.Label, item.Action);
                var now = Environment.TickCount64;
                if (foregroundDirty)
                {
                    foregroundDirty = false;
                    foreach (var ticker in tickers.ToArray()) Guard("foreground", () => ticker.ForegroundChanged(now));
                }
                if (now >= nextTick)
                {
                    nextTick = now + TickMs;
                    foreach (var ticker in tickers.ToArray()) Guard("tick", () => ticker.Tick(now));
                }
            }
            while (work.TryDequeue(out var item)) Guard(item.Label, item.Action);
        }
        finally { if (events != IntPtr.Zero) UnhookWinEvent(events); }
    }

    void Guard(string label, Action action)
    {
        Interlocked.Exchange(ref currentStart, Environment.TickCount64);
        currentWork = label;
        try { action(); }
        catch (Exception error) { DebugNdjson.Write("F", "FocusDispatcher", "work_failed", new { type = error.GetType().Name, code = (error as CoreFault)?.Code }); }
        finally { currentWork = null; }
    }

    public void Dispose()
    {
        if (stopping) return;
        stopping = true;
        wake.Set();
        CoreTrace.Mark("dispatcher", "stopping");
        if (!IsDispatcherThread && !thread.Join(TimeSpan.FromSeconds(3))) CoreTrace.Mark("dispatcher", "join-timeout");
    }

    delegate void WinEventProc(IntPtr hook, uint eventType, IntPtr window, int objectId, int childId, uint thread, uint time);

    [StructLayout(LayoutKind.Sequential)]
    struct Message { public IntPtr Window; public uint Id; public IntPtr WParam, LParam; public uint Time; public int X, Y; }

    [DllImport("user32.dll")] static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr module, WinEventProc callback, uint process, uint thread, uint flags);
    [DllImport("user32.dll")] static extern bool UnhookWinEvent(IntPtr hook);
    [DllImport("user32.dll")] static extern uint MsgWaitForMultipleObjectsEx(uint count, IntPtr[] handles, uint milliseconds, uint wakeMask, uint flags);
    [DllImport("user32.dll")] static extern bool PeekMessage(out Message message, IntPtr window, uint min, uint max, uint remove);
    [DllImport("user32.dll")] static extern bool TranslateMessage(ref Message message);
    [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref Message message);
}
