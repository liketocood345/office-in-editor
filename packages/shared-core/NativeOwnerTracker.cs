using System.Collections.Concurrent;
using System.Runtime.InteropServices;

namespace OfficeInEditor.Core;

/// <summary>
/// Follows one proven owner HWND without entering Office, UI Automation, or its STA.
/// Call Update after committing a freshly measured physical rectangle. Callbacks only
/// request deferred calibration; they run on the shared WinEvent thread, not an OLE STA.
/// </summary>
internal sealed class NativeOwnerTracker : IDisposable
{
    internal readonly record struct Bounds(int X, int Y, int Width, int Height);
    private sealed record Binding(IntPtr Owner, IntPtr Popup, uint Process, uint Thread,
        uint PopupThread, uint Dpi, Bounds AnchorOwner, Bounds AnchorPopup, bool DesiredVisible,
        bool HasAnchor, bool AnchorInvalidated, long Generation);

    private static readonly object SharedGate = new();
    private static EventThread? shared;
    private static long nextId;
    private static int activeTrackers, runningThreads;
    private readonly object gate = new();
    private readonly EventThread thread;
    private readonly long id = Interlocked.Increment(ref nextId);
    private readonly Action<bool>? requestLayoutAfterMove;
    private Binding binding;
    private int disposed, queued, inMoveSize;
    // The fields below are only touched by the WinEvent thread.
    private readonly List<IntPtr> hooks = [];
    private Binding? applied;
    private int events;
    private bool minimized, ownerHidden, ownerDestroyed, needsAnchor, invalidationNotified, hookFailureNotified;
    private bool recoveryPending;
    private bool? pendingMinimized, pendingHidden;

    private const int Location = 1, MoveStart = 2, MoveEnd = 4, MinimizeStart = 8,
        MinimizeEnd = 16, Destroy = 32, Hide = 64, Show = 128;

    private NativeOwnerTracker(EventThread thread, Binding binding, Action<bool>? requestLayoutAfterMove)
    {
        this.thread = thread;
        this.binding = binding;
        this.requestLayoutAfterMove = requestLayoutAfterMove;
    }

    internal static int ActiveTrackerCount => Volatile.Read(ref activeTrackers);
    internal static bool EventThreadRunning => Volatile.Read(ref runningThreads) != 0;
    internal bool IsInMoveSize => Volatile.Read(ref inMoveSize) != 0;

    internal static NativeOwnerTracker Create(IntPtr owner, IntPtr popup, Action<bool>? requestLayoutAfterMove = null)
    {
        var binding = ReadBinding(owner, popup, default, default, false, false, 1);
        lock (SharedGate)
        {
            shared ??= new EventThread();
            var tracker = new NativeOwnerTracker(shared, binding, requestLayoutAfterMove);
            shared.Add(tracker);
            Interlocked.Increment(ref activeTrackers);
            tracker.Schedule();
            return tracker;
        }
    }

    internal void Update(Bounds anchorOwner, Bounds anchorPopup, bool desiredVisible)
    {
        Binding current;
        lock (gate) current = binding;
        Update(current.Owner, current.Popup, anchorOwner, anchorPopup, desiredVisible);
    }

    internal void Update(IntPtr owner, IntPtr popup, Bounds anchorOwner, Bounds anchorPopup, bool desiredVisible)
    {
        if (Volatile.Read(ref disposed) != 0) return;
        if (!Valid(anchorOwner) || !Valid(anchorPopup)) throw new ArgumentOutOfRangeException(nameof(anchorOwner));
        var next = ReadBinding(owner, popup, anchorOwner, anchorPopup, desiredVisible, true, 0);
        lock (gate)
        {
            if (disposed != 0) return;
            binding = next with { Generation = binding.Generation + 1 };
        }
        Schedule();
    }

    /// <summary>
    /// Discards geometry without discarding the owner subscription or the tab's
    /// desired visibility. Only a later Update may authorize showing the popup.
    /// </summary>
    internal void InvalidateAnchor()
    {
        lock (gate)
        {
            if (disposed != 0) return;
            binding = binding with { HasAnchor = false, AnchorInvalidated = true, Generation = binding.Generation + 1 };
        }
        Schedule();
    }

    private static Binding ReadBinding(IntPtr owner, IntPtr popup, Bounds anchorOwner, Bounds anchorPopup,
        bool visible, bool hasAnchor, long generation)
    {
        if (!Native.IsWindow(owner) || !Native.IsWindow(popup) || owner == popup)
            throw new ArgumentException("A live owner and a distinct owned popup are required.");
        var ownerThread = Native.GetWindowThreadProcessId(owner, out var ownerProcess);
        var popupThread = Native.GetWindowThreadProcessId(popup, out var popupProcess);
        if (ownerThread == 0 || ownerProcess == 0 || popupThread == 0 || popupProcess != (uint)Environment.ProcessId)
            throw new ArgumentException("Only a popup in this process may be moved.");
        return new(owner, popup, ownerProcess, ownerThread, popupThread, Native.GetDpiForWindow(owner),
            anchorOwner, anchorPopup, visible, hasAnchor, false, generation);
    }

    internal static bool AcceptsEvent(IntPtr expectedOwner, uint expectedThread, IntPtr hwnd,
        int objectId, int childId, uint eventThread) =>
        expectedOwner != IntPtr.Zero && hwnd == expectedOwner && expectedThread != 0
        && eventThread == expectedThread && objectId == 0 && childId == 0;

    internal static bool TryTranslate(Bounds anchorOwner, Bounds anchorPopup, Bounds currentOwner, out Bounds popup)
    {
        popup = default;
        if (!Valid(anchorOwner) || !Valid(anchorPopup) || !Valid(currentOwner)
            || anchorOwner.Width != currentOwner.Width || anchorOwner.Height != currentOwner.Height) return false;
        long x = (long)anchorPopup.X + currentOwner.X - anchorOwner.X;
        long y = (long)anchorPopup.Y + currentOwner.Y - anchorOwner.Y;
        if (x < int.MinValue || x > int.MaxValue || y < int.MinValue || y > int.MaxValue
            || x + anchorPopup.Width > int.MaxValue || y + anchorPopup.Height > int.MaxValue) return false;
        popup = new((int)x, (int)y, anchorPopup.Width, anchorPopup.Height);
        return true;
    }

    private static bool Valid(Bounds bounds) => bounds.Width > 0 && bounds.Height > 0
        && (long)bounds.X + bounds.Width <= int.MaxValue && (long)bounds.Y + bounds.Height <= int.MaxValue;

    private void Schedule()
    {
        if (Interlocked.CompareExchange(ref queued, 1, 0) != 0) return;
        if (!thread.Post(id)) Interlocked.Exchange(ref queued, 0);
    }

    private void OnEvent(uint eventType, IntPtr hwnd, int objectId, int childId, uint eventThread)
    {
        if (Volatile.Read(ref disposed) != 0 || applied is not { } current
            || !AcceptsEvent(current.Owner, current.Thread, hwnd, objectId, childId, eventThread)) return;
        if (eventType == Native.EventMoveStart) Volatile.Write(ref inMoveSize, 1);
        else if (eventType == Native.EventMoveEnd) Volatile.Write(ref inMoveSize, 0);
        if (eventType == Native.EventMinimizeStart) pendingMinimized = true;
        else if (eventType == Native.EventMinimizeEnd) pendingMinimized = false;
        if (eventType == Native.EventHide) pendingHidden = true;
        else if (eventType == Native.EventShow) pendingHidden = false;
        events |= eventType switch
        {
            Native.EventLocation => Location,
            Native.EventMoveStart => MoveStart,
            Native.EventMoveEnd => MoveEnd,
            Native.EventMinimizeStart => MinimizeStart,
            Native.EventMinimizeEnd => MinimizeEnd,
            Native.EventDestroy => Destroy,
            Native.EventHide => Hide,
            Native.EventShow => Show,
            _ => 0
        };
        Schedule();
    }

    private void Process()
    {
        Interlocked.Exchange(ref queued, 0);
        if (Volatile.Read(ref disposed) != 0) { Unhook(); thread.Remove(this); return; }
        Binding current;
        lock (gate) current = binding;
        var pending = events;
        events = 0;
        var nextMinimized = pendingMinimized;
        var nextHidden = pendingHidden;
        pendingMinimized = pendingHidden = null;
        var changedBinding = applied is null || applied.Owner != current.Owner || applied.Popup != current.Popup
            || applied.Process != current.Process || applied.Thread != current.Thread;
        var rebinding = changedBinding || (hooks.Count == 0 && !ownerDestroyed);
        if (rebinding)
        {
            Unhook();
            if (changedBinding)
            {
                minimized = ownerHidden = ownerDestroyed = hookFailureNotified = recoveryPending = false;
                Volatile.Write(ref inMoveSize, 0);
            }
            if (!IdentityMatches(current)) { applied = current; ownerDestroyed = true; HidePopup(current); return; }
            foreach (var range in new (uint First, uint Last)[]
            {
                (Native.EventLocation, Native.EventLocation),
                (Native.EventMoveStart, Native.EventMoveEnd),
                (Native.EventMinimizeStart, Native.EventMinimizeEnd),
                (Native.EventDestroy, Native.EventHide)
            })
            {
                var hook = Native.SetWinEventHook(range.First, range.Last, IntPtr.Zero, thread.Callback,
                    current.Process, current.Thread, 0 /* WINEVENT_OUTOFCONTEXT; exact owner filter */);
                if (hook == IntPtr.Zero)
                {
                    Unhook(); applied = current; needsAnchor = true; HidePopup(current);
                    if (!hookFailureNotified) { hookFailureNotified = true; Notify(false); }
                    return;
                }
                hooks.Add(hook);
                thread.Register(hook, this);
            }
            hookFailureNotified = false;
            pending = 0; // Old owner's queued events do not apply to the new binding.
            nextMinimized = nextHidden = null;
        }
        if (applied?.Generation != current.Generation)
        {
            needsAnchor = invalidationNotified = false;
            if (current.HasAnchor) recoveryPending = false;
        }
        applied = current;
        if ((pending & Destroy) != 0) ownerDestroyed = true;
        if (nextMinimized.HasValue) minimized = nextMinimized.Value;
        if (nextHidden.HasValue) ownerHidden = nextHidden.Value;
        if (ownerDestroyed || !IdentityMatches(current))
        {
            ownerDestroyed = true;
            HidePopup(current);
            Unhook();
            return;
        }
        if (!current.HasAnchor)
        {
            if (current.AnchorInvalidated)
            {
                HidePopup(current);
                if (!current.DesiredVisible) recoveryPending = false;
                else if ((pending & (MoveEnd | MinimizeEnd | Show)) != 0) recoveryPending = true;
                // MINIMIZEEND may precede the owner's actual restore. Retain its
                // one-shot request until a later native event confirms readiness.
                // Neither this notification nor an owner SHOW authorizes revealing
                // the popup: only Update can install another trusted anchor.
                if (recoveryPending && current.DesiredVisible && !minimized && !ownerHidden
                    && !IsInMoveSize && !Native.IsIconic(current.Owner) && Native.IsWindowVisible(current.Owner))
                {
                    var currentGeneration = false;
                    lock (gate) currentGeneration = disposed == 0 && binding.Generation == current.Generation;
                    if (currentGeneration) { recoveryPending = false; Notify(true); }
                }
            }
            return;
        }
        if (!current.DesiredVisible || minimized || ownerHidden || Native.IsIconic(current.Owner)
            || !Native.IsWindowVisible(current.Owner))
        {
            HidePopup(current);
            return;
        }
        using var physical = new PhysicalCoordinates();
        if (HostedChild(current))
        {
            if (!Native.GetClientRect(current.Owner, out var client)) return;
            var clientBounds = client.Bounds;
            var childDpi = Native.GetDpiForWindow(current.Owner);
            if (needsAnchor || childDpi == 0 || current.Dpi != childDpi
                || clientBounds.Width != current.AnchorOwner.Width || clientBounds.Height != current.AnchorOwner.Height)
            {
                needsAnchor = true;
                HidePopup(current);
                if (!invalidationNotified) { invalidationNotified = true; Notify(false); }
                if ((pending & (MoveEnd | MinimizeEnd)) != 0) Notify(true);
                return;
            }
            lock (gate)
            {
                if (disposed != 0 || binding.Generation != current.Generation) return;
                if (current.DesiredVisible && !Native.IsWindowVisible(current.Popup))
                    Native.SetWindowPos(current.Popup, IntPtr.Zero, current.AnchorPopup.X, current.AnchorPopup.Y,
                        current.AnchorPopup.Width, current.AnchorPopup.Height,
                        Native.NoActivate | Native.NoZOrder | Native.AsyncPosition | Native.ShowWindow);
                // The parent already moves its child. Touching a visible child's
                // Z order here competes with sibling surfaces and native overlays.
            }
            if ((pending & (MoveEnd | MinimizeEnd)) != 0) Notify(true);
            return;
        }
        if (!Native.GetWindowRect(current.Owner, out var owner)) return;
        var ownerBounds = owner.Bounds;
        var dpi = Native.GetDpiForWindow(current.Owner);
        Bounds target = default;
        var invalid = needsAnchor || dpi == 0 || current.Dpi != dpi
            || !TryTranslate(current.AnchorOwner, current.AnchorPopup, ownerBounds, out target);
        if (invalid)
        {
            needsAnchor = true;
            HidePopup(current);
            if (!invalidationNotified) { invalidationNotified = true; Notify(false); }
        }
        else
        {
            // Re-read visibility/generation before native work: a queued event never
            // overrides a newer tab hide or a replacement anchor.
            lock (gate)
            {
                if (disposed != 0 || binding.Generation != current.Generation) return;
                if (Native.GetWindowRect(current.Popup, out var before))
                {
                    var show = !Native.IsWindowVisible(current.Popup);
                    if (before.Left != target.X || before.Top != target.Y || show)
                        Native.SetWindowPos(current.Popup, IntPtr.Zero, target.X, target.Y, 0, 0,
                            Native.NoSize | Native.NoActivate | Native.NoZOrder | Native.AsyncPosition
                            | (show ? Native.ShowWindow : 0));
                }
            }
        }
        if ((pending & (MoveEnd | MinimizeEnd)) != 0) Notify(true);
    }

    private static bool HostedChild(Binding current) =>
        (Native.GetWindowLong(current.Popup, Native.GwlStyle) & Native.WsChild) != 0
        && Native.GetParent(current.Popup) == current.Owner;

    private bool IdentityMatches(Binding current) =>
        Native.IsWindow(current.Owner) && Native.IsWindow(current.Popup)
        && Native.GetWindowThreadProcessId(current.Owner, out var process) == current.Thread && process == current.Process
        && Native.GetWindowThreadProcessId(current.Popup, out var popupProcess) == current.PopupThread
        && popupProcess == (uint)Environment.ProcessId;

    private void HidePopup(Binding current)
    {
        lock (gate)
        {
            if (disposed != 0 || binding.Generation != current.Generation || !Native.IsWindow(current.Popup)) return;
            if (Native.GetWindowThreadProcessId(current.Popup, out var process) != current.PopupThread
                || process != (uint)Environment.ProcessId) return;
            if (Native.IsWindowVisible(current.Popup))
                Native.SetWindowPos(current.Popup, IntPtr.Zero, 0, 0, 0, 0,
                    Native.NoMove | Native.NoSize | Native.NoActivate | Native.NoZOrder | Native.AsyncPosition | Native.HideWindow);
        }
    }

    private void Notify(bool final)
    {
        if (Volatile.Read(ref disposed) != 0) return;
        try { requestLayoutAfterMove?.Invoke(final); }
        catch { /* Never let consumer code unwind through a Win32 callback/message loop. */ }
    }

    private void Unhook()
    {
        foreach (var hook in hooks) { thread.Unregister(hook); Native.UnhookWinEvent(hook); }
        hooks.Clear();
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref disposed, 1) != 0) return;
        Volatile.Write(ref inMoveSize, 0);
        Interlocked.Decrement(ref activeTrackers);
        Schedule(); // Unhook and final worker shutdown happen on the registration thread.
    }

    private sealed class EventThread
    {
        private const uint ProcessTracker = 0x8000 + 117;
        private readonly ConcurrentDictionary<long, NativeOwnerTracker> trackers = new();
        private readonly Dictionary<IntPtr, NativeOwnerTracker> hookOwners = [];
        private readonly ManualResetEventSlim ready = new(false);
        private readonly Thread worker;
        internal readonly Native.WinEventProc Callback;
        private uint threadId;
        private bool stopping;
        internal EventThread()
        {
            Callback = HandleEvent;
            worker = new Thread(Run) { IsBackground = true, Name = "OieOwnerEvents" };
            worker.Start();
            if (!ready.Wait(TimeSpan.FromSeconds(5)) || threadId == 0)
                throw new InvalidOperationException("The owner event thread could not start.");
        }
        internal void Add(NativeOwnerTracker tracker) => trackers.TryAdd(tracker.id, tracker);
        internal void Register(IntPtr hook, NativeOwnerTracker tracker) => hookOwners[hook] = tracker;
        internal void Unregister(IntPtr hook) => hookOwners.Remove(hook);
        internal bool Post(long id) => threadId != 0 && Native.PostThreadMessage(threadId, ProcessTracker, new IntPtr(id), IntPtr.Zero);
        internal void Remove(NativeOwnerTracker tracker)
        {
            lock (SharedGate)
            {
                trackers.TryRemove(tracker.id, out _);
                if (!trackers.IsEmpty) return;
                stopping = true;
                if (ReferenceEquals(shared, this)) shared = null;
            }
        }
        private void HandleEvent(IntPtr hook, uint eventType, IntPtr hwnd, int objectId, int childId, uint eventThread, uint timestamp)
        {
            try
            {
                if (hookOwners.TryGetValue(hook, out var tracker)) tracker.OnEvent(eventType, hwnd, objectId, childId, eventThread);
            }
            catch { /* OS callbacks must never propagate managed exceptions. */ }
        }
        private void Run()
        {
            Interlocked.Increment(ref runningThreads);
            var callbackRoot = GCHandle.Alloc(Callback);
            try
            {
                Native.PeekMessage(out _, IntPtr.Zero, 0, 0, 0); // Create the thread message queue before publishing its ID.
                threadId = Native.GetCurrentThreadId();
                ready.Set();
                while (!stopping && Native.GetMessage(out var message, IntPtr.Zero, 0, 0) > 0)
                {
                    if (message.Id == ProcessTracker && trackers.TryGetValue(message.WParam.ToInt64(), out var tracker))
                    {
                        try { tracker.Process(); }
                        catch { Interlocked.Exchange(ref tracker.queued, 0); /* A later event can recover. */ }
                    }
                    else { Native.TranslateMessage(ref message); Native.DispatchMessage(ref message); }
                }
            }
            finally
            {
                foreach (var tracker in trackers.Values) tracker.Unhook();
                trackers.Clear();
                callbackRoot.Free();
                threadId = 0;
                ready.Set();
                Interlocked.Decrement(ref runningThreads);
            }
        }
    }

    private readonly struct PhysicalCoordinates : IDisposable
    {
        private readonly IntPtr previous;
        public PhysicalCoordinates() => previous = Native.SetThreadDpiAwarenessContext(new IntPtr(-4));
        public void Dispose() { if (previous != IntPtr.Zero) Native.SetThreadDpiAwarenessContext(previous); }
    }

    private static class Native
    {
        internal const uint EventMoveStart = 0x000A, EventMoveEnd = 0x000B,
            EventMinimizeStart = 0x0016, EventMinimizeEnd = 0x0017,
            EventDestroy = 0x8001, EventShow = 0x8002, EventHide = 0x8003, EventLocation = 0x800B;
        internal const uint NoSize = 0x0001, NoMove = 0x0002, NoZOrder = 0x0004,
            NoActivate = 0x0010, ShowWindow = 0x0040, HideWindow = 0x0080, AsyncPosition = 0x4000;
        internal const int GwlStyle = -16, WsChild = 0x40000000;
        internal delegate void WinEventProc(IntPtr hook, uint eventType, IntPtr hwnd, int objectId, int childId, uint eventThread, uint timestamp);
        [StructLayout(LayoutKind.Sequential)] internal struct Rect
        {
            internal int Left, Top, Right, Bottom;
            internal readonly Bounds Bounds => new(Left, Top, Right - Left, Bottom - Top);
        }
        [StructLayout(LayoutKind.Sequential)] internal struct Message
        {
            internal IntPtr Hwnd;
            internal uint Id;
            internal IntPtr WParam, LParam;
            internal uint Time;
            internal int X, Y;
            internal uint Private;
        }
        [DllImport("user32.dll")] internal static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr module, WinEventProc callback, uint process, uint thread, uint flags);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool UnhookWinEvent(IntPtr hook);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool IsIconic(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint process);
        [DllImport("user32.dll")] internal static extern uint GetDpiForWindow(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool GetClientRect(IntPtr hwnd, out Rect rect);
        [DllImport("user32.dll")] internal static extern IntPtr GetParent(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern int GetWindowLong(IntPtr hwnd, int index);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int width, int height, uint flags);
        [DllImport("user32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool PostThreadMessage(uint thread, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] internal static extern int GetMessage(out Message message, IntPtr hwnd, uint min, uint max);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool PeekMessage(out Message message, IntPtr hwnd, uint min, uint max, uint remove);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool TranslateMessage(ref Message message);
        [DllImport("user32.dll")] internal static extern IntPtr DispatchMessage(ref Message message);
        [DllImport("kernel32.dll")] internal static extern uint GetCurrentThreadId();
    }
}
