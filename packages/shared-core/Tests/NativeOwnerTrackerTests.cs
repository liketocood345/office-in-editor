using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using OfficeInEditor.Core;

internal static class NativeOwnerTrackerTests
{
    internal static void Run(Action<bool, string> check)
    {
        var owner = new NativeOwnerTracker.Bounds(-1000, 100, 800, 600);
        var popup = new NativeOwnerTracker.Bounds(-880, 210, 600, 450);
        check(NativeOwnerTracker.TryTranslate(owner, popup, new(-700, -200, 800, 600), out var moved)
            && moved == new NativeOwnerTracker.Bounds(-580, -90, 600, 450),
            "owner translation follows both coordinates without changing popup size");

        var exact = true;
        for (var i = 0; i < 1000; i++)
        {
            var dx = i % 101 - 50;
            var dy = 50 - i % 101;
            exact &= NativeOwnerTracker.TryTranslate(owner, popup, new(owner.X + dx, owner.Y + dy, 800, 600), out moved)
                && moved == new NativeOwnerTracker.Bounds(popup.X + dx, popup.Y + dy, 600, 450);
        }
        check(exact && NativeOwnerTracker.TryTranslate(owner, popup, owner, out moved) && moved == popup,
            "repeated owner moves are computed from absolute anchors without cumulative drift");
        check(!NativeOwnerTracker.TryTranslate(owner, popup, new(-1000, 100, 801, 600), out _)
            && !NativeOwnerTracker.TryTranslate(owner, popup, new(-1000, 100, 800, 601), out _),
            "owner size changes require recalibration rather than guessed popup dimensions");
        check(!NativeOwnerTracker.TryTranslate(new(0, 0, 0, 600), popup, new(1, 1, 0, 600), out _)
            && !NativeOwnerTracker.TryTranslate(owner, new(0, 0, 0, 50), owner, out _),
            "invalid owner and popup dimensions cannot approve a move");
        check(!NativeOwnerTracker.TryTranslate(new(0, 0, 800, 600), new(int.MaxValue - 20, 0, 10, 10), new(50, 0, 800, 600), out _)
            && !NativeOwnerTracker.TryTranslate(new(0, 0, 800, 600), new(int.MinValue + 20, 0, 10, 10), new(-50, 0, 800, 600), out _)
            && !NativeOwnerTracker.TryTranslate(new(0, 0, 800, 600), new(int.MaxValue - 20, 0, 40, 10), new(0, 0, 800, 600), out _),
            "translation rejects coordinate and extent overflow");

        var ownerHwnd = new IntPtr(101);
        check(NativeOwnerTracker.AcceptsEvent(ownerHwnd, 7, ownerHwnd, 0, 0, 7),
            "only the bound owner window event is accepted");
        check(!NativeOwnerTracker.AcceptsEvent(ownerHwnd, 7, new IntPtr(102), 0, 0, 7)
            && !NativeOwnerTracker.AcceptsEvent(ownerHwnd, 7, ownerHwnd, -4, 0, 7)
            && !NativeOwnerTracker.AcceptsEvent(ownerHwnd, 7, ownerHwnd, 0, 1, 7)
            && !NativeOwnerTracker.AcceptsEvent(ownerHwnd, 7, ownerHwnd, 0, 0, 8)
            && !NativeOwnerTracker.AcceptsEvent(IntPtr.Zero, 7, IntPtr.Zero, 0, 0, 7),
            "other windows, caret/client objects, child objects and stale owner threads are ignored");
    }

    // Explicit opt-in only. These windows are synthetic User32 fixtures: no Office,
    // editor process, COM object, installed component or document is opened.
    internal static void RunNative(Action<bool, string> check)
    {
        if (!OperatingSystem.IsWindows()) return;
        var before = NativeOwnerTracker.ActiveTrackerCount;
        using var windows = new SyntheticWindows();
        var anchorOwner = windows.OwnerBounds;
        var anchorPopup = windows.PopupBounds;
        var invalidations = 0;
        var finalCalibrations = 0;
        var tracker = NativeOwnerTracker.Create(windows.Owner, windows.Popup, final => {
            Interlocked.Increment(ref invalidations);
            if (final) Interlocked.Increment(ref finalCalibrations);
        });
        try
        {
            tracker.Update(anchorOwner, anchorPopup, true);
            check(Eventually(() => NativeOwnerTracker.ActiveTrackerCount == before + 1 && NativeOwnerTracker.EventThreadRunning),
                "native owner tracking starts one shared event pump for a synthetic owner");
            windows.MoveOwner(anchorOwner.X + 135, anchorOwner.Y + 71);
            var expected = new NativeOwnerTracker.Bounds(anchorPopup.X + 135, anchorPopup.Y + 71, anchorPopup.Width, anchorPopup.Height);
            check(Eventually(() => windows.PopupBounds == expected),
                "native location events move a popup when the owner moves without any resize");
            windows.MoveOwner(anchorOwner.X - 27, anchorOwner.Y + 43);
            expected = new(anchorPopup.X - 27, anchorPopup.Y + 43, anchorPopup.Width, anchorPopup.Height);
            check(Eventually(() => windows.PopupBounds == expected),
                "a second native move uses the original anchor and keeps popup dimensions");
            windows.MoveOwner(anchorOwner.X, anchorOwner.Y);
            check(Eventually(() => windows.PopupBounds == anchorPopup),
                "returning the owner to its anchor restores the exact popup rectangle");

            windows.MinimizeOwner();
            check(Eventually(() => !windows.PopupVisible), "minimizing a visible owner hides its popup");
            windows.RestoreOwner();
            check(Eventually(() => windows.PopupVisible && windows.PopupBounds == anchorPopup),
                "restoring a visible owner restores its popup at the calibrated rectangle without a new anchor");

            tracker.InvalidateAnchor();
            check(Eventually(() => !windows.PopupVisible), "invalidating a visible anchor hides the popup without releasing its owner tracker");
            var beforeInvalidMoveEnd = Volatile.Read(ref finalCalibrations);
            windows.MoveOwner(anchorOwner.X + 20, anchorOwner.Y + 10);
            windows.EmitMoveEnd();
            check(Eventually(() => Volatile.Read(ref finalCalibrations) > beforeInvalidMoveEnd && !windows.PopupVisible),
                "an invalid anchor still requests final calibration through a real move-end hook without revealing its popup");
            var beforeInvalidRestore = Volatile.Read(ref finalCalibrations);
            windows.MinimizeOwner();
            windows.RestoreOwner();
            check(Eventually(() => Volatile.Read(ref finalCalibrations) > beforeInvalidRestore && !windows.PopupVisible),
                "restoring an invalidated visible tab requests another calibration while leaving its popup hidden");
            check(StaysHidden(windows), "late location and restore events cannot reveal a popup whose anchor was invalidated");
            anchorOwner = windows.OwnerBounds;
            windows.PlacePopup(new(anchorOwner.X + 80, anchorOwner.Y + 90, anchorPopup.Width, anchorPopup.Height));
            anchorPopup = windows.PopupBounds;
            tracker.Update(anchorOwner, anchorPopup, true);
            windows.MoveOwner(anchorOwner.X + 10, anchorOwner.Y + 15);
            expected = new(anchorPopup.X + 10, anchorPopup.Y + 15, anchorPopup.Width, anchorPopup.Height);
            check(Eventually(() => windows.PopupVisible && windows.PopupBounds == expected),
                "only a fresh calibrated anchor restores visibility and reuses the existing owner hooks");

            tracker.Update(windows.OwnerBounds, windows.PopupBounds, false);
            check(Eventually(() => !windows.PopupVisible), "a hidden tab hides its synthetic popup");
            tracker.InvalidateAnchor();
            var beforeHiddenRecovery = Volatile.Read(ref finalCalibrations);
            windows.MoveOwner(anchorOwner.X + 30, anchorOwner.Y + 30);
            windows.EmitMoveEnd();
            windows.MinimizeOwner();
            windows.RestoreOwner();
            check(StaysHidden(windows),
                "owner movement and minimize/restore cannot reveal a tab whose desired visibility is false");
            check(Volatile.Read(ref finalCalibrations) == beforeHiddenRecovery,
                "a hidden tab with an invalid anchor never requests calibration for move-end or restore events");

            anchorOwner = windows.OwnerBounds;
            windows.PlacePopup(new(anchorOwner.X + 80, anchorOwner.Y + 90, anchorPopup.Width, anchorPopup.Height));
            anchorPopup = windows.PopupBounds;
            tracker.Update(anchorOwner, anchorPopup, true);
            check(Eventually(() => windows.PopupVisible), "a new calibrated visible anchor permits popup visibility");
            var beforeResizeInvalidations = Volatile.Read(ref invalidations);
            windows.ResizeOwner(anchorOwner.Width + 60, anchorOwner.Height + 20);
            check(Eventually(() => !windows.PopupVisible && Volatile.Read(ref invalidations) > beforeResizeInvalidations),
                "owner resize hides the uncalibrated popup and requests layout without resizing Office content");

            using var secondWindows = new SyntheticWindows();
            var secondOwner = secondWindows.OwnerBounds;
            var secondPopup = secondWindows.PopupBounds;
            using var secondTracker = NativeOwnerTracker.Create(secondWindows.Owner, secondWindows.Popup);
            secondTracker.Update(secondOwner, secondPopup, true);
            secondWindows.MoveOwner(secondOwner.X + 45, secondOwner.Y + 70);
            var secondExpected = new NativeOwnerTracker.Bounds(secondPopup.X + 45, secondPopup.Y + 70, secondPopup.Width, secondPopup.Height);
            check(Eventually(() => NativeOwnerTracker.ActiveTrackerCount == before + 2 && secondWindows.PopupBounds == secondExpected),
                "a second live tracker independently follows its synthetic owner on the shared event thread");
            tracker.Dispose();
            check(Eventually(() => NativeOwnerTracker.ActiveTrackerCount == before + 1 && NativeOwnerTracker.EventThreadRunning),
                "disposing one of two trackers keeps the shared event thread alive");
            secondWindows.MoveOwner(secondOwner.X - 25, secondOwner.Y + 15);
            secondExpected = new(secondPopup.X - 25, secondPopup.Y + 15, secondPopup.Width, secondPopup.Height);
            check(Eventually(() => secondWindows.PopupBounds == secondExpected),
                "the remaining tracker continues following after the first tracker is disposed");
        }
        finally { tracker.Dispose(); }
        check(Eventually(() => NativeOwnerTracker.ActiveTrackerCount == before
            && (before != 0 || !NativeOwnerTracker.EventThreadRunning)),
            "disposing the last tracker releases hooks and stops the shared event thread");
        var detached = windows.PopupBounds;
        windows.MoveOwner(windows.OwnerBounds.X + 25, windows.OwnerBounds.Y + 25);
        windows.Flush();
        check(windows.PopupBounds == detached, "disposed tracking cannot move the former popup");
    }

    private static bool Eventually(Func<bool> condition)
    {
        var deadline = Environment.TickCount64 + 4000;
        do
        {
            if (condition()) return true;
            Thread.Sleep(10);
        } while (Environment.TickCount64 < deadline);
        return condition();
    }

    private static bool StaysHidden(SyntheticWindows windows)
    {
        // Observe after the owner operations as well as after queued native events;
        // an immediately true visibility sample would miss a late restore bug.
        var deadline = Environment.TickCount64 + 150;
        do
        {
            if (windows.PopupVisible) return false;
            Thread.Sleep(10);
        } while (Environment.TickCount64 < deadline);
        return !windows.PopupVisible;
    }

    private sealed class SyntheticWindows : IDisposable
    {
        private const uint ExecuteMessage = 0x8000 + 73;
        private const uint SwpNoZOrder = 0x0004, SwpNoActivate = 0x0010;
        private readonly ConcurrentQueue<Action> commands = new();
        private readonly TaskCompletionSource ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly Thread thread;
        private readonly WindowProc windowProc = DefWindowProc;
        private uint threadId;
        private string className = "";
        private bool disposed;
        internal IntPtr Owner { get; private set; }
        internal IntPtr Popup { get; private set; }
        internal NativeOwnerTracker.Bounds OwnerBounds => Invoke(() => ReadBounds(Owner));
        internal NativeOwnerTracker.Bounds PopupBounds => Invoke(() => ReadBounds(Popup));
        internal bool PopupVisible => Invoke(() => IsWindowVisible(Popup));

        internal SyntheticWindows()
        {
            thread = new Thread(Run) { IsBackground = true, Name = "OieSyntheticOwnerTests" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!ready.Task.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException("Synthetic window fixture did not start.");
            ready.Task.GetAwaiter().GetResult();
        }
        private void Run()
        {
            threadId = GetCurrentThreadId();
            var priorDpi = SetThreadDpiAwarenessContext(new IntPtr(-4));
            try
            {
                className = "OieSyntheticOwner_" + Guid.NewGuid().ToString("N");
                var cls = new WindowClass { Size = (uint)Marshal.SizeOf<WindowClass>(), Proc = windowProc, Instance = GetModuleHandle(null), Name = className };
                if (RegisterClassEx(ref cls) == 0) throw new InvalidOperationException("Synthetic window class registration failed.");
                Owner = CreateWindowEx(0x08000000, className, "Synthetic owner", 0x00CF0000, 100, 100, 640, 480, IntPtr.Zero, IntPtr.Zero, cls.Instance, IntPtr.Zero);
                Popup = CreateWindowEx(0x08000080, className, "Synthetic popup", 0x80000000, 180, 190, 400, 300, Owner, IntPtr.Zero, cls.Instance, IntPtr.Zero);
                if (Owner == IntPtr.Zero || Popup == IntPtr.Zero) throw new InvalidOperationException("Synthetic windows could not be created.");
                ShowWindow(Owner, 4); ShowWindow(Popup, 4);
                ready.TrySetResult();
                while (GetMessage(out var message, IntPtr.Zero, 0, 0) > 0)
                {
                    if (message.Id == ExecuteMessage)
                    {
                        while (commands.TryDequeue(out var command)) command();
                    }
                    else { TranslateMessage(ref message); DispatchMessage(ref message); }
                }
            }
            catch (Exception error) { ready.TrySetException(error); }
            finally
            {
                if (Popup != IntPtr.Zero) DestroyWindow(Popup);
                if (Owner != IntPtr.Zero) DestroyWindow(Owner);
                if (className.Length > 0) UnregisterClass(className, GetModuleHandle(null));
                if (priorDpi != IntPtr.Zero) SetThreadDpiAwarenessContext(priorDpi);
            }
        }
        private T Invoke<T>(Func<T> operation)
        {
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            commands.Enqueue(() => { try { completion.SetResult(operation()); } catch (Exception error) { completion.SetException(error); } });
            if (!PostThreadMessage(threadId, ExecuteMessage, IntPtr.Zero, IntPtr.Zero)) throw new InvalidOperationException("Synthetic event thread is unavailable.");
            if (!completion.Task.Wait(TimeSpan.FromSeconds(4))) throw new TimeoutException("Synthetic event thread did not respond.");
            return completion.Task.GetAwaiter().GetResult();
        }
        internal void MoveOwner(int x, int y) => Invoke(() => SetWindowPos(Owner, IntPtr.Zero, x, y, 0, 0, 0x0001 | SwpNoZOrder | SwpNoActivate));
        internal void ResizeOwner(int width, int height) => Invoke(() => SetWindowPos(Owner, IntPtr.Zero, 0, 0, width, height, 0x0002 | SwpNoZOrder | SwpNoActivate));
        internal void PlacePopup(NativeOwnerTracker.Bounds bounds) => Invoke(() => SetWindowPos(Popup, IntPtr.Zero, bounds.X, bounds.Y, bounds.Width, bounds.Height, SwpNoZOrder | SwpNoActivate));
        internal void MinimizeOwner() => Invoke(() => ShowWindow(Owner, 6));
        internal void RestoreOwner() => Invoke(() => ShowWindow(Owner, 4));
        // Deliver the synthetic terminal event through User32 on the actual
        // owner's thread; this exercises the WinEvent hook and strict TID filter.
        internal void EmitMoveEnd() => Invoke(() => { NotifyWinEvent(0x000B, Owner, 0, 0); return true; });
        internal void Flush() => Invoke(() => true);
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            if (thread.IsAlive) PostThreadMessage(threadId, 0x0012, IntPtr.Zero, IntPtr.Zero);
            if (!thread.Join(TimeSpan.FromSeconds(4))) throw new TimeoutException("Synthetic event thread did not stop.");
            GC.KeepAlive(windowProc);
        }
        private static NativeOwnerTracker.Bounds ReadBounds(IntPtr window)
        {
            if (!GetWindowRect(window, out var rect)) throw new InvalidOperationException("Synthetic window rectangle is unavailable.");
            return new(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
        }
    }

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate IntPtr WindowProc(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WindowClass
    {
        internal uint Size, Style;
        internal WindowProc Proc;
        internal int ClassExtra, WindowExtra;
        internal IntPtr Instance, Icon, Cursor, Background;
        internal string? Menu;
        internal string Name;
        internal IntPtr SmallIcon;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct Message { internal IntPtr Window; internal uint Id; internal UIntPtr WParam; internal IntPtr LParam; internal uint Time; internal int X, Y; internal uint Private; }
    [StructLayout(LayoutKind.Sequential)]
    private struct Rect { internal int Left, Top, Right, Bottom; }
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string? module);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern ushort RegisterClassEx(ref WindowClass windowClass);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool UnregisterClass(string name, IntPtr instance);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateWindowEx(uint exStyle, string cls, string title, uint style, int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);
    [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window, int command);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern void NotifyWinEvent(uint eventType, IntPtr window, int objectId, int childId);
    [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] private static extern bool PostThreadMessage(uint thread, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern int GetMessage(out Message message, IntPtr window, uint min, uint max);
    [DllImport("user32.dll")] private static extern bool TranslateMessage(ref Message message);
    [DllImport("user32.dll")] private static extern IntPtr DispatchMessage(ref Message message);
    [DllImport("user32.dll")] private static extern IntPtr DefWindowProc(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
}
