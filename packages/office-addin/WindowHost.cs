using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OfficeInEditor.OfficeAddin
{
    internal static class NativeWindow
    {
        internal const int SwpNoZOrder = 0x0004;
        internal const int SwpNoActivate = 0x0010;
        internal const uint SwpNoOwnerZOrder = 0x0200;
        internal const uint SwShowNoActivate = 4;
        internal const uint GaRoot = 2;

        [DllImport("user32.dll")] internal static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern IntPtr GetParent(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] internal static extern bool ShowWindow(IntPtr hwnd, uint command);
        [DllImport("user32.dll")] internal static extern bool MoveWindow(IntPtr hwnd, int x, int y, int width, int height, bool repaint);
        [DllImport("user32.dll")] internal static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
        [DllImport("user32.dll")] internal static extern bool EnumWindows(EnumWindowsProc callback, IntPtr data);
        internal delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr data);

        [StructLayout(LayoutKind.Sequential)]
        internal struct Rect { public int Left, Top, Right, Bottom; }
    }

    internal sealed class WindowHost
    {
        private readonly object sync = new object();
        private readonly Dictionary<string, Placement> placements = new Dictionary<string, Placement>(StringComparer.Ordinal);

        private sealed class Placement
        {
            internal IntPtr Window;
            internal int X, Y, Width, Height;
            internal bool Attached;
        }

        internal object DescribeWindow(dynamic application)
        {
            IntPtr hwnd = RootWindow(application);
            return Json.Object(
                "hwnd", hwnd.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture),
                "processId", ProcessId(hwnd),
                "attached", IsAttached(hwnd));
        }

        internal object Attach(object parameters, dynamic application)
        {
            var bounds = Bounds(parameters);
            IntPtr hwnd = RootWindow(application);
            if (hwnd == IntPtr.Zero || !NativeWindow.IsWindow(hwnd)) throw new BridgeException("OFFICE_WINDOW_UNAVAILABLE");
            lock (sync)
            {
                if (!placements.TryGetValue(Key(hwnd), out var placement))
                {
                    NativeWindow.Rect original;
                    if (!NativeWindow.GetWindowRect(hwnd, out original)) throw new BridgeException("OFFICE_WINDOW_UNAVAILABLE");
                    placement = new Placement
                    {
                        Window = hwnd,
                        X = original.Left,
                        Y = original.Top,
                        Width = Math.Max(1, original.Right - original.Left),
                        Height = Math.Max(1, original.Bottom - original.Top)
                    };
                    placements[Key(hwnd)] = placement;
                }
                Place(hwnd, bounds);
                placement.Attached = true;
            }
            return Json.Object("attached", true, "hwnd", hwnd.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        internal object Layout(object parameters, dynamic application)
        {
            var bounds = Bounds(parameters);
            IntPtr hwnd = RootWindow(application);
            lock (sync)
            {
                if (!placements.TryGetValue(Key(hwnd), out var placement) || !placement.Attached)
                    throw new BridgeException("OFFICE_NOT_ATTACHED");
                Place(hwnd, bounds);
            }
            return Json.Object("attached", true);
        }

        internal object Detach(dynamic application)
        {
            IntPtr hwnd = RootWindow(application);
            lock (sync)
            {
                if (!placements.TryGetValue(Key(hwnd), out var placement) || !placement.Attached)
                    return Json.Object("attached", false);
                Restore(placement);
            }
            return Json.Object("attached", false);
        }

        internal void RestoreAll()
        {
            lock (sync)
            {
                foreach (var placement in new List<Placement>(placements.Values)) Restore(placement);
                placements.Clear();
            }
        }

        private bool IsAttached(IntPtr hwnd)
        {
            return placements.TryGetValue(Key(hwnd), out var placement) && placement.Attached && NativeWindow.IsWindow(hwnd);
        }

        private static void Place(IntPtr hwnd, HostBounds bounds)
        {
            NativeWindow.ShowWindow(hwnd, NativeWindow.SwShowNoActivate);
            NativeWindow.MoveWindow(hwnd, bounds.X, bounds.Y, bounds.Width, bounds.Height, true);
            NativeWindow.SetWindowPos(hwnd, IntPtr.Zero, bounds.X, bounds.Y, bounds.Width, bounds.Height,
                NativeWindow.SwpNoZOrder | NativeWindow.SwpNoActivate | NativeWindow.SwpNoOwnerZOrder);
        }

        private void Restore(Placement placement)
        {
            if (!NativeWindow.IsWindow(placement.Window)) { placement.Attached = false; return; }
            NativeWindow.MoveWindow(placement.Window, placement.X, placement.Y, placement.Width, placement.Height, true);
            NativeWindow.ShowWindow(placement.Window, NativeWindow.SwShowNoActivate);
            placement.Attached = false;
        }

        private static IntPtr RootWindow(dynamic application)
        {
            try
            {
                object raw = application.Hwnd;
                IntPtr hwnd = ToHandle(raw);
                if (hwnd != IntPtr.Zero && NativeWindow.IsWindow(hwnd))
                    return NativeWindow.GetAncestor(hwnd, NativeWindow.GaRoot);
            }
            catch { }
            uint pid = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            IntPtr found = IntPtr.Zero;
            NativeWindow.EnumWindows((hwnd, _) =>
            {
                NativeWindow.GetWindowThreadProcessId(hwnd, out uint windowPid);
                if (windowPid == pid && NativeWindow.IsWindow(hwnd) && NativeWindow.GetParent(hwnd) == IntPtr.Zero)
                {
                    found = hwnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        private static int ProcessId(IntPtr hwnd)
        {
            NativeWindow.GetWindowThreadProcessId(hwnd, out uint pid);
            return unchecked((int)pid);
        }

        private static string Key(IntPtr hwnd) => hwnd.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture);

        private static HostBounds Bounds(object parameters)
        {
            int x = Integer(parameters, "x");
            int y = Integer(parameters, "y");
            int width = Integer(parameters, "width");
            int height = Integer(parameters, "height");
            if (width < 32 || height < 32) throw new BridgeException("INVALID_HOST_WINDOW");
            return new HostBounds { X = x, Y = y, Width = width, Height = height };
        }

        private static int Integer(object parameters, string name)
        {
            object raw = Json.Get(parameters, name);
            try { return Convert.ToInt32(raw, System.Globalization.CultureInfo.InvariantCulture); }
            catch { throw new BridgeException("INVALID_HOST_WINDOW"); }
        }

        private struct HostBounds { internal int X, Y, Width, Height; }

        private static IntPtr ToHandle(object raw)
        {
            if (raw == null) return IntPtr.Zero;
            if (raw is IntPtr pointer) return pointer;
            try { return new IntPtr(Convert.ToInt64(raw, System.Globalization.CultureInfo.InvariantCulture)); }
            catch { return IntPtr.Zero; }
        }
    }
}
