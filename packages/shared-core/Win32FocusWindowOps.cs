using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

/// <summary>Win32 window operations for one focus session. Runs on the focus dispatcher STA thread only.</summary>
internal sealed class Win32FocusWindowOps : IFocusWindowOps
{
    const uint GaRootOwner = 3;
    const uint GwOwner = 4;
    const int SmXVirtualScreen = 76, SmYVirtualScreen = 77, SmCxVirtualScreen = 78, SmCyVirtualScreen = 79;
    const uint SwpNoSize = 0x0001, SwpNoZOrder = 0x0004, SwpNoActivate = 0x0010;
    const int SwShowNoActivate = 4, SwShowMinimized = 2, SwShowNormal = 1;
    const int DwmwaExtendedFrameBounds = 9;
    const int ResolveIntervalMs = 300;
    const uint WmMouseWheel = 0x020A, WmMouseHWheel = 0x020E;
    const int MkShift = 0x0004, MkControl = 0x0008;
    const uint CwpSkipInvisible = 0x0001, CwpSkipDisabled = 0x0002, CwpSkipTransparent = 0x0004;

    readonly IntPtr office;
    readonly uint officePid;
    string surfaceToken;
    readonly FocusRecovery recovery;
    readonly WindowPlacement original;
    IntPtr editorRoot;
    uint editorPid;
    FocusContract.SurfaceSize size;
    long nextResolve;

    public Win32FocusWindowOps(IntPtr office, uint officePid, string surfaceToken, FocusContract.SurfaceSize size, FocusRecovery recovery)
    {
        this.office = office; this.officePid = officePid; this.surfaceToken = surfaceToken; this.size = size; this.recovery = recovery;
        // A frame moved out of view at cold start keeps the placement it had before that move as its original.
        if (PreParked.TryRemove(office.ToInt64(), out var saved)) original = saved;
        else
        {
            original = new WindowPlacement { Length = Marshal.SizeOf<WindowPlacement>() };
            if (!GetWindowPlacement(office, ref original)) throw new CoreFault("office_window_unavailable", "The Office window is unavailable.");
            if (original.ShowCmd == SwShowMinimized) original.ShowCmd = SwShowNormal;
        }
        recovery.Remember(office, officePid, original.NormalLeft, original.NormalTop, original.NormalRight, original.NormalBottom, original.ShowCmd);
    }

    static readonly System.Collections.Concurrent.ConcurrentDictionary<long, WindowPlacement> PreParked = new();

    /// <summary>
    /// Moves the main frame of a freshly started Office process out of view as soon as it appears, before the add-in
    /// reports the document window. Returns the frame, or zero while it has not appeared yet. Its placement is kept
    /// as the original, and crash recovery knows it.
    /// </summary>
    internal static IntPtr PreParkNewFrame(int pid, string application, FocusRecovery recovery)
    {
        var found = IntPtr.Zero;
        NativeWindow.EnumWindows((window, _) =>
        {
            if (Pid(window) != pid || !NativeWindow.IsWindowVisible(window) || IsIconic(window) || !IsOfficeFrame(window, pid, application)) return true;
            found = window;
            return false;
        }, IntPtr.Zero);
        if (found == IntPtr.Zero) return IntPtr.Zero;
        var placement = new WindowPlacement { Length = Marshal.SizeOf<WindowPlacement>() };
        if (!GetWindowPlacement(found, ref placement)) return IntPtr.Zero;
        if (placement.ShowCmd == SwShowMinimized) placement.ShowCmd = SwShowNormal;
        if (!PreParked.TryAdd(found.ToInt64(), placement)) return found;
        recovery.Remember(found, (uint)pid, placement.NormalLeft, placement.NormalTop, placement.NormalRight, placement.NormalBottom, placement.ShowCmd);
        using var dpi = WebviewSurface.EnterPhysicalCoordinates();
        NativeWindow.GetWindowRect(found, out var rect);
        // SetWindowPos, not SetWindowPlacement: the latter pulls a window back onto a monitor.
        var left = NativeWindow.GetSystemMetrics(SmXVirtualScreen) - (rect.Right - rect.Left) - 128;
        NativeWindow.SetWindowPos(found, IntPtr.Zero, left, rect.Top, 0, 0, SwpNoSize | SwpNoZOrder | SwpNoActivate);
        DebugNdjson.Write("F", "FocusWindow.PrePark", "office_frame_preparked", null);
        return found;
    }

    /// <summary>Puts a pre-parked frame back where it was when no session takes it (open failed or another window won).</summary>
    internal static void ReleasePrePark(IntPtr window, FocusRecovery recovery)
    {
        if (!PreParked.TryRemove(window.ToInt64(), out var placement)) return;
        if (NativeWindow.IsWindow(window))
        {
            placement.Length = Marshal.SizeOf<WindowPlacement>();
            if (placement.ShowCmd != 3) placement.ShowCmd = SwShowNoActivate;
            SetWindowPlacement(window, ref placement);
        }
        recovery.Forget(window);
    }

    /// <summary>Checks that <paramref name="hwnd"/> is a top-level frame of <paramref name="application"/> owned by <paramref name="pid"/>.</summary>
    public static bool IsOfficeFrame(IntPtr hwnd, int pid, string application)
    {
        if (hwnd == IntPtr.Zero || !NativeWindow.IsWindow(hwnd) || NativeWindow.GetAncestor(hwnd, GaRootOwner) != hwnd) return false;
        NativeWindow.GetWindowThreadProcessId(hwnd, out var owner);
        if (owner != pid) return false;
        var name = new StringBuilder(64);
        NativeWindow.GetClassName(hwnd, name, name.Capacity);
        return (application, name.ToString()) is ("word", "OpusApp") or ("excel", "XLMAIN") or ("powerpoint", "PPTFrameClass");
    }

    public void UpdateSize(FocusContract.SurfaceSize value) => size = value;

    /// <summary>Switches the tab Office docks into (a file shown in several tabs docks into the one last clicked).</summary>
    public void UseSurface(string token, FocusContract.SurfaceSize value)
    {
        if (token != surfaceToken) { WebviewSurface.ReleaseAnchor(surfaceToken); surfaceToken = token; nextResolve = 0; }
        size = value;
    }
    public void ForgetRecovery() => recovery.Forget(office);

    public bool OfficeAlive => NativeWindow.IsWindow(office) && Pid(office) == officePid;
    public bool OfficeEnabled => IsWindowEnabled(office);
    public bool EditorIconic => editorRoot != IntPtr.Zero && IsIconic(editorRoot);
    public int VirtualScreenLeft => NativeWindow.GetSystemMetrics(SmXVirtualScreen);

    public FocusOwner Foreground
    {
        get
        {
            var foreground = NativeWindow.GetForegroundWindow();
            if (foreground == IntPtr.Zero) return FocusOwner.None;
            var root = NativeWindow.GetAncestor(foreground, GaRootOwner);
            if (root == office) return FocusOwner.Office;
            var pid = Pid(root);
            if (root == editorRoot || editorPid != 0 && pid == editorPid) return FocusOwner.Editor;
            return pid == 0 ? FocusOwner.None : FocusOwner.Other;
        }
    }

    public FocusRect OfficeRect
    {
        get
        {
            using var dpi = WebviewSurface.EnterPhysicalCoordinates();
            return NativeWindow.GetWindowRect(office, out var rect) ? new FocusRect(rect.Left, rect.Top, rect.Right, rect.Bottom) : default;
        }
    }

    public FocusRect? SurfaceBounds()
    {
        using var dpi = WebviewSurface.EnterPhysicalCoordinates();
        var placement = WebviewSurface.TryRefresh(surfaceToken, size.CssWidth, size.CssHeight, size.PixelRatio);
        if (placement is null)
        {
            // Discovery enumerates windows; throttle it while the tab is hidden.
            if (Environment.TickCount64 < nextResolve) return null;
            nextResolve = Environment.TickCount64 + ResolveIntervalMs;
            placement = WebviewSurface.Resolve(surfaceToken, 0, size.CssWidth, size.CssHeight, size.PixelRatio);
        }
        if (placement is not { } found) return null;
        editorRoot = found.Parent;
        editorPid = Pid(found.Parent);
        return new FocusRect(found.X, found.Y, found.X + found.Width, found.Y + found.Height);
    }

    public bool Fit(FocusRect target)
    {
        using var dpi = WebviewSurface.EnterPhysicalCoordinates();
        if (IsZoomed(office)) ShowWindow(office, SwShowNoActivate);
        var before = Frame(office);
        int left = before.Visible.Left - before.Outer.Left, top = before.Visible.Top - before.Outer.Top;
        int right = before.Outer.Right - before.Visible.Right, bottom = before.Outer.Bottom - before.Visible.Bottom;
        if (!NativeWindow.SetWindowPos(office, IntPtr.Zero, target.Left - left, target.Top - top,
                target.Width + left + right, target.Height + top + bottom, SwpNoActivate | SwpNoZOrder)) return false;
        DwmFlush();
        var actual = Frame(office).Visible;
        var error = new[] { actual.Left - target.Left, actual.Top - target.Top, actual.Right - target.Right, actual.Bottom - target.Bottom }.Max(Math.Abs);
        DebugNdjson.Write("F", "FocusWindow.Fit", "alignment", new { width = target.Width, height = target.Height, alignmentErrorPx = error });
        return error <= 2;
    }

    public void RestoreOriginal()
    {
        var placement = original;
        placement.Length = Marshal.SizeOf<WindowPlacement>();
        if (placement.ShowCmd != 3) placement.ShowCmd = SwShowNoActivate;
        SetWindowPlacement(office, ref placement);
    }

    public bool RequestOfficeFocus() => RequestFocus(office);
    public bool RequestEditorFocus() => editorRoot != IntPtr.Zero && RequestFocus(editorRoot);

    public void RecenterOwnedDialog(FocusRect area)
    {
        using var dpi = WebviewSurface.EnterPhysicalCoordinates();
        var foreground = NativeWindow.GetForegroundWindow();
        if (foreground == office || NativeWindow.GetAncestor(foreground, GaRootOwner) != office) return;
        if (NativeWindow.GetWindowRect(foreground, out var dialog) && dialog.Right <= VirtualScreenLeft) Center(foreground, dialog, area);
    }

    public bool RevealModal(FocusRect area)
    {
        using var dpi = WebviewSurface.EnterPhysicalCoordinates();
        var found = IntPtr.Zero;
        NativeWindow.EnumWindows((window, _) =>
        {
            if (window == office || Pid(window) != officePid || !NativeWindow.IsWindowVisible(window) || !IsWindowEnabled(window)) return true;
            for (var owner = NativeWindow.GetWindow(window, GwOwner); owner != IntPtr.Zero; owner = NativeWindow.GetWindow(owner, GwOwner))
                if (owner == office) { found = window; return false; }
            return true;
        }, IntPtr.Zero);
        if (found == IntPtr.Zero) return false;
        if (NativeWindow.GetWindowRect(found, out var dialog)) Center(found, dialog, area);
        return RequestFocus(found);
    }

    /// <summary>
    /// Intentionally does nothing. RISK_BOUNDARIES S05 forbids SendInput and other input injection, so the click that
    /// activated Office is not replayed; the user's next click lands in Office directly. A compliant caret placement
    /// through the Office object model (Window.RangeFromPoint) is tracked as a P5 option in ADR 0003.
    /// </summary>
    public void PlaceCaret(int x, int y) { }

    public bool PostWheel(double pointX, double pointY, int deltaY, int deltaX, bool ctrl, bool shift)
    {
        using var dpi = WebviewSurface.EnterPhysicalCoordinates();
        if (!OfficeAlive) return false;
        // The frame is the parked window's visible (DWM) area 1:1, so the tab point maps onto it directly.
        var visible = Frame(office).Visible;
        int width = visible.Right - visible.Left, height = visible.Bottom - visible.Top;
        if (width <= 0 || height <= 0) return false;
        var x = visible.Left + (int)Math.Round(pointX * (width - 1));
        var y = visible.Top + (int)Math.Round(pointY * (height - 1));
        // Word, Excel and PowerPoint only scroll when the content child gets the message, not the frame.
        var target = DeepestChild(office, x, y);
        if (target == IntPtr.Zero || NativeWindow.GetAncestor(target, GaRootOwner) != office) return false;
        var keys = (ctrl ? MkControl : 0) | (shift ? MkShift : 0);
        var position = (IntPtr)(((y & 0xffff) << 16) | (x & 0xffff));
        var posted = true;
        if (deltaY != 0) posted &= PostMessage(target, WmMouseWheel, (IntPtr)(((deltaY & 0xffff) << 16) | keys), position);
        if (deltaX != 0) posted &= PostMessage(target, WmMouseHWheel, (IntPtr)(((deltaX & 0xffff) << 16) | keys), position);
        return posted;
    }

    /// <summary>Deepest visible, enabled descendant under a screen point; works while the window is off screen.</summary>
    static IntPtr DeepestChild(IntPtr top, int screenX, int screenY)
    {
        var current = top;
        for (var depth = 0; depth < 32; depth++)
        {
            var point = new NativeWindow.Point { X = screenX, Y = screenY };
            if (!NativeWindow.ScreenToClient(current, ref point)) return current;
            var child = ChildWindowFromPointEx(current, point, CwpSkipInvisible | CwpSkipDisabled | CwpSkipTransparent);
            if (child == IntPtr.Zero || child == current) return current;
            current = child;
        }
        return current;
    }

    public (int X, int Y)? CursorPosition
    {
        get
        {
            using var dpi = WebviewSurface.EnterPhysicalCoordinates();
            return GetCursorPos(out var point) ? (point.X, point.Y) : null;
        }
    }

    bool RequestFocus(IntPtr wanted)
    {
        bool Confirmed() => NativeWindow.GetAncestor(NativeWindow.GetForegroundWindow(), GaRootOwner) == NativeWindow.GetAncestor(wanted, GaRootOwner);
        if (SetForegroundWindow(wanted)) for (var i = 0; i < 20 && !Confirmed(); i++) Thread.Sleep(10);
        if (Confirmed()) return true;
        // Foreground lock: only steal focus between the editor and this Office window, never from another app.
        if (Foreground is not (FocusOwner.Editor or FocusOwner.Office)) return false;
        var result = UiaSetFocus(wanted);
        for (var i = 0; result == 0 && i < 20 && !Confirmed(); i++) Thread.Sleep(10);
        return Confirmed();
    }

    static int UiaSetFocus(IntPtr window)
    {
        var automation = UiaCom.Create();
        if (automation is null) return UiaCom.LastCreateHr == 0 ? -1 : UiaCom.LastCreateHr;
        try
        {
            var hr = automation.ElementFromHandle(window, out var element);
            if (hr != 0 || element is null) return hr == 0 ? -1 : hr;
            try { return ((IFocusableElement)element).SetFocus(); }
            finally { Marshal.ReleaseComObject(element); }
        }
        catch (Exception error) { return error.HResult; }
        finally { Marshal.ReleaseComObject(automation); }
    }

    void Center(IntPtr window, OleDocumentHost.NativeRect dialog, FocusRect area)
    {
        int width = dialog.Right - dialog.Left, height = dialog.Bottom - dialog.Top;
        var left = Math.Max(area.Left, (area.Left + area.Right - width) / 2);
        var top = Math.Max(area.Top, (area.Top + area.Bottom - height) / 2);
        var moved = NativeWindow.SetWindowPos(window, IntPtr.Zero, left, top, 0, 0, SwpNoSize | SwpNoZOrder | SwpNoActivate);
        DebugNdjson.Write("F", "FocusWindow.Dialog", moved ? "office_dialog_recentered" : "office_dialog_recenter_failed", null);
    }

    static uint Pid(IntPtr window)
    {
        if (window == IntPtr.Zero) return 0;
        NativeWindow.GetWindowThreadProcessId(window, out var pid);
        return pid;
    }

    static (OleDocumentHost.NativeRect Outer, OleDocumentHost.NativeRect Visible) Frame(IntPtr window)
    {
        NativeWindow.GetWindowRect(window, out var outer);
        var visible = DwmGetWindowAttribute(window, DwmwaExtendedFrameBounds, out var extended, Marshal.SizeOf<OleDocumentHost.NativeRect>()) == 0 ? extended : outer;
        return (outer, visible);
    }


    [ComImport, Guid("D22108AA-8AC5-49A5-837B-37BBB3D7591E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IFocusableElement { [PreserveSig] int SetFocus(); }

    [StructLayout(LayoutKind.Sequential)]
    struct WindowPlacement
    {
        public int Length, Flags, ShowCmd;
        public int MinX, MinY, MaxX, MaxY;
        public int NormalLeft, NormalTop, NormalRight, NormalBottom;
    }

    [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern IntPtr ChildWindowFromPointEx(IntPtr parent, NativeWindow.Point point, uint flags);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr window);
    [DllImport("user32.dll")] static extern bool IsZoomed(IntPtr window);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr window, int command);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] static extern bool GetCursorPos(out NativeWindow.Point point);
    [DllImport("user32.dll")] static extern bool GetWindowPlacement(IntPtr window, ref WindowPlacement placement);
    [DllImport("user32.dll")] static extern bool SetWindowPlacement(IntPtr window, ref WindowPlacement placement);
    [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr window, int attribute, out OleDocumentHost.NativeRect value, int size);
    [DllImport("dwmapi.dll")] static extern int DwmFlush();

    /// <summary>Puts an Office window back on screen; used by crash recovery.</summary>
    internal static bool RestorePlacement(IntPtr window, uint pid, int left, int top, int right, int bottom, int showCmd)
    {
        if (!NativeWindow.IsWindow(window) || Pid(window) != pid) return false;
        if (!NativeWindow.GetWindowRect(window, out var rect) || rect.Right > NativeWindow.GetSystemMetrics(SmXVirtualScreen)) return false;
        var placement = new WindowPlacement
        {
            Length = Marshal.SizeOf<WindowPlacement>(), ShowCmd = showCmd == 3 ? 3 : SwShowNoActivate,
            NormalLeft = left, NormalTop = top, NormalRight = right, NormalBottom = bottom,
        };
        return SetWindowPlacement(window, ref placement);
    }
}

/// <summary>
/// D2: remembers the original placement of every Office window Core moves, so a Core crash never leaves a document
/// parked off-screen. Written only when a session starts or ends; holds window handles and rectangles, no paths.
/// </summary>
internal sealed class FocusRecovery
{
    readonly string file;
    readonly object gate = new();
    readonly Dictionary<long, JsonObject> entries = new();

    public FocusRecovery(string dataRoot) => file = Path.Combine(Path.GetFullPath(dataRoot), "focus-recovery.json");

    public void Remember(IntPtr window, uint pid, int left, int top, int right, int bottom, int showCmd)
    {
        lock (gate)
        {
            entries[window.ToInt64()] = new JsonObject { ["hwnd"] = window.ToInt64(), ["pid"] = pid, ["left"] = left, ["top"] = top, ["right"] = right, ["bottom"] = bottom, ["show"] = showCmd };
            Save();
        }
    }

    public void Forget(IntPtr window)
    {
        lock (gate) if (entries.Remove(window.ToInt64())) Save();
    }

    /// <summary>Restores windows left parked by a previous Core process. Returns how many were moved back.</summary>
    public int RestoreOrphans()
    {
        JsonArray? saved;
        try { saved = File.Exists(file) ? JsonNode.Parse(File.ReadAllText(file)) as JsonArray : null; }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException) { saved = null; }
        var restored = 0;
        foreach (var item in saved ?? [])
        {
            if (item is not JsonObject entry || entry["hwnd"] is not JsonValue handle || !handle.TryGetValue<long>(out var hwnd)) continue;
            var pid = entry["pid"] is JsonValue p && p.TryGetValue<uint>(out var parsed) ? parsed : 0;
            if (Win32FocusWindowOps.RestorePlacement(new IntPtr(hwnd), pid, Wire.Number(entry, "left"), Wire.Number(entry, "top"),
                    Wire.Number(entry, "right"), Wire.Number(entry, "bottom"), Wire.Number(entry, "show"))) restored++;
        }
        lock (gate) { entries.Clear(); Save(); }
        if (restored > 0) DebugNdjson.Write("F", "FocusRecovery", "restored_orphans", new { restored });
        return restored;
    }

    void Save()
    {
        try
        {
            if (entries.Count == 0) { if (File.Exists(file)) File.Delete(file); return; }
            Directory.CreateDirectory(Path.GetDirectoryName(file)!);
            File.WriteAllText(file, new JsonArray(entries.Values.Select(item => (JsonNode)item.DeepClone()).ToArray()).ToJsonString());
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { /* Recovery is best effort. */ }
    }
}
