using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeInEditor.Core;

/// <summary>
/// Resolves a planted editor token to its own IDE root and physical screen rectangle.
/// No foreground-window or size-based window selection is used.
/// </summary>
internal static class WebviewSurface
{
    private static readonly string[] IdeProcessNames = ["Cursor", "Code"];
    private static readonly string[] CornerSuffixes = [":tl", ":tr", ":bl", ":br"];

    // Parent is the token's top-level IDE owner, not a browser child HWND.
    // Every coordinate and extent here is in physical screen pixels.
    internal readonly record struct Placement(IntPtr Parent, int X, int Y, int Width, int Height, string How);
    internal readonly record struct ScreenBounds(int X, int Y, int Width, int Height);

    private sealed class Search(int width, int height, double pixelRatio)
    {
        internal readonly int Width = width, Height = height;
        internal readonly double PixelRatio = pixelRatio;
        internal bool Poked;
        internal int Nodes;
    }

    internal static bool IsToken(string? token)
    {
        if (token is not { Length: 44 } || !token.StartsWith("oie-surface-", StringComparison.OrdinalIgnoreCase)) return false;
        for (var i = 12; i < token.Length; i++) if (!char.IsAsciiHexDigit(token[i])) return false;
        return true;
    }

    internal static bool IsTrustedHow(string? how)
    {
        if (string.IsNullOrEmpty(how)) return false;
        while (how.EndsWith("-cache", StringComparison.Ordinal)) how = how[..^6];
        return how is "uia-beacon-corners" or "uia-token" or "acc-token" or "token-hwnd";
    }

    internal static bool IsValidPixelRatio(double pixelRatio) =>
        double.IsFinite(pixelRatio) && pixelRatio is >= 0.5 and <= 8;

    // The wire width/height are CSS pixels and are used only to reject stale or
    // workbench-sized provider rectangles. Never use them to move a native HWND.
    internal static bool MatchesExpectedSize(ScreenBounds bounds, int width, int height, double pixelRatio = 1)
    {
        if (!ValidBounds(bounds) || !IsValidPixelRatio(pixelRatio)) return false;
        if (width < 32 || height < 32) return width == 0 && height == 0;
        double physicalWidth = width * pixelRatio, physicalHeight = height * pixelRatio;
        double tolerance = Math.Max(4, 2 * pixelRatio);
        return Math.Abs(bounds.Width - physicalWidth) <= tolerance
            && Math.Abs(bounds.Height - physicalHeight) <= tolerance;
    }

    internal static bool ContainsBounds(ScreenBounds outer, ScreenBounds inner) =>
        ValidBounds(outer) && ValidBounds(inner)
        && inner.X >= outer.X && inner.Y >= outer.Y
        && (long)inner.X + inner.Width <= (long)outer.X + outer.Width
        && (long)inner.Y + inner.Height <= (long)outer.Y + outer.Height;

    private static bool ValidBounds(ScreenBounds bounds) =>
        bounds.Width >= 32 && bounds.Height >= 32
        && (long)bounds.X + bounds.Width <= int.MaxValue
        && (long)bounds.Y + bounds.Height <= int.MaxValue;

    internal static bool TryBeaconBounds(IReadOnlyList<ScreenBounds> corners, out ScreenBounds bounds)
    {
        bounds = default;
        if (corners.Count != 4) return false;
        foreach (var dot in corners)
            if (dot.Width is < 1 or > 16 || dot.Height is < 1 or > 16) return false;
        var tl = corners[0]; var tr = corners[1]; var bl = corners[2]; var br = corners[3];
        long Right(ScreenBounds r) => (long)r.X + r.Width;
        long Bottom(ScreenBounds r) => (long)r.Y + r.Height;
        var tolerance = Math.Max(2, corners.Max(r => Math.Max(r.Width, r.Height)));
        if (Math.Abs((long)tl.X - bl.X) > tolerance || Math.Abs(Right(tr) - Right(br)) > tolerance
            || Math.Abs((long)tl.Y - tr.Y) > tolerance || Math.Abs(Bottom(bl) - Bottom(br)) > tolerance
            || Right(tl) >= tr.X || Right(bl) >= br.X || Bottom(tl) >= bl.Y || Bottom(tr) >= br.Y)
            return false;
        long left = Math.Min(tl.X, bl.X), top = Math.Min(tl.Y, tr.Y);
        long right = Math.Max(Right(tr), Right(br)), bottom = Math.Max(Bottom(bl), Bottom(br));
        if (right > int.MaxValue || bottom > int.MaxValue || right - left > int.MaxValue || bottom - top > int.MaxValue)
            return false;
        bounds = new ScreenBounds((int)left, (int)top, (int)(right - left), (int)(bottom - top));
        return ValidBounds(bounds);
    }

    internal static (int X, int Y, int Width, int Height) ToClient(int screenX, int screenY, int width, int height, int parentScreenX, int parentScreenY) =>
        (screenX - parentScreenX, screenY - parentScreenY, width, height);

    internal static bool IsEditorWebview(IntPtr hwnd) =>
        hwnd != IntPtr.Zero && NativeWindow.IsWindow(hwnd) && NativeWindow.GetParent(hwnd) != IntPtr.Zero
        && (ClassName(hwnd).StartsWith("Chrome_RenderWidgetHostHWND", StringComparison.Ordinal)
            || ClassName(hwnd).StartsWith("Chrome_WidgetWin_0", StringComparison.Ordinal));

    // Kept for callers that inspect legacy browser-child candidates. Resolution
    // itself always returns the owning IDE root and the exact token rectangle.
    internal static bool IsTightEditorHost(Placement placement)
    {
        if (!IsEditorWebview(placement.Parent)) return false;
        NativeWindow.GetClientRect(placement.Parent, out var client);
        return placement.X == 0 && placement.Y == 0
            && placement.Width == client.Right - client.Left && placement.Height == client.Bottom - client.Top;
    }

    internal static string ClassName(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return "";
        var name = new StringBuilder(256);
        NativeWindow.GetClassName(hwnd, name, name.Capacity);
        return name.ToString();
    }

    internal static bool FitsEditor(Placement placement, int expectedWidth, int expectedHeight, double pixelRatio = 1)
    {
        using var dpi = EnterPhysicalCoordinates();
        if (!IsTrustedHow(placement.How) || !IsIdeRoot(placement.Parent)) return false;
        if (!NativeWindow.GetWindowRect(placement.Parent, out var owner)) return false;
        return ContainsBounds(Bounds(owner), new(placement.X, placement.Y, placement.Width, placement.Height))
            && MatchesExpectedSize(new(placement.X, placement.Y, placement.Width, placement.Height),
                expectedWidth, expectedHeight, pixelRatio);
    }

    internal static Placement? Resolve(string token, int timeoutMs, int expectedWidth = 0, int expectedHeight = 0, double pixelRatio = 1)
    {
        if (!IsToken(token) || !IsValidPixelRatio(pixelRatio)) return null;
        using var dpi = EnterPhysicalCoordinates();
        var search = new Search(expectedWidth, expectedHeight, pixelRatio);
        var until = Environment.TickCount64 + Math.Max(0, timeoutMs);
        while (true)
        {
            var found = TryResolve(token, search);
            if (found is not null) return found;
            if (Environment.TickCount64 >= until) return null;
            NativeWindow.PumpWait(50);
        }
    }

    internal static Placement? TryResolve(string token) => Resolve(token, 0);

    private static Placement? TryResolve(string token, Search search)
    {
        // Each native root supplies both the search scope and the eventual owner.
        // A token discovered under one IDE can never be attached to another IDE.
        var owners = IdeFrames().ToArray();
        var uia = UiaCom.Create();
        try
        {
            foreach (var owner in owners)
            {
                foreach (var hwnd in SurfaceWindows(owner))
                {
                    if (!search.Poked) PokeUia(hwnd);
                    UiaCom.IElement? root = null;
                    try
                    {
                        if (uia is not null && uia.ElementFromHandle(hwnd, out root) == 0 && root is not null)
                        {
                            var hit = TryBeacons(uia, root, token, owner, search)
                                ?? TryUiaToken(uia, root, token, owner, search);
                            if (hit is not null) return hit;
                        }
                        // A top-level title can describe a tab while its HWND covers
                        // the whole workbench. Only token-bearing browser children qualify.
                        if (hwnd != owner && IsEditorWebview(hwnd) && HasToken(WindowText(hwnd), token)
                            && NativeWindow.GetWindowRect(hwnd, out var rect))
                        {
                            var hit = FromBounds(owner, Bounds(rect), "token-hwnd", search);
                            if (hit is not null) return hit;
                        }
                    }
                    catch (COMException) { }
                    finally { Release(root); }
                }
            }
            search.Poked = true;
            foreach (var owner in owners)
            {
                search.Nodes = 0;
                foreach (var hwnd in SurfaceWindows(owner))
                {
                    object? accessible = null;
                    try
                    {
                        accessible = AccFromWindow(hwnd);
                        if (accessible is null) continue;
                        var hit = Walk(accessible, token, owner, search, 0);
                        if (hit is not null) return hit;
                    }
                    catch (COMException) { }
                    finally { Release(accessible); }
                }
            }
            return null;
        }
        finally { Release(uia); }
    }

    private static Placement? TryBeacons(UiaCom.IAutomation uia, UiaCom.IElement root, string token, IntPtr owner, Search search)
    {
        var corners = new ScreenBounds[4];
        for (var i = 0; i < CornerSuffixes.Length; i++)
        {
            UiaCom.IElement? element = null;
            UiaCom.ICondition? condition = null;
            try
            {
                if (uia.CreatePropertyCondition(UiaCom.Name, token + CornerSuffixes[i], out condition) != 0 || condition is null
                    || root.FindFirst(UiaCom.TreeDescendants, condition, out element) != 0 || element is null
                    || !IsVisible(element) || !TryElementBounds(element, 1, out corners[i]))
                    return null;
            }
            finally { Release(element); Release(condition); }
        }
        return TryBeaconBounds(corners, out var bounds) ? FromBounds(owner, bounds, "uia-beacon-corners", search) : null;
    }

    private static Placement? TryUiaToken(UiaCom.IAutomation uia, UiaCom.IElement root, string token, IntPtr owner, Search search)
    {
        if (ElementHasToken(root, token) && IsVisible(root) && TryElementBounds(root, 32, out var rootBounds))
        {
            var hit = FromBounds(owner, rootBounds, "uia-token", search);
            if (hit is not null) return hit;
        }
        foreach (var property in (ReadOnlySpan<int>)[UiaCom.Name, UiaCom.AutomationId])
        {
            UiaCom.ICondition? condition = null;
            UiaCom.IElement? element = null;
            try
            {
                if (uia.CreatePropertyCondition(property, token, out condition) != 0 || condition is null
                    || root.FindFirst(UiaCom.TreeDescendants, condition, out element) != 0 || element is null
                    || !IsVisible(element) || !TryElementBounds(element, 32, out var bounds))
                    continue;
                var hit = FromBounds(owner, bounds, "uia-token", search);
                if (hit is not null) return hit;
            }
            finally { Release(element); Release(condition); }
        }
        return null;
    }

    private static bool ElementHasToken(UiaCom.IElement element, string token)
    {
        foreach (var property in (ReadOnlySpan<int>)[UiaCom.Name, UiaCom.AutomationId])
            if (element.GetCurrentPropertyValue(property, out var value) == 0 && HasToken(value as string, token)) return true;
        return false;
    }

    private static bool IsVisible(UiaCom.IElement element) =>
        element.GetCurrentPropertyValue(UiaCom.IsOffscreen, out var value) == 0 && value is bool offscreen && !offscreen;

    private static bool TryElementBounds(UiaCom.IElement element, int minimum, out ScreenBounds bounds)
    {
        bounds = default;
        return element.GetCurrentPropertyValue(UiaCom.BoundingRectangle, out var value) == 0
            && TryReadBounds(value, minimum, out bounds);
    }

    private static Placement? FromBounds(IntPtr owner, ScreenBounds bounds, string how, Search search)
    {
        if (!IsIdeRoot(owner) || !NativeWindow.GetWindowRect(owner, out var ownerRect)
            || !ContainsBounds(Bounds(ownerRect), bounds)
            || !MatchesExpectedSize(bounds, search.Width, search.Height, search.PixelRatio))
            return null;
        // Without a caller size, never accept an entire workbench as an editor.
        if (search.Width == 0 && search.Height == 0
            && Math.Abs((long)bounds.Width - (ownerRect.Right - ownerRect.Left)) <= 8
            && Math.Abs((long)bounds.Height - (ownerRect.Bottom - ownerRect.Top)) <= 8)
            return null;
        return new Placement(owner, bounds.X, bounds.Y, bounds.Width, bounds.Height, how);
    }

    private static ScreenBounds Bounds(OleDocumentHost.NativeRect rect) =>
        new(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);

    internal static bool TryBounds(object? value, out int x, out int y, out int width, out int height)
    {
        var valid = TryReadBounds(value, 32, out var bounds);
        x = bounds.X; y = bounds.Y; width = bounds.Width; height = bounds.Height;
        return valid;
    }

    private static bool TryReadBounds(object? value, int minimum, out ScreenBounds bounds)
    {
        bounds = default;
        try
        {
            double x, y, width, height;
            if (value is Array array && array.Rank == 1 && array.Length == 4)
            {
                var start = array.GetLowerBound(0);
                x = Convert.ToDouble(array.GetValue(start));
                y = Convert.ToDouble(array.GetValue(start + 1));
                width = Convert.ToDouble(array.GetValue(start + 2));
                height = Convert.ToDouble(array.GetValue(start + 3));
            }
            else
            {
                if (value is null) return false;
                var type = value.GetType();
                double? Read(string name)
                {
                    var property = type.GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                    var field = type.GetField(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                    var raw = property is not null ? property.GetValue(value) : field?.GetValue(value);
                    return raw is null ? null : Convert.ToDouble(raw);
                }
                var left = Read("left") ?? Read("x");
                var top = Read("top") ?? Read("y");
                if (left is null || top is null) return false;
                x = left.Value; y = top.Value;
                width = Read("width") ?? (Read("right") - x) ?? 0;
                height = Read("height") ?? (Read("bottom") - y) ?? 0;
            }
            if (!double.IsFinite(x) || !double.IsFinite(y) || !double.IsFinite(width) || !double.IsFinite(height)
                || width < minimum || height < minimum || x < int.MinValue || y < int.MinValue)
                return false;
            var leftPixel = Math.Floor(x); var topPixel = Math.Floor(y);
            var rightPixel = Math.Ceiling(x + width); var bottomPixel = Math.Ceiling(y + height);
            if (rightPixel > int.MaxValue || bottomPixel > int.MaxValue
                || rightPixel - leftPixel > int.MaxValue || bottomPixel - topPixel > int.MaxValue)
                return false;
            bounds = new((int)leftPixel, (int)topPixel, (int)(rightPixel - leftPixel), (int)(bottomPixel - topPixel));
            return bounds.Width >= minimum && bounds.Height >= minimum;
        }
        catch (Exception error) when (error is ArgumentException or InvalidCastException or FormatException or OverflowException or TargetInvocationException)
        {
            return false;
        }
    }

    private static Placement? Walk(object accessible, string token, IntPtr owner, Search search, int depth)
    {
        if (depth > 24 || ++search.Nodes > 4000) return null;
        if (HasToken(AccName(accessible, 0), token) || HasToken(AccName(accessible, Type.Missing), token))
        {
            var hit = FromAccessible(accessible, 0, owner, search);
            if (hit is not null) return hit;
        }
        var count = Math.Min(AccChildCount(accessible), 256);
        if (count <= 0) return null;
        var children = new object[count];
        if (AccessibleChildren(accessible, 0, count, children, out var got) != 0) return null;
        for (var i = 0; i < Math.Min(got, count); i++)
        {
            var child = children[i];
            if (child is int or short or byte)
            {
                if (HasToken(AccName(accessible, child), token))
                {
                    var hit = FromAccessible(accessible, child, owner, search);
                    if (hit is not null) return hit;
                }
            }
            else if (child is not null && child is not DBNull)
            {
                try
                {
                    var hit = Walk(child, token, owner, search, depth + 1);
                    if (hit is not null) return hit;
                }
                finally { Release(child); }
            }
        }
        return null;
    }

    private static Placement? FromAccessible(object accessible, object child, IntPtr owner, Search search)
    {
        try
        {
            object[] args = [0, 0, 0, 0, child];
            var modifier = new ParameterModifier(5);
            modifier[0] = modifier[1] = modifier[2] = modifier[3] = true;
            accessible.GetType().InvokeMember("accLocation", BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod,
                null, accessible, args, [modifier], null, null);
            var bounds = new ScreenBounds(Convert.ToInt32(args[0]), Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToInt32(args[3]));
            return FromBounds(owner, bounds, "acc-token", search);
        }
        catch { return null; }
    }

    private static bool HasToken(string? name, string token) => string.Equals(name, token, StringComparison.OrdinalIgnoreCase);

    private static string AccName(object accessible, object child)
    {
        try
        {
            return accessible.GetType().InvokeMember("accName", BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty,
                null, accessible, [child]) as string ?? "";
        }
        catch { return ""; }
    }

    private static int AccChildCount(object accessible)
    {
        try
        {
            return Convert.ToInt32(accessible.GetType().InvokeMember("accChildCount",
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty, null, accessible, null));
        }
        catch { return 0; }
    }

    private static object? AccFromWindow(IntPtr hwnd)
    {
        var iid = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
        return AccessibleObjectFromWindow(hwnd, unchecked((uint)0xFFFFFFFC), ref iid, out var accessible) == 0 ? accessible : null;
    }

    private static string WindowText(IntPtr hwnd)
    {
        var text = new StringBuilder(128);
        NativeWindow.GetWindowText(hwnd, text, text.Capacity);
        return text.ToString();
    }

    private static bool IsIdeRoot(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeWindow.IsWindow(hwnd) || !NativeWindow.IsWindowVisible(hwnd)
            || NativeWindow.GetAncestor(hwnd, 2) != hwnd
            || !ClassName(hwnd).StartsWith("Chrome_WidgetWin_", StringComparison.Ordinal)) return false;
        NativeWindow.GetWindowThreadProcessId(hwnd, out var pid);
        try
        {
            using var process = Process.GetProcessById((int)pid);
            return IdeProcessNames.Contains(process.ProcessName, StringComparer.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static IEnumerable<IntPtr> IdeFrames()
    {
        var frames = new List<IntPtr>();
        NativeWindow.EnumProc callback = (window, _) => { if (IsIdeRoot(window)) frames.Add(window); return true; };
        NativeWindow.EnumWindows(callback, IntPtr.Zero);
        GC.KeepAlive(callback);
        return frames;
    }

    private static IEnumerable<IntPtr> SurfaceWindows(IntPtr owner)
    {
        yield return owner;
        var children = new List<IntPtr>();
        NativeWindow.EnumChildProc callback = (window, _) =>
        {
            if (NativeWindow.IsWindowVisible(window)
                && NativeWindow.GetAncestor(window, 2) == owner
                && ClassName(window).StartsWith("Chrome_", StringComparison.Ordinal)) children.Add(window);
            return true;
        };
        NativeWindow.EnumChildWindows(owner, callback, IntPtr.Zero);
        GC.KeepAlive(callback);
        foreach (var child in children) yield return child;
    }

    private static void PokeUia(IntPtr hwnd)
    {
        // A hung IDE window must not indefinitely block another tab's resolution.
        SendMessageTimeout(hwnd, 0x003D, IntPtr.Zero, new IntPtr(-25), 3, 100, out _);
    }

    internal static IDisposable EnterPhysicalCoordinates() => new PhysicalCoordinates();

    private sealed class PhysicalCoordinates : IDisposable
    {
        private readonly IntPtr previous;
        internal PhysicalCoordinates()
        {
            previous = SetThreadDpiAwarenessContext(new IntPtr(-4));
            if (previous == IntPtr.Zero)
                throw new CoreFault("invalid_host_window", "The editor's native coordinate context is unavailable.");
        }
        public void Dispose() { SetThreadDpiAwarenessContext(previous); }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            try { Marshal.ReleaseComObject(value); } catch { }
    }

    [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object accessible);
    [DllImport("oleacc.dll")] private static extern int AccessibleChildren([MarshalAs(UnmanagedType.IUnknown)] object container, int start, int count, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] children, out int obtained);
}

