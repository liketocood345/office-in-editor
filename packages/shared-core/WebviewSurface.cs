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
    private const int AnchorCapacity = 16;
    private const int ElementNotAvailable = unchecked((int)0x80040201);
    [ThreadStatic] private static Dictionary<string, Anchor>? anchors;
    [ThreadStatic] private static long anchorClock;

    // Raw IUnknown pointers, rather than the search RCWs, own these references.
    // Search finally blocks may release their RCWs without invalidating anchors.
    // The thread-local cache must be acquired, refreshed and released on one STA.
    private sealed class Anchor(IntPtr owner, uint pid, uint tid, IntPtr[] nodes, string how)
    {
        internal readonly IntPtr Owner = owner;
        internal readonly uint Pid = pid, Tid = tid;
        internal readonly IntPtr[] Nodes = nodes;
        internal readonly string How = how;
        internal IntPtr ChildWindow { get; init; }
        internal uint ChildPid { get; init; }
        internal uint ChildTid { get; init; }
        internal object AccessibleChild { get; init; } = 0;
        internal object AccessibleNameChild { get; init; } = 0;
        internal long Used;
    }

    private enum AnchorRead { Success, Transient, Invalid }

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
    internal static int PhysicalFromCss(int cssPixels, double pixelRatio)
    {
        if (cssPixels < 32 || !IsValidPixelRatio(pixelRatio)) return 0;
        var physical = (int)Math.Round(cssPixels * pixelRatio, MidpointRounding.AwayFromZero);
        return physical < 32 ? 0 : physical;
    }

    internal static bool MatchesExpectedSize(ScreenBounds bounds, int width, int height, double pixelRatio = 1)
    {
        if (!ValidBounds(bounds) || !IsValidPixelRatio(pixelRatio)) return false;
        if (width < 32 || height < 32) return width == 0 && height == 0;
        double physicalWidth = width * pixelRatio, physicalHeight = height * pixelRatio;
        double tolerance = Math.Max(4, 2 * pixelRatio);
        return Math.Abs(bounds.Width - physicalWidth) <= tolerance
            && Math.Abs(bounds.Height - physicalHeight) <= tolerance;
    }

    internal static bool TryCssIdleNested(ScreenBounds owner, int insetLeft, int insetTop, int cssWidth, int cssHeight, double pixelRatio, out ScreenBounds nested)
    {
        nested = default;
        var width = PhysicalFromCss(cssWidth, pixelRatio);
        var height = PhysicalFromCss(cssHeight, pixelRatio);
        if (width < 32 || height < 32 || insetLeft < 0 || insetTop < 0) return false;
        var x = owner.X + insetLeft;
        var y = owner.Y + insetTop;
        long maxRight = (long)owner.X + owner.Width;
        long maxBottom = (long)owner.Y + owner.Height;
        if (x < owner.X || y < owner.Y || x >= maxRight || y >= maxBottom) return false;
        var maxW = (int)Math.Min(int.MaxValue, maxRight - x);
        var maxH = (int)Math.Min(int.MaxValue, maxBottom - y);
        nested = new(x, y, Math.Min(width, maxW), Math.Min(height, maxH));
        return ValidBounds(nested) && ContainsBounds(owner, nested);
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
        var inside = ContainsBounds(Bounds(owner), new(placement.X, placement.Y, placement.Width, placement.Height));
        if (!inside) return false;
        var how = placement.How;
        while (how.EndsWith("-cache", StringComparison.Ordinal)) how = how[..^6];
        if (how == "uia-beacon-corners") return true;
        return MatchesExpectedSize(new(placement.X, placement.Y, placement.Width, placement.Height),
            expectedWidth, expectedHeight, pixelRatio);
    }

    internal static Placement? Resolve(string token, int timeoutMs, int expectedWidth = 0, int expectedHeight = 0, double pixelRatio = 1)
    {
        if (!IsToken(token) || !IsValidPixelRatio(pixelRatio)) return null;
        // Explicit resolution replaces this token's old provider/owner binding.
        ReleaseAnchor(token);
        using var dpi = EnterPhysicalCoordinates();
        var search = new Search(expectedWidth, expectedHeight, pixelRatio);
        var until = Environment.TickCount64 + Math.Max(0, timeoutMs);
        var started = Stopwatch.GetTimestamp();
        var attempts = 0;
        while (true)
        {
            attempts++;
            var found = TryResolve(token, search);
            if (found is not null)
            {
                var elapsedMs = (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds;
                if (elapsedMs >= 100)
                {
                    // #region agent log
                    DebugNdjson.Write("T", "WebviewSurface.Resolve", "resolve_timing", new { hypothesisId = "T1", runId = "post-fix", elapsedMs, attempts, timeoutMs, how = found.Value.How, expectedWidth, expectedHeight });
                    // #endregion
                }
                return found;
            }
            if (Environment.TickCount64 >= until)
            {
                // #region agent log
                DebugNdjson.Write("T", "WebviewSurface.Resolve", "resolve_miss", new { hypothesisId = "T1", runId = "post-fix", elapsedMs = (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds, attempts, timeoutMs, expectedWidth, expectedHeight });
                // #endregion
                return null;
            }
            NativeWindow.PumpWait(50);
        }
    }

    internal static Placement? TryResolve(string token) => Resolve(token, 0);

    // Refresh is deliberately a fail-closed, bounded read of previously resolved
    // nodes. It must never fall back to discovery, window enumeration or a pump.
    internal static Placement? TryRefresh(string token, int cssW, int cssH, double ratio = 1)
    {
        if (!IsToken(token) || !IsValidPixelRatio(ratio)
            || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA
            || anchors is null || !anchors.TryGetValue(token, out var anchor)) return null;
        try
        {
            using var dpi = EnterPhysicalCoordinates();
            if (!SameAnchorOwner(anchor)) return DropAnchor(token);
            // Hidden/minimized IDEs can temporarily stop exposing usable bounds.
            // Keep their owned references so the next explicit refresh can recover.
            if (!NativeWindow.IsWindowVisible(anchor.Owner) || IsIconic(anchor.Owner)) return null;
            ScreenBounds bounds;
            if (anchor.How == "uia-beacon-corners" && anchor.Nodes.Length == 4)
            {
                var corners = new ScreenBounds[4];
                for (var i = 0; i < corners.Length; i++)
                {
                    var read = ReadAnchorNode(anchor.Nodes[i], token + CornerSuffixes[i], 1, out corners[i]);
                    if (read != AnchorRead.Success) return RefreshMiss(token, anchor, read == AnchorRead.Invalid);
                }
                if (!TryBeaconBounds(corners, out bounds)) return RefreshMiss(token, anchor);
            }
            else if (anchor.How == "uia-token" && anchor.Nodes.Length == 1)
            {
                var read = ReadAnchorNode(anchor.Nodes[0], token, 32, out bounds);
                if (read != AnchorRead.Success) return RefreshMiss(token, anchor, read == AnchorRead.Invalid);
                if (!MatchesExpectedSize(bounds, cssW, cssH, ratio)) return RefreshMiss(token, anchor);
            }
            else if (anchor.How == "acc-token" && anchor.Nodes.Length == 1)
            {
                var read = ReadAccessibleAnchor(anchor, token, out bounds);
                if (read != AnchorRead.Success) return RefreshMiss(token, anchor, read == AnchorRead.Invalid);
                if (!MatchesExpectedSize(bounds, cssW, cssH, ratio)) return RefreshMiss(token, anchor);
            }
            else if (anchor.How == "token-hwnd" && anchor.Nodes.Length == 0)
            {
                var read = ReadWindowAnchor(anchor, token, out bounds);
                if (read != AnchorRead.Success) return RefreshMiss(token, anchor, read == AnchorRead.Invalid);
                if (!MatchesExpectedSize(bounds, cssW, cssH, ratio)) return RefreshMiss(token, anchor);
            }
            else return DropAnchor(token);
            // Recheck after provider calls: an owner may have gone away while a
            // cached provider was answering. Never accept its stale rectangle.
            if (!SameAnchorOwner(anchor)) return DropAnchor(token);
            if (!NativeWindow.IsWindowVisible(anchor.Owner) || IsIconic(anchor.Owner)) return null;
            if (!NativeWindow.GetWindowRect(anchor.Owner, out var ownerRect)
                || !ContainsBounds(Bounds(ownerRect), bounds)) return null;
            anchor.Used = ++anchorClock;
            return new Placement(anchor.Owner, bounds.X, bounds.Y, bounds.Width, bounds.Height, anchor.How + "-cache");
        }
        catch (Exception error) when (error is COMException or InvalidComObjectException or ArgumentException or CoreFault
            or TargetInvocationException or InvalidCastException or FormatException or OverflowException or MissingMethodException)
        {
            var cause = error;
            while (cause is TargetInvocationException { InnerException: not null } invocation) cause = invocation.InnerException!;
            return RefreshMiss(token, anchor, cause.HResult == ElementNotAvailable);
        }
    }

    internal static void ReleaseAnchor(string token)
    {
        if (string.IsNullOrEmpty(token) || anchors is null || !anchors.Remove(token, out var anchor)) return;
        ReleasePointers(anchor.Nodes);
    }

    private static Placement? DropAnchor(string token)
    {
        ReleaseAnchor(token);
        return null;
    }

    private static Placement? RefreshMiss(string token, Anchor anchor, bool invalid = false) =>
        invalid || !SameAnchorOwner(anchor) ? DropAnchor(token) : null;

    private static bool SameAnchorOwner(Anchor anchor)
    {
        if (!NativeWindow.IsWindow(anchor.Owner)) return false;
        var tid = NativeWindow.GetWindowThreadProcessId(anchor.Owner, out var pid);
        return tid != 0 && pid != 0 && tid == anchor.Tid && pid == anchor.Pid;
    }

    private static AnchorRead ReadAnchorNode(IntPtr pointer, string token, int minimum, out ScreenBounds bounds)
    {
        bounds = default;
        if (pointer == IntPtr.Zero) return AnchorRead.Transient;
        object? temporary = null;
        try
        {
            // An isolated RCW avoids decrementing a search caller's shared RCW.
            temporary = Marshal.GetUniqueObjectForIUnknown(pointer);
            if (temporary is not UiaCom.IElement element) return AnchorRead.Transient;
            object? value = null;
            try
            {
                var hr = element.GetCurrentPropertyValue(UiaCom.BoundingRectangle, out value);
                if (hr == ElementNotAvailable) return AnchorRead.Invalid;
                if (hr != 0 || !TryReadBounds(value, minimum, out bounds)) return AnchorRead.Transient;
            }
            finally { Release(value); }
            var complete = true;
            foreach (var property in (ReadOnlySpan<int>)[UiaCom.Name, UiaCom.AutomationId])
            {
                value = null;
                try
                {
                    var hr = element.GetCurrentPropertyValue(property, out value);
                    if (hr == ElementNotAvailable) return AnchorRead.Invalid;
                    if (hr != 0 || value is not string name) { complete = false; continue; }
                    if (HasToken(name, token)) return AnchorRead.Success;
                }
                finally { Release(value); }
            }
            // A property/provider error is not proof that a planted node died.
            return complete ? AnchorRead.Invalid : AnchorRead.Transient;
        }
        finally { Release(temporary); }
    }

    private static AnchorRead ReadAccessibleAnchor(Anchor anchor, string token, out ScreenBounds bounds)
    {
        bounds = default;
        object? temporary = null;
        object? name = null;
        try
        {
            temporary = Marshal.GetUniqueObjectForIUnknown(anchor.Nodes[0]);
            if (!TryAccessibleBounds(temporary, anchor.AccessibleChild, out bounds)) return AnchorRead.Transient;
            name = temporary.GetType().InvokeMember("accName", BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty,
                null, temporary, [anchor.AccessibleNameChild]);
            return name is not string text ? AnchorRead.Transient
                : HasToken(text, token) ? AnchorRead.Success : AnchorRead.Invalid;
        }
        finally { Release(name); Release(temporary); }
    }

    private static bool SameAnchorChild(Anchor anchor)
    {
        if (anchor.ChildWindow == IntPtr.Zero || !NativeWindow.IsWindow(anchor.ChildWindow)
            || NativeWindow.GetAncestor(anchor.ChildWindow, 2) != anchor.Owner) return false;
        var tid = NativeWindow.GetWindowThreadProcessId(anchor.ChildWindow, out var pid);
        return tid != 0 && pid != 0 && tid == anchor.ChildTid && pid == anchor.ChildPid;
    }

    private static AnchorRead ReadWindowAnchor(Anchor anchor, string token, out ScreenBounds bounds)
    {
        bounds = default;
        if (!SameAnchorChild(anchor)) return AnchorRead.Invalid;
        if (!NativeWindow.IsWindowVisible(anchor.ChildWindow)
            || !TryClientBounds(anchor.ChildWindow, out bounds)) return AnchorRead.Transient;
        var title = new StringBuilder(128);
        Marshal.SetLastPInvokeError(0);
        var length = GetWindowTextForAnchor(anchor.ChildWindow, title, title.Capacity);
        if (length == 0 && Marshal.GetLastPInvokeError() != 0) return AnchorRead.Transient;
        if (!SameAnchorChild(anchor)) return AnchorRead.Invalid;
        if (!NativeWindow.IsWindowVisible(anchor.ChildWindow)) return AnchorRead.Transient;
        return HasToken(title.ToString(), token) ? AnchorRead.Success : AnchorRead.Invalid;
    }

    private static bool TryClientBounds(IntPtr hwnd, out ScreenBounds bounds)
    {
        bounds = default;
        if (!NativeWindow.GetClientRect(hwnd, out var client)) return false;
        var origin = new NativeWindow.Point { X = client.Left, Y = client.Top };
        if (!NativeWindow.ClientToScreen(hwnd, ref origin)) return false;
        bounds = new(origin.X, origin.Y, client.Right - client.Left, client.Bottom - client.Top);
        return ValidBounds(bounds);
    }

    // On success the cache takes ownership of every pointer in nodes. On failure
    // ownership remains with the caller, including partial acquisition failures.
    private static bool RememberAnchor(string token, IntPtr owner, IntPtr[] nodes, string how,
        IntPtr childWindow = default, object? accessibleChild = null, object? accessibleNameChild = null)
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA
            || nodes.Any(node => node == IntPtr.Zero) || !NativeWindow.IsWindow(owner)) return false;
        var tid = NativeWindow.GetWindowThreadProcessId(owner, out var pid);
        if (tid == 0 || pid == 0) return false;
        uint childTid = 0, childPid = 0;
        if (how == "token-hwnd")
        {
            if (childWindow == IntPtr.Zero || !NativeWindow.IsWindow(childWindow)
                || NativeWindow.GetAncestor(childWindow, 2) != owner) return false;
            childTid = NativeWindow.GetWindowThreadProcessId(childWindow, out childPid);
            if (childTid == 0 || childPid == 0) return false;
        }
        anchors ??= new(StringComparer.OrdinalIgnoreCase);
        ReleaseAnchor(token);
        // Prune stale owners while already handling an explicit resolution.
        foreach (var stale in anchors.Where(item => !SameAnchorOwner(item.Value)).Select(item => item.Key).ToArray())
            ReleaseAnchor(stale);
        if (anchors.Count >= AnchorCapacity)
            ReleaseAnchor(anchors.MinBy(item => item.Value.Used).Key);
        anchors.Add(token, new(owner, pid, tid, nodes, how)
        {
            Used = ++anchorClock, ChildWindow = childWindow, ChildPid = childPid, ChildTid = childTid,
            AccessibleChild = accessibleChild ?? 0, AccessibleNameChild = accessibleNameChild ?? accessibleChild ?? 0
        });
        return true;
    }

    private static void ReleasePointers(IntPtr[] pointers)
    {
        for (var i = 0; i < pointers.Length; i++)
        {
            var pointer = pointers[i];
            pointers[i] = IntPtr.Zero;
            if (pointer != IntPtr.Zero) Marshal.Release(pointer);
        }
    }

    private static Placement? KeepTokenAnchor(string token, IntPtr owner, UiaCom.IElement element, Placement? placement)
    {
        if (placement is null) return null;
        var nodes = new IntPtr[1];
        var kept = false;
        try
        {
            nodes[0] = Marshal.GetIUnknownForObject(element);
            kept = RememberAnchor(token, owner, nodes, "uia-token");
            return placement;
        }
        finally { if (!kept) ReleasePointers(nodes); }
    }

    private static Placement? KeepAccessibleAnchor(string token, IntPtr owner, object accessible, object child,
        object nameChild, Placement? placement)
    {
        if (placement is null) return null;
        var nodes = new IntPtr[1];
        var kept = false;
        try
        {
            nodes[0] = Marshal.GetIUnknownForObject(accessible);
            kept = RememberAnchor(token, owner, nodes, "acc-token", accessibleChild: child, accessibleNameChild: nameChild);
            return placement;
        }
        finally { if (!kept) ReleasePointers(nodes); }
    }

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
                            && TryClientBounds(hwnd, out var bounds))
                        {
                            var hit = FromBounds(owner, bounds, "token-hwnd", search);
                            if (hit is not null)
                            {
                                RememberAnchor(token, owner, [], "token-hwnd", childWindow: hwnd);
                                return hit;
                            }
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
        var nodes = new IntPtr[4];
        var kept = false;
        try
        {
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
                    nodes[i] = Marshal.GetIUnknownForObject(element);
                }
                finally { Release(element); Release(condition); }
            }
            var placement = TryBeaconBounds(corners, out var bounds)
                ? FromBounds(owner, bounds, "uia-beacon-corners", search, requireSize: false) : null;
            if (placement is not null) kept = RememberAnchor(token, owner, nodes, "uia-beacon-corners");
            return placement;
        }
        finally { if (!kept) ReleasePointers(nodes); }
    }

    private static Placement? TryUiaToken(UiaCom.IAutomation uia, UiaCom.IElement root, string token, IntPtr owner, Search search)
    {
        if (ElementHasToken(root, token) && IsVisible(root) && TryElementBounds(root, 32, out var rootBounds))
        {
            var hit = FromBounds(owner, rootBounds, "uia-token", search);
            if (hit is not null) return KeepTokenAnchor(token, owner, root, hit);
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
                if (hit is not null) return KeepTokenAnchor(token, owner, element, hit);
            }
            finally { Release(element); Release(condition); }
        }
        return null;
    }

    private static bool ElementHasToken(UiaCom.IElement element, string token)
    {
        foreach (var property in (ReadOnlySpan<int>)[UiaCom.Name, UiaCom.AutomationId])
        {
            object? value = null;
            try { if (element.GetCurrentPropertyValue(property, out value) == 0 && HasToken(value as string, token)) return true; }
            finally { Release(value); }
        }
        return false;
    }

    private static bool IsVisible(UiaCom.IElement element) =>
        element.GetCurrentPropertyValue(UiaCom.IsOffscreen, out var value) == 0 && value is bool offscreen && !offscreen;

    private static bool TryElementBounds(UiaCom.IElement element, int minimum, out ScreenBounds bounds)
    {
        bounds = default;
        object? value = null;
        try { return element.GetCurrentPropertyValue(UiaCom.BoundingRectangle, out value) == 0 && TryReadBounds(value, minimum, out bounds); }
        finally { Release(value); }
    }

    private static Placement? FromBounds(IntPtr owner, ScreenBounds bounds, string how, Search search, bool requireSize = true)
    {
        if (!IsIdeRoot(owner) || !NativeWindow.GetWindowRect(owner, out var ownerRect)
            || !ContainsBounds(Bounds(ownerRect), bounds)
            || (requireSize && !MatchesExpectedSize(bounds, search.Width, search.Height, search.PixelRatio)))
            return null;
        // Without a caller size, never accept an entire workbench as an editor.
        if (search.Width == 0 && search.Height == 0
            && Math.Abs((long)bounds.Width - (ownerRect.Right - ownerRect.Left)) <= 8
            && Math.Abs((long)bounds.Height - (ownerRect.Bottom - ownerRect.Top)) <= 8)
            return null;
        return new Placement(owner, bounds.X, bounds.Y, bounds.Width, bounds.Height, how);
    }

    internal static ScreenBounds Bounds(OleDocumentHost.NativeRect rect) =>
        new(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);

    internal static void ReadOwnerInsets(ScreenBounds owner, ScreenBounds nested, out int insetLeft, out int insetTop, out int insetRight, out int insetBottom)
    {
        insetLeft = nested.X - owner.X;
        insetTop = nested.Y - owner.Y;
        insetRight = owner.X + owner.Width - nested.X - nested.Width;
        insetBottom = owner.Y + owner.Height - nested.Y - nested.Height;
    }

    internal static ScreenBounds ApplyOwnerInsets(ScreenBounds owner, int insetLeft, int insetTop, int insetRight, int insetBottom)
    {
        var width = Math.Max(32, owner.Width - insetLeft - insetRight);
        var height = Math.Max(32, owner.Height - insetTop - insetBottom);
        return new(owner.X + insetLeft, owner.Y + insetTop, width, height);
    }

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
        if (HasToken(AccName(accessible, 0), token))
        {
            var hit = FromAccessible(accessible, 0, owner, search, token);
            if (hit is not null) return hit;
        }
        if (HasToken(AccName(accessible, Type.Missing), token))
        {
            // Preserve providers whose self-name is exposed only via the optional
            // property argument, while their accLocation still needs CHILDID_SELF.
            var hit = FromAccessible(accessible, 0, owner, search, token, Type.Missing);
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
                    var hit = FromAccessible(accessible, child, owner, search, token);
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

    private static Placement? FromAccessible(object accessible, object child, IntPtr owner, Search search, string token, object? nameChild = null)
    {
        try
        {
            if (!TryAccessibleBounds(accessible, child, out var bounds)) return null;
            return KeepAccessibleAnchor(token, owner, accessible, child, nameChild ?? child,
                FromBounds(owner, bounds, "acc-token", search));
        }
        catch { return null; }
    }

    private static bool TryAccessibleBounds(object accessible, object child, out ScreenBounds bounds)
    {
        object[] args = [0, 0, 0, 0, child];
        var modifier = new ParameterModifier(5);
        modifier[0] = modifier[1] = modifier[2] = modifier[3] = true;
        accessible.GetType().InvokeMember("accLocation", BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod,
            null, accessible, args, [modifier], null, null);
        bounds = new(Convert.ToInt32(args[0]), Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToInt32(args[3]));
        return ValidBounds(bounds);
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
    [DllImport("user32.dll", EntryPoint = "GetWindowTextW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetWindowTextForAnchor(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object accessible);
    [DllImport("oleacc.dll")] private static extern int AccessibleChildren([MarshalAs(UnmanagedType.IUnknown)] object container, int start, int count, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] children, out int obtained);
}

