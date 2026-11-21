using OfficeInEditor.Core;

internal static class SurfaceGeometryTests
{
    internal static void Run(Action<bool, string> check)
    {
        var corners = new WebviewSurface.ScreenBounds[]
        {
            new(-1800, 120, 2, 2), new(-602, 120, 2, 2),
            new(-1800, 1018, 2, 2), new(-602, 1018, 2, 2)
        };
        check(WebviewSurface.TryBeaconBounds(corners, out var rectangle)
            && rectangle == new WebviewSurface.ScreenBounds(-1800, 120, 1200, 900),
            "four token corners preserve physical coordinates on a negative-origin display");
        check(WebviewSurface.MatchesExpectedSize(rectangle, 800, 600, 1.5)
            && !WebviewSurface.MatchesExpectedSize(rectangle, 800, 600, 1),
            "CSS dimensions are validated using the reported pixel ratio");
        check(!WebviewSurface.MatchesExpectedSize(rectangle, 900, 600, 1.5)
            && !WebviewSurface.MatchesExpectedSize(rectangle, 800, 600, double.NaN),
            "stale dimensions and invalid scale cannot approve a native rectangle");
        check(!WebviewSurface.TryBeaconBounds(corners.Take(2).ToArray(), out _),
            "partial corner evidence cannot expand into an editor rectangle");
        var crossed = (WebviewSurface.ScreenBounds[])corners.Clone();
        (crossed[1], crossed[2]) = (crossed[2], crossed[1]);
        check(!WebviewSurface.TryBeaconBounds(crossed, out _),
            "corner labels must describe a consistent rectangle");
        var mixed = (WebviewSurface.ScreenBounds[])corners.Clone();
        mixed[3] = new WebviewSurface.ScreenBounds(400, 1018, 2, 2);
        check(!WebviewSurface.TryBeaconBounds(mixed, out _),
            "corners from different editor positions cannot be combined");
        var owner = new WebviewSurface.ScreenBounds(-1920, 0, 1920, 1080);
        check(WebviewSurface.ContainsBounds(owner, rectangle)
            && !WebviewSurface.ContainsBounds(new(0, 0, 1920, 1080), rectangle),
            "a token rectangle is constrained to its own IDE window");
        WebviewSurface.ReadOwnerInsets(owner, rectangle, out var insetL, out var insetT, out var insetR, out var insetB);
        check(insetL == 120 && insetT == 120 && insetR == 600 && insetB == 60
            && WebviewSurface.ApplyOwnerInsets(new(0, 40, 1920, 1080), insetL, insetT, insetR, insetB)
                == new WebviewSurface.ScreenBounds(120, 160, 1200, 900),
            "owner-inset follow preserves the beacon rectangle when the IDE window moves or resizes");
        check(WebviewSurface.ShiftedTowardOwnerOrigin(owner, rectangle, new(owner.X + 8, owner.Y + 8, 1200, 900))
            && !WebviewSurface.ShiftedTowardOwnerOrigin(owner, rectangle, new(rectangle.X, rectangle.Y, 1100, 900)),
            "a tab-reveal beacon union that walks toward the IDE origin is rejected; a same-origin shrink is not");
        var client = WebviewSurface.ClientOrigin(860, 916);
        WebviewSurface.ReadOwnerInsets(client, new(0, 0, 860, 916), out var cx, out var cy, out var cr, out var cb);
        check(cx == 0 && cy == 0 && cr == 0 && cb == 0
            && WebviewSurface.ApplyOwnerInsets(WebviewSurface.ClientOrigin(800, 900), 0, 0, 0, 0)
                == new WebviewSurface.ScreenBounds(0, 0, 800, 900),
            "child-hosted follow uses client origin insets, not screen overlay coordinates");
        check(WebviewSurface.IsTightSlot(0, 0, 966, 916, 966, 916)
            && !WebviewSurface.IsTightSlot(214, 102, 966, 916, 1920, 1080),
            "a workbench GPU surface offset is not a tight editor container");
        check(WebviewSurface.PhysicalFromCss(1107, 1) == 1107 && WebviewSurface.PhysicalFromCss(800, 1.5) == 1200,
            "CSS idle mapping uses the reported pixel ratio");
        check(WebviewSurface.TryCssIdleNested(new(0, 0, 1935, 1056), 222, 110, 807, 853, 1, out var cssIdle)
            && cssIdle == new WebviewSurface.ScreenBounds(222, 110, 807, 853),
            "CSS idle keeps beacon left/top insets when the editor pane shrinks inside an unchanged owner");
        check(WebviewSurface.IsTrustedHow("uia-beacon-corners-cache")
            && !WebviewSurface.IsTrustedHow("webview-size-match")
            && !WebviewSurface.IsTrustedHow("uia-js-fallback"),
            "only token-derived native geometry is accepted as trusted");
        check(!WebviewSurface.TryBounds(new[] { double.NaN, 0, 800, 600 }, out _, out _, out _, out _)
            && !WebviewSurface.TryBounds(new[] { (double)int.MaxValue, 0, 800, 600 }, out _, out _, out _, out _),
            "non-finite and overflowing provider bounds are rejected");
        check(WebviewSurface.TryBounds(new[] { -10.25, 20.25, 800.5, 600.5 }, out var x, out var y, out var width, out var height)
            && x == -11 && y == 20 && width == 802 && height == 601,
            "fractional provider rectangles round outward without clipping");

        AppContext.SetSwitch("OfficeInEditor.Diagnostics", false);
        var before = DebugNdjson.Snapshot().Length;
        DebugNdjson.Write("test", "test", "disabled", new { secret = "document content" });
        DebugNdjson.Probe("test", "test", "disabled", new { secret = "document content" });
        check(DebugNdjson.Snapshot().Length == before, "both diagnostic entry points are disabled by default");
        try
        {
            AppContext.SetSwitch("OfficeInEditor.Diagnostics", true);
            Parallel.For(0, 10000, i =>
            {
                if (i % 2 == 0) DebugNdjson.Write("test", "test", "stage", new { secret = "document content" });
                else DebugNdjson.Probe("test", "test", "stage", new { secret = "document content" });
            });
            check(DebugNdjson.Snapshot().Length == 64
                && !System.Text.Json.JsonSerializer.Serialize(DebugNdjson.Snapshot()).Contains("document content", StringComparison.Ordinal),
                "concurrent diagnostic floods stay bounded and discard payloads through both entry points");
            DebugNdjson.Probe(new string('s', 4096), "private/path", "private\ntext", new { secret = "document content" });
            var diagnosticJson = System.Text.Json.JsonSerializer.Serialize(DebugNdjson.Snapshot());
            check(diagnosticJson.Length < 32768 && diagnosticJson.Contains("invalid-stage", StringComparison.Ordinal)
                && !diagnosticJson.Contains(new string('s', 256), StringComparison.Ordinal)
                && !diagnosticJson.Contains("private", StringComparison.Ordinal),
                "diagnostic stages reject oversized or non-identifier strings instead of retaining arbitrary text");
        }
        finally { AppContext.SetSwitch("OfficeInEditor.Diagnostics", false); }
    }
}
