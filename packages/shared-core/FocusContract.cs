using System.Buffers.Binary;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace OfficeInEditor.Core;

/// <summary>
/// Wire contract for the focus-hybrid route: the native Office window is shown over the tab only while it holds
/// the foreground, and the tab shows captured frames otherwise. Control uses the v1 request pipe (focus.* methods);
/// state events and frames use one per-session stream pipe. See packages/contracts/IMPLEMENTATION.md.
/// </summary>
public static partial class FocusContract
{
    public const string Open = "focus.open";
    public const string Activate = "focus.activate";
    public const string Layout = "focus.layout";
    public const string Policy = "focus.policy";
    public const string Close = "focus.close";
    public const string Status = "focus.status";
    /// <summary>Wheel over a parked tab, forwarded to the parked Office window without activating it (user-approved S05 exception, ADR 0003).</summary>
    public const string Wheel = "focus.wheel";
    public static readonly IReadOnlySet<string> Methods = new HashSet<string>(StringComparer.Ordinal) { Open, Activate, Layout, Policy, Close, Status, Wheel };
    /// <summary>Bound for one forwarded wheel step, in WHEEL_DELTA units (120 = one notch).</summary>
    public const int MaxWheelDelta = 2400;

    public static readonly IReadOnlySet<string> States = new HashSet<string>(StringComparer.Ordinal) { "parked", "activating", "editing", "releasing", "lost" };
    public static readonly IReadOnlySet<string> Causes = new HashSet<string>(StringComparer.Ordinal) { "open", "click", "foreground", "layout", "timer", "policy", "dialog", "close", "recovery" };
    public static readonly IReadOnlySet<string> PixelPolicies = new HashSet<string>(StringComparer.Ordinal) { "always", "switch", "mouse" };
    public static readonly IReadOnlySet<string> NoticeCodes = new HashSet<string>(StringComparer.Ordinal)
        { "modal_blocking", "modal_cleared", "alignment_unconfirmed", "activation_denied", "office_exited", "capture_unavailable" };
    /// <summary>Frame image formats: jpeg for the 30 fps stream of a visible tab, png when lossless frames are wanted.</summary>
    public static readonly IReadOnlySet<string> FrameFormats = new HashSet<string>(StringComparer.Ordinal) { "png", "jpeg" };
    public static readonly IReadOnlySet<string> CloseDecisions = new HashSet<string>(StringComparer.Ordinal) { "save", "discard" };

    public const int MinSurfacePx = 32;
    public const int MaxSurfacePx = 16384;
    public const double MinPixelRatio = 0.5;
    public const double MaxPixelRatio = 8;
    /// <summary>Upper bound for one stream message (kind byte + body). Frames above it are dropped, never truncated.</summary>
    public const int MaxStreamMessageBytes = 32 * 1024 * 1024;
    public const int MaxFrameHeaderBytes = 4096;
    public const byte EventKind = 1;
    public const byte FrameKind = 2;

    [GeneratedRegex("^oie-surface-[a-f0-9]{32}$")] private static partial Regex SurfaceTokenPattern();
    [GeneratedRegex("^[a-f0-9]{32}$")] private static partial Regex SessionIdPattern();
    [GeneratedRegex("^[a-f0-9]{64}$")] private static partial Regex StreamTokenPattern();

    public sealed record OpenRequest(string Path, string SurfaceToken, int EditorPid, string PixelPolicy, SurfaceSize Size);
    public sealed record ActivateRequest(string SessionId, string SurfaceToken, int ActivationId, double? PointX, double? PointY);
    public sealed record LayoutRequest(string SessionId, string SurfaceToken, bool Visible, SurfaceSize Size);
    public sealed record PolicyRequest(string SessionId, string PixelPolicy);
    public sealed record CloseRequest(string SessionId, string SurfaceToken, string? Decision);
    public sealed record WheelRequest(string SessionId, string SurfaceToken, double PointX, double PointY, int DeltaY, int DeltaX, bool Ctrl, bool Shift);
    public sealed record SurfaceSize(int CssWidth, int CssHeight, double PixelRatio);
    public sealed record StreamHello(string Token, string SurfaceToken);
    public sealed record FrameHeader(long FrameSeq, int ParkId, int Width, int Height, string Format);

    public static OpenRequest ParseOpen(JsonNode? parameters)
    {
        var pid = Wire.Number(parameters, "editorPid");
        if (pid <= 0) throw Invalid();
        var policy = Wire.Text(parameters, "pixelPolicy", 16, false);
        if (policy.Length == 0) policy = "always";
        else if (!PixelPolicies.Contains(policy)) throw Invalid();
        return new OpenRequest(Wire.Text(parameters, "path", 32767), SurfaceToken(parameters), pid, policy, Size(parameters));
    }

    public static ActivateRequest ParseActivate(JsonNode? parameters)
    {
        // Only a trusted pointer event in the tab may move the native window over the editor.
        if (Wire.OptionalBoolean(parameters, "trusted") != true) throw new CoreFault("untrusted_activation", "Activation requires a user click.");
        var activationId = Wire.Number(parameters, "activationId");
        if (activationId <= 0) throw Invalid();
        double? x = null, y = null;
        if (parameters?["point"] is JsonObject point)
        {
            x = Fraction(point, "x");
            y = Fraction(point, "y");
        }
        else if (parameters?["point"] is not null) throw Invalid();
        return new ActivateRequest(SessionId(parameters), SurfaceToken(parameters), activationId, x, y);
    }

    /// <summary>
    /// A wheel step from a trusted wheel event in the tab. Deltas use Win32 conventions: <c>deltaY</c> positive scrolls
    /// up (away from the user), <c>deltaX</c> positive scrolls right; 120 is one notch.
    /// </summary>
    public static WheelRequest ParseWheel(JsonNode? parameters)
    {
        if (Wire.OptionalBoolean(parameters, "trusted") != true) throw new CoreFault("untrusted_activation", "Wheel forwarding requires a user wheel event.");
        if (parameters?["point"] is not JsonObject point) throw Invalid();
        var deltaY = Wire.Number(parameters, "deltaY");
        var deltaX = Wire.Number(parameters, "deltaX");
        if (Math.Abs(deltaY) > MaxWheelDelta || Math.Abs(deltaX) > MaxWheelDelta || deltaY == 0 && deltaX == 0) throw Invalid();
        return new WheelRequest(SessionId(parameters), SurfaceToken(parameters), Fraction(point, "x"), Fraction(point, "y"), deltaY, deltaX,
            Wire.OptionalBoolean(parameters, "ctrl") == true, Wire.OptionalBoolean(parameters, "shift") == true);
    }

    public static LayoutRequest ParseLayout(JsonNode? parameters) =>
        new(SessionId(parameters), SurfaceToken(parameters), Wire.OptionalBoolean(parameters, "visible") ?? throw Invalid(), Size(parameters));

    public static PolicyRequest ParsePolicy(JsonNode? parameters)
    {
        var policy = Wire.Text(parameters, "pixelPolicy", 16);
        if (!PixelPolicies.Contains(policy)) throw Invalid();
        return new PolicyRequest(SessionId(parameters), policy);
    }

    public static CloseRequest ParseClose(JsonNode? parameters)
    {
        var decision = Wire.Text(parameters, "decision", 16, false);
        if (decision.Length > 0 && !CloseDecisions.Contains(decision)) throw Invalid();
        return new CloseRequest(SessionId(parameters), SurfaceToken(parameters), decision.Length == 0 ? null : decision);
    }

    public static string ParseStatus(JsonNode? parameters) => SessionId(parameters);

    public static StreamHello ParseHello(string line)
    {
        JsonNode? node;
        try { node = JsonNode.Parse(line); } catch (Exception) { throw Invalid(); }
        if (Wire.Number(node, "v") != 1) throw Invalid();
        var token = Wire.Text(node, "token", 64);
        if (!StreamTokenPattern().IsMatch(token)) throw Invalid();
        return new StreamHello(token, SurfaceToken(node));
    }

    public static string NewSessionId() => Wire.NewToken()[..32];

    /// <summary>Per-session stream pipe; shares the per-user prefix so the pipe ACL rules match the broker pipe.</summary>
    public static string StreamPipeName(string brokerPipeName, string sessionId)
    {
        if (!SessionIdPattern().IsMatch(sessionId)) throw Invalid();
        return brokerPipeName + "-focus-" + sessionId;
    }

    public static JsonObject StateEvent(string state, string cause, int activationId, int parkId)
    {
        if (!States.Contains(state) || !Causes.Contains(cause) || activationId < 0 || parkId < 0) throw new ArgumentException("Unknown focus state.");
        return new JsonObject { ["type"] = "state", ["state"] = state, ["cause"] = cause, ["activationId"] = activationId, ["parkId"] = parkId };
    }

    public static JsonObject NoticeEvent(string code)
    {
        if (!NoticeCodes.Contains(code)) throw new ArgumentException("Unknown focus notice.");
        return new JsonObject { ["type"] = "notice", ["code"] = code };
    }

    public static JsonObject ClosedEvent(string reason)
    {
        if (reason is not ("closed" or "office_exited" or "superseded")) throw new ArgumentException("Unknown close reason.");
        return new JsonObject { ["type"] = "closed", ["reason"] = reason };
    }

    // Stream message: u32 LE length of (kind + body), u8 kind, body.
    // Event body: UTF-8 JSON. Frame body: u16 LE header length, header JSON, image bytes.
    public static byte[] EncodeEvent(JsonObject value)
    {
        var body = Encoding.UTF8.GetBytes(value.ToJsonString());
        return Envelope(EventKind, body.Length, span => body.CopyTo(span));
    }

    public static byte[]? EncodeFrame(FrameHeader header, ReadOnlySpan<byte> image)
    {
        if (!FrameFormats.Contains(header.Format) || header.Width <= 0 || header.Height <= 0 || header.FrameSeq <= 0 || header.ParkId < 0) throw new ArgumentException("Invalid frame header.");
        var json = Encoding.UTF8.GetBytes(new JsonObject
        {
            ["frameSeq"] = header.FrameSeq, ["parkId"] = header.ParkId, ["width"] = header.Width, ["height"] = header.Height, ["format"] = header.Format,
        }.ToJsonString());
        var bodyLength = 2 + json.Length + image.Length;
        if (json.Length > MaxFrameHeaderBytes || 1 + bodyLength > MaxStreamMessageBytes) return null;
        var copy = image.ToArray();
        return Envelope(FrameKind, bodyLength, span =>
        {
            BinaryPrimitives.WriteUInt16LittleEndian(span, (ushort)json.Length);
            json.CopyTo(span[2..]);
            copy.CopyTo(span[(2 + json.Length)..]);
        });
    }

    public abstract record StreamMessage;
    public sealed record EventMessage(JsonObject Value) : StreamMessage;
    public sealed record FrameMessage(FrameHeader Header, byte[] Image) : StreamMessage;

    /// <summary>Decodes one complete message from the start of <paramref name="buffer"/>; returns consumed bytes, or 0 when more input is needed.</summary>
    public static int TryDecode(ReadOnlySpan<byte> buffer, out StreamMessage? message)
    {
        message = null;
        if (buffer.Length < 4) return 0;
        var length = BinaryPrimitives.ReadUInt32LittleEndian(buffer);
        if (length < 1 || length > MaxStreamMessageBytes) throw new CoreFault("invalid_stream", "Stream message length is invalid.");
        if (buffer.Length < 4 + length) return 0;
        var kind = buffer[4];
        var body = buffer.Slice(5, (int)length - 1);
        if (kind == EventKind)
        {
            message = new EventMessage(JsonNode.Parse(body) as JsonObject ?? throw new CoreFault("invalid_stream", "Event is not an object."));
        }
        else if (kind == FrameKind)
        {
            if (body.Length < 2) throw new CoreFault("invalid_stream", "Frame is truncated.");
            var headerLength = BinaryPrimitives.ReadUInt16LittleEndian(body);
            if (headerLength > MaxFrameHeaderBytes || body.Length < 2 + headerLength) throw new CoreFault("invalid_stream", "Frame header is invalid.");
            var header = JsonNode.Parse(body.Slice(2, headerLength));
            var frameSeq = header?["frameSeq"] is JsonValue seq && seq.TryGetValue<long>(out var parsed) ? parsed : 0;
            message = new FrameMessage(new FrameHeader(frameSeq, Wire.Number(header, "parkId", -1),
                Wire.Number(header, "width"), Wire.Number(header, "height"), Wire.Text(header, "format", 16)), body[(2 + headerLength)..].ToArray());
        }
        else throw new CoreFault("invalid_stream", "Unknown stream message kind.");
        return 4 + (int)length;
    }

    static byte[] Envelope(byte kind, int bodyLength, SpanAction write)
    {
        var bytes = new byte[5 + bodyLength];
        BinaryPrimitives.WriteUInt32LittleEndian(bytes, (uint)(1 + bodyLength));
        bytes[4] = kind;
        write(bytes.AsSpan(5));
        return bytes;
    }

    delegate void SpanAction(Span<byte> span);

    static string SurfaceToken(JsonNode? parameters)
    {
        var token = Wire.Text(parameters, "surfaceToken", 80);
        return SurfaceTokenPattern().IsMatch(token) ? token : throw Invalid();
    }

    static string SessionId(JsonNode? parameters)
    {
        var id = Wire.Text(parameters, "sessionId", 32);
        return SessionIdPattern().IsMatch(id) ? id : throw Invalid();
    }

    static SurfaceSize Size(JsonNode? parameters)
    {
        var width = Wire.Number(parameters, "cssWidth");
        var height = Wire.Number(parameters, "cssHeight");
        var ratio = parameters?["pixelRatio"] is JsonValue scalar && scalar.TryGetValue<double>(out var value) ? value : double.NaN;
        if (width < MinSurfacePx || height < MinSurfacePx || width > MaxSurfacePx || height > MaxSurfacePx ||
            !(ratio >= MinPixelRatio && ratio <= MaxPixelRatio)) throw Invalid();
        return new SurfaceSize(width, height, ratio);
    }

    static double Fraction(JsonNode point, string key)
    {
        if (point[key] is not JsonValue scalar || !scalar.TryGetValue<double>(out var value) || !(value >= 0 && value <= 1)) throw Invalid();
        return value;
    }

    static CoreFault Invalid() => new("invalid_params", "A required field is invalid.");
}
