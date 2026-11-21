using System.Text.Json.Nodes;
using OfficeInEditor.Core;

static class FocusContractTests
{
    const string Surface = "oie-surface-0123456789abcdef0123456789abcdef";
    const string Session = "fedcba9876543210fedcba9876543210";

    public static int Run(Action<bool, string> check)
    {
        var before = 0;
        void Check(bool condition, string name) { check(condition, name); before++; }
        bool Rejects(Action action, string code = "invalid_params")
        {
            try { action(); return false; } catch (CoreFault fault) { return fault.Code == code; }
        }
        JsonNode Node(string json) => JsonNode.Parse(json)!;

        var open = FocusContract.ParseOpen(Node($$"""{"path":"C:\\x\\a.docx","surfaceToken":"{{Surface}}","editorPid":42,"cssWidth":800,"cssHeight":600,"pixelRatio":1.25}"""));
        Check(open.PixelPolicy == "always" && open.EditorPid == 42 && open.Size.PixelRatio == 1.25, "focus open defaults pixel policy");
        Check(Rejects(() => FocusContract.ParseOpen(Node($$"""{"path":"a.docx","surfaceToken":"oie-surface-XYZ","editorPid":42,"cssWidth":800,"cssHeight":600,"pixelRatio":1}"""))), "focus open rejects malformed surface token");
        Check(Rejects(() => FocusContract.ParseOpen(Node($$"""{"path":"a.docx","surfaceToken":"{{Surface}}","editorPid":42,"cssWidth":8,"cssHeight":600,"pixelRatio":1}"""))), "focus open rejects tiny surface");
        Check(Rejects(() => FocusContract.ParseOpen(Node($$"""{"path":"a.docx","surfaceToken":"{{Surface}}","editorPid":42,"cssWidth":800,"cssHeight":600,"pixelRatio":1,"pixelPolicy":"fast"}"""))), "focus open rejects unknown policy");

        var activate = FocusContract.ParseActivate(Node($$$"""{"sessionId":"{{{Session}}}","surfaceToken":"{{{Surface}}}","activationId":3,"trusted":true,"point":{"x":0.5,"y":1}}"""));
        Check(activate.ActivationId == 3 && activate.PointX == 0.5 && activate.PointY == 1, "focus activate keeps normalized point");
        Check(Rejects(() => FocusContract.ParseActivate(Node($$"""{"sessionId":"{{Session}}","surfaceToken":"{{Surface}}","activationId":3}""")), "untrusted_activation"), "focus activate requires trusted click");
        Check(Rejects(() => FocusContract.ParseActivate(Node($$$"""{"sessionId":"{{{Session}}}","surfaceToken":"{{{Surface}}}","activationId":3,"trusted":true,"point":{"x":1.5,"y":0}}"""))), "focus activate rejects point outside surface");

        var wheel = FocusContract.ParseWheel(Node($$$"""{"sessionId":"{{{Session}}}","surfaceToken":"{{{Surface}}}","trusted":true,"point":{"x":0.2,"y":0.8},"deltaY":-120,"shift":true}"""));
        Check(wheel.DeltaY == -120 && wheel.DeltaX == 0 && wheel.Shift && !wheel.Ctrl && wheel.PointY == 0.8, "focus wheel keeps deltas, point and modifiers");
        Check(Rejects(() => FocusContract.ParseWheel(Node($$$"""{"sessionId":"{{{Session}}}","surfaceToken":"{{{Surface}}}","point":{"x":0.2,"y":0.8},"deltaY":-120}""")), "untrusted_activation"), "focus wheel requires a user wheel event");
        Check(Rejects(() => FocusContract.ParseWheel(Node($$$"""{"sessionId":"{{{Session}}}","surfaceToken":"{{{Surface}}}","trusted":true,"point":{"x":0.2,"y":0.8},"deltaY":99999}"""))), "focus wheel bounds the delta");
        Check(Rejects(() => FocusContract.ParseWheel(Node($$$"""{"sessionId":"{{{Session}}}","surfaceToken":"{{{Surface}}}","trusted":true,"point":{"x":0.2,"y":0.8},"deltaY":0}"""))), "focus wheel rejects an empty step");

        Check(Rejects(() => FocusContract.ParseLayout(Node($$"""{"sessionId":"{{Session}}","surfaceToken":"{{Surface}}","cssWidth":800,"cssHeight":600,"pixelRatio":1}"""))), "focus layout requires visibility");
        Check(FocusContract.ParseClose(Node($$"""{"sessionId":"{{Session}}","surfaceToken":"{{Surface}}"}""")).Decision is null, "focus close without decision");
        Check(Rejects(() => FocusContract.ParseClose(Node($$"""{"sessionId":"{{Session}}","surfaceToken":"{{Surface}}","decision":"maybe"}"""))), "focus close rejects unknown decision");

        Check(FocusContract.StreamPipeName("office-in-editor-v1-0123456789abcdef", Session) == "office-in-editor-v1-0123456789abcdef-focus-" + Session, "focus stream pipe shares broker prefix");
        Check(Rejects(() => FocusContract.ParseHello("""{"v":1,"token":"short","surfaceToken":"x"}""")), "focus stream hello rejects bad token");

        var eventBytes = FocusContract.EncodeEvent(FocusContract.StateEvent("editing", "foreground", 4, 2));
        var consumed = FocusContract.TryDecode(eventBytes, out var decoded);
        Check(consumed == eventBytes.Length && decoded is FocusContract.EventMessage { Value: var state } &&
            state["state"]!.GetValue<string>() == "editing" && state["cause"]!.GetValue<string>() == "foreground", "focus stream event round trip");
        Check(FocusContract.TryDecode(eventBytes.AsSpan(0, eventBytes.Length - 1), out _) == 0, "focus stream waits for complete message");

        var image = new byte[] { 0x89, 0x50, 0x4e, 0x47, 1, 2, 3 };
        var frame = FocusContract.EncodeFrame(new FocusContract.FrameHeader(7, 2, 640, 480, "png"), image)!;
        var both = eventBytes.Concat(frame).ToArray();
        var first = FocusContract.TryDecode(both, out _);
        FocusContract.TryDecode(both.AsSpan(first), out var frameMessage);
        Check(frameMessage is FocusContract.FrameMessage { Header: { FrameSeq: 7, ParkId: 2, Width: 640 } } f && f.Image.SequenceEqual(image), "focus stream frame round trip after event");
        Check(FocusContract.EncodeFrame(new FocusContract.FrameHeader(8, 0, 1, 1, "png"), new byte[FocusContract.MaxStreamMessageBytes]) is null, "focus stream drops oversize frame instead of truncating");
        var jpeg = FocusContract.EncodeFrame(new FocusContract.FrameHeader(9, 3, 320, 200, "jpeg"), new byte[] { 0xff, 0xd8, 0xff })!;
        FocusContract.TryDecode(jpeg, out var jpegMessage);
        Check(jpegMessage is FocusContract.FrameMessage { Header: { Format: "jpeg", FrameSeq: 9 } }, "focus stream carries jpeg frames");
        var refused = false;
        try { FocusContract.EncodeFrame(new FocusContract.FrameHeader(10, 3, 1, 1, "bmp"), new byte[] { 1 }); } catch (ArgumentException) { refused = true; }
        Check(refused, "focus stream refuses unknown frame formats");
        Check(Rejects(() => FocusContract.TryDecode(new byte[] { 0xff, 0xff, 0xff, 0x7f, 1 }, out _), "invalid_stream"), "focus stream rejects oversize length");
        return before;
    }
}
