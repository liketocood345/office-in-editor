using System.Text.Json.Nodes;
using OfficeInEditor.Core;

/// <summary>Focus-hybrid state machine against a fake window layer: no Office, no IDE, no lab machine.</summary>
static class FocusStateMachineTests
{
    sealed class FakeOps : IFocusWindowOps
    {
        public bool OfficeAlive { get; set; } = true;
        public bool OfficeEnabled { get; set; } = true;
        public FocusOwner Foreground { get; set; } = FocusOwner.Editor;
        public bool EditorIconic { get; set; }
        public int VirtualScreenLeft => 0;
        public FocusRect OfficeRect { get; set; } = new(100, 100, 900, 700);
        public FocusRect? Surface { get; set; } = new(200, 150, 1000, 750);
        public bool FitWorks { get; set; } = true;
        public bool FocusWorks { get; set; } = true;
        public bool ModalFound { get; set; } = true;
        public int Restored, Recentered;
        public (int X, int Y)? Caret;
        public (int X, int Y)? CursorPosition { get; set; } = (500, 500);
        public FocusRect? SurfaceBounds() => Surface;
        public bool Fit(FocusRect target) { if (!FitWorks) return false; OfficeRect = target; return true; }
        public void RestoreOriginal() { Restored++; OfficeRect = new(100, 100, 900, 700); }
        public bool RequestOfficeFocus() { if (FocusWorks) Foreground = FocusOwner.Office; return FocusWorks; }
        public bool RequestEditorFocus() { Foreground = FocusOwner.Editor; return true; }
        public void RecenterOwnedDialog(FocusRect area) => Recentered++;
        public bool RevealModal(FocusRect area) { if (ModalFound) Foreground = FocusOwner.Office; return ModalFound; }
        public void PlaceCaret(int x, int y) => Caret = (x, y);
        public List<(double X, double Y, int DeltaY, int DeltaX, bool Ctrl, bool Shift)> Wheels { get; } = [];
        public bool PostWheel(double pointX, double pointY, int deltaY, int deltaX, bool ctrl, bool shift)
        {
            Wheels.Add((pointX, pointY, deltaY, deltaX, ctrl, shift));
            return true;
        }
        public bool Offscreen => OfficeRect.Right <= VirtualScreenLeft;
    }

    sealed class FakeCapture : IFocusCapture
    {
        public bool Available => true;
        public readonly List<string> Calls = [];
        public void Prepare(string reason) => Calls.Add("prepare:" + reason);
        public void Publish(int parkId) => Calls.Add("publish:" + parkId);
        public void Stop() => Calls.Add("stop");
        public void SetFrameRate(int fps) => Calls.Add("fps:" + fps);
    }

    sealed record Rig(FakeOps Ops, FakeCapture Capture, List<JsonObject> Events, FocusStateMachine Machine)
    {
        public IEnumerable<string> States => Events.Where(item => item["type"]!.GetValue<string>() == "state").Select(item => item["state"] + "/" + item["cause"]);
        public bool Noticed(string code) => Events.Any(item => item["type"]!.GetValue<string>() == "notice" && item["code"]!.GetValue<string>() == code);
        public int NoticeCount(string code) => Events.Count(item => item["type"]!.GetValue<string>() == "notice" && item["code"]!.GetValue<string>() == code);
    }

    static Rig Create(string policy = "always", Action<FakeOps>? setup = null, bool initialize = true)
    {
        var ops = new FakeOps();
        setup?.Invoke(ops);
        var capture = new FakeCapture();
        var events = new List<JsonObject>();
        var machine = new FocusStateMachine(ops, capture, events.Add, policy, 0);
        if (initialize) machine.Tick(FocusStateMachine.StartupDelayMs);
        return new Rig(ops, capture, events, machine);
    }

    public static void Run(Action<bool, string> check)
    {
        var early = Create(initialize: false);
        early.Machine.Tick(FocusStateMachine.StartupDelayMs - 1);
        check(early.Machine.State == "releasing" && early.Events.Count == 0, "focus machine waits for the startup delay before parking");

        var rig = Create();
        check(rig.Machine.State == "parked" && rig.Machine.Cause == "open" && rig.Machine.ParkId == 1 && rig.Ops.Offscreen &&
            rig.Capture.Calls.SequenceEqual(["fps:30", "prepare:switch", "publish:1"]), "focus machine parks off-screen on open and publishes frames at 30 fps");

        // Only the tab the user can see is streamed at 30 fps; a hidden tab or a minimized editor pauses frames.
        var rate = Create();
        rate.Machine.Layout(false, 2000);
        check(rate.Capture.Calls[^1] == "fps:0", "focus hidden tab pauses frames");
        rate.Machine.Layout(true, 2100);
        check(rate.Capture.Calls.Count(call => call == "fps:30") == 2, "focus visible tab resumes 30 fps");
        var callsBefore = rate.Capture.Calls.Count;
        rate.Machine.Tick(2200);
        check(rate.Capture.Calls.Count == callsBefore, "focus frame rate is only sent when it changes");
        rate.Ops.EditorIconic = true;
        rate.Machine.Tick(2300);
        check(rate.Capture.Calls[^1] == "fps:0", "focus minimized editor pauses frames");
        check(rig.Ops.OfficeRect.Width == 800 && rig.Ops.OfficeRect.Right <= -128, "parked frame keeps the tab size and a gap left of the virtual screen");

        var state = rig.Machine.Activate(3, 0.5, 0, 2000);
        check(state == "editing" && rig.Ops.OfficeRect == rig.Ops.Surface && rig.Machine.ActivationId == 3 &&
            rig.States.Skip(1).SequenceEqual(["activating/click", "editing/click"]), "focus click docks Office over the tab");
        check(rig.Ops.Caret == (600, 150) && rig.Capture.Calls[^1] == "prepare:always", "focus click places the caret and prewarms capture under policy always");

        rig.Ops.Foreground = FocusOwner.Editor;
        rig.Machine.ForegroundChanged(2100);
        check(rig.Machine.State == "parked" && rig.Machine.Cause == "foreground" && rig.Machine.ParkId == 2 && rig.Ops.Offscreen,
            "focus leaving Office parks it again with a new park id");

        // D1: Alt+Tab or the taskbar raises the parked window and no foreground event arrives.
        rig.Ops.Foreground = FocusOwner.Office;
        rig.Machine.Tick(2200);
        check(rig.Machine.State == "editing" && rig.Machine.Cause == "foreground" && !rig.Ops.Offscreen && rig.Ops.Recentered > 0,
            "focus watchdog adopts a parked window that took the foreground (D1)");
        rig.Ops.Foreground = FocusOwner.Other;
        rig.Machine.Tick(2300);
        check(rig.Machine.State == "parked" && rig.Machine.Cause == "timer", "focus watchdog parks when a foreground loss was missed");

        var hidden = Create(setup: ops => ops.EditorIconic = true);
        hidden.Ops.Foreground = FocusOwner.Office;
        hidden.Machine.ForegroundChanged(2000);
        check(hidden.Machine.State == "editing" && hidden.Ops.Restored == 1 && !hidden.Ops.Offscreen, "focus adoption with a minimized editor restores Office where it was");

        // Office raises itself right after the first park (Word does once the file finishes loading) and the bounds
        // read misses at that instant: Office must dock over the last known tab rectangle, not go back outside.
        var selfRaise = Create();
        var knownPane = selfRaise.Ops.Surface!.Value;
        selfRaise.Ops.Surface = null;
        selfRaise.Ops.Foreground = FocusOwner.Office;
        selfRaise.Machine.Tick(1400);
        check(selfRaise.Machine.State == "editing" && selfRaise.Ops.Restored == 0 && selfRaise.Ops.OfficeRect == knownPane,
            "focus adoption with a missed bounds read docks at the last known tab, not outside the editor");

        // Wheel over a parked tab scrolls the parked window in place; while editing Office gets real input itself.
        var wheel = Create();
        check(wheel.Machine.Wheel(0.5, 0.25, -240, 0, false, true) && wheel.Ops.Wheels.Single() == (0.5, 0.25, -240, 0, false, true) &&
            wheel.Machine.State == "parked" && wheel.Ops.Foreground == FocusOwner.Editor, "focus wheel scrolls the parked window without activating it");
        wheel.Machine.Layout(false, 2000);
        check(!wheel.Machine.Wheel(0.5, 0.5, -120, 0, false, false) && wheel.Ops.Wheels.Count == 1, "focus wheel is ignored for a hidden tab");
        wheel.Machine.Layout(true, 2100);
        wheel.Machine.Activate(1, null, null, 2200);
        check(!wheel.Machine.Wheel(0.5, 0.5, -120, 0, false, false) && wheel.Ops.Wheels.Count == 1, "focus wheel is not forwarded while editing");

        var startFocused = Create(setup: ops => ops.Foreground = FocusOwner.Office);
        check(startFocused.Machine.State == "editing" && startFocused.Machine.Cause == "open" && startFocused.Machine.ParkId == 0,
            "focus open docks Office when it already holds the foreground");

        var denied = Create(setup: ops => ops.FocusWorks = false);
        denied.Machine.Activate(1, null, null, 2000);
        check(denied.Machine.State == "parked" && denied.Noticed("activation_denied") && denied.Ops.Offscreen, "focus denied activation stays parked with a notice");

        var modal = Create();
        modal.Ops.OfficeEnabled = false;
        modal.Machine.Tick(1400);
        modal.Machine.Tick(1500);
        check(modal.NoticeCount("modal_blocking") == 1, "focus reports a parked modal dialog once");
        modal.Machine.Activate(2, null, null, 1600);
        check(modal.Machine.State == "editing" && modal.Machine.Cause == "dialog", "focus click on a modal-blocked tab shows the dialog over the tab (D3)");
        modal.Ops.OfficeEnabled = true;
        modal.Machine.Tick(1700);
        check(modal.Noticed("modal_cleared"), "focus reports the modal dialog cleared");

        var layout = Create();
        layout.Machine.Activate(1, null, null, 2000);
        layout.Machine.Layout(false, 2100);
        check(layout.Machine.State == "parked" && layout.Machine.Cause == "layout", "focus hidden tab while editing returns to parked");
        layout.Ops.Surface = new FocusRect(200, 150, 700, 550);
        layout.Machine.Layout(true, 2200);
        check(layout.Machine.ParkId == 3 && layout.Ops.OfficeRect.Width == 500, "focus resized tab re-parks at the new size");

        // Resize while parked: UIA bounds lag the CSS size, so the first read fails. The parked window (the capture
        // source) must still follow, or the tab shows a stretched old-size frame.
        var lagging = Create();
        var firstPark = lagging.Machine.ParkId;
        lagging.Ops.Surface = null;
        lagging.Machine.Layout(true, 2000);
        check(lagging.Machine.ParkId == firstPark, "focus layout without readable bounds keeps the current park");
        lagging.Ops.Surface = new FocusRect(200, 150, 1200, 850);
        lagging.Machine.Tick(2100);
        check(lagging.Machine.ParkId == firstPark + 1 && lagging.Ops.OfficeRect.Width == 1000 && lagging.Ops.OfficeRect.Height == 700 && lagging.Ops.Offscreen,
            "focus re-parks at the new size once lagging bounds arrive");
        lagging.Ops.Surface = new FocusRect(200, 150, 900, 650);
        lagging.Machine.Tick(2000 + FocusStateMachine.LayoutSettleMs + 100);
        check(lagging.Ops.OfficeRect.Width == 1000, "focus stops re-reading bounds after the settle window");
        var dockedResize = Create();
        dockedResize.Machine.Activate(1, null, null, 2000);
        dockedResize.Ops.Surface = null;
        dockedResize.Machine.Layout(true, 2100);
        dockedResize.Ops.Surface = new FocusRect(200, 150, 1100, 800);
        dockedResize.Machine.Tick(2200);
        check(dockedResize.Machine.State == "editing" && dockedResize.Ops.OfficeRect == new FocusRect(200, 150, 1100, 800),
            "focus docked window follows lagging bounds while editing");

        var unconfirmed = Create(setup: ops => ops.FitWorks = false);
        check(unconfirmed.Machine.State == "releasing" && unconfirmed.Noticed("alignment_unconfirmed") && unconfirmed.Capture.Calls[^1] == "stop",
            "focus unconfirmed park reports releasing and stops capture");
        unconfirmed.Ops.FitWorks = true;
        unconfirmed.Machine.Tick(FocusStateMachine.StartupDelayMs + FocusStateMachine.ParkRetryMs);
        check(unconfirmed.Machine.State == "parked" && unconfirmed.Ops.Offscreen, "focus retries an unconfirmed park");

        var mouse = Create("mouse");
        mouse.Machine.Activate(1, null, null, 2000);
        check(!mouse.Capture.Calls.Contains("prepare:mouse"), "focus mouse policy waits for pointer movement");
        mouse.Ops.CursorPosition = (502, 501);
        mouse.Machine.Tick(2100);
        check(!mouse.Capture.Calls.Contains("prepare:mouse"), "focus mouse policy ignores jitter below 4 px");
        mouse.Ops.CursorPosition = (510, 500);
        mouse.Machine.Tick(2200);
        check(mouse.Capture.Calls[^1] == "prepare:mouse", "focus mouse policy prewarms after pointer movement");

        var switching = Create("switch");
        switching.Machine.Activate(1, null, null, 2000);
        check(switching.Capture.Calls[^1] == "stop", "focus switch policy does not capture while editing");
        switching.Machine.SetPolicy("always");
        check(switching.Capture.Calls[^1] == "prepare:always" && switching.Machine.PixelPolicy == "always", "focus policy change applies while editing");

        var exited = Create();
        exited.Ops.OfficeAlive = false;
        exited.Machine.Tick(2000);
        check(exited.Machine.State == "lost" && exited.Noticed("office_exited") &&
            exited.Events.Any(item => item["type"]!.GetValue<string>() == "closed" && item["reason"]!.GetValue<string>() == "office_exited"), "focus Office exit marks the session lost");
        var threw = false;
        try { exited.Machine.Activate(9, null, null, 2100); } catch (CoreFault fault) { threw = fault.Code == "session_lost"; }
        check(threw, "focus lost session rejects activation");

        var restore = Create();
        restore.Machine.Restore();
        check(restore.Ops.Restored == 1 && !restore.Ops.Offscreen && restore.Capture.Calls[^1] == "stop", "focus restore puts Office back and stops capture");
        check(rig.Events.All(item => !item.ToJsonString().Contains(":\\", StringComparison.Ordinal)), "focus events carry no paths");
    }
}
