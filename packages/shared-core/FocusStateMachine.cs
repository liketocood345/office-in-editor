using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

/// <summary>Physical screen rectangle (left, top, right, bottom).</summary>
internal readonly record struct FocusRect(int Left, int Top, int Right, int Bottom)
{
    public int Width => Right - Left;
    public int Height => Bottom - Top;
    public bool IsEmpty => Width <= 0 || Height <= 0;
}

internal enum FocusOwner { None, Office, Editor, Other }

/// <summary>Window operations for one focus session. Real calls are in <see cref="Win32FocusWindowOps"/>; tests use a fake.</summary>
internal interface IFocusWindowOps
{
    bool OfficeAlive { get; }
    /// <summary>False while Office shows an application-modal dialog.</summary>
    bool OfficeEnabled { get; }
    /// <summary>Owner of the foreground root; Office-owned dialogs count as <see cref="FocusOwner.Office"/>.</summary>
    FocusOwner Foreground { get; }
    bool EditorIconic { get; }
    int VirtualScreenLeft { get; }
    FocusRect OfficeRect { get; }
    /// <summary>Physical bounds of the tab surface, or null when the tab is hidden or its bounds cannot be read.</summary>
    FocusRect? SurfaceBounds();
    /// <summary>Moves the visible frame of the Office window to <paramref name="target"/>; true when within 2 px.</summary>
    bool Fit(FocusRect target);
    /// <summary>Puts the Office window back where it was before the session took it.</summary>
    void RestoreOriginal();
    bool RequestOfficeFocus();
    bool RequestEditorFocus();
    /// <summary>Moves a foreground Office-owned dialog that sits off-screen into <paramref name="area"/>.</summary>
    void RecenterOwnedDialog(FocusRect area);
    /// <summary>Shows an enabled Office-owned modal dialog in <paramref name="area"/> and focuses it; true when one was found.</summary>
    bool RevealModal(FocusRect area);
    /// <summary>Places the caret at a screen point inside the docked Office window.</summary>
    void PlaceCaret(int x, int y);
    /// <summary>
    /// Posts a wheel step to the content window under the normalized point of the parked window. Office handles it
    /// without becoming active (measured for Word, Excel and PowerPoint). False when no target was found.
    /// </summary>
    bool PostWheel(double pointX, double pointY, int deltaY, int deltaX, bool ctrl, bool shift);
    /// <summary>
    /// Shows or hides the Office window's taskbar button. Hosted in a tab, Office is part of the editor and has no
    /// button of its own. Called every tick with the wanted state; the implementation re-applies a hidden state
    /// periodically, since Explorer can add the button back.
    /// </summary>
    void SetTaskbarButton(bool visible);
    (int X, int Y)? CursorPosition { get; }
}

/// <summary>Frame source for the parked window. The capture subprocess implements it.</summary>
internal interface IFocusCapture
{
    bool Available { get; }
    /// <summary>Starts capture without publishing frames (pre-warm).</summary>
    void Prepare(string reason);
    /// <summary>The window is parked: publish frames tagged with <paramref name="parkId"/>.</summary>
    void Publish(int parkId);
    void Stop();
    /// <summary>Frame-rate cap; 0 pauses sending while keeping capture ready (the tab is hidden).</summary>
    void SetFrameRate(int fps);
}

internal sealed class NoFocusCapture : IFocusCapture
{
    public static readonly NoFocusCapture Instance = new();
    public bool Available => false;
    public void Prepare(string reason) { }
    public void Publish(int parkId) { }
    public void Stop() { }
    public void SetFrameRate(int fps) { }
}

/// <summary>
/// Focus-hybrid state machine, ported from the earlier focus-hybrid lab prototype. Every call must come from the
/// focus dispatcher thread. Office is visible over the tab only while it owns the foreground; otherwise it is
/// parked left of the virtual screen and the tab shows frames. A parked window that takes the foreground by any
/// route (Alt+Tab, taskbar, its own dialog) is adopted into editing (D1), so it is never focused off-screen.
/// </summary>
internal sealed class FocusStateMachine
{
    /// <summary>
    /// Short settle before the first park. A cold-started frame is already out of view (pre-parked when it appeared),
    /// the tab bounds and Office modal dialogs have their own checks, so this no longer needs to cover Office startup.
    /// </summary>
    public const int StartupDelayMs = 300;
    public const int ParkRetryMs = 500;
    /// <summary>
    /// After a layout change the tab's UIA bounds can lag the reported CSS size, so the first read may fail or still
    /// show the old rectangle. Bounds are re-read on every tick for this long and the window follows any change.
    /// </summary>
    public const int LayoutSettleMs = 1500;
    /// <summary>Frame rate for a tab the user can see. A hidden tab, or a minimized editor, gets 0 (paused).</summary>
    public const int ForegroundFps = 30;
    const int PointerThresholdPx = 4;
    int appliedFps = -1;
    readonly IFocusWindowOps ops;
    readonly IFocusCapture capture;
    readonly Action<JsonObject> emit;
    readonly long earliestInit;
    long nextParkRetry, layoutSettleUntil;
    bool initialized, editing, parked, parkPending, modalReported, layoutUnavailableReported, captureReported, lostReported;
    // Office was put back where it was (hidden tab or minimized editor): it is a normal window and keeps its button.
    bool standalone;
    bool surfaceVisible = true;
    FocusRect pane;
    (int X, int Y)? pointerBaseline;
    bool pointerWatch;

    public FocusStateMachine(IFocusWindowOps ops, IFocusCapture capture, Action<JsonObject> emit, string pixelPolicy, long now)
    {
        this.ops = ops; this.capture = capture; this.emit = emit;
        PixelPolicy = FocusContract.PixelPolicies.Contains(pixelPolicy) ? pixelPolicy : "always";
        earliestInit = now + StartupDelayMs;
        State = "releasing"; Cause = "open";
    }

    public string State { get; private set; }
    public string Cause { get; private set; }
    public int ActivationId { get; private set; }
    public int ParkId { get; private set; }
    public string PixelPolicy { get; private set; }
    public bool AdoptForeground { get; set; } = true;
    public bool Lost => State == "lost";
    public JsonObject Snapshot() => FocusContract.StateEvent(State, Cause, ActivationId, ParkId);

    /// <summary>Startup poll, watchdog for missed foreground events, park retry and pointer watch. Called every 100 ms.</summary>
    public void Tick(long now)
    {
        if (Lost) return;
        if (!ops.OfficeAlive) { MarkLost(); return; }
        ops.SetTaskbarButton(standalone);
        if (!initialized) { TryInitialize(now); return; }
        ApplyFrameRate();
        var owner = ops.Foreground;
        ReportModal(!ops.OfficeEnabled && !editing);
        // Some foreground changes (seen with UIA SetFocus) never deliver EVENT_SYSTEM_FOREGROUND.
        if (!editing && owner == FocusOwner.Office) { if (AdoptForeground) Adopt(watchdog: true, now); return; }
        if (editing && owner != FocusOwner.Office) { Release("timer"); return; }
        if (now < layoutSettleUntil && FollowSurface(now)) return;
        if (parkPending && now >= nextParkRetry) { Park(Cause, now); return; }
        if (editing && pointerWatch && ops.CursorPosition is { } point && pointerBaseline is { } baseline &&
            (Math.Abs(point.X - baseline.X) >= PointerThresholdPx || Math.Abs(point.Y - baseline.Y) >= PointerThresholdPx))
        {
            pointerWatch = false;
            capture.Prepare("mouse");
        }
    }

    public void ForegroundChanged(long now)
    {
        if (Lost || !initialized) return;
        if (!ops.OfficeAlive) { MarkLost(); return; }
        var owner = ops.Foreground;
        if (!editing && owner == FocusOwner.Office && AdoptForeground) Adopt(watchdog: false, now);
        else if (editing && owner != FocusOwner.Office) Release("foreground", now);
    }

    /// <summary>A trusted click in the tab. Returns the state after the attempt.</summary>
    public string Activate(int activationId, double? pointX, double? pointY, long now)
    {
        if (Lost) throw new CoreFault("session_lost", "The Office window is gone. Close the tab.");
        if (activationId > ActivationId) ActivationId = activationId;
        if (!initialized) { Notice("activation_denied"); return State; }
        if (editing && ops.Foreground == FocusOwner.Office) return State;
        var surface = ops.SurfaceBounds();
        if (surface is null) { Notice("activation_denied"); return State; }
        pane = surface.Value;
        if (!ops.OfficeEnabled)
        {
            // D3: the modal dialog belongs to a parked window; show it over the tab instead of off-screen.
            ReportModal(true);
            if (ops.RevealModal(pane)) { parked = false; editing = true; parkPending = false; StopPointerWatch(); capture.Stop(); SetState("editing", "dialog"); }
            return State;
        }
        SetState("activating", "click");
        StopPointerWatch();
        capture.Stop();
        if (!ops.RequestOfficeFocus())
        {
            Notice("activation_denied");
            Park("click", now);
            return State;
        }
        parked = false; parkPending = false; standalone = false;
        if (!ops.Fit(pane))
        {
            Notice("alignment_unconfirmed");
            ops.RequestEditorFocus();
            if (ops.Foreground != FocusOwner.Office) { editing = false; Park("click", now); return State; }
        }
        editing = true;
        if (pointX is { } x && pointY is { } y)
            ops.PlaceCaret(pane.Left + (int)Math.Round(x * Math.Max(0, pane.Width - 1)), pane.Top + (int)Math.Round(y * Math.Max(0, pane.Height - 1)));
        SetState("editing", "click");
        PrepareWhileEditing();
        return State;
    }

    /// <summary>
    /// Wheel over the tab while Office is parked: scroll the parked window in place. It stays inactive and parked;
    /// the new pixels come back as frames. While editing, Office covers the tab and gets real wheel input itself.
    /// </summary>
    public bool Wheel(double pointX, double pointY, int deltaY, int deltaX, bool ctrl, bool shift)
    {
        if (Lost || !initialized || editing || !parked || !surfaceVisible)
        {
            DebugTrail.Write("wheel", "rejected", new { lost = Lost, initialized, editing, parked, surfaceVisible, standalone });
            return false;
        }
        return ops.PostWheel(pointX, pointY, deltaY, deltaX, ctrl, shift);
    }

    public void Layout(bool visible, long now)
    {
        if (Lost) return;
        surfaceVisible = visible;
        if (!initialized) return;
        ApplyFrameRate();
        if (!visible)
        {
            layoutSettleUntil = 0;
            if (editing) Release("layout", now);
            return;
        }
        layoutSettleUntil = now + LayoutSettleMs;
        FollowSurface(now);
    }

    /// <summary>Only the tab the user can see is streamed at <see cref="ForegroundFps"/>; others are paused.</summary>
    void ApplyFrameRate()
    {
        var wanted = surfaceVisible && !ops.EditorIconic ? ForegroundFps : 0;
        if (wanted == appliedFps) return;
        appliedFps = wanted;
        capture.SetFrameRate(wanted);
    }

    /// <summary>Re-reads the tab bounds; when they changed, the parked or docked window follows. True when it moved.</summary>
    bool FollowSurface(long now)
    {
        var updated = ops.SurfaceBounds();
        if (updated is null || updated.Value == pane) return false;
        pane = updated.Value;
        if (editing && ops.Foreground == FocusOwner.Office) { ops.Fit(pane); return true; }
        if (parked || parkPending) { Park("layout", now); return true; }
        return false;
    }

    public void SetPolicy(string policy)
    {
        if (!FocusContract.PixelPolicies.Contains(policy)) throw new CoreFault("invalid_params", "Pixel policy is unsupported.");
        if (policy == PixelPolicy) return;
        PixelPolicy = policy;
        if (editing && !Lost) { capture.Stop(); PrepareWhileEditing(); }
    }

    /// <summary>Stops capture and puts Office back at its original position (close, Core shutdown).</summary>
    public void Restore()
    {
        StopPointerWatch();
        capture.Stop();
        if (ops.OfficeAlive) { ops.RestoreOriginal(); ops.SetTaskbarButton(true); }
        parked = false; editing = false; parkPending = false; standalone = true;
    }

    void TryInitialize(long now)
    {
        if (now < earliestInit) return;
        if (!ops.OfficeEnabled) { ReportModal(true); return; }
        var surface = surfaceVisible ? ops.SurfaceBounds() : null;
        if (surface is null)
        {
            if (!layoutUnavailableReported) DebugNdjson.Write("F", "FocusStateMachine.Init", "surface_unavailable", null);
            layoutUnavailableReported = true;
            return;
        }
        pane = surface.Value;
        initialized = true;
        ApplyFrameRate();
        ReportModal(false);
        if (!capture.Available && !captureReported) { captureReported = true; Notice("capture_unavailable"); }
        // Office usually holds the foreground right after it opens the file: dock it into the tab.
        if (ops.Foreground == FocusOwner.Office) Adopt(watchdog: false, now, cause: "open");
        else Park("open", now);
    }

    void Adopt(bool watchdog, long now, string cause = "foreground")
    {
        StopPointerWatch();
        capture.Stop();
        var tabShown = !ops.EditorIconic && surfaceVisible;
        var aligned = false;
        if (tabShown)
        {
            // A fresh read can miss (UIA lag, or the discovery throttle right after a park); the last known tab
            // rectangle is still where the tab is, so dock there rather than leaving Office outside the editor.
            if (ops.SurfaceBounds() is { } current) pane = current;
            if (!pane.IsEmpty) aligned = ops.Fit(pane);
            layoutSettleUntil = Math.Max(layoutSettleUntil, now + LayoutSettleMs);
        }
        // Only a hidden tab or a minimized editor puts Office back where it was before docking.
        if (!aligned) ops.RestoreOriginal();
        standalone = !aligned;
        ops.SetTaskbarButton(standalone);
        parked = false; parkPending = false; editing = true;
        ops.RecenterOwnedDialog(aligned ? pane : ops.OfficeRect);
        DebugNdjson.Write("F", "FocusStateMachine.Adopt", "office_foreground_adopted", new { docked = aligned, watchdog, tabShown, surfaceVisible, paneEmpty = pane.IsEmpty });
        SetState("editing", cause);
        PrepareWhileEditing();
    }

    void Release(string cause, long now = 0)
    {
        editing = false;
        SetState("releasing", cause);
        if (ops.SurfaceBounds() is { } current) pane = current;
        Park(cause, now);
    }

    void Park(string cause, long now)
    {
        StopPointerWatch();
        capture.Prepare("switch");
        standalone = false;
        ops.SetTaskbarButton(false);
        var width = Math.Max(FocusContract.MinSurfacePx, pane.Width);
        var height = Math.Max(FocusContract.MinSurfacePx, pane.Height);
        // Keep at least 128 px (or the original width) between the parked frame and the virtual screen.
        var right = ops.VirtualScreenLeft - Math.Max(128, ops.OfficeRect.Width);
        ops.Fit(new FocusRect(right - width, pane.Top, right, pane.Top + height));
        parked = ops.OfficeRect.Right <= ops.VirtualScreenLeft;
        if (!parked)
        {
            capture.Stop();
            if (!parkPending) Notice("alignment_unconfirmed");
            parkPending = true;
            nextParkRetry = now + ParkRetryMs;
            SetState("releasing", cause);
            return;
        }
        parkPending = false;
        ParkId++;
        // State first: frames of this park are only forwarded once the session reports it parked.
        SetState("parked", cause);
        capture.Publish(ParkId);
    }

    void PrepareWhileEditing()
    {
        StopPointerWatch();
        if (PixelPolicy == "always") capture.Prepare("always");
        else if (PixelPolicy == "mouse") { pointerBaseline = ops.CursorPosition; pointerWatch = pointerBaseline is not null; }
    }

    void StopPointerWatch() { pointerWatch = false; pointerBaseline = null; }

    void ReportModal(bool blocking)
    {
        if (blocking == modalReported) return;
        modalReported = blocking;
        Notice(blocking ? "modal_blocking" : "modal_cleared");
    }

    void MarkLost()
    {
        StopPointerWatch();
        capture.Stop();
        editing = false; parked = false; parkPending = false;
        SetState("lost", "recovery");
        if (!lostReported) { lostReported = true; Notice("office_exited"); emit(FocusContract.ClosedEvent("office_exited")); }
    }

    void SetState(string state, string cause)
    {
        if (state == State && cause == Cause && state != "parked") return;
        State = state; Cause = cause;
        DebugTrail.Write("focus-state", state + ":" + cause, new { ParkId, ActivationId, standalone, surfaceVisible });
        emit(Snapshot());
    }

    void Notice(string code)
    {
        DebugTrail.Write("focus-notice", code);
        emit(FocusContract.NoticeEvent(code));
    }
}
