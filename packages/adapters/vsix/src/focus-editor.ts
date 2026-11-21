import { randomBytes } from "node:crypto";
import { MAX_WHEEL_DELTA, errorCode, isRecord } from "../../../adapter-sdk/src";
import type {
  FocusCloseResult, FocusEvent, FocusFrame, FocusOpenResult, FocusState, OfficeBroker, PixelPolicy,
} from "../../../adapter-sdk/src";
import { FOCUS_TEXT, openFailureText } from "./focus-ui";
import type { FocusStreamHandlers } from "./focus-stream";

/** Extensions the focus-hybrid route serves; other Office formats keep the OLE route. */
export const FOCUS_EXTENSIONS = new Set([
  ".docx", ".doc", ".docm", ".dotx", ".dotm", ".rtf",
  ".xlsx", ".xls", ".xlsm", ".xlsb", ".xltx", ".xltm", ".csv",
  ".pptx", ".ppt", ".pptm", ".potx", ".potm", ".ppsx",
]);

interface Disposable { dispose(): unknown }
type Event<T> = (listener: (value: T) => unknown) => Disposable;

/** The part of vscode.WebviewPanel the controller uses (kept structural so tests need no VS Code). */
export interface FocusPanel {
  readonly visible: boolean;
  readonly webview: { postMessage(message: unknown): Thenable<boolean> | Promise<boolean>; onDidReceiveMessage: Event<unknown> };
  onDidChangeViewState: Event<unknown>;
  onDidDispose: Event<void>;
}

export interface FocusStreamLike { connect(): Promise<void>; dispose(): void }

export interface FocusEditorDeps {
  broker: Pick<OfficeBroker, "openFocusSession" | "activateFocusSession" | "layoutFocusSession" | "closeFocusSession" | "wheelFocusSession">;
  stream(pipeName: string, token: string, surfaceToken: string, handlers: FocusStreamHandlers): FocusStreamLike;
  /** Asks whether to save before the last tab of a dirty document closes. */
  askSave(): Promise<"save" | "discard" | "cancel">;
  /** Opens the document in a new tab (used when the user cancels closing). */
  reopen(): Promise<unknown>;
  trace(stage: string): void;
  error(code: string): void;
  pixelPolicy: PixelPolicy;
  editorPid: number;
  geometryTimeoutMs?: number;
  rejoinTimeoutMs?: number;
}

export function newSurfaceToken(): string {
  return "oie-surface-" + randomBytes(16).toString("hex");
}

/**
 * One focus-hybrid tab. Core owns the window; this controller only forwards the tab's size, visibility and trusted
 * clicks, and draws what the stream sends. It never sees a window handle or a screen coordinate.
 */
export class FocusEditor {
  private session?: FocusOpenResult;
  private stream?: FocusStreamLike;
  private size?: { cssWidth: number; cssHeight: number; pixelRatio: number };
  private sizeWaiter?: () => void;
  private state: FocusState = "releasing";
  private parkId = 0;
  private activationId = 0;
  private framePosting = false;
  private pendingFrame?: FocusFrame;
  private layoutRunning = false;
  private layoutDirty = false;
  /** Wheel steps that arrived while a forward was in flight; sent together as one step. */
  private wheelPending?: { point: { x: number; y: number }; deltaY: number; deltaX: number; ctrl: boolean; shift: boolean };
  private wheelRunning = false;
  private reconnects = 0;
  private closed = false;
  private closing?: Promise<void>;
  private readonly subscriptions: Disposable[] = [];

  constructor(
    private readonly panel: FocusPanel,
    private readonly path: string,
    readonly surfaceToken: string,
    private readonly deps: FocusEditorDeps,
  ) {
    this.subscriptions.push(panel.webview.onDidReceiveMessage(message => this.receive(message)));
    this.subscriptions.push(panel.onDidChangeViewState(() => {
      void this.post({ kind: "layout-request" });
      this.scheduleLayout();
    }));
    this.subscriptions.push(panel.onDidDispose(() => { void this.close(); }));
  }

  get currentState(): FocusState { return this.state; }

  async start(): Promise<void> {
    this.deps.trace("focus_tab_created");
    if (!this.size) {
      await new Promise<void>(resolve => {
        const timer = setTimeout(resolve, this.deps.geometryTimeoutMs ?? 10_000);
        this.sizeWaiter = () => { clearTimeout(timer); resolve(); };
        void this.post({ kind: "layout-request" });
      });
      this.sizeWaiter = undefined;
    }
    if (this.closed) return;
    if (!this.size) { this.fail("geometry_timeout"); return; }
    let opened: FocusOpenResult;
    try {
      opened = await this.deps.broker.openFocusSession({
        path: this.path, surfaceToken: this.surfaceToken, editorPid: this.deps.editorPid,
        pixelPolicy: this.deps.pixelPolicy, ...this.size,
      });
    } catch (error) { this.fail("open_" + errorCode(error), errorCode(error)); return; }
    this.session = opened;
    this.deps.trace(opened.shared ? "focus_session_joined" : "focus_session_opened");
    // The tab closed while Core was opening: release the surface right away.
    if (this.closed) { await this.closeSession(); return; }
    if (!(await this.connectStream())) return;
    this.scheduleLayout();
  }

  private async connectStream(): Promise<boolean> {
    const session = this.session;
    if (!session || this.closed) return false;
    // Core creates the stream pipe while answering focus.open; allow it a moment to listen.
    for (let attempt = 0; attempt < 5 && !this.closed; attempt++) {
      const stream = this.deps.stream(session.stream.pipeName, session.stream.token, this.surfaceToken, {
        event: event => this.onEvent(event),
        frame: frame => this.onFrame(frame),
        ended: error => this.onStreamEnded(error),
      });
      try {
        await stream.connect();
        if (this.closed) { stream.dispose(); return false; }
        this.stream = stream;
        this.deps.trace("focus_stream_connected");
        return true;
      } catch {
        stream.dispose();
        await new Promise(resolve => setTimeout(resolve, 100 * (attempt + 1)));
      }
    }
    if (!this.closed) this.fail("stream_unavailable");
    return false;
  }

  private onStreamEnded(error?: unknown): void {
    this.stream = undefined;
    if (this.closed || this.state === "lost") return;
    this.deps.trace("focus_stream_ended_" + (error ? errorCode(error) : "eof"));
    if (++this.reconnects > 3) { this.fail("stream_lost"); return; }
    void this.connectStream();
  }

  private onEvent(event: FocusEvent): void {
    if (event.type === "state") {
      this.state = event.state;
      if (event.parkId > this.parkId) this.parkId = event.parkId;
      void this.post({ kind: "state", state: event.state, cause: event.cause, parkId: event.parkId, activationId: event.activationId });
      if (event.state === "editing") void this.post({ kind: "clear-status" });
      if (event.state === "lost") this.showStatus(FOCUS_TEXT.lost, { sticky: true, opaque: true });
      return;
    }
    if (event.type === "notice") {
      this.deps.trace("focus_notice_" + event.code);
      switch (event.code) {
        case "modal_blocking": this.showStatus(FOCUS_TEXT.modal, { sticky: true }); break;
        case "modal_cleared": void this.post({ kind: "clear-status" }); break;
        case "capture_unavailable": this.showStatus(FOCUS_TEXT.captureUnavailable, { sticky: true }); break;
        case "activation_denied": this.showStatus(FOCUS_TEXT.activationDenied); break;
        case "office_exited": this.state = "lost"; this.showStatus(FOCUS_TEXT.lost, { sticky: true, opaque: true }); break;
        default: break;
      }
      return;
    }
    if (event.reason === "office_exited") { this.state = "lost"; this.showStatus(FOCUS_TEXT.lost, { sticky: true, opaque: true }); }
  }

  /** Only the newest frame is posted; frames of an older park are dropped. */
  private onFrame(frame: FocusFrame): void {
    if (frame.parkId < this.parkId || this.closed) return;
    if (this.framePosting) { this.pendingFrame = frame; return; }
    this.framePosting = true;
    void (async () => {
      let next: FocusFrame | undefined = frame;
      while (next && !this.closed) {
        const current: FocusFrame = next;
        this.pendingFrame = undefined;
        if (current.parkId >= this.parkId) {
          await this.post({ kind: "frame", parkId: current.parkId, frameSeq: current.frameSeq, width: current.width, height: current.height, format: current.format, image: current.image });
        }
        next = this.pendingFrame;
      }
      this.framePosting = false;
    })();
  }

  private receive(message: unknown): void {
    if (!isRecord(message) || this.closed) return;
    if (message.action === "layout" && message.surfaceToken === this.surfaceToken) {
      const width = message.width, height = message.height, ratio = message.dpr;
      if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || (width as number) < 32 || (height as number) < 32
          || (width as number) > 16384 || (height as number) > 16384) return;
      const pixelRatio = typeof ratio === "number" && Number.isFinite(ratio) && ratio >= 0.5 && ratio <= 8 ? ratio : 1;
      this.size = { cssWidth: width as number, cssHeight: height as number, pixelRatio };
      this.sizeWaiter?.();
      this.scheduleLayout();
      return;
    }
    // Only the page's own trusted pointer or key event may move Office over the editor.
    if (message.action === "activate" && message.surfaceToken === this.surfaceToken && message.trusted === true) {
      void this.activate(message.point);
      return;
    }
    // A trusted wheel over the parked preview scrolls the parked Office window without activating it.
    if (message.action === "wheel" && message.surfaceToken === this.surfaceToken && message.trusted === true) {
      this.queueWheel(message);
      return;
    }
    if (message.action === "script-error" || message.action === "frame-error") this.deps.trace("focus_page_" + message.action);
  }

  private queueWheel(message: Record<string, unknown>): void {
    const point = message.point, deltaY = message.deltaY, deltaX = message.deltaX;
    if (!isRecord(point) || typeof point.x !== "number" || typeof point.y !== "number" || !(point.x >= 0 && point.x <= 1 && point.y >= 0 && point.y <= 1)) return;
    const y = Number.isSafeInteger(deltaY) ? deltaY as number : 0, x = Number.isSafeInteger(deltaX) ? deltaX as number : 0;
    if (y === 0 && x === 0) return;
    const pending = this.wheelPending;
    this.wheelPending = {
      point: { x: point.x, y: point.y },
      deltaY: (pending?.deltaY ?? 0) + y,
      deltaX: (pending?.deltaX ?? 0) + x,
      ctrl: message.ctrl === true,
      shift: message.shift === true,
    };
    if (!this.wheelRunning) void this.flushWheel();
  }

  private async flushWheel(): Promise<void> {
    this.wheelRunning = true;
    try {
      while (this.wheelPending && !this.closed) {
        const session = this.session;
        const step = this.wheelPending;
        this.wheelPending = undefined;
        if (!session || this.state !== "parked" || step.deltaY === 0 && step.deltaX === 0) continue;
        const clamp = (value: number) => Math.max(-MAX_WHEEL_DELTA, Math.min(MAX_WHEEL_DELTA, value));
        try {
          await this.deps.broker.wheelFocusSession({ sessionId: session.sessionId, surfaceToken: this.surfaceToken, trusted: true,
            point: step.point, deltaY: clamp(step.deltaY), deltaX: clamp(step.deltaX), ctrl: step.ctrl, shift: step.shift });
        } catch (error) { this.deps.trace("focus_wheel_failed_" + errorCode(error)); }
      }
    } finally { this.wheelRunning = false; }
  }

  private async activate(point: unknown): Promise<void> {
    const session = this.session;
    if (!session || this.state === "lost") return;
    const normalized = isRecord(point) && typeof point.x === "number" && typeof point.y === "number"
      && point.x >= 0 && point.x <= 1 && point.y >= 0 && point.y <= 1 ? { x: point.x, y: point.y } : undefined;
    try {
      const result = await this.deps.broker.activateFocusSession({
        sessionId: session.sessionId, surfaceToken: this.surfaceToken, activationId: ++this.activationId, trusted: true,
        ...(normalized ? { point: normalized } : {}),
      });
      this.deps.trace("focus_activate_" + result.state);
    } catch (error) { this.deps.trace("focus_activate_failed_" + errorCode(error)); }
  }

  private scheduleLayout(): void {
    if (!this.session || !this.size || this.closed) return;
    this.layoutDirty = true;
    if (this.layoutRunning) return;
    this.layoutRunning = true;
    void (async () => {
      try {
        while (this.layoutDirty && !this.closed && this.session && this.size) {
          this.layoutDirty = false;
          try {
            await this.deps.broker.layoutFocusSession({ sessionId: this.session.sessionId, surfaceToken: this.surfaceToken, visible: this.panel.visible, ...this.size });
          } catch (error) { this.deps.trace("focus_layout_failed_" + errorCode(error)); }
        }
      } finally { this.layoutRunning = false; }
    })();
  }

  /** Runs when the tab is closed. Unsaved changes on the last tab are asked about here; cancel reopens the tab. */
  close(): Promise<void> {
    if (this.closing) return this.closing;
    this.closed = true;
    this.sizeWaiter?.();
    for (const subscription of this.subscriptions.splice(0)) subscription.dispose();
    this.closing = this.closeSession().finally(() => { this.stream?.dispose(); this.stream = undefined; });
    return this.closing;
  }

  private async closeSession(): Promise<void> {
    const session = this.session;
    if (!session) return;
    const params = { sessionId: session.sessionId, surfaceToken: this.surfaceToken };
    let result: FocusCloseResult;
    try { result = await this.deps.broker.closeFocusSession(params); }
    catch (error) { this.deps.error(errorCode(error)); return; }
    if (result.closed || !result.dirty) { this.deps.trace("focus_closed"); return; }
    const choice = await this.deps.askSave();
    if (choice !== "cancel") {
      try { await this.deps.broker.closeFocusSession({ ...params, decision: choice }); this.deps.trace("focus_closed_" + choice); }
      catch (error) { this.deps.error(errorCode(error)); }
      return;
    }
    // Cancel: a new tab joins the session, then this surface detaches (it is no longer the last one).
    this.deps.trace("focus_close_cancelled");
    try { await this.deps.reopen(); } catch (error) { this.deps.error(errorCode(error)); return; }
    const deadline = Date.now() + (this.deps.rejoinTimeoutMs ?? 15_000);
    while (Date.now() < deadline) {
      await new Promise(resolve => setTimeout(resolve, 250));
      try {
        const retry = await this.deps.broker.closeFocusSession(params);
        if (retry.closed) { this.deps.trace("focus_surface_detached"); return; }
      } catch (error) { this.deps.trace("focus_detach_failed_" + errorCode(error)); return; }
    }
  }

  private showStatus(text: string, options: { sticky?: boolean; opaque?: boolean } = {}): void {
    void this.post({ kind: "status", text, sticky: options.sticky === true, opaque: options.opaque === true });
  }

  private fail(stage: string, code = ""): void {
    this.deps.trace("focus_" + stage);
    this.showStatus(openFailureText(code.toLowerCase()), { sticky: true, opaque: true });
  }

  private async post(message: unknown): Promise<void> {
    try { await this.panel.webview.postMessage(message); } catch { /* The tab is gone. */ }
  }
}
