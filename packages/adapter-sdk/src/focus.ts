import { OfficeBridgeError } from "./errors";
import { isRecord } from "./protocol";
import type { OfficeApplication } from "./types";

/**
 * Focus-hybrid contract (mirrors packages/shared-core/FocusContract.cs).
 * Control calls go over the v1 broker pipe; state events and frames arrive on a per-session stream pipe.
 */
export type FocusState = "parked" | "activating" | "editing" | "releasing" | "lost";
export type FocusCause = "open" | "click" | "foreground" | "layout" | "timer" | "policy" | "dialog" | "close" | "recovery";
export type PixelPolicy = "always" | "switch" | "mouse";
export type FocusNoticeCode =
  | "modal_blocking" | "modal_cleared" | "alignment_unconfirmed" | "activation_denied" | "office_exited" | "capture_unavailable";
export type FocusCloseDecision = "save" | "discard";

export const FOCUS_STATES: readonly FocusState[] = ["parked", "activating", "editing", "releasing", "lost"];
export const PIXEL_POLICIES: readonly PixelPolicy[] = ["always", "switch", "mouse"];
export type FrameFormat = "png" | "jpeg";
export const FRAME_FORMATS: readonly FrameFormat[] = ["png", "jpeg"];
export const SURFACE_TOKEN_PATTERN = /^oie-surface-[a-f0-9]{32}$/u;
export const FOCUS_SESSION_PATTERN = /^[a-f0-9]{32}$/u;
export const FOCUS_STREAM_PIPE_PATTERN = /^office-in-editor-v1-[a-f0-9]{16}-focus-[a-f0-9]{32}$/u;
export const MIN_SURFACE_PX = 32;
export const MAX_SURFACE_PX = 16384;
export const MAX_STREAM_MESSAGE_BYTES = 32 * 1024 * 1024;
export const MAX_FRAME_HEADER_BYTES = 4096;
const EVENT_KIND = 1;
const FRAME_KIND = 2;

/** CSS size of the preview surface; Core reads physical bounds itself through UI Automation. */
export interface SurfaceSize {
  cssWidth: number;
  cssHeight: number;
  pixelRatio: number;
}

export interface FocusOpenParams extends SurfaceSize {
  path: string;
  surfaceToken: string;
  /** Main process id of this IDE window, used only to pick the matching top-level editor window. */
  editorPid: number;
  pixelPolicy?: PixelPolicy;
}

export interface FocusOpenResult {
  sessionId: string;
  application: OfficeApplication;
  /** True when another tab already opened the same file and this surface joined that session. */
  shared: boolean;
  state: FocusState;
  stream: { pipeName: string; token: string };
}

export interface FocusActivateParams {
  sessionId: string;
  surfaceToken: string;
  activationId: number;
  /** Must come from a trusted pointer event; Core refuses anything else. */
  trusted: true;
  /** Click position normalized to the surface, used to place the caret (P5). */
  point?: { x: number; y: number };
}

/**
 * One wheel step over a parked tab. Win32 conventions: deltaY > 0 scrolls up (away from the user), deltaX > 0 scrolls
 * right, 120 is one notch. Core forwards it to the parked Office window without activating it.
 */
export interface FocusWheelParams {
  sessionId: string;
  surfaceToken: string;
  /** Must come from a trusted wheel event in the tab. */
  trusted: true;
  point: { x: number; y: number };
  deltaY: number;
  deltaX: number;
  ctrl?: boolean;
  shift?: boolean;
}

export const MAX_WHEEL_DELTA = 2400;

export interface FocusLayoutParams extends SurfaceSize {
  sessionId: string;
  surfaceToken: string;
  visible: boolean;
}

export interface FocusCloseParams {
  sessionId: string;
  surfaceToken: string;
  decision?: FocusCloseDecision;
}

/**
 * closed=false with dirty=true means the last tab is closing over unsaved changes: the host asks the user,
 * then calls focus.close again with a decision, or keeps the tab open on cancel.
 */
export interface FocusCloseResult {
  closed: boolean;
  dirty: boolean;
  remainingSurfaces: number;
}

export interface FocusStatus {
  state: FocusState;
  application: OfficeApplication;
  surfaces: number;
  /** null when the Office add-in could not be reached. */
  dirty: boolean | null;
  activationId: number;
  parkId: number;
  pixelPolicy: PixelPolicy;
}

export type FocusEvent =
  | { type: "state"; state: FocusState; cause: FocusCause; activationId: number; parkId: number }
  | { type: "notice"; code: FocusNoticeCode }
  | { type: "closed"; reason: "closed" | "office_exited" | "superseded" };

export interface FocusFrame {
  frameSeq: number;
  parkId: number;
  width: number;
  height: number;
  /** jpeg for the 30 fps stream of a visible tab, png for lossless frames. */
  format: FrameFormat;
  image: Uint8Array;
}

export type FocusStreamMessage = { kind: "event"; event: FocusEvent } | { kind: "frame"; frame: FocusFrame };

const isInt = (value: unknown, min = 0): value is number => Number.isSafeInteger(value) && (value as number) >= min;

export function validateSurfaceSize(size: SurfaceSize): void {
  const valid = [size.cssWidth, size.cssHeight].every(value => isInt(value, MIN_SURFACE_PX) && value <= MAX_SURFACE_PX)
    && Number.isFinite(size.pixelRatio) && size.pixelRatio >= 0.5 && size.pixelRatio <= 8;
  if (!valid) throw new OfficeBridgeError("INVALID_PARAMS");
}

export function parseFocusOpenResult(value: unknown): FocusOpenResult {
  if (!isRecord(value) || typeof value.sessionId !== "string" || !FOCUS_SESSION_PATTERN.test(value.sessionId)
      || !["word", "excel", "powerpoint"].includes(value.application as string) || typeof value.shared !== "boolean"
      || !FOCUS_STATES.includes(value.state as FocusState) || !isRecord(value.stream)
      || typeof value.stream.pipeName !== "string" || !FOCUS_STREAM_PIPE_PATTERN.test(value.stream.pipeName)
      || typeof value.stream.token !== "string" || !/^[a-f0-9]{64}$/u.test(value.stream.token)) {
    throw new OfficeBridgeError("INVALID_RESPONSE");
  }
  return value as unknown as FocusOpenResult;
}

export function parseFocusCloseResult(value: unknown): FocusCloseResult {
  if (!isRecord(value) || typeof value.closed !== "boolean" || typeof value.dirty !== "boolean" || !isInt(value.remainingSurfaces)) {
    throw new OfficeBridgeError("INVALID_RESPONSE");
  }
  return value as unknown as FocusCloseResult;
}

/** First line a surface writes after connecting to the stream pipe. */
export function encodeStreamHello(token: string, surfaceToken: string): string {
  if (!/^[a-f0-9]{64}$/u.test(token) || !SURFACE_TOKEN_PATTERN.test(surfaceToken)) throw new OfficeBridgeError("INVALID_PARAMS");
  return JSON.stringify({ v: 1, token, surfaceToken }) + "\n";
}

function parseEvent(value: unknown): FocusEvent {
  if (isRecord(value)) {
    if (value.type === "state" && FOCUS_STATES.includes(value.state as FocusState) && typeof value.cause === "string"
        && isInt(value.activationId) && isInt(value.parkId)) return value as unknown as FocusEvent;
    if (value.type === "notice" && typeof value.code === "string") return value as unknown as FocusEvent;
    if (value.type === "closed" && ["closed", "office_exited", "superseded"].includes(value.reason as string)) return value as unknown as FocusEvent;
  }
  throw new OfficeBridgeError("INVALID_STREAM");
}

/**
 * Incremental decoder for the stream pipe: u32 LE length of (kind + body), u8 kind, body.
 * Event body is UTF-8 JSON; frame body is u16 LE header length, header JSON, image bytes.
 */
export class FocusStreamDecoder {
  private pending: Buffer = Buffer.alloc(0);
  // Chunks of a large frame are held apart until the frame is complete, so a multi-MB frame is copied once.
  private parts: Buffer[] = [];
  private partsLength = 0;
  private needed = 0;

  push(chunk: Uint8Array): FocusStreamMessage[] {
    if (this.needed && this.pending.length + this.partsLength + chunk.length < this.needed) {
      this.parts.push(Buffer.from(chunk));
      this.partsLength += chunk.length;
      return [];
    }
    this.pending = Buffer.concat([this.pending, ...this.parts, chunk]);
    this.parts = []; this.partsLength = 0; this.needed = 0;
    const messages: FocusStreamMessage[] = [];
    while (this.pending.length >= 4) {
      const length = this.pending.readUInt32LE(0);
      if (length < 1 || length > MAX_STREAM_MESSAGE_BYTES) throw new OfficeBridgeError("INVALID_STREAM");
      if (this.pending.length < 4 + length) { this.needed = 4 + length; break; }
      const kind = this.pending[4];
      const body = this.pending.subarray(5, 4 + length);
      if (kind === EVENT_KIND) {
        messages.push({ kind: "event", event: parseEvent(JSON.parse(body.toString("utf8"))) });
      } else if (kind === FRAME_KIND) {
        if (body.length < 2) throw new OfficeBridgeError("INVALID_STREAM");
        const headerLength = body.readUInt16LE(0);
        if (headerLength > MAX_FRAME_HEADER_BYTES || body.length < 2 + headerLength) throw new OfficeBridgeError("INVALID_STREAM");
        const header: unknown = JSON.parse(body.subarray(2, 2 + headerLength).toString("utf8"));
        if (!isRecord(header) || !isInt(header.frameSeq, 1) || !isInt(header.parkId) || !isInt(header.width, 1)
            || !isInt(header.height, 1) || !FRAME_FORMATS.includes(header.format as FrameFormat)) throw new OfficeBridgeError("INVALID_STREAM");
        // Copy so the frame stays valid after the pending buffer is trimmed.
        messages.push({ kind: "frame", frame: { frameSeq: header.frameSeq, parkId: header.parkId, width: header.width,
          height: header.height, format: header.format as FrameFormat, image: new Uint8Array(body.subarray(2 + headerLength)) } });
      } else {
        throw new OfficeBridgeError("INVALID_STREAM");
      }
      this.pending = this.pending.subarray(4 + length);
    }
    return messages;
  }
}

/** Encoders used by tests and by any in-process fake Core. */
export function encodeFocusEvent(event: FocusEvent): Buffer {
  const body = Buffer.from(JSON.stringify(event), "utf8");
  const out = Buffer.alloc(5 + body.length);
  out.writeUInt32LE(1 + body.length, 0);
  out[4] = EVENT_KIND;
  body.copy(out, 5);
  return out;
}

export function encodeFocusFrame(frame: Omit<FocusFrame, "image">, image: Uint8Array): Buffer {
  const header = Buffer.from(JSON.stringify({ frameSeq: frame.frameSeq, parkId: frame.parkId, width: frame.width,
    height: frame.height, format: frame.format }), "utf8");
  const out = Buffer.alloc(7 + header.length + image.length);
  out.writeUInt32LE(3 + header.length + image.length, 0);
  out[4] = FRAME_KIND;
  out.writeUInt16LE(header.length, 5);
  header.copy(out, 7);
  Buffer.from(image).copy(out, 7 + header.length);
  return out;
}
