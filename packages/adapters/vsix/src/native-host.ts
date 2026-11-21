import { createConnection, type Socket } from "node:net";
import { isRecord, OfficeBridgeError } from "../../../adapter-sdk/src";

interface SlotIdentity {
  hwnd: string;
  generation: string;
  pid: number;
}

interface PendingRequest {
  complete: (value: unknown) => void;
  fail: (error: Error) => void;
}

export interface NativeSurfaceInspection {
  wordChild: boolean;
  visible: boolean;
  slotIsVerifiedChild?: boolean;
  officeApplications: ("word" | "excel" | "powerpoint")[];
  viewport?: { x: number; y: number; width: number; height: number };
  geometryAttempts?: number;
  hitTest?: { samples: number; inSlot: number; officeHits: number; foreground: boolean };
}

/** A local connection to the UI-thread provider. It never opens a document. */
export class NativeHostClient {
  private readonly pipe = process.env.OIE_NATIVE_PIPE ?? "";
  private readonly secret = process.env.OIE_NATIVE_TOKEN ?? "";
  readonly available = /^\\\\\.\\pipe\\office-in-editor-slot-[0-9a-f]{32}$/.test(this.pipe)
    && /^[0-9a-f]{64}$/.test(this.secret);
  private readonly pending = new Map<number, PendingRequest>();
  private readonly closers = new Map<string, () => Promise<void>>();
  private socket?: Socket;
  private handshake?: Promise<Socket>;
  private sequence = 0;

  async create(token: string, visible: boolean, close: () => Promise<void>): Promise<SlotIdentity> {
    this.closers.set(token, close);
    try {
      const value = await this.call("create", token, visible);
      if (!isRecord(value) || typeof value.hwnd !== "string" || !/^\d+$/.test(value.hwnd)
        || typeof value.generation !== "string" || !/^\d+$/.test(value.generation)
        || typeof value.pid !== "number" || !Number.isSafeInteger(value.pid) || value.pid <= 0) {
        throw new OfficeBridgeError("INVALID_NATIVE_SURFACE");
      }
      return { hwnd: value.hwnd, generation: value.generation, pid: value.pid };
    } catch (error) {
      this.closers.delete(token);
      throw error;
    }
  }

  async layout(token: string, visible: boolean): Promise<boolean> {
    const value = await this.call("layout", token, visible);
    return isRecord(value) && value.visible === true;
  }

  async release(token: string): Promise<void> {
    await this.call("release", token);
    this.closers.delete(token);
  }

  async inspect(token: string): Promise<NativeSurfaceInspection> {
    const value = await this.call("inspect", token);
    if (!isRecord(value) || typeof value.wordChild !== "boolean" || typeof value.visible !== "boolean")
      throw new OfficeBridgeError("INVALID_NATIVE_SURFACE");
    const flags = typeof value.officeApplications === "number" ? value.officeApplications : 0;
    const officeApplications: NativeSurfaceInspection["officeApplications"] = [];
    if (flags & 1) officeApplications.push("word");
    if (flags & 2) officeApplications.push("excel");
    if (flags & 4) officeApplications.push("powerpoint");
    const result: NativeSurfaceInspection = { wordChild: value.wordChild, visible: value.visible, officeApplications };
    if (typeof value.slotIsVerifiedChild === "boolean") result.slotIsVerifiedChild = value.slotIsVerifiedChild;
    if (Array.isArray(value.viewport) && value.viewport.length === 4 &&
      value.viewport.every(n => typeof n === "number" && Number.isInteger(n) && n >= 0 && n <= 32768)) {
      const [x, y, width, height] = value.viewport as number[];
      result.viewport = { x, y, width, height };
    }
    if (typeof value.geometryAttempts === "number" && Number.isInteger(value.geometryAttempts) &&
      value.geometryAttempts >= 1 && value.geometryAttempts <= 6) result.geometryAttempts = value.geometryAttempts;
    const count = (n: unknown): n is number => typeof n === "number" && Number.isInteger(n) && n >= 0 && n <= 5;
    if (count(value.samples) && count(value.inSlot) && count(value.officeHits) &&
      value.officeHits <= value.inSlot && value.inSlot <= value.samples && typeof value.foreground === "boolean") {
      result.hitTest = { samples: value.samples, inSlot: value.inSlot, officeHits: value.officeHits, foreground: value.foreground };
    }
    return result;
  }

  private async call(op: string, token: string, visible?: boolean): Promise<unknown> {
    this.handshake ??= this.connect();
    return this.request(await this.handshake, { op, token, visible });
  }

  private async connect(): Promise<Socket> {
    if (!this.available) throw new OfficeBridgeError("NATIVE_HOST_UNAVAILABLE");
    const socket = this.socket = createConnection(this.pipe);
    socket.setEncoding("utf8");
    let buffer = "";
    const disconnected = () => {
      for (const request of this.pending.values()) request.fail(new OfficeBridgeError("NATIVE_HOST_DISCONNECTED"));
      this.pending.clear();
    };
    socket.on("error", () => socket.destroy());
    socket.on("close", disconnected);
    socket.on("data", (chunk: string) => {
      buffer += chunk;
      if (buffer.length > 32_768) { socket.destroy(); return; }
      let newline: number;
      while ((newline = buffer.indexOf("\n")) !== -1) {
        const line = buffer.slice(0, newline);
        buffer = buffer.slice(newline + 1);
        let message: unknown;
        try { message = JSON.parse(line); } catch { socket.destroy(); return; }
        if (!isRecord(message)) continue;
        if (Array.isArray(message.close) && message.close.length <= 32) {
          for (const token of message.close) {
            if (typeof token === "string") void this.closers.get(token)?.().catch(() => {});
          }
        } else if (typeof message.id === "number") {
          const request = this.pending.get(message.id);
          if (typeof message.error === "string") request?.fail(new OfficeBridgeError(message.error));
          else request?.complete(message.value);
        }
      }
    });
    const value = await this.request(socket, { secret: this.secret });
    if (!isRecord(value) || value.protocol !== 1) {
      socket.destroy();
      throw new OfficeBridgeError("INVALID_NATIVE_HOST_PROTOCOL");
    }
    return socket;
  }

  private request(socket: Socket, body: Record<string, unknown>): Promise<unknown> {
    if (socket.destroyed || this.pending.size >= 16 || socket.writableLength >= 16_384)
      return Promise.reject(new OfficeBridgeError("NATIVE_HOST_UNAVAILABLE"));
    const id = ++this.sequence;
    return new Promise((resolve, reject) => {
      const finish = () => { clearTimeout(deadline); this.pending.delete(id); };
      const deadline = setTimeout(() => {
        finish();
        socket.destroy();
        reject(new OfficeBridgeError("NATIVE_HOST_TIMEOUT"));
      }, 3000);
      this.pending.set(id, {
        complete: value => { finish(); resolve(value); },
        fail: error => { finish(); reject(error); },
      });
      socket.write(JSON.stringify({ id, ...body }) + "\n");
    });
  }

  dispose(): void {
    this.socket?.destroy();
    this.closers.clear();
  }
}
