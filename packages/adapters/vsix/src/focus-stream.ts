import * as net from "node:net";
import {
  FOCUS_STREAM_PIPE_PATTERN, FocusStreamDecoder, OfficeBridgeError, encodeStreamHello,
  type FocusEvent, type FocusFrame,
} from "../../../adapter-sdk/src";

export interface FocusStreamHandlers {
  event(event: FocusEvent): void;
  frame(frame: FocusFrame): void;
  /** The pipe ended. `error` is set when it ended because of a protocol or I/O failure. */
  ended(error?: unknown): void;
}

/** Connects one tab surface to its session's stream pipe and decodes events and frames. */
export class FocusStreamClient {
  private socket?: net.Socket;
  private disposed = false;

  constructor(
    private readonly pipeName: string,
    private readonly token: string,
    private readonly surfaceToken: string,
    private readonly handlers: FocusStreamHandlers,
    private readonly connectTimeoutMs = 5_000,
  ) {
    if (!FOCUS_STREAM_PIPE_PATTERN.test(pipeName)) throw new OfficeBridgeError("INVALID_PARAMS");
  }

  connect(): Promise<void> {
    if (this.disposed) return Promise.reject(new OfficeBridgeError("STREAM_CLOSED"));
    const hello = encodeStreamHello(this.token, this.surfaceToken);
    const decoder = new FocusStreamDecoder();
    return new Promise((resolve, reject) => {
      const socket = net.connect("\\\\.\\pipe\\" + this.pipeName);
      this.socket = socket;
      let connected = false;
      let ended = false;
      const end = (error?: unknown): void => {
        if (ended) return;
        ended = true;
        socket.destroy();
        if (!connected) reject(error ?? new OfficeBridgeError("STREAM_UNAVAILABLE"));
        else if (!this.disposed) this.handlers.ended(error);
      };
      const timer = setTimeout(() => end(new OfficeBridgeError("STREAM_TIMEOUT")), this.connectTimeoutMs);
      socket.once("connect", () => {
        clearTimeout(timer);
        connected = true;
        socket.write(hello);
        resolve();
      });
      socket.on("data", chunk => {
        try {
          for (const message of decoder.push(chunk)) {
            if (message.kind === "event") this.handlers.event(message.event);
            else this.handlers.frame(message.frame);
          }
        } catch (error) { end(error); }
      });
      socket.on("error", error => { clearTimeout(timer); end(error); });
      socket.on("close", () => { clearTimeout(timer); end(); });
    });
  }

  dispose(): void {
    this.disposed = true;
    this.socket?.destroy();
    this.socket = undefined;
  }
}
