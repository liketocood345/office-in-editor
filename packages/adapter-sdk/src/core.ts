import { execFile } from "node:child_process";
import { randomUUID } from "node:crypto";
import { stat } from "node:fs/promises";
import { homedir } from "node:os";
import * as net from "node:net";
import * as path from "node:path";
import { TextDecoder } from "node:util";
import { OfficeBridgeError } from "./errors";
import { CORE_START_TIMEOUT_MS, DEFAULT_TIMEOUT_MS, MAX_LINE_BYTES, MAX_PIPE_TIMEOUT_MS, encodeRequest, parseEndpoint, parseResponse } from "./protocol";
import type { BrokerMethod, BrokerTransport, CoreEndpoint } from "./types";

export interface CoreLocationOptions {
  extensionPath: string;
  platform?: NodeJS.Platform;
  localAppData?: string;
}

export async function resolveCoreExecutable(options: CoreLocationOptions): Promise<string> {
  if ((options.platform ?? process.platform) !== "win32") throw new OfficeBridgeError("UNSUPPORTED_PLATFORM");
  const appData = options.localAppData ?? process.env.LOCALAPPDATA ?? path.join(homedir(), "AppData", "Local");
  const candidates = [
    path.join(appData, "Programs", "OfficeInEditor", "current", "core", "OfficeInEditor.Core.exe"),
    path.join(options.extensionPath, "payload", "core", "OfficeInEditor.Core.exe"),
  ];
  for (const candidate of candidates) {
    try { if ((await stat(candidate)).isFile()) return candidate; } catch { /* Try the next owned product location. */ }
  }
  throw new OfficeBridgeError("CORE_NOT_FOUND");
}

export async function readCoreEndpoint(executable: string, ensureBroker = false): Promise<CoreEndpoint> {
  const stdout = await new Promise<string>((resolve, reject) => {
    execFile(executable, [ensureBroker ? "--ensure-broker" : "--endpoint"], {
      windowsHide: true,
      timeout: ensureBroker ? CORE_START_TIMEOUT_MS : DEFAULT_TIMEOUT_MS,
      maxBuffer: 16 * 1024,
      encoding: "utf8",
      cwd: path.dirname(executable),
    }, (error, output) => {
      if (error) {
        // Ignore stderr and native error strings; both can contain local paths.
        reject(new OfficeBridgeError(error.killed ? "REQUEST_TIMEOUT" : "CORE_START_FAILED"));
      } else resolve(output);
    });
  });
  return parseEndpoint(stdout);
}

/** One request per connection, with no write retry after the request may have reached Core. */
export class NamedPipeBrokerTransport implements BrokerTransport {
  private endpointValue?: CoreEndpoint;
  private executableValue?: string;

  constructor(
    private readonly location: CoreLocationOptions,
    private readonly onBrokerReady?: () => Promise<void>,
  ) {}

  async executable(): Promise<string> {
    this.executableValue ??= await resolveCoreExecutable(this.location);
    return this.executableValue;
  }

  async endpoint(ensureBroker = false): Promise<CoreEndpoint> {
    if (!this.endpointValue || ensureBroker) {
      this.endpointValue = await readCoreEndpoint(await this.executable(), ensureBroker);
    }
    return this.endpointValue;
  }

  async request<T>(
    method: BrokerMethod,
    params: Record<string, unknown> = {},
    options: { ensureBroker?: boolean; timeoutMs?: number } = {},
  ): Promise<T> {
    const endpoint = await this.endpoint();
    try {
      return await requestPipe<T>(endpoint.pipeName, method, params, options.timeoutMs);
    } catch (error) {
      if (!(error instanceof OfficeBridgeError)
        || error.code !== "BROKER_UNAVAILABLE"
        || options.ensureBroker === false) throw error;
      if (method === "host.register" && this.recovering) throw error;
      const ready = await this.recover(method);
      return requestPipe<T>(ready.pipeName, method, params, options.timeoutMs);
    }
  }

  private recovering?: Promise<CoreEndpoint>;

  private async recover(method: BrokerMethod): Promise<CoreEndpoint> {
    if (this.recovering) return this.recovering;
    const pending = (async () => {
      const endpoint = await this.endpoint(true);
      if (method !== "host.register") await this.onBrokerReady?.();
      return endpoint;
    })();
    this.recovering = pending;
    try { return await pending; }
    finally { if (this.recovering === pending) this.recovering = undefined; }
  }
}

export async function requestPipe<T>(
  pipeName: string,
  method: BrokerMethod,
  params: Record<string, unknown>,
  timeoutMs = DEFAULT_TIMEOUT_MS,
): Promise<T> {
  const id = randomUUID();
  const request = encodeRequest(id, method, params);
  return new Promise<T>((resolve, reject) => {
    let connected = false;
    let settled = false;
    let bytes = Buffer.alloc(0);
    const socket = net.createConnection({ path: "\\\\.\\pipe\\" + pipeName });
    const finish = (error?: OfficeBridgeError, result?: T): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      socket.destroy();
      if (error) reject(error); else resolve(result as T);
    };
    const timer = setTimeout(
      () => finish(new OfficeBridgeError("REQUEST_TIMEOUT", connected)),
      Math.max(250, Math.min(timeoutMs, MAX_PIPE_TIMEOUT_MS)),
    );
    socket.on("connect", () => {
      connected = true;
      socket.write(request, "utf8", (error) => {
        if (error) finish(new OfficeBridgeError("REQUEST_WRITE_FAILED", true));
      });
    });
    socket.on("data", (chunk: Buffer) => {
      bytes = Buffer.concat([bytes, chunk]);
      const newline = bytes.indexOf(0x0a);
      if ((newline < 0 ? bytes.length : newline) > MAX_LINE_BYTES) {
        finish(new OfficeBridgeError("RESPONSE_TOO_LARGE", true));
        return;
      }
      if (newline < 0) return;
      let line: string;
      try { line = new TextDecoder("utf-8", { fatal: true }).decode(bytes.subarray(0, newline)); }
      catch { finish(new OfficeBridgeError("INVALID_RESPONSE", true)); return; }
      try { finish(undefined, parseResponse<T>(line, id)); }
      catch (error) {
        finish(error instanceof OfficeBridgeError ? error : new OfficeBridgeError("INVALID_RESPONSE", true));
      }
    });
    socket.on("error", () => finish(new OfficeBridgeError(
      connected ? "CONNECTION_LOST" : "BROKER_UNAVAILABLE", connected,
    )));
    socket.on("end", () => finish(new OfficeBridgeError("CONNECTION_CLOSED", connected)));
  });
}
