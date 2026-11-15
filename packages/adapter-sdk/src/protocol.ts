import { OfficeBridgeError } from "./errors";
import type { BrokerMethod, CoreEndpoint } from "./types";

export const MAX_LINE_BYTES = 2 * 1024 * 1024;
export const DEFAULT_TIMEOUT_MS = 10_000;
export const CORE_START_TIMEOUT_MS = 25_000;
export const OPEN_DOCUMENT_TIMEOUT_MS = 60_000;
export const MAX_PIPE_TIMEOUT_MS = 60_000;
export const PIPE_NAME_PATTERN = /^office-in-editor-v1-[a-f0-9]{16}$/u;

export function parseEndpoint(stdout: string): CoreEndpoint {
  let value: unknown;
  try { value = JSON.parse(stdout.trim()); } catch { throw new OfficeBridgeError("INVALID_ENDPOINT"); }
  if (!isRecord(value)
    || value.protocolVersion !== 1
    || typeof value.pipeName !== "string"
    || !PIPE_NAME_PATTERN.test(value.pipeName)
    || typeof value.version !== "string"
    || !/^[a-z0-9.+-]{1,64}$/iu.test(value.version)) {
    throw new OfficeBridgeError("INVALID_ENDPOINT");
  }
  return { pipeName: value.pipeName, protocolVersion: 1, version: value.version };
}

export function encodeRequest(id: string, method: BrokerMethod, params: Record<string, unknown>): string {
  const line = JSON.stringify({ v: 1, id, method, params }) + "\n";
  if (Buffer.byteLength(line, "utf8") > MAX_LINE_BYTES) throw new OfficeBridgeError("REQUEST_TOO_LARGE");
  return line;
}

export function parseResponse<T>(line: string, expectedId: string): T {
  let value: unknown;
  try { value = JSON.parse(line); } catch { throw new OfficeBridgeError("INVALID_RESPONSE", true); }
  if (!isRecord(value) || value.id !== expectedId || typeof value.ok !== "boolean") {
    throw new OfficeBridgeError("INVALID_RESPONSE", true);
  }
  if (value.ok) return value.result as T;
  const code = isRecord(value.error) && typeof value.error.code === "string"
    ? value.error.code
    : "BROKER_ERROR";
  // Do not retain the worker's message: it may contain a file path or document text.
  throw new OfficeBridgeError(code);
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}
