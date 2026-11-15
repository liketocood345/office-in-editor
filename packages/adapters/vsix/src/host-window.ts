import { OfficeBridgeError } from "../../../adapter-sdk/src";
import type { HostBounds } from "../../../adapter-sdk/src";

const SURFACE_TOKEN = /^oie-surface-[a-f0-9]{32}$/iu;

function integer(value: unknown): number | undefined {
  return typeof value === "number" && Number.isInteger(value) ? value : undefined;
}

function surfaceToken(value: unknown, fallback?: string): string | undefined {
  const token = typeof value === "string" ? value : fallback;
  return token && SURFACE_TOKEN.test(token) ? token : undefined;
}

export function tryHostBounds(message: unknown, fallback?: HostBounds): HostBounds | undefined {
  const record = message && typeof message === "object" ? message as Record<string, unknown> : {};
  const token = surfaceToken(record.surfaceToken, fallback?.surfaceToken);
  const width = integer(record.width) ?? fallback?.width;
  const height = integer(record.height) ?? fallback?.height;
  if (!token || width == null || height == null || width < 32 || height < 32) {
    return undefined;
  }
  return {
    x: integer(record.x) ?? fallback?.x ?? 0,
    y: integer(record.y) ?? fallback?.y ?? 0,
    width,
    height,
    surfaceToken: token,
  };
}

export function hostBounds(message: unknown, fallback?: HostBounds): HostBounds {
  return tryHostBounds(message, fallback) ?? (() => { throw new OfficeBridgeError("HOST_WINDOW_UNAVAILABLE"); })();
}
