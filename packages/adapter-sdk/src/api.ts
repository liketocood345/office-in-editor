import { randomUUID } from "node:crypto";
import { OfficeBridgeError } from "./errors";
import { isRecord } from "./protocol";
import { OPEN_DOCUMENT_TIMEOUT_MS } from "./protocol";

/** focus.open upper bound: Core resolves or fails within 25 s, so a focus tab never waits more than 30 s. */
export const FOCUS_OPEN_TIMEOUT_MS = 30_000;
import {
  parseFocusCloseResult, parseFocusOpenResult, validateSurfaceSize,
  type FocusActivateParams, type FocusCloseParams, type FocusCloseResult, type FocusLayoutParams,
  type FocusOpenParams, type FocusOpenResult, type FocusState, type FocusStatus, type FocusWheelParams, type PixelPolicy,
} from "./focus";
import type {
  BrokerTransport, CoreHealth, DashboardSnapshot, DeliveryReceipt, HostLease,
  HostBounds, HostRegistration, OfficeInstance, OfficeReference, ReferenceMetadata, RegisteredChat,
} from "./types";

function arrayResult<T>(value: unknown): T[] {
  if (!Array.isArray(value)) throw new OfficeBridgeError("INVALID_RESPONSE");
  return value as T[];
}

/** Typed local-only feature boundary. Reference text is retrieved only on explicit user action. */
export class OfficeBroker {
  private host?: HostLease;
  constructor(private readonly transport: BrokerTransport) {}

  bindHost(lease?: HostLease): void {
    this.host = lease;
  }

  private withHost(params: Record<string, unknown>): Record<string, unknown> {
    const host = this.host;
    if (!host) return params;
    return { ...params, hostId: host.hostId, hostLease: host.lease, hostEpoch: host.epoch };
  }

  async health(): Promise<CoreHealth> {
    const value = await this.transport.request<unknown>("health");
    if (!isRecord(value) || value.protocolVersion !== 1 || !isRecord(value.capabilities)) {
      throw new OfficeBridgeError("INVALID_RESPONSE");
    }
    return value as unknown as CoreHealth;
  }

  registerHost(registration: HostRegistration): Promise<HostLease> {
    return this.transport.request("host.register", { ...registration });
  }

  registerNativeSurface(surfaceToken: string, slot: { hwnd: string; generation: string; pid: number }): Promise<unknown> {
    return this.transport.request("host.native_surface.register", this.withHost({ surfaceToken, ...slot }));
  }

  releaseNativeSurface(surfaceToken: string): Promise<unknown> {
    return this.transport.request("host.native_surface.release", this.withHost({ surfaceToken }), { ensureBroker: false });
  }

  unregisterHost(lease: HostLease): Promise<unknown> {
    return this.transport.request("host.unregister", {
      hostId: lease.hostId, lease: lease.lease,
    }, { ensureBroker: false, timeoutMs: 1_500 });
  }

  async listOfficeInstances(): Promise<OfficeInstance[]> {
    return arrayResult(await this.transport.request("office.list"));
  }

  captureSelection(officeId: string): Promise<ReferenceMetadata> {
    return this.transport.request("office.capture", { officeId });
  }

  async listReferences(): Promise<ReferenceMetadata[]> {
    return arrayResult(await this.transport.request("reference.list"));
  }

  async getReference(refId: string): Promise<OfficeReference> {
    const value = await this.transport.request<unknown>("reference.get", { refId });
    if (!isRecord(value) || typeof value.refId !== "string" || typeof value.text !== "string") {
      throw new OfficeBridgeError("INVALID_RESPONSE");
    }
    return value as unknown as OfficeReference;
  }

  releaseReference(refId: string): Promise<unknown> {
    return this.transport.request("reference.release", { refId, confirm: true });
  }

  async listRegisteredChats(): Promise<RegisteredChat[]> {
    return arrayResult(await this.transport.request("chat.list"));
  }

  deliverReference(refId: string, targetId: string, requestId = randomUUID()): Promise<DeliveryReceipt> {
    return this.transport.request("reference.deliver", { refId, targetId, requestId });
  }

  openInOffice(
    documentPath: string,
    mode: "external" | "embedded" = "external",
    options: { fallback?: boolean; x?: number; y?: number; width?: number; height?: number; sessionId?: string; surfaceToken?: string; timeoutMs?: number; visible?: boolean; pixelRatio?: number; ackAlerts?: boolean } = {},
  ): Promise<{ opened: true; mode: "external" | "embedded"; application: string; pid?: number; inplace?: boolean; officeId?: string; hwnd?: string; shared?: boolean; viewMode?: string }> {
    const { timeoutMs = OPEN_DOCUMENT_TIMEOUT_MS, ...payload } = options;
    return this.transport.request("office.open", this.withHost({ path: documentPath, mode, ...payload }), { timeoutMs: Math.max(250, Math.min(timeoutMs, OPEN_DOCUMENT_TIMEOUT_MS)) });
  }

  saveOfficeWindow(officeId: string): Promise<{ dirty?: boolean; saved?: boolean }> {
    return this.transport.request("office.save_window", this.withHost({ officeId }));
  }

  waitHostDelivery(lease: HostLease, timeoutMs = 25_000): Promise<{ waiting?: boolean; text?: string; documentName?: string; kind?: string; requestId?: string; refId?: string }> {
    return this.transport.request("host.wait_delivery", { hostId: lease.hostId, lease: lease.lease }, { timeoutMs });
  }

  describeOfficeWindow(officeId: string): Promise<{ hwnd: string; processId: number; attached: boolean }> {
    return this.transport.request("office.describe_window", this.withHost({ officeId }));
  }

  attachOfficeWindow(officeId: string, bounds: HostBounds): Promise<{ attached: true; hwnd: string }> {
    return this.transport.request("office.attach_window", this.withHost({ officeId, ...bounds }));
  }

  layoutOfficeWindow(officeId: string, bounds: HostBounds): Promise<{ attached: true }> {
    return this.transport.request("office.layout_window", this.withHost({ officeId, ...bounds }), { ensureBroker: false, timeoutMs: 1_500 });
  }

  detachOfficeWindow(officeId: string): Promise<{ attached: false }> {
    return this.transport.request("office.detach_window", this.withHost({ officeId }), { ensureBroker: true, timeoutMs: 20_000 });
  }

  configureOfficeWindow(officeId: string, payload: { enabled?: boolean; seconds?: number; saveOnClose?: boolean; ackAlerts?: boolean }): Promise<{ ackAlerts?: boolean; saveOnClose?: boolean }> {
    return this.transport.request("office.configure_save_window", this.withHost({ officeId, ...payload }));
  }

  /** Opens (or joins) the focus-hybrid session for a file; frames and state arrive on the returned stream pipe. */
  async openFocusSession(params: FocusOpenParams): Promise<FocusOpenResult> {
    validateSurfaceSize(params);
    // Core answers within 25 s even on a cold Office start; the tab never waits longer than 30 s.
    const value = await this.transport.request<unknown>("focus.open", this.withHost({ ...params }), { timeoutMs: FOCUS_OPEN_TIMEOUT_MS });
    return parseFocusOpenResult(value);
  }

  activateFocusSession(params: FocusActivateParams): Promise<{ state: FocusState; activationId: number }> {
    return this.transport.request("focus.activate", this.withHost({ ...params }), { timeoutMs: 3_000 });
  }

  layoutFocusSession(params: FocusLayoutParams): Promise<{ state: FocusState }> {
    validateSurfaceSize(params);
    return this.transport.request("focus.layout", this.withHost({ ...params }), { ensureBroker: false, timeoutMs: 1_500 });
  }

  /** Wheel over a parked tab; false when Core did not forward it (not parked, hidden tab, or another tab is active). */
  wheelFocusSession(params: FocusWheelParams): Promise<{ forwarded: boolean }> {
    return this.transport.request("focus.wheel", this.withHost({ ...params }), { ensureBroker: false, timeoutMs: 1_000 });
  }

  setFocusPixelPolicy(sessionId: string, pixelPolicy: PixelPolicy): Promise<{ pixelPolicy: PixelPolicy }> {
    return this.transport.request("focus.policy", this.withHost({ sessionId, pixelPolicy }));
  }

  async closeFocusSession(params: FocusCloseParams): Promise<FocusCloseResult> {
    return parseFocusCloseResult(await this.transport.request<unknown>("focus.close", this.withHost({ ...params }), { timeoutMs: 20_000 }));
  }

  focusStatus(sessionId: string): Promise<FocusStatus> {
    return this.transport.request("focus.status", this.withHost({ sessionId }), { ensureBroker: false });
  }

  async dashboard(): Promise<DashboardSnapshot> {
    const [health, offices, references, chats] = await Promise.all([
      this.health(), this.listOfficeInstances(), this.listReferences(), this.listRegisteredChats(),
    ]);
    return { health, offices, references, chats };
  }
}
