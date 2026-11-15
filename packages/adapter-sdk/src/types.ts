export const PROTOCOL_VERSION = 1 as const;

export interface CoreEndpoint {
  pipeName: string;
  protocolVersion: 1;
  version: string;
}

export interface CoreHealth {
  nativeHost?: { lastFailure: string; trace: { stage: string; elapsedMs: number }[] };
  version: string;
  protocolVersion: 1;
  capabilities: {
    nativeOfficeEmbedding: boolean;
    nativeAutoSaveLocalRedirect: false;
    documentWrite: boolean;
    [key: string]: unknown;
  };
  appPort?: number | null;
}

export type OfficeApplication = "word" | "excel" | "powerpoint";

export interface HostBounds {
  x: number;
  y: number;
  width: number;
  height: number;
  visible?: boolean;
  surfaceToken?: string;
}

export interface OfficeInstance {
  officeId: string;
  application: OfficeApplication;
  pid: number;
  registeredAt?: string;
}

export interface ReferenceMetadata {
  refId: string;
  id?: string;
  application: OfficeApplication;
  documentId: string;
  documentName: string;
  kind: string;
  createdAt: string;
  textLength: number;
  sha256: string;
  dirty: boolean;
  autoSaveAvailable: boolean | null;
  autoSaveEnabled: boolean | null;
  autoSaveState?: string;
}

export interface OfficeReference extends ReferenceMetadata {
  text: string;
  locator: unknown;
  sourcePath?: string;
}

export interface RegisteredChat {
  id: string;
  label: string;
  hostId?: string;
  epoch: string;
  ownerId?: string;
  online: boolean;
}

export interface HostLease {
  hostId: string;
  lease: string;
  epoch: string;
}

export interface DeliveryReceipt {
  requestId: string;
  refId: string;
  targetId: string;
  status: "queued";
  createdAt?: string;
}

export interface DashboardSnapshot {
  health: CoreHealth;
  offices: OfficeInstance[];
  references: ReferenceMetadata[];
  chats: RegisteredChat[];
}

export interface HostRegistration {
  hostId: string;
  kind: "vscode" | "cursor";
  label: string;
}

export type BrokerMethod =
  | "health"
  | "host.register"
  | "host.unregister"
  | "host.wait_delivery"
  | "office.list"
  | "office.capture"
  | "office.open"
  | "office.describe_window"
  | "office.attach_window"
  | "office.layout_window"
  | "office.detach_window"
  | "office.save_window"
  | "office.configure_save_window"
  | "office.status_window"
  | "reference.list"
  | "reference.get"
  | "reference.release"
  | "reference.deliver"
  | "chat.list";

export interface BrokerTransport {
  request<T = unknown>(
    method: BrokerMethod,
    params?: Record<string, unknown>,
    options?: { ensureBroker?: boolean; timeoutMs?: number },
  ): Promise<T>;
}
