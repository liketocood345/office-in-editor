import { randomUUID } from "node:crypto";
import { OfficeBridgeError } from "./errors";
import { isRecord } from "./protocol";
import type {
  BrokerTransport, CoreHealth, DashboardSnapshot, DeliveryReceipt, HostLease,
  HostRegistration, OfficeInstance, OfficeReference, ReferenceMetadata, RegisteredChat,
} from "./types";

function arrayResult<T>(value: unknown): T[] {
  if (!Array.isArray(value)) throw new OfficeBridgeError("INVALID_RESPONSE");
  return value as T[];
}

/** Typed local-only feature boundary. Reference text is retrieved only on explicit user action. */
export class OfficeBroker {
  constructor(private readonly transport: BrokerTransport) {}

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

  openInOffice(documentPath: string): Promise<{ opened: true; mode: "external"; application: string }> {
    return this.transport.request("office.open", { path: documentPath });
  }

  async dashboard(): Promise<DashboardSnapshot> {
    const [health, offices, references, chats] = await Promise.all([
      this.health(), this.listOfficeInstances(), this.listReferences(), this.listRegisteredChats(),
    ]);
    return { health, offices, references, chats };
  }
}
