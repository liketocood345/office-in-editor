import * as vscode from "vscode";
import { randomUUID } from "node:crypto";
import { stat } from "node:fs/promises";
import * as path from "node:path";
import {
  NamedPipeBrokerTransport, OfficeBroker, OfficeBridgeError, errorCode, errorDescription, isRecord,
  referenceAsJson, referenceAsMarkdown,
} from "../../../adapter-sdk/src";
import type {
  DashboardSnapshot, HostLease, OfficeInstance, OfficeReference, ReferenceMetadata, RegisteredChat,
} from "../../../adapter-sdk/src";
import { McpRegistration, hostKind, localAccessAllowed } from "./mcp";
import { InboxTree, displayText } from "./tree";
import { ErrorNoticeBudget } from "./notices";
import { actionButton, escapeHtml, hostHtml, panelHtml } from "./ui";
import { tryHostBounds } from "./host-window";
import { officeTabAppearance } from "./office-theme";

const REF_ID = /^(?:[a-f0-9]{32}|[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})$/iu;
const DOCUMENT_EXTENSIONS = new Set([
  ".docx", ".doc", ".docm", ".dotx", ".dotm", ".rtf",
  ".xlsx", ".xls", ".xlsm", ".xlsb", ".xltx", ".xltm", ".csv",
  ".pptx", ".ppt", ".pptm", ".potx", ".potm", ".ppsx",
  ".vsdx", ".vsd", ".vstx", ".vssx", ".vstm",
  ".one", ".onepkg",
  ".pub", ".mpp",
]);

function isOfficeCacheFile(fsPath: string): boolean {
  return path.basename(fsPath).startsWith("~");
}

function assertLocalAccess(): void {
  if (process.platform !== "win32") throw new OfficeBridgeError("UNSUPPORTED_PLATFORM");
  if (vscode.env.remoteName) throw new OfficeBridgeError("REMOTE_NOT_SUPPORTED");
  if (!vscode.workspace.isTrusted) throw new OfficeBridgeError("WORKSPACE_NOT_TRUSTED");
}

function referenceId(value: unknown): string | undefined {
  if (typeof value === "string" && REF_ID.test(value)) return value;
  if (isRecord(value)) {
    if (typeof value.refId === "string" && REF_ID.test(value.refId)) return value.refId;
    if (isRecord(value.reference)) return referenceId(value.reference);
  }
  return undefined;
}

async function waitForOffice(
  broker: OfficeBroker,
  application: string,
  pid?: number,
): Promise<string> {
  const deadline = Date.now() + 20_000;
  while (Date.now() < deadline) {
    const offices = await broker.listOfficeInstances();
    const match = offices.find(office => office.application === application && (pid == null || office.pid === pid))
      ?? offices.find(office => office.application === application);
    if (match) return match.officeId;
    await new Promise(resolve => setTimeout(resolve, 250));
  }
  throw new OfficeBridgeError("OFFICE_NOT_FOUND");
}

function officeId(value: unknown): string | undefined {
  if (typeof value === "string" && value.length > 0 && value.length < 256) return value;
  if (isRecord(value) && isRecord(value.office)) return officeId(value.office.officeId);
  return undefined;
}

async function validateLocalDocument(uri: vscode.Uri): Promise<void> {
  if (uri.scheme !== "file" || uri.authority
    || uri.fsPath.startsWith("\\\\")
    || !DOCUMENT_EXTENSIONS.has(path.extname(uri.fsPath).toLowerCase())) {
    throw new OfficeBridgeError("INVALID_DOCUMENT");
  }
  if (isOfficeCacheFile(uri.fsPath)) throw new OfficeBridgeError("office_cache_file");
  try {
    if (!(await stat(uri.fsPath)).isFile()) throw new OfficeBridgeError("INVALID_DOCUMENT");
  } catch { throw new OfficeBridgeError("INVALID_DOCUMENT"); }
}

class OfficeExtension implements vscode.Disposable {
  private readonly tree = new InboxTree();
  private readonly transport: NamedPipeBrokerTransport;
  private readonly broker: OfficeBroker;
  private readonly mcp: McpRegistration;
  private readonly output = vscode.window.createOutputChannel("Office In Editor");
  private readonly hostId = hostKind() + ":" + randomUUID();
  private readonly panels = new Map<string, vscode.WebviewPanel>();
  private readonly officeClosers = new Set<() => Promise<void>>();
  private readonly officeAligners = new Map<string, { panel: vscode.WebviewPanel; realign: () => Promise<void> }>();
  private traceEvents = 0;
  private loading?: { stage: string; elapsedMs: number };
  private readonly view: vscode.TreeView<import("./tree").InboxNode>;
  private lease?: HostLease;
  private snapshot?: DashboardSnapshot;
  private lastError?: string;
  private readonly errorBudget = new ErrorNoticeBudget();
  private statusMessage?: vscode.Disposable;
  private connectPromise?: Promise<OfficeBroker>;
  private readonly convertingPreview = new Set<string>();
  private disposed = false;
  private shutdownPromise?: Promise<void>;
  private deliveryPump?: Promise<void>;
  readonly ready: Promise<void>;

  constructor(private readonly context: vscode.ExtensionContext) {
    this.transport = new NamedPipeBrokerTransport(
      { extensionPath: context.extensionPath },
      () => this.registerHost(),
    );
    this.broker = new OfficeBroker(this.transport);
    this.mcp = new McpRegistration(context.extensionPath, this.hostId);
    this.view = vscode.window.createTreeView("officeInEditor.inbox", {
      treeDataProvider: this.tree,
      showCollapseAll: false,
    });
    context.subscriptions.push(this.tree, this.view, this.output, this.mcp);
    const commands: Record<string, (argument?: unknown) => Promise<unknown>> = {
      refresh: () => this.refresh(),
      status: () => this.showStatus(),
      diagnostics: () => this.showDiagnostics(),
      openDocument: argument => this.openDocument(argument),
      openInOffice: argument => this.openInOffice(argument),
      realignWindow: () => this.realignOfficeWindow(),
      listOfficeInstances: () => this.listOfficeInstances(),
      captureSelection: argument => this.captureSelection(argument),
      listRegisteredChats: () => this.listRegisteredChats(),
      deliverReference: argument => this.deliverReference(argument),
      releaseReference: argument => this.releaseReference(argument),
      exportReference: argument => this.exportReference(argument),
      showReference: argument => this.showReference(argument),
    };
    for (const [name, handler] of Object.entries(commands)) {
      context.subscriptions.push(vscode.commands.registerCommand(
        "officeInEditor." + name,
        (argument?: unknown) => this.run(() => handler(argument)),
      ));
    }
    context.subscriptions.push(this.view.onDidChangeVisibility(event => {
      if (event.visible) void this.run(() => this.refresh(), false);
    }));
    const extension = this;
    context.subscriptions.push(vscode.window.registerCustomEditorProvider(
      "officeInEditor.document",
      {
        openCustomDocument(uri: vscode.Uri) {
          return { uri, dispose() { /* The editor tab holds no separate Office document model. */ } };
        },
        resolveCustomEditor(document: vscode.CustomDocument, webviewPanel: vscode.WebviewPanel) {
          // Resolve immediately: the iframe needs the editor to be resolved to lay out.
          void extension.hostDocument(document.uri, webviewPanel).catch(error => {
            extension.output.appendLine(`host_initialization ${errorCode(error)}`);
          });
        },
      },
      { webviewOptions: { retainContextWhenHidden: true }, supportsMultipleEditorsPerDocument: true },
    ));
    context.subscriptions.push(vscode.window.tabGroups.onDidChangeTabs(event => {
      void this.retainOfficePreviewTabs([...event.opened, ...event.changed]);
    }));
    context.subscriptions.push(vscode.workspace.onDidGrantWorkspaceTrust(() => {
      void this.run(async () => {
        await this.mcp.refresh();
        if (this.view.visible) await this.refresh();
      }, false);
    }));
    // Registration only resolves our owned Core path. --mcp remains a passive bridge.
    this.ready = this.run(async () => {
      await this.mcp.refresh();
      if (this.view.visible) await this.refresh();
    }, false);
  }

  async run(operation: () => Promise<unknown>, notify = true): Promise<void> {
    try { await operation(); }
    catch (error) {
      const code = errorCode(error);
      this.lastError = code;
      this.tree.setMessage(errorDescription(error) + " [" + code + "]");
      const decision = this.errorBudget.observe(code, notify);
      if (decision.log) this.output.appendLine(new Date().toISOString() + " " + code);
      this.announce(errorDescription(error) + " [" + code + "]");
      if (decision.notify && !this.disposed) void vscode.window.showErrorMessage(errorDescription(error) + " [" + code + "]");
    }
  }

  private announce(message: string): void {
    if (this.disposed) return;
    this.statusMessage?.dispose();
    this.statusMessage = vscode.window.setStatusBarMessage(message, 8_000);
  }

  private async registerHost(): Promise<void> {
    if (this.disposed) return;
    const lease = await this.broker.registerHost({
      hostId: this.hostId,
      kind: hostKind(),
      label: hostKind() === "cursor" ? "Cursor" : "VS Code",
    });
    if (this.disposed) {
      try { await this.broker.unregisterHost(lease); } catch { /* The old worker may already have stopped. */ }
      this.broker.bindHost(undefined);
      return;
    }
    this.lease = lease;
    this.broker.bindHost(lease);
    this.startDeliveryPump();
  }

  private startDeliveryPump(): void {
    if (this.deliveryPump || this.disposed) return;
    this.deliveryPump = this.pumpHostDeliveries();
  }

  private async pumpHostDeliveries(): Promise<void> {
    while (!this.disposed) {
      const lease = this.lease;
      if (!lease) {
        await new Promise(resolve => setTimeout(resolve, 500));
        continue;
      }
      try {
        const event = await this.broker.waitHostDelivery(lease, 25_000);
        if (this.disposed) return;
        if (event.waiting || typeof event.text !== "string") continue;
        const placed = await this.presentInChat(event.documentName ?? "Word", event.text);
        this.announce(placed ? "已放入聊天输入框，未自动发送。" : "已复制选区，聊天输入不可用。");
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "C", location: "extension.ts:pumpHostDeliveries", message: "host_delivery", data: { placed, textLength: event.text.length, kind: event.kind ?? "" }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
      } catch {
        await new Promise(resolve => setTimeout(resolve, 1_000));
      }
    }
  }

  private async presentInChat(documentName: string, text: string): Promise<boolean> {
    const query = documentName + "\n\n" + text;
    const cursor = hostKind() === "cursor";
    let placed = false;
    let used = "";
    // #region agent log
    fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "present_begin", data: { cursor, skipStartPrompt: true, exactResource: true, neverClose: cursor, textLength: text.length }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    try {
      await vscode.workspace.fs.createDirectory(this.context.globalStorageUri);
      const file = vscode.Uri.file(vscode.Uri.joinPath(this.context.globalStorageUri, "oie-chat-ref.md").fsPath);
      await vscode.workspace.fs.writeFile(file, new TextEncoder().encode(query));
      if (cursor) {
        try {
          await vscode.commands.executeCommand("composer.startComposerPrompt", { neverClose: true });
          // #region agent log
          fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "never_close_ok", data: { textLength: text.length }, timestamp: Date.now() }) }).catch(() => {});
          // #endregion
        } catch (error) {
          // #region agent log
          fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "never_close_fail", data: { error: error instanceof Error ? error.message : "unknown" }, timestamp: Date.now() }) }).catch(() => {});
          // #endregion
        }
      }
      try {
        await vscode.commands.executeCommand("composer.addfilestocomposer", file, { useExactResource: true });
        placed = true;
        used = "addfile-exact";
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "addfile_ok", data: { as: "exact", textLength: text.length, skipStartPrompt: true }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
      } catch {
        await vscode.commands.executeCommand("composer.addfilestocomposer", file);
        placed = true;
        used = "addfile-uri";
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "addfile_ok", data: { as: "uri", textLength: text.length, skipStartPrompt: true }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
      }
    } catch (error) {
      // #region agent log
      fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "addfile_fail", data: { error: error instanceof Error ? error.message : "unknown" }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
    }
    if (!cursor && !placed) {
      try {
        await vscode.commands.executeCommand("workbench.action.chat.open", { query, isPartialQuery: true });
        placed = true;
        used = "chat.open";
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "command_ok", data: { command: "workbench.action.chat.open", cursor, shape: "partial", textLength: text.length }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
      } catch (error) {
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "command_fail", data: { command: "workbench.action.chat.open", cursor, error: error instanceof Error ? error.message : "unknown" }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
      }
    }
    // #region agent log
    fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H-chat1", location: "extension.ts:presentInChat", message: "present_done", data: { placed, used, cursor, skipStartPrompt: true, textLength: text.length }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    return placed;
  }

  private async connect(): Promise<OfficeBroker> {
    assertLocalAccess();
    if (this.disposed) throw new OfficeBridgeError("EXTENSION_DISPOSED");
    if (this.connectPromise) return this.connectPromise;
    const pending = (async () => {
      await this.mcp.refresh();
      if (!this.lease) await this.registerHost();
      if (this.disposed) throw new OfficeBridgeError("EXTENSION_DISPOSED");
      return this.broker;
    })();
    this.connectPromise = pending;
    try { return await pending; }
    finally { if (this.connectPromise === pending) this.connectPromise = undefined; }
  }

  private async refresh(): Promise<void> {
    const broker = await this.connect();
    const snapshot = await broker.dashboard();
    if (this.disposed) return;
    this.snapshot = snapshot;
    this.lastError = undefined;
    this.tree.setSnapshot(snapshot);
  }

  private async showStatus(): Promise<void> {
    await this.refresh();
    const snapshot = this.snapshot!;
    const panel = this.createPanel("officeInEditor.status", "Office In Editor · 状态");
    panel.webview.html = panelHtml("Office 协作状态",
      "<p>Office " + snapshot.offices.length
      + " · 引用 " + snapshot.references.length
      + " · 聊天 " + snapshot.chats.length + "</p>");
  }

  private async showDiagnostics(): Promise<void> {
    if (localAccessAllowed()) {
      try { await this.refresh(); } catch (error) { this.lastError = errorCode(error); }
    }
    const report = {
      extensionVersion: displayText(this.context.extension.packageJSON.version, 64),
      platform: process.platform,
      architecture: process.arch,
      trustedWorkspace: vscode.workspace.isTrusted,
      remoteWorkspace: Boolean(vscode.env.remoteName),
      mcpRegistration: this.mcp.status.mode,
      protocolVersion: this.snapshot?.health.protocolVersion ?? 1,
      coreVersion: this.snapshot ? displayText(this.snapshot.health.version, 64) : "unavailable",
      nativeOfficeEmbedding: true,
      nativeAutoSaveLocalRedirect: false,
      panelEditing: false,
      agentOperations: "core-whitelist",
      coreDocumentWrite: this.snapshot?.health.capabilities.documentWrite ?? null,
      registeredOfficeCount: this.snapshot?.offices.length ?? 0,
      referenceCount: this.snapshot?.references.length ?? 0,
      registeredSpecificChatCount: this.snapshot?.chats.length ?? 0,
      lastErrorCode: this.lastError ?? null,
    };
    const panel = this.createPanel("officeInEditor.diagnostics", "Office In Editor · 脱敏诊断");
    panel.webview.html = panelHtml("脱敏诊断",
      "<pre>" + escapeHtml(JSON.stringify(report, null, 2)) + "</pre>");
  }

  private async chooseDocument(argument?: unknown): Promise<vscode.Uri | undefined> {
    const provided = argument instanceof vscode.Uri ? argument : undefined;
    if (provided) { await validateLocalDocument(provided); return provided; }
    const selection = await vscode.window.showOpenDialog({
      canSelectFiles: true,
      canSelectFolders: false,
      canSelectMany: false,
      openLabel: "选择 Office 文件",
      filters: { "Office 文档": ["docx", "doc", "xlsx", "xls", "pptx", "ppt", "vsdx", "vsd", "one", "pub", "mpp", "rtf"] },
    });
    const uri = selection?.[0];
    if (uri) await validateLocalDocument(uri);
    return uri;
  }

  private async retainOfficePreviewTabs(tabs: readonly vscode.Tab[]): Promise<void> {
    for (const tab of tabs) {
      if (!(tab.input instanceof vscode.TabInputCustom)) continue;
      if (tab.input.viewType !== "officeInEditor.document" || !tab.isPreview) continue;
      const key = tab.input.uri.toString();
      if (this.convertingPreview.has(key)) continue;
      this.convertingPreview.add(key);
      try {
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "B", location: "extension.ts:retainOfficePreviewTabs", message: "convert_preview", data: { preview: tab.isPreview, active: tab.isActive }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
        await vscode.commands.executeCommand("vscode.openWith", tab.input.uri, "officeInEditor.document", { preview: false, preserveFocus: true });
      } catch { /* Office tabs must not stay in the default preview-replace slot. */ }
      finally {
        setTimeout(() => this.convertingPreview.delete(key), 2000);
      }
    }
  }

  private async openDocument(argument?: unknown): Promise<void> {
    assertLocalAccess();
    const uri = await this.chooseDocument(argument);
    if (!uri) return;
    await vscode.commands.executeCommand("vscode.openWith", uri, "officeInEditor.document", { preview: false });
  }

  private async hostDocument(uri: vscode.Uri, panel: vscode.WebviewPanel): Promise<void> {
    panel.webview.options = { enableScripts: true, localResourceRoots: [] };
    const name = displayText(path.basename(uri.fsPath));
    if (isOfficeCacheFile(uri.fsPath)) {
      panel.webview.html = hostHtml(name, "这是 Office 缓存文件，不要打开。", "", officeTabAppearance());
      // #region agent log
      fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "C1", location: "extension.ts:hostDocument", message: "office_cache_file", data: { nameLen: name.length }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
      return;
    }
    const sessionId = "ole-" + randomUUID().replaceAll("-", "");
    const started = performance.now();
    const trace = (stage: string): void => {
      this.loading = { stage, elapsedMs: Math.round(performance.now() - started) };
      if (++this.traceEvents <= 1024) this.output.appendLine(`open=${sessionId.slice(4, 12)} stage=${stage} elapsedMs=${Math.round(performance.now() - started)}`);
    };
    trace("tab_created");
    // #region agent log
    fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "C", location: "extension.ts:hostDocument", message: "tab_created", data: { sessionPrefix: sessionId.slice(4, 12) }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    let officeId: string | undefined;
    let attached = false;
    let closed = false;
    let bounds: import("../../../adapter-sdk/src").HostBounds | undefined;
    let layoutRunning = false;
    let layoutVersion = 0;
    this.officeAligners.set(sessionId, { panel, realign: async () => {
      if (!officeId || closed || !bounds) throw new OfficeBridgeError("OFFICE_NOT_FOUND");
      await this.broker.attachOfficeWindow(officeId, { ...bounds, visible: panel.visible });
    } });
    let lastLayoutKey = "";
    const layout = async (message?: unknown): Promise<void> => {
      const next = tryHostBounds(message ?? bounds, bounds);
      if (!next) return;
      bounds = next;
      ++layoutVersion;
      if (!officeId || closed || layoutRunning) return;
      layoutRunning = true;
      try {
        let sent: number;
        do {
          sent = layoutVersion;
          await this.broker.layoutOfficeWindow(officeId, { ...bounds, visible: panel.visible });
          attached = panel.visible;
        } while (!closed && sent !== layoutVersion);
      } finally { layoutRunning = false; }
    };
    const close = async (): Promise<void> => {
      if (closed) return;
      closed = true;
      attached = false;
      this.officeClosers.delete(close);
      this.officeAligners.delete(sessionId);
      if (officeId && bounds) {
        try { await this.broker.layoutOfficeWindow(officeId, { ...bounds, visible: false }); }
        catch { /* Hide is best-effort; detach still destroys the HWND. */ }
      }
      // #region agent log
      fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "C", location: "extension.ts:close", message: "close_hide", data: { hasOffice: Boolean(officeId), hasBounds: Boolean(bounds) }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
      try { await this.broker.detachOfficeWindow(officeId ?? sessionId); }
      catch { /* Closing the tab must not kill Word/Excel/PowerPoint. */ }
    };
    this.officeClosers.add(close);
    panel.webview.onDidReceiveMessage((message: unknown) => {
      if (isRecord(message) && message.action === "layout") {
        if (!tryHostBounds(message, bounds)) return;
        if (!bounds) {
          trace("geometry_received");
          const first = tryHostBounds(message);
          // #region agent log
          fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "T1", location: "extension.ts:layout", message: "geometry_received", data: { width: first?.width, height: first?.height, x: first?.x, y: first?.y, tokenLen: first?.surfaceToken?.length ?? 0, hasToken: Boolean(first?.surfaceToken), elapsedMs: Math.round(performance.now() - started) }, timestamp: Date.now() }) }).catch(() => {});
          // #endregion
        }
        void panel.webview.postMessage({ kind: "layout-ack" });
        const sized = tryHostBounds(message, bounds);
        const layoutKey = sized ? sized.width + "x" + sized.height : "";
        if (sized && layoutKey !== lastLayoutKey) {
          lastLayoutKey = layoutKey;
          // #region agent log
          fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "G2", location: "extension.ts:layout", message: "layout_size", data: { w: sized.width, h: sized.height, dpr: sized.pixelRatio ?? message.dpr, innerW: message.innerW, vvW: message.vvW, vvScale: message.vvScale, source: "ole-root" }, timestamp: Date.now() }) }).catch(() => {});
          // #endregion
        }
        void this.run(() => layout(message), false);
      }
      if (isRecord(message) && message.action === "ready") trace("webview_ready");
      if (isRecord(message) && message.action === "script-error") trace("webview_script_error");
    }, undefined, this.context.subscriptions);
    const keepOfficeTab = async (): Promise<void> => {
      const tabOf = (): vscode.Tab | undefined => vscode.window.tabGroups.all
        .flatMap(group => group.tabs)
        .find(tab => tab.input instanceof vscode.TabInputCustom && tab.input.uri.toString() === uri.toString());
      const before = tabOf();
      try {
        if (panel.active) {
          await vscode.commands.executeCommand("workbench.action.keepEditor");
        }
      } catch { /* keepEditor is best-effort against preview replacement; do not pin. */ }
      const after = tabOf();
      // #region agent log
      fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "B", location: "extension.ts:keepOfficeTab", message: "tab_keep", data: { panelActive: panel.active, panelVisible: panel.visible, beforePreview: before?.isPreview ?? null, beforePinned: before?.isPinned ?? null, afterPreview: after?.isPreview ?? null, afterPinned: after?.isPinned ?? null }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
    };
    void keepOfficeTab();
    void this.retainOfficePreviewTabs(vscode.window.tabGroups.all.flatMap(group => group.tabs));
    panel.onDidChangeViewState(event => {
      // #region agent log
      fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H", location: "extension.ts:viewState", message: "view_state", data: { visible: event.webviewPanel.visible, active: event.webviewPanel.active, attached, closed, hasOffice: Boolean(officeId) }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
      void panel.webview.postMessage({ kind: "layout-request" });
      if (event.webviewPanel.active && attached && !closed) void keepOfficeTab();
      if (!officeId || closed || !bounds) return;
      void this.run(() => layout(bounds), false);
    }, undefined, this.context.subscriptions);
    panel.onDidDispose(() => {
      // #region agent log
      fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "H", location: "extension.ts:dispose", message: "tab_disposed", data: { attached, closed, hasOffice: Boolean(officeId) }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
      void close();
    }, undefined, this.context.subscriptions);
    const surfaceToken = "oie-surface-" + sessionId.slice(4);
    panel.webview.html = hostHtml(name, "正在打开 Office…", surfaceToken, officeTabAppearance());
    await this.run(async () => {
      const opening = this.connect();
      trace("connect_started");
      void opening.catch(() => {});
      const overallDeadline = performance.now() + 60_000;
      const deadline = Date.now() + 8_000;
      while (!closed && !bounds && Date.now() < deadline) {
        await panel.webview.postMessage({ kind: "layout-request" });
        await new Promise(resolve => setTimeout(resolve, 250));
      }
      if (closed) { await opening; return; }
      if (!bounds) {
        trace("geometry_timeout");
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "C", location: "extension.ts:geometry_timeout", message: "fail_no_external", data: { launchedExternal: false }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
        await panel.webview.postMessage({ kind: "status", text: "标签几何初始化失败。" });
        return;
      }
      const broker = await opening;
      trace("core_connected");
      if (closed) return;
      // Test-only handshake probe: deliberately never enters Office/OLE.
      if (this.context.extensionMode === vscode.ExtensionMode.Test && process.env.OIE_GEOMETRY_ONLY === "1") {
        trace("geometry_verified");
        closed = true;
        this.officeClosers.delete(close);
        return;
      }
      let opened;
      try {
        const remaining = overallDeadline - performance.now();
        if (remaining < 250) throw new OfficeBridgeError("HOST_WINDOW_UNAVAILABLE");
        trace("office_open_started");
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "T1", location: "extension.ts:office_open", message: "embedded_request", data: { fallback: false, remainingMs: Math.round(remaining), visible: panel.visible, elapsedMs: Math.round(performance.now() - started), ...bounds }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
        opened = await broker.openInOffice(uri.fsPath, "embedded", { fallback: false, ...bounds, visible: panel.visible, sessionId, timeoutMs: remaining });
        trace("office_open_completed");
      }
      catch (error) {
        const code = errorCode(error);
        trace(`open_failed_${code}`);
        // #region agent log
        fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "F", location: "extension.ts:open_failed", message: "fail_no_external", data: { code, launchedExternal: false, visible: panel.visible, elapsedMs: Math.round(performance.now() - started), remainingMs: Math.round(overallDeadline - performance.now()) }, timestamp: Date.now() }) }).catch(() => {});
        // #endregion
        const text = code === "REQUEST_TIMEOUT" ? "打开超时。" : "原位打开未完成。";
        void panel.webview.postMessage({ kind: "status", text });
        void broker.health().then(health => {
          if (health.nativeHost && ++this.traceEvents <= 1024) this.output.appendLine(JSON.stringify({ nativeHost: health.nativeHost }));
        }).catch(() => {});
        void close();
        return;
      }
      if (opened.opened !== true || opened.mode !== "embedded" || opened.inplace !== true || typeof opened.officeId !== "string") {
        void panel.webview.postMessage({ kind: "status", text: "原位打开未完成。" });
        void close();
        return;
      }
      officeId = opened.officeId;
      // #region agent log
      fetch("http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093", { method: "POST", headers: { "Content-Type": "application/json", "X-Debug-Session-Id": "4dd893" }, body: JSON.stringify({ sessionId: "4dd893", hypothesisId: "T1", location: "extension.ts:attached", message: "embedded_opened", data: { mode: opened.mode, inplace: opened.inplace, application: opened.application, elapsedMs: Math.round(performance.now() - started), loading: this.loading, shared: opened.shared === true, viewMode: opened.viewMode }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
      if (closed || this.disposed) { await broker.detachOfficeWindow(officeId); return; }
      await layout(bounds);
      await panel.webview.postMessage({ kind: "attached" });
      await keepOfficeTab();
      trace("attached");
    }, false);
  }

  private async realignOfficeWindow(): Promise<void> {
    const active = [...this.officeAligners.values()].find(entry => entry.panel.active);
    if (!active) return;
    await active.realign();
  }

  private async openInOffice(argument?: unknown): Promise<void> {
    assertLocalAccess();
    const uri = await this.chooseDocument(argument);
    if (!uri) return;
    await (await this.connect()).openInOffice(uri.fsPath, "external");
  }

  private async pickOffice(argument?: unknown): Promise<OfficeInstance | undefined> {
    const broker = await this.connect();
    const offices = await broker.listOfficeInstances();
    const requestedId = officeId(argument);
    if (requestedId) {
      const selected = offices.find(office => office.officeId === requestedId);
      if (!selected) throw new OfficeBridgeError("OFFICE_NOT_FOUND");
      return selected;
    }
    if (offices.length === 0) {
      this.announce("尚无 Office Add-in 实例；请先在 Microsoft Office 中启用 Office In Editor 加载项。");
      return undefined;
    }
    const selected = await vscode.window.showQuickPick(offices.map(office => ({
      label: displayText(office.application),
      description: "PID " + Number(office.pid),
      detail: "实例 ID：" + displayText(office.officeId),
      office,
    })), {
      title: "选择已连接的 Office 实例",
      placeHolder: "将读取该实例当前文档中的选区",
      ignoreFocusOut: true,
    });
    return selected?.office;
  }

  private async listOfficeInstances(): Promise<void> { await this.pickOffice(); }

  private async captureSelection(argument?: unknown): Promise<void> {
    const office = await this.pickOffice(argument);
    if (!office) return;
    const broker = await this.connect();
    const captured = await broker.captureSelection(office.officeId);
    await this.refresh();
    this.announce("已捕获不可变引用：" + displayText(captured.documentName) + "。请从收件箱查看；尚未投递或提交聊天。");
  }

  private async listRegisteredChats(): Promise<void> {
    const broker = await this.connect();
    await this.pickChat(await broker.listRegisteredChats(), false);
  }

  private async pickChat(chats: RegisteredChat[], delivery = true): Promise<RegisteredChat | undefined> {
    const available = chats;
    if (!available.length) {
      this.announce("没有可用的已登记聊天；请先在目标聊天调用 office_register_chat。");
      return undefined;
    }
    const selected = await vscode.window.showQuickPick(available.map(chat => ({
      label: displayText(chat.label),
      description: "ID " + displayText(chat.id, 16),
      detail: "仅为 MCP App 显式登记的具体聊天 · epoch " + displayText(chat.epoch, 16),
      chat,
    })), {
      title: delivery ? "选择接收引用的已登记具体聊天" : "已登记具体聊天",
      placeHolder: delivery ? "引用只排队，不自动提交提示词" : "此列表不代表 IDE 中所有聊天",
      ignoreFocusOut: true,
    });
    return selected?.chat;
  }

  private async pickReference(argument?: unknown): Promise<ReferenceMetadata | undefined> {
    const broker = await this.connect();
    const references = await broker.listReferences();
    const id = referenceId(argument);
    if (id) {
      const selected = references.find(reference => reference.refId === id);
      if (!selected) throw new OfficeBridgeError("REFERENCE_NOT_FOUND");
      return selected;
    }
    if (!references.length) {
      this.announce("暂无引用；请先在 Office 中选中内容并捕获选区。");
      return undefined;
    }
    const selected = await vscode.window.showQuickPick(references.map(reference => ({
      label: displayText(reference.documentName),
      description: displayText(reference.kind) + " · " + Number(reference.textLength || 0) + " 字符",
      detail: "引用 ID：" + displayText(reference.refId),
      reference,
    })), { title: "选择引用快照", ignoreFocusOut: true });
    return selected?.reference;
  }

  private async deliverReference(argument?: unknown): Promise<void> {
    const reference = await this.pickReference(argument);
    if (!reference) return;
    const broker = await this.connect();
    const target = await this.pickChat(await broker.listRegisteredChats());
    if (!target) return;
    const result = await broker.deliverReference(reference.refId, target.id);
    if (result.status !== "queued") throw new OfficeBridgeError("INVALID_DELIVERY_RESPONSE", true);
    this.announce("引用已向“" + displayText(target.label) + "”排队；尚未确认附加，不会自动发送。");
    await this.refresh();
  }

  private async showReference(argument?: unknown): Promise<void> {
    const metadata = await this.pickReference(argument);
    if (!metadata) return;
    const broker = await this.connect();
    const reference = await broker.getReference(metadata.refId);
    const key = "reference:" + reference.refId;
    const existing = this.panels.get(key);
    if (existing) { existing.reveal(); return; }
    const panel = this.createPanel("officeInEditor.reference", "Office 引用 · " + displayText(reference.documentName), key);
    panel.webview.html = this.referenceHtml(reference);
    panel.webview.onDidReceiveMessage((message: unknown) => {
      if (!isRecord(message)) return;
      if (message.action === "deliver") void this.run(() => this.deliverReference(reference.refId));
      else if (message.action === "export") void this.run(() => this.exportReference(reference.refId));
      else if (message.action === "release") void this.run(() => this.releaseReference(reference.refId));
    }, undefined, this.context.subscriptions);
  }

  private referenceHtml(reference: OfficeReference): string {
    return panelHtml("Office 引用快照",
      '<p>此处显示捕获时的不可变选区；不修改原文档。</p>'
      + '<dl><dt>文档</dt><dd>' + escapeHtml(reference.documentName)
      + '</dd><dt>类型</dt><dd>' + escapeHtml(reference.application + " / " + reference.kind)
      + '</dd><dt>引用 ID</dt><dd>' + escapeHtml(reference.refId)
      + '</dd><dt>捕获时文档状态</dt><dd>' + (reference.dirty ? "有未保存更改" : "无未保存更改")
      + '</dd></dl><div class="actions">'
      + actionButton("deliver", "向已登记聊天排队")
      + actionButton("export", "导出引用")
      + actionButton("release", "释放引用快照")
      + '</div><h2>选区文字</h2><pre>' + escapeHtml(reference.text)
      + '</pre><h2>语义位置</h2><pre>' + escapeHtml(JSON.stringify(reference.locator, null, 2)) + "</pre>");
  }

  private async releaseReference(argument?: unknown): Promise<void> {
    const reference = await this.pickReference(argument);
    if (!reference) return;
    const confirmed = await vscode.window.showWarningMessage("释放此引用快照？", {
      modal: true,
      detail: "这只删除本机协作引用，不删除或修改原始 Office 文件。已有聊天中对该引用的后续读取可能失效。",
    }, "释放引用");
    if (confirmed !== "释放引用") return;
    const broker = await this.connect();
    await broker.releaseReference(reference.refId);
    this.panels.get("reference:" + reference.refId)?.dispose();
    await this.refresh();
    this.announce("引用快照已释放。");
  }

  private async exportReference(argument?: unknown): Promise<void> {
    const metadata = await this.pickReference(argument);
    if (!metadata) return;
    const destination = await vscode.window.showSaveDialog({
      defaultUri: vscode.Uri.file(path.join(
        this.context.globalStorageUri.fsPath, "office-reference-" + metadata.refId.slice(0, 8) + ".md",
      )),
      filters: { "Markdown 引用": ["md"], "JSON 引用": ["json"] },
      saveLabel: "导出引用快照",
    });
    if (!destination) return;
    if (destination.scheme !== "file" || destination.authority
      || ![".md", ".json"].includes(path.extname(destination.fsPath).toLowerCase())) {
      throw new OfficeBridgeError("INVALID_EXPORT_PATH");
    }
    const broker = await this.connect();
    const reference = await broker.getReference(metadata.refId);
    const content = path.extname(destination.fsPath).toLowerCase() === ".json"
      ? referenceAsJson(reference)
      : referenceAsMarkdown(reference);
    await vscode.workspace.fs.writeFile(destination, Buffer.from(content, "utf8"));
    this.announce("引用快照已导出。");
  }

  private createPanel(viewType: string, title: string, key?: string): vscode.WebviewPanel {
    const panel = vscode.window.createWebviewPanel(viewType, title, vscode.ViewColumn.Active, {
      enableScripts: true,
      localResourceRoots: [],
      retainContextWhenHidden: false,
    });
    if (key) {
      this.panels.set(key, panel);
      panel.onDidDispose(() => {
        if (this.panels.get(key) === panel) this.panels.delete(key);
      }, undefined, this.context.subscriptions);
    }
    this.context.subscriptions.push(panel);
    return panel;
  }

  /** Read-only export for diagnostics and extension-host tests. It never contacts Core. */
  getDiagnostics(): OfficeExtensionDiagnostics {
    return {
      version: displayText(this.context.extension.packageJSON.version, 64),
      loading: this.loading,
      hostKind: hostKind(),
      mcpMode: this.mcp.status.mode,
      capabilities: {
        nativeOfficeEmbedding: true,
        nativeAutoSaveLocalRedirect: false,
        panelEditing: false,
        agentOperations: "core-whitelist",
        coreDocumentWrite: this.snapshot?.health.capabilities.documentWrite ?? null,
        automaticChatSubmit: false,
      },
      trustedWorkspace: vscode.workspace.isTrusted,
      remoteWorkspace: Boolean(vscode.env.remoteName),
    };
  }

  shutdown(): Promise<void> {
    this.shutdownPromise ??= (async () => {
      this.disposed = true;
      await Promise.allSettled([...this.officeClosers].map(close => close()));
      this.statusMessage?.dispose();
      this.mcp.dispose();
      if (this.lease) {
        try { await this.broker.unregisterHost(this.lease); }
        catch { /* A closed broker is already detached. Never start a worker while shutting down. */ }
        this.broker.bindHost(undefined);
      }
    })();
    return this.shutdownPromise;
  }

  dispose(): void { void this.shutdown(); }
}

let activeExtension: OfficeExtension | undefined;

export interface OfficeExtensionDiagnostics {
  loading?: { stage: string; elapsedMs: number };
  version: string;
  hostKind: "vscode" | "cursor";
  mcpMode: import("./mcp").McpRegistrationStatus["mode"];
  capabilities: {
    nativeOfficeEmbedding: boolean;
    nativeAutoSaveLocalRedirect: false;
    panelEditing: false;
    agentOperations: "core-whitelist";
    coreDocumentWrite: boolean | null;
    automaticChatSubmit: false;
  };
  trustedWorkspace: boolean;
  remoteWorkspace: boolean;
}

export interface OfficeExtensionApi {
  getDiagnostics(): OfficeExtensionDiagnostics;
}

export async function activate(context: vscode.ExtensionContext): Promise<OfficeExtensionApi> {
  const extension = new OfficeExtension(context);
  activeExtension = extension;
  context.subscriptions.push(extension);
  await extension.ready;
  return Object.freeze({ getDiagnostics: () => extension.getDiagnostics() });
}

export function deactivate(): Promise<void> | undefined {
  return activeExtension?.shutdown();
}
