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
import { actionButton, capabilitiesNotice, escapeHtml, panelHtml } from "./ui";

const REF_ID = /^(?:[a-f0-9]{32}|[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})$/iu;
const DOCUMENT_EXTENSIONS = new Set([".docx", ".xlsx", ".pptx"]);

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
  private readonly view: vscode.TreeView<import("./tree").InboxNode>;
  private lease?: HostLease;
  private snapshot?: DashboardSnapshot;
  private lastError?: string;
  private readonly errorBudget = new ErrorNoticeBudget();
  private statusMessage?: vscode.Disposable;
  private connectPromise?: Promise<OfficeBroker>;
  private disposed = false;
  private shutdownPromise?: Promise<void>;
  readonly ready: Promise<void>;

  constructor(private readonly context: vscode.ExtensionContext) {
    this.transport = new NamedPipeBrokerTransport(
      { extensionPath: context.extensionPath },
      () => this.registerHost(),
    );
    this.broker = new OfficeBroker(this.transport);
    this.mcp = new McpRegistration(context.extensionPath);
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
      return;
    }
    this.lease = lease;
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
    panel.webview.html = panelHtml("Office 协作状态", capabilitiesNotice()
      + "<p>本机 Office 实例：" + snapshot.offices.length
      + " · 引用：" + snapshot.references.length
      + " · 已登记具体聊天：" + snapshot.chats.length + "</p>" + this.operationStatusHtml());
  }

  private operationStatusHtml(): string {
    const capability = this.snapshot?.health.capabilities.documentWrite;
    const text = capability === undefined
      ? "MCP 写能力：尚未读取 Core capabilities。"
      : capability
        ? "MCP 写能力：Core 已声明；具体操作受原生白名单、权限与可恢复事务约束。"
        : "MCP 写能力：Core 当前未声明可用。";
    return '<p class="muted">' + escapeHtml(text) + "</p>";
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
      nativeOfficeEmbedding: false,
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
      capabilitiesNotice()
      + '<p>诊断仅包含能力、版本、数量和错误代码；不包含文档文字、文件路径、聊天名称、令牌或用户标识。</p><pre>'
      + escapeHtml(JSON.stringify(report, null, 2)) + "</pre>");
  }

  private async chooseDocument(argument?: unknown): Promise<vscode.Uri | undefined> {
    const provided = argument instanceof vscode.Uri ? argument : undefined;
    if (provided) { await validateLocalDocument(provided); return provided; }
    const selection = await vscode.window.showOpenDialog({
      canSelectFiles: true,
      canSelectFolders: false,
      canSelectMany: false,
      openLabel: "选择 Office 文件",
      filters: { "Office 文档": ["docx", "xlsx", "pptx"] },
    });
    const uri = selection?.[0];
    if (uri) await validateLocalDocument(uri);
    return uri;
  }

  private async openDocument(argument?: unknown): Promise<void> {
    assertLocalAccess();
    const uri = await this.chooseDocument(argument);
    if (!uri) return;
    const key = "session:" + uri.toString();
    const existing = this.panels.get(key);
    if (existing) { existing.reveal(); return; }
    const name = displayText(path.basename(uri.fsPath));
    const panel = this.createPanel("officeInEditor.session", name + " · Office 会话控制", key);
    const render = (): void => {
      panel.webview.html = panelHtml(name,
        capabilitiesNotice()
        + this.operationStatusHtml()
        + '<p>所选文件：<strong>' + escapeHtml(name) + '</strong></p>'
        + '<p>在独立 Microsoft Office 窗口中编辑。此标签页用于捕获选区、查看引用和选择已登记的具体聊天。</p>'
        + '<div class="actions">'
        + actionButton("open", "在已安装的 Microsoft Office 中打开")
        + actionButton("capture", "选择 Office 实例并捕获选区")
        + actionButton("chats", "查看已登记具体聊天")
        + actionButton("refresh", "刷新状态") + "</div>"
        + '<p class="muted">捕获前请在 Office 中选中需要引用的内容。收件箱只在明确刷新或操作后更新，不轮询文件。</p>');
    };
    panel.webview.onDidReceiveMessage((message: unknown) => {
      if (!isRecord(message)) return;
      const action = message.action;
      if (action === "open") void this.run(() => this.openInOffice(uri));
      else if (action === "capture") void this.run(() => this.captureSelection());
      else if (action === "chats") void this.run(() => this.listRegisteredChats());
      else if (action === "refresh") void this.run(async () => { await this.refresh(); render(); });
    }, undefined, this.context.subscriptions);
    render();
  }

  private async openInOffice(argument?: unknown): Promise<void> {
    assertLocalAccess();
    const uri = await this.chooseDocument(argument);
    if (!uri) return;
    const broker = await this.connect();
    const result = await broker.openInOffice(uri.fsPath);
    if (result.opened !== true || result.mode !== "external") throw new OfficeBridgeError("INVALID_RESPONSE");
    this.announce("已请求本机已安装的 Microsoft Office 在独立窗口打开所选文件。");
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
    const available = delivery ? chats.filter(chat => chat.online !== false) : chats;
    if (!available.length) {
      this.announce("没有可用的已登记聊天；请先在目标聊天调用 office_register_chat 并打开 MCP App。");
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
      hostKind: hostKind(),
      mcpMode: this.mcp.status.mode,
      capabilities: {
        nativeOfficeEmbedding: false,
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
      this.statusMessage?.dispose();
      this.mcp.dispose();
      if (this.lease) {
        try { await this.broker.unregisterHost(this.lease); }
        catch { /* A closed broker is already detached. Never start a worker while shutting down. */ }
      }
    })();
    return this.shutdownPromise;
  }

  dispose(): void { void this.shutdown(); }
}

let activeExtension: OfficeExtension | undefined;

export interface OfficeExtensionDiagnostics {
  version: string;
  hostKind: "vscode" | "cursor";
  mcpMode: import("./mcp").McpRegistrationStatus["mode"];
  capabilities: {
    nativeOfficeEmbedding: false;
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
