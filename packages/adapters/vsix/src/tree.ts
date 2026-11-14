import * as vscode from "vscode";
import type {
  DashboardSnapshot, OfficeInstance, ReferenceMetadata, RegisteredChat,
} from "../../../adapter-sdk/src";

export type InboxNode =
  | { kind: "group"; key: "references" | "offices" | "chats" | "capabilities" }
  | { kind: "reference"; reference: ReferenceMetadata }
  | { kind: "office"; office: OfficeInstance }
  | { kind: "chat"; chat: RegisteredChat }
  | { kind: "info"; label: string };

export function displayText(value: unknown, limit = 120): string {
  return (typeof value === "string" ? value : "").replace(/[\u0000-\u001f\u007f]/gu, " ").slice(0, limit) || "未命名";
}

export class InboxTree implements vscode.TreeDataProvider<InboxNode>, vscode.Disposable {
  private readonly changes = new vscode.EventEmitter<InboxNode | undefined>();
  readonly onDidChangeTreeData = this.changes.event;
  private snapshot?: DashboardSnapshot;
  private message = "点击刷新，读取本机 Office 与引用状态。";

  setSnapshot(snapshot: DashboardSnapshot): void {
    this.snapshot = snapshot;
    this.message = "";
    this.changes.fire(undefined);
  }

  setMessage(message: string): void {
    this.message = message;
    this.changes.fire(undefined);
  }

  getTreeItem(node: InboxNode): vscode.TreeItem {
    if (node.kind === "group") {
      const counts = {
        references: this.snapshot?.references.length ?? 0,
        offices: this.snapshot?.offices.length ?? 0,
        chats: this.snapshot?.chats.length ?? 0,
        capabilities: undefined,
      };
      const labels = {
        references: "引用收件箱",
        offices: "已连接 Office 实例",
        chats: "已登记具体聊天",
        capabilities: "当前能力",
      };
      const item = new vscode.TreeItem(labels[node.key], vscode.TreeItemCollapsibleState.Expanded);
      item.id = node.key;
      item.description = counts[node.key] === undefined ? undefined : String(counts[node.key]);
      return item;
    }
    if (node.kind === "reference") {
      const ref = node.reference;
      const item = new vscode.TreeItem(displayText(ref.documentName));
      item.id = "reference:" + ref.refId;
      item.description = displayText(ref.kind) + " · " + Number(ref.textLength || 0) + " 字符";
      item.tooltip = "不可变选区快照；点击查看。引用 ID：" + displayText(ref.refId);
      item.contextValue = "officeInEditor.reference";
      item.iconPath = new vscode.ThemeIcon("quote");
      item.command = { command: "officeInEditor.showReference", title: "查看引用", arguments: [ref.refId] };
      return item;
    }
    if (node.kind === "office") {
      const item = new vscode.TreeItem(displayText(node.office.application));
      item.id = "office:" + node.office.officeId;
      item.description = "PID " + Number(node.office.pid);
      item.iconPath = new vscode.ThemeIcon("window");
      item.contextValue = "officeInEditor.office";
      item.command = { command: "officeInEditor.captureSelection", title: "捕获选区", arguments: [node.office.officeId] };
      return item;
    }
    if (node.kind === "chat") {
      const item = new vscode.TreeItem(displayText(node.chat.label));
      item.id = "chat:" + node.chat.id;
      item.description = node.chat.online ? "已登记" : "当前不可用";
      item.tooltip = "此项仅表示经 MCP App 显式登记的具体聊天。\nID：" + displayText(node.chat.id);
      item.iconPath = new vscode.ThemeIcon("comment-discussion");
      item.contextValue = "officeInEditor.chat";
      return item;
    }
    const item = new vscode.TreeItem(node.label);
    item.iconPath = new vscode.ThemeIcon("info");
    return item;
  }

  getChildren(node?: InboxNode): InboxNode[] {
    if (!node) {
      if (!this.snapshot) return [{ kind: "info", label: this.message }];
      return [
        ...(this.message ? [{ kind: "info" as const, label: this.message }] : []),
        { kind: "group", key: "references" },
        { kind: "group", key: "offices" },
        { kind: "group", key: "chats" },
        { kind: "group", key: "capabilities" },
      ];
    }
    if (node.kind !== "group" || !this.snapshot) return [];
    if (node.key === "references") return this.snapshot.references.length
      ? this.snapshot.references.map(reference => ({ kind: "reference", reference }))
      : [{ kind: "info", label: "暂无引用；在 Office 中选中内容后捕获。" }];
    if (node.key === "offices") return this.snapshot.offices.length
      ? this.snapshot.offices.map(office => ({ kind: "office", office }))
      : [{ kind: "info", label: "暂无 Office Add-in 实例。" }];
    if (node.key === "chats") return this.snapshot.chats.length
      ? this.snapshot.chats.map(chat => ({ kind: "chat", chat }))
      : [{ kind: "info", label: "请在目标聊天调用 office_register_chat。" }];
    return [
      { kind: "info", label: "Office 原生标签页嵌入：未实现" },
      { kind: "info", label: "Office AutoSave 本地重定向：未实现" },
      { kind: "info", label: "控制标签页编辑：未提供" },
      { kind: "info", label: this.snapshot.health.capabilities.documentWrite
        ? "MCP 写操作：按 Core 已声明的白名单与事务开放"
        : "MCP 写操作：Core 当前未声明写能力" },
      { kind: "info", label: "聊天自动提交：不支持" },
      { kind: "info", label: "引用仅排队；需 MCP App 确认附加。" },
    ];
  }

  dispose(): void { this.changes.dispose(); }
}
