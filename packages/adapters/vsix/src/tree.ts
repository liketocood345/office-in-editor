import * as vscode from "vscode";
import type {
  DashboardSnapshot, OfficeInstance, ReferenceMetadata, RegisteredChat,
} from "../../../adapter-sdk/src";

export type InboxNode =
  | { kind: "group"; key: "references" | "offices" | "chats" }
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
  private message = "";

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
      };
      const labels = {
        references: "引用收件箱",
        offices: "已连接 Office 实例",
        chats: "已登记具体聊天",
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
      item.tooltip = displayText(ref.refId);
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
      item.tooltip = displayText(node.chat.id);
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
      if (!this.snapshot) return this.message ? [{ kind: "info", label: this.message }] : [];
      return [
        ...(this.message ? [{ kind: "info" as const, label: this.message }] : []),
        { kind: "group", key: "references" },
        { kind: "group", key: "offices" },
        { kind: "group", key: "chats" },
      ];
    }
    if (node.kind !== "group" || !this.snapshot) return [];
    if (node.key === "references") return this.snapshot.references.length
      ? this.snapshot.references.map(reference => ({ kind: "reference", reference }))
      : [];
    if (node.key === "offices") return this.snapshot.offices.length
      ? this.snapshot.offices.map(office => ({ kind: "office", office }))
      : [];
    if (node.key === "chats") return this.snapshot.chats.length
      ? this.snapshot.chats.map(chat => ({ kind: "chat", chat }))
      : [];
    return [];
  }

  dispose(): void { this.changes.dispose(); }
}
