import * as vscode from "vscode";
import { resolveCoreExecutable } from "../../../adapter-sdk/src";

interface McpDefinitionProvider {
  provideMcpServerDefinitions(): Promise<unknown[]>;
  resolveMcpServerDefinition?(definition: unknown): unknown;
}
interface PublicHostApi {
  lm?: {
    registerMcpServerDefinitionProvider?: (id: string, provider: McpDefinitionProvider) => vscode.Disposable;
  };
  McpStdioServerDefinition?: new (label: string, command: string, args?: string[], env?: Record<string, string>) => unknown;
  cursor?: {
    mcp?: {
      registerServer?: (config: {
        name: string;
        server: { command: string; args: string[]; env: Record<string, string> };
      }) => void;
      unregisterServer?: (name: string) => void;
    };
  };
}

export interface McpRegistrationStatus {
  mode: "cursor-official" | "vscode-provider" | "unavailable" | "blocked" | "core-missing";
}

export function localAccessAllowed(): boolean {
  return process.platform === "win32" && !vscode.env.remoteName && vscode.workspace.isTrusted;
}

export function hostKind(): "cursor" | "vscode" {
  const api = vscode as unknown as PublicHostApi;
  return api.cursor?.mcp || /cursor/iu.test(vscode.env.appName) ? "cursor" : "vscode";
}

/** Public capability checks only. No IDE database, private commands, or configuration edits. */
export class McpRegistration implements vscode.Disposable {
  private registration?: vscode.Disposable;
  private cursorRegistered = false;
  private statusValue: McpRegistrationStatus = { mode: "unavailable" };
  private disposed = false;

  constructor(private readonly extensionPath: string) {}

  get status(): McpRegistrationStatus { return { ...this.statusValue }; }

  async refresh(): Promise<void> {
    if (this.disposed) return;
    if (!localAccessAllowed()) {
      this.statusValue = { mode: "blocked" };
      return;
    }
    const api = vscode as unknown as PublicHostApi;
    if (api.cursor?.mcp?.registerServer && api.cursor.mcp.unregisterServer) {
      if (this.cursorRegistered) return;
      let command: string;
      try { command = await resolveCoreExecutable({ extensionPath: this.extensionPath }); }
      catch { this.statusValue = { mode: "core-missing" }; return; }
      if (this.disposed) return;
      api.cursor.mcp.registerServer({
        name: "office-in-editor",
        server: { command, args: ["--mcp"], env: { OFFICE_IN_EDITOR_HOST_KIND: "cursor" } },
      });
      this.cursorRegistered = true;
      this.statusValue = { mode: "cursor-official" };
      return;
    }
    const register = api.lm?.registerMcpServerDefinitionProvider;
    const StdioDefinition = api.McpStdioServerDefinition;
    if (register && StdioDefinition) {
      if (!this.registration) {
        this.registration = register.call(api.lm, "office-in-editor", {
          provideMcpServerDefinitions: async () => {
            if (!localAccessAllowed() || this.disposed) return [];
            try {
              const command = await resolveCoreExecutable({ extensionPath: this.extensionPath });
              if (!localAccessAllowed() || this.disposed) return [];
              return [new StdioDefinition("Office In Editor", command, ["--mcp"], { OFFICE_IN_EDITOR_HOST_KIND: "vscode" })];
            } catch { return []; }
          },
          resolveMcpServerDefinition: definition => {
            if (!localAccessAllowed() || this.disposed) throw new Error("Office In Editor is unavailable in this workspace.");
            return definition;
          },
        });
      }
      this.statusValue = { mode: "vscode-provider" };
    } else this.statusValue = { mode: "unavailable" };
  }

  dispose(): void {
    this.disposed = true;
    this.registration?.dispose();
    this.registration = undefined;
    if (this.cursorRegistered) {
      const api = vscode as unknown as PublicHostApi;
      api.cursor?.mcp?.unregisterServer?.("office-in-editor");
      this.cursorRegistered = false;
    }
  }
}
