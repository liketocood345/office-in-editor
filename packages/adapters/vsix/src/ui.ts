import { randomBytes } from "node:crypto";
import { escapeHtml } from "../../../adapter-sdk/src";

export { escapeHtml };

export function actionButton(action: string, label: string): string {
  return '<button type="button" data-action="' + escapeHtml(action) + '">' + escapeHtml(label) + "</button>";
}

/** No runtime value enters script source, CSS, or an executable URL. */
export function panelHtml(title: string, body: string): string {
  const nonce = randomBytes(18).toString("base64");
  return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}'; base-uri 'none'; form-action 'none'">
<title>${escapeHtml(title)}</title>
<style nonce="${nonce}">
body{font-family:var(--vscode-font-family);padding:24px;max-width:960px;line-height:1.65;color:var(--vscode-foreground);background:var(--vscode-editor-background)}
h1{font-size:1.6em}h2{font-size:1.1em;margin-top:24px}
.notice{border-left:3px solid var(--vscode-focusBorder);padding:10px 16px;background:var(--vscode-textBlockQuote-background)}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin:20px 0}
button{cursor:pointer;font:inherit;padding:7px 12px;background:var(--vscode-button-background);color:var(--vscode-button-foreground);border:1px solid transparent;border-radius:3px}
button:focus-visible{outline:2px solid var(--vscode-focusBorder);outline-offset:3px}
button:hover{background:var(--vscode-button-hoverBackground)}
pre{white-space:pre-wrap;overflow-wrap:anywhere;background:var(--vscode-textCodeBlock-background);padding:14px}
dt{font-weight:600}dd{margin:0 0 10px;overflow-wrap:anywhere}.muted{color:var(--vscode-descriptionForeground)}
</style></head><body><h1>${escapeHtml(title)}</h1>${body}
<script nonce="${nonce}">
const api=acquireVsCodeApi();
for(const button of document.querySelectorAll('button[data-action]')){
  button.addEventListener('click',()=>api.postMessage({action:button.getAttribute('data-action')}));
}
</script></body></html>`;
}

export function capabilitiesNotice(): string {
  return '<p class="notice">本版本提供 Office 会话与协作控制标签页。Office 编辑窗口尚未嵌入编辑器；Office AutoSave 尚未重定向到本地。此控制标签页不直接编辑文档；模型可通过 MCP 调用 Core 已声明的 Office 原生白名单操作。具体能力以 Core 为准，引用不会自动提交聊天。</p>';
}
