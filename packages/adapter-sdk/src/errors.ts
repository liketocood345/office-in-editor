/** Errors contain stable codes only; private worker messages are never logged. */
export class OfficeBridgeError extends Error {
  readonly code: string;
  readonly outcomeUnknown: boolean;

  constructor(code: string, outcomeUnknown = false) {
    const safeCode = /^[a-z0-9_.:/-]{1,80}$/iu.test(code) ? code : "UNKNOWN_ERROR";
    super(safeCode);
    this.name = "OfficeBridgeError";
    this.code = safeCode;
    this.outcomeUnknown = outcomeUnknown;
  }
}

export function errorCode(error: unknown): string {
  return error instanceof OfficeBridgeError ? error.code : "UNEXPECTED_ERROR";
}

export function errorDescription(error: unknown): string {
  if (error instanceof OfficeBridgeError && error.outcomeUnknown) {
    return "操作结果尚未确认。请先刷新状态，再决定下一步；本插件不会自动重试写入。";
  }
  const code = errorCode(error).toUpperCase();
  if (code === "CORE_NOT_FOUND") return "未找到 Office In Editor Core。请安装产品运行时后重试。";
  if (code === "UNSUPPORTED_PLATFORM") return "当前版本仅支持本机 Windows。";
  if (code === "WORKSPACE_NOT_TRUSTED") return "请先信任此工作区，再使用本机 Office 协作功能。";
  if (code === "REMOTE_NOT_SUPPORTED") return "当前版本不支持远程工作区中的 Office 控制。";
  if (code === "BROKER_UNAVAILABLE") return "本机协作服务尚不可用。请刷新重试。";
  if (code === "REQUEST_TIMEOUT") return "本机协作服务响应超时。";
  if (code === "INVALID_DOCUMENT") return "请选择本机现有的 Office 文件。";
  if (code === "OFFICE_CACHE_FILE") return "这是 Office 缓存文件，不要打开。";
  if (code.includes("NOT_FOUND") || code.includes("EXPIRED")) return "所选会话、引用或 Office 实例已不可用，请刷新后重新选择。";
  return "操作未完成。可查看脱敏诊断中的错误代码。";
}
