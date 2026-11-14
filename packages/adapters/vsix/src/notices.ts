/** A finite, per-extension-session budget for redacted failures. */
export class ErrorNoticeBudget {
  private readonly loggedCodes = new Set<string>();
  private readonly notifiedCodes = new Set<string>();
  private noticesUsed = 0;

  constructor(private readonly maximumNotices = 1, private readonly maximumLogs = 32) {}

  observe(code: string, explicitUserAction: boolean): { log: boolean; notify: boolean } {
    const log = !this.loggedCodes.has(code) && this.loggedCodes.size < this.maximumLogs;
    if (log) this.loggedCodes.add(code);
    const notify = explicitUserAction
      && !this.notifiedCodes.has(code)
      && this.noticesUsed < this.maximumNotices;
    if (notify) {
      this.notifiedCodes.add(code);
      this.noticesUsed += 1;
    }
    return { log, notify };
  }
}
