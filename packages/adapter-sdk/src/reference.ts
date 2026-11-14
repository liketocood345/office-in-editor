import type { OfficeReference } from "./types";

/** Explicit export deliberately omits source paths, transport identities and chat tokens. */
export function referenceExportValue(reference: OfficeReference): Record<string, unknown> {
  return {
    schemaVersion: 1,
    refId: reference.refId,
    application: reference.application,
    documentName: reference.documentName,
    kind: reference.kind,
    createdAt: reference.createdAt,
    text: reference.text,
    locator: reference.locator,
    sha256: reference.sha256,
    dirtyAtCapture: reference.dirty,
  };
}

export function referenceAsJson(reference: OfficeReference): string {
  return JSON.stringify(referenceExportValue(reference), null, 2) + "\n";
}

export function referenceAsMarkdown(reference: OfficeReference): string {
  const title = String(reference.documentName).replace(/[\r\n]+/gu, " ");
  const quote = reference.text.split(/\r?\n/u).map(line => "> " + line).join("\n");
  const metadata = { ...referenceExportValue(reference) };
  delete metadata.text;
  return "# Office reference: " + title + "\n\n" + quote
    + "\n\nMetadata (JSON):\n\n" + JSON.stringify(metadata, null, 2) + "\n";
}

export function escapeHtml(value: unknown): string {
  return String(value ?? "").replace(/[&<>"']/gu, character => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[character] ?? character);
}
