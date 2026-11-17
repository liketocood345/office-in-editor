import * as path from "node:path";

/** Office lock/temp files that must not be hosted as documents. */
export function isOfficeCacheFileName(name: string): boolean {
  return name.startsWith("~");
}

export function isOfficeCacheFile(fsPath: string): boolean {
  return isOfficeCacheFileName(path.basename(fsPath));
}

/**
 * Map `~$文档.docx` to `文档.docx` in the same directory.
 * Other `~` temps (such as `~WRL0001.tmp`) have no sibling document.
 */
export function officeCacheOwnerName(name: string): string | undefined {
  if (!name.startsWith("~$") || name.length <= 2) return undefined;
  const owner = name.slice(2);
  if (!owner || isOfficeCacheFileName(owner)) return undefined;
  return owner;
}

export function officeCacheOwnerPath(fsPath: string): string | undefined {
  const ownerName = officeCacheOwnerName(path.basename(fsPath));
  if (!ownerName) return undefined;
  return path.join(path.dirname(fsPath), ownerName);
}
