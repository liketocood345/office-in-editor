import { access } from 'node:fs/promises';
import path from 'node:path';

const root = process.cwd();
const required = [
  'payload/core/OfficeInEditor.Core.exe',
  'payload/core/mcp-app.html',
  'payload/office/OfficeInEditor.Addin.dll',
  'payload/LICENSE',
  'payload/THIRD_PARTY_LICENSES.md'
];
const missing = [];
for (const relative of required) {
  try { await access(path.join(root, relative)); }
  catch { missing.push(relative); }
}
if (missing.length) {
  console.error('Setup payload is incomplete. Run npm run build first.');
  for (const relative of missing) console.error(relative);
  process.exitCode = 1;
} else {
  console.log('Setup payload is present. Install with scripts/install.ps1 after packaging the VSIX. This script does not register Office or editors.');
}
