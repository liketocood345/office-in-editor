import { build } from 'esbuild';
import { cp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { writeLicenseNotices } from './license-notices.mjs';
const root = process.cwd();
function run(command, args) {
  const result = spawnSync(command, args, { cwd: root, stdio: 'inherit', shell: false, windowsHide: true });
  if (result.status !== 0) throw new Error(`Build failed: ${command}`);
}
await mkdir('artifacts', { recursive: true });
await import('./build-extension.mjs');
const app = await build({ entryPoints: ['apps/mcp-app/main.ts'], bundle: true, platform: 'browser', target: 'es2022',
  format: 'iife', write: false, minify: true, sourcemap: false, legalComments: 'eof', metafile: true });
await writeLicenseNotices(Object.keys(app.metafile.inputs));
const template = await readFile('apps/mcp-app/template.html', 'utf8');
const html = template.replace('__APP_SCRIPT__', app.outputFiles[0].text.replace(/<\/script/gi, '<\\/script'));
run('dotnet', ['publish', 'packages/shared-core/OfficeInEditor.Core.csproj', '-c', 'Release', '-r', 'win-x64', '--self-contained', 'true',
  '-p:DebugType=None', '-p:DebugSymbols=false', '-o', path.join(root, 'payload/core')]);
await writeFile('payload/core/mcp-app.html', html);
run('dotnet', ['build', 'packages/office-addin/OfficeInEditor.OfficeAddin.csproj', '-c', 'Release', '-p:DebugType=None', '-p:DebugSymbols=false',
  '-o', path.join(root, 'payload/office')]);
// Include the existing license and explicit third-party notices in local payloads.
await cp('LICENSE', 'payload/LICENSE');
if (await readFile('THIRD_PARTY_NOTICES.md', 'utf8').catch(() => null)) await cp('THIRD_PARTY_NOTICES.md', 'payload/THIRD_PARTY_NOTICES.md');
await cp('THIRD_PARTY_LICENSES.md', 'payload/THIRD_PARTY_LICENSES.md');
console.log('Built one VSIX branch, shared core, Office add-in and MCP App. DSH was not built.');
