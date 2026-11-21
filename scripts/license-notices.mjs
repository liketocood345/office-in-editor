import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

/**
 * Copies the upstream .NET and ASP.NET Core third-party notices next to the self-contained runtime in payload/core.
 * The runtime packs are found through the Core project's restore output and the versions recorded in its deps file.
 */
export async function copyRuntimeNotices(coreDir, assetsFile) {
  const assets = JSON.parse(await readFile(assetsFile, 'utf8'));
  const folders = Object.keys(assets.packageFolders ?? {});
  const deps = JSON.parse(await readFile(path.join(coreDir, 'OfficeInEditor.Core.deps.json'), 'utf8'));
  const packs = Object.keys(deps.libraries ?? {}).filter(key => key.startsWith('runtimepack.'));
  const wanted = [['Microsoft.NETCore.App.Runtime.', 'dotnet-THIRD-PARTY-NOTICES.TXT'], ['Microsoft.AspNetCore.App.Runtime.', 'aspnetcore-THIRD-PARTY-NOTICES.TXT']];
  for (const [prefix, target] of wanted) {
    const pack = packs.find(key => key.startsWith('runtimepack.' + prefix));
    if (!pack) throw new Error(`Missing runtime pack in Core deps: ${prefix}`);
    const [name, version] = pack.slice('runtimepack.'.length).split('/');
    let text = '';
    for (const folder of folders) {
      text = await readFile(path.join(folder, name.toLowerCase(), version, 'THIRD-PARTY-NOTICES.TXT'), 'utf8').catch(() => '');
      if (text) break;
    }
    if (!text) throw new Error(`Missing upstream notices for ${name} ${version}`);
    await writeFile(path.join(coreDir, target), text);
  }
}

export async function writeLicenseNotices(inputs) {
  const packageRoots = new Set();
  for (const input of inputs) {
    const normalized = input.replaceAll('\\', '/');
    const marker = normalized.lastIndexOf('node_modules/');
    if (marker < 0) continue;
    const remainder = normalized.slice(marker + 'node_modules/'.length).split('/');
    const name = remainder[0].startsWith('@') ? remainder.slice(0, 2).join('/') : remainder[0];
    packageRoots.add(normalized.slice(0, marker) + 'node_modules/' + name);
  }
  const sections = ['# Bundled JavaScript dependency licenses\n\nGenerated from the MCP App bundle inputs. No machine paths or user data are included.'];
  for (const root of [...packageRoots].sort()) {
    const pkg = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
    let license = '';
    for (const name of ['LICENSE', 'LICENSE.txt', 'LICENSE.md', 'LICENSE-MIT', 'COPYING', 'license']) {
      license = await readFile(path.join(root, name), 'utf8').catch(() => '');
      if (license) break;
    }
    if (!license) throw new Error(`Missing bundled dependency license: ${pkg.name}`);
    sections.push(`## ${pkg.name} ${pkg.version}\n\nLicense: ${pkg.license ?? 'see notice'}\n\n${license.trim()}`);
  }
  await writeFile('THIRD_PARTY_LICENSES.md', sections.join('\n\n') + '\n');
}
