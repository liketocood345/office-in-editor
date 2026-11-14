import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

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
