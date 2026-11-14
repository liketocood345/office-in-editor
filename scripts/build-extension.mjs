import { build } from 'esbuild';
import { mkdir } from 'node:fs/promises';
await mkdir('dist', { recursive: true });
await build({ entryPoints: ['packages/adapters/vsix/src/extension.ts'], outfile: 'dist/extension.js', bundle: true,
  platform: 'node', target: 'node18', format: 'cjs', external: ['vscode'], sourcemap: false, legalComments: 'eof' });
