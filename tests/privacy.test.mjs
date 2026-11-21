import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { inspectText, scanRepository } from '../scripts/privacy-scan.mjs';
test('privacy gate identifies a machine path and a key without returning contents', () => {
  const sample = ['Q', ':', String.fromCharCode(92), 'private', String.fromCharCode(92), 'file.txt'].join('');
  const key = '-----BEGIN ' + 'PRIVATE KEY-----';
  const found = inspectText(`${sample}\n${key}`);
  assert.equal(found.length, 2);
  assert.ok(found.every(item => !Object.hasOwn(item, 'text')));
});
test('public URLs and dynamic path construction are allowed', () => {
  assert.equal(inspectText('https://github.com/example/project\nPath.Combine(root, "data")').length, 0);
});

test('deleted indexed files do not prevent scanning remaining tracked and untracked content', async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), 'oie-privacy-'));
  try {
    execFileSync('git', ['init', '--quiet'], { cwd: root, windowsHide: true });
    await writeFile(path.join(root, 'deleted.txt'), 'indexed, then deleted');
    const machinePath = ['Q', ':', String.fromCharCode(92), 'private', String.fromCharCode(92), 'file.txt'].join('');
    await writeFile(path.join(root, 'remaining.txt'), machinePath);
    execFileSync('git', ['add', 'deleted.txt', 'remaining.txt'], { cwd: root, windowsHide: true });
    await rm(path.join(root, 'deleted.txt'));
    await writeFile(path.join(root, 'untracked.txt'), 'safe content');

    const result = await scanRepository(root);
    assert.equal(result.fileCount, 2);
    assert.deepEqual(result.failures, [{ file: 'remaining.txt', rule: 'absolute-windows-path', line: 1 }]);
  } finally {
    // Delete only this fixture's resolved directory inside the temporary root.
    assert.equal(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
    assert.ok(path.basename(root).startsWith('oie-privacy-'));
    await rm(root, { recursive: true, force: true });
  }
});
