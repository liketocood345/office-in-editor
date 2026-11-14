import test from 'node:test';
import assert from 'node:assert/strict';
import { inspectText } from '../scripts/privacy-scan.mjs';
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
