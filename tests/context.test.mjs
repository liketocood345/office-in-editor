import test from 'node:test';
import assert from 'node:assert/strict';
import { prepareContextUpdate } from '../apps/mcp-app/context.mjs';

test('context updates retain earlier references without mutating previous state', () => {
  const first = prepareContextUpdate([], { refId: 'ref-a', text: 'Example selection A', documentName: 'Example.docx' });
  const second = prepareContextUpdate(first.items, { refId: 'ref-b', text: 'Example selection B' });
  assert.equal(first.items.length, 1);
  assert.equal(second.items.length, 2);
  assert.match(second.payload.content[0].text, /Example selection A/);
  assert.match(second.payload.content[0].text, /Example selection B/);
});
test('duplicate delivery does not duplicate or replace an immutable reference', () => {
  const first = prepareContextUpdate([], { refId: 'ref-a', text: 'Original' });
  const duplicate = prepareContextUpdate(first.items, { refId: 'ref-a', text: 'Changed' });
  assert.equal(duplicate.items.length, 1);
  assert.match(duplicate.payload.content[0].text, /Original/);
  assert.doesNotMatch(duplicate.payload.content[0].text, /Changed/);
});
test('context limits reject rather than silently truncate document text', () => {
  assert.throws(() => prepareContextUpdate([], { refId: 'ref-a', text: 'x'.repeat(1000) }, { maxBytes: 100 }), /context_capacity/);
  assert.throws(() => prepareContextUpdate([{ refId: 'old', text: 'old' }], { refId: 'new', text: 'new' }, { maxCount: 1 }), /context_capacity/);
});
test('empty selection is not attached', () => {
  assert.throws(() => prepareContextUpdate([], { refId: 'empty', text: ' ' }), /invalid_reference/);
});
