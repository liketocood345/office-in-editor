import assert from "node:assert/strict";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const root = fileURLToPath(new URL("../../../../", import.meta.url));
const compiled = await build({
  entryPoints: [fileURLToPath(new URL("../../../adapter-sdk/src/index.ts", import.meta.url))],
  bundle: true, platform: "node", format: "esm", target: "node18", write: false, absWorkingDir: root,
});
const sdk = await import("data:text/javascript;base64," + Buffer.from(compiled.outputFiles[0].text).toString("base64"));

const surfaceToken = "oie-surface-" + "a".repeat(32);
const sessionId = "b".repeat(32);
const openResult = {
  sessionId, application: "word", shared: false, state: "parked",
  stream: { pipeName: "office-in-editor-v1-0123456789abcdef-focus-" + sessionId, token: "c".repeat(64) },
};

function brokerWith(reply) {
  const calls = [];
  const broker = new sdk.OfficeBroker({ request: async (method, params, options) => { calls.push({ method, params, options }); return reply(method, params); } });
  return { broker, calls };
}

test("focus stream decoder reassembles events and frames split across chunks", () => {
  const image = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 9, 8, 7]);
  const bytes = Buffer.concat([
    sdk.encodeFocusEvent({ type: "state", state: "editing", cause: "click", activationId: 2, parkId: 1 }),
    sdk.encodeFocusFrame({ frameSeq: 5, parkId: 1, width: 640, height: 480, format: "png" }, image),
    sdk.encodeFocusEvent({ type: "notice", code: "modal_blocking" }),
  ]);
  const decoder = new sdk.FocusStreamDecoder();
  const messages = [];
  for (let offset = 0; offset < bytes.length; offset += 3) messages.push(...decoder.push(bytes.subarray(offset, offset + 3)));
  assert.equal(messages.length, 3);
  assert.deepEqual(messages[0], { kind: "event", event: { type: "state", state: "editing", cause: "click", activationId: 2, parkId: 1 } });
  assert.equal(messages[1].kind, "frame");
  assert.equal(messages[1].frame.frameSeq, 5);
  assert.deepEqual([...messages[1].frame.image], [...image]);
  assert.equal(messages[2].event.code, "modal_blocking");
});

test("focus stream decoder rejects oversize lengths, unknown kinds and bad frame headers", () => {
  const oversize = Buffer.alloc(5); oversize.writeUInt32LE(sdk.MAX_STREAM_MESSAGE_BYTES + 1, 0);
  assert.throws(() => new sdk.FocusStreamDecoder().push(oversize), { code: "INVALID_STREAM" });
  const unknown = Buffer.from([1, 0, 0, 0, 9]);
  assert.throws(() => new sdk.FocusStreamDecoder().push(unknown), { code: "INVALID_STREAM" });
  const jpeg = new sdk.FocusStreamDecoder().push(sdk.encodeFocusFrame({ frameSeq: 2, parkId: 1, width: 8, height: 8, format: "jpeg" }, Uint8Array.from([0xff, 0xd8])));
  assert.equal(jpeg[0].frame.format, "jpeg");
  const bmp = sdk.encodeFocusFrame({ frameSeq: 3, parkId: 1, width: 8, height: 8, format: "bmp" }, new Uint8Array(1));
  assert.throws(() => new sdk.FocusStreamDecoder().push(bmp), { code: "INVALID_STREAM" });
  const badFrame = sdk.encodeFocusFrame({ frameSeq: 1, parkId: 0, width: 1, height: 1, format: "png" }, new Uint8Array(1));
  badFrame.write('{"frameSeq":0', 7, "utf8");
  assert.throws(() => new sdk.FocusStreamDecoder().push(badFrame), { code: "INVALID_STREAM" });
  assert.throws(() => new sdk.FocusStreamDecoder().push(sdk.encodeFocusEvent({ type: "state", state: "flying", cause: "click", activationId: 0, parkId: 0 })), { code: "INVALID_STREAM" });
});

test("openFocusSession validates size locally and the response shape from Core", async () => {
  const { broker, calls } = brokerWith(() => openResult);
  broker.bindHost({ hostId: "h1", lease: "l1", epoch: "e1" });
  const opened = await broker.openFocusSession({ path: "a.docx", surfaceToken, editorPid: 7, cssWidth: 800, cssHeight: 600, pixelRatio: 1.5 });
  assert.equal(opened.sessionId, sessionId);
  assert.equal(calls[0].method, "focus.open");
  assert.equal(calls[0].params.hostLease, "l1");
  await assert.rejects(broker.openFocusSession({ path: "a.docx", surfaceToken, editorPid: 7, cssWidth: 10, cssHeight: 600, pixelRatio: 1 }), { code: "INVALID_PARAMS" });
  const bad = brokerWith(() => ({ ...openResult, stream: { pipeName: "\\\\.\\pipe\\elsewhere", token: "c".repeat(64) } }));
  await assert.rejects(bad.broker.openFocusSession({ path: "a.docx", surfaceToken, editorPid: 7, cssWidth: 800, cssHeight: 600, pixelRatio: 1 }), { code: "INVALID_RESPONSE" });
});

test("closeFocusSession surfaces the dirty prompt result and stream hello is strict", async () => {
  const { broker } = brokerWith(() => ({ closed: false, dirty: true, remainingSurfaces: 1 }));
  assert.deepEqual(await broker.closeFocusSession({ sessionId, surfaceToken }), { closed: false, dirty: true, remainingSurfaces: 1 });
  assert.equal(sdk.encodeStreamHello("c".repeat(64), surfaceToken), JSON.stringify({ v: 1, token: "c".repeat(64), surfaceToken }) + "\n");
  assert.throws(() => sdk.encodeStreamHello("short", surfaceToken), { code: "INVALID_PARAMS" });
});
