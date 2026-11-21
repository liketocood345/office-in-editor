import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import * as net from "node:net";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const root = fileURLToPath(new URL("../../../../", import.meta.url));
async function load(entry) {
  const compiled = await build({ entryPoints: [fileURLToPath(new URL(entry, import.meta.url))], bundle: true, platform: "node",
    format: "esm", target: "node18", write: false, absWorkingDir: root });
  return import("data:text/javascript;base64," + Buffer.from(compiled.outputFiles[0].text).toString("base64"));
}
const editorModule = await load("../src/focus-editor.ts");
const streamModule = await load("../src/focus-stream.ts");
const uiModule = await load("../src/focus-ui.ts");
const sdk = await load("../../../adapter-sdk/src/index.ts");

const tick = () => new Promise(resolve => setTimeout(resolve, 5));
async function until(condition, ms = 2000) {
  const deadline = Date.now() + ms;
  while (!condition()) { if (Date.now() > deadline) throw new Error("timed out"); await tick(); }
}

function emitter() {
  const listeners = [];
  const event = listener => { listeners.push(listener); return { dispose() { listeners.splice(listeners.indexOf(listener), 1); } }; };
  return { event, fire: value => [...listeners].forEach(listener => listener(value)) };
}

function rig(overrides = {}) {
  const sessionId = "a".repeat(32);
  const received = emitter(), viewState = emitter(), disposed = emitter();
  const posted = [];
  const panel = {
    visible: true,
    webview: { postMessage: async message => { posted.push(message); return true; }, onDidReceiveMessage: received.event },
    onDidChangeViewState: viewState.event,
    onDidDispose: disposed.event,
  };
  const calls = [];
  let handlers;
  const closeReplies = overrides.closeReplies ?? [{ closed: true, dirty: false, remainingSurfaces: 0 }];
  const broker = {
    openFocusSession: async params => { calls.push(["open", params]); await overrides.openGate; return { sessionId, application: "word", shared: false, state: "releasing",
      stream: { pipeName: "office-in-editor-v1-0123456789abcdef-focus-" + sessionId, token: "b".repeat(64) } }; },
    activateFocusSession: async params => { calls.push(["activate", params]); return { state: "editing", activationId: params.activationId }; },
    wheelFocusSession: async params => { calls.push(["wheel", params]); await overrides.wheelGate; return { forwarded: true }; },
    layoutFocusSession: async params => { calls.push(["layout", params]); return { state: "parked" }; },
    closeFocusSession: async params => { calls.push(["close", params]); return closeReplies.length > 1 ? closeReplies.shift() : closeReplies[0]; },
  };
  const traces = [];
  const deps = {
    broker,
    stream: (pipeName, token, surfaceToken, h) => { handlers = h; calls.push(["stream", { pipeName, token, surfaceToken }]); return { connect: async () => {}, dispose() {} }; },
    askSave: overrides.askSave ?? (async () => "save"),
    reopen: overrides.reopen ?? (async () => {}),
    trace: stage => traces.push(stage),
    error: code => traces.push("error:" + code),
    pixelPolicy: "mouse", editorPid: 4242, geometryTimeoutMs: 500, rejoinTimeoutMs: 2000,
  };
  const surfaceToken = editorModule.newSurfaceToken();
  const editor = new editorModule.FocusEditor(panel, "C:\\docs\\a.docx", surfaceToken, deps);
  return { editor, panel, posted, calls, traces, received, viewState, disposed, surfaceToken, handlers: () => handlers };
}

test("focus tab opens with the page size, connects the stream and reports visibility", async () => {
  const r = rig();
  const started = r.editor.start();
  r.received.fire({ action: "layout", surfaceToken: r.surfaceToken, width: 800, height: 600, dpr: 1.5 });
  await started;
  const open = r.calls.find(call => call[0] === "open")[1];
  assert.deepEqual({ ...open }, { path: "C:\\docs\\a.docx", surfaceToken: r.surfaceToken, editorPid: 4242, pixelPolicy: "mouse", cssWidth: 800, cssHeight: 600, pixelRatio: 1.5 });
  assert.equal(r.calls.find(call => call[0] === "stream")[1].surfaceToken, r.surfaceToken);
  await until(() => r.calls.some(call => call[0] === "layout"));
  assert.equal(r.calls.find(call => call[0] === "layout")[1].visible, true);
  r.received.fire({ action: "layout", surfaceToken: "oie-surface-" + "f".repeat(32), width: 10, height: 10 });
  r.received.fire({ action: "layout", surfaceToken: r.surfaceToken, width: 8, height: 600 });
  assert.equal(r.calls.filter(call => call[0] === "layout").length, 1, "foreign or tiny layouts are ignored");
});

test("focus tab forwards only trusted clicks and only frames of the current park", async () => {
  const r = rig();
  const started = r.editor.start();
  r.received.fire({ action: "layout", surfaceToken: r.surfaceToken, width: 800, height: 600, dpr: 1 });
  await started;
  r.received.fire({ action: "activate", surfaceToken: r.surfaceToken, trusted: false });
  r.received.fire({ action: "activate", surfaceToken: r.surfaceToken, trusted: true, point: { x: 0.25, y: 2 } });
  r.received.fire({ action: "activate", surfaceToken: r.surfaceToken, trusted: true, point: { x: 0.25, y: 0.5 } });
  await until(() => r.calls.filter(call => call[0] === "activate").length === 2);
  const activations = r.calls.filter(call => call[0] === "activate").map(call => call[1]);
  assert.equal(activations[0].point, undefined, "an out-of-range point is dropped, the click still activates");
  assert.deepEqual(activations[1].point, { x: 0.25, y: 0.5 });
  assert.deepEqual(activations.map(item => item.activationId), [1, 2]);

  const h = r.handlers();
  h.event({ type: "state", state: "parked", cause: "foreground", activationId: 2, parkId: 3 });
  h.frame({ frameSeq: 9, parkId: 2, width: 4, height: 4, format: "png", image: new Uint8Array([1]) });
  h.frame({ frameSeq: 10, parkId: 3, width: 4, height: 4, format: "jpeg", image: new Uint8Array([2]) });
  await until(() => r.posted.some(message => message.kind === "frame"));
  const frames = r.posted.filter(message => message.kind === "frame");
  assert.deepEqual(frames.map(frame => frame.frameSeq), [10]);
  assert.equal(frames[0].format, "jpeg", "the page learns the image format from the frame");
  h.event({ type: "notice", code: "modal_blocking" });
  await tick();
  assert.ok(r.posted.some(message => message.kind === "status" && message.sticky === true && message.text === uiModule.FOCUS_TEXT.modal));
  h.event({ type: "closed", reason: "office_exited" });
  await tick();
  assert.equal(r.editor.currentState, "lost");
});

test("wheel over a parked tab is forwarded, merged while busy, and dropped when not parked", async () => {
  let release;
  const r = rig({ wheelGate: new Promise(resolve => { release = resolve; }) });
  const started = r.editor.start();
  r.received.fire({ action: "layout", surfaceToken: r.surfaceToken, width: 800, height: 600, dpr: 1 });
  await started;
  const h = r.handlers();
  r.received.fire({ action: "wheel", surfaceToken: r.surfaceToken, trusted: true, point: { x: 0.5, y: 0.5 }, deltaY: -120, deltaX: 0 });
  assert.equal(r.calls.filter(call => call[0] === "wheel").length, 0, "nothing is forwarded before the tab is parked");
  h.event({ type: "state", state: "parked", cause: "open", activationId: 0, parkId: 1 });
  r.received.fire({ action: "wheel", surfaceToken: r.surfaceToken, trusted: false, point: { x: 0.5, y: 0.5 }, deltaY: -120, deltaX: 0 });
  r.received.fire({ action: "wheel", surfaceToken: r.surfaceToken, trusted: true, point: { x: 0.5, y: 0.5 }, deltaY: -120, deltaX: 0 });
  r.received.fire({ action: "wheel", surfaceToken: r.surfaceToken, trusted: true, point: { x: 0.5, y: 0.6 }, deltaY: -120, deltaX: 0, shift: true });
  r.received.fire({ action: "wheel", surfaceToken: r.surfaceToken, trusted: true, point: { x: 0.5, y: 0.7 }, deltaY: -240, deltaX: 30 });
  release();
  await until(() => r.calls.filter(call => call[0] === "wheel").length === 2);
  const wheels = r.calls.filter(call => call[0] === "wheel").map(call => call[1]);
  assert.equal(wheels[0].deltaY, -120, "a step from before parking is dropped, not saved up");
  assert.deepEqual([wheels[1].deltaY, wheels[1].deltaX, wheels[1].point.y], [-360, 30, 0.7], "steps that arrive while busy are merged");
  assert.ok(wheels.every(step => step.trusted === true));
  h.event({ type: "state", state: "editing", cause: "click", activationId: 1, parkId: 1 });
  r.received.fire({ action: "wheel", surfaceToken: r.surfaceToken, trusted: true, point: { x: 0.5, y: 0.5 }, deltaY: -120, deltaX: 0 });
  await tick(); await tick();
  assert.equal(r.calls.filter(call => call[0] === "wheel").length, 2, "nothing is forwarded while editing");
});

test("closing the last dirty tab asks, then closes with the decision", async () => {
  const r = rig({ closeReplies: [{ closed: false, dirty: true, remainingSurfaces: 1 }, { closed: true, dirty: false, remainingSurfaces: 0 }],
    askSave: async () => "discard" });
  const started = r.editor.start();
  r.received.fire({ action: "layout", surfaceToken: r.surfaceToken, width: 800, height: 600, dpr: 1 });
  await started;
  r.disposed.fire();
  await r.editor.close();
  const closes = r.calls.filter(call => call[0] === "close").map(call => call[1]);
  assert.equal(closes.length, 2);
  assert.equal(closes[0].decision, undefined);
  assert.equal(closes[1].decision, "discard");
});

test("cancelling close reopens the tab and detaches this surface once the new one joined", async () => {
  let reopened = 0;
  const r = rig({ closeReplies: [{ closed: false, dirty: true, remainingSurfaces: 1 }, { closed: false, dirty: true, remainingSurfaces: 1 },
    { closed: true, dirty: false, remainingSurfaces: 1 }], askSave: async () => "cancel", reopen: async () => { reopened++; } });
  const started = r.editor.start();
  r.received.fire({ action: "layout", surfaceToken: r.surfaceToken, width: 800, height: 600, dpr: 1 });
  await started;
  await r.editor.close();
  assert.equal(reopened, 1);
  const closes = r.calls.filter(call => call[0] === "close").map(call => call[1]);
  assert.ok(closes.length >= 3 && closes.every(item => item.decision === undefined));
  assert.ok(r.traces.includes("focus_surface_detached"));
});

test("a tab closed before Core answers still releases its surface", async () => {
  let release;
  const r = rig({ openGate: new Promise(resolve => { release = resolve; }) });
  const started = r.editor.start();
  r.received.fire({ action: "layout", surfaceToken: r.surfaceToken, width: 800, height: 600, dpr: 1 });
  await until(() => r.calls.some(call => call[0] === "open"));
  await r.editor.close();
  assert.ok(!r.calls.some(call => call[0] === "close"), "nothing to close before Core answered");
  release();
  await started;
  assert.ok(r.calls.some(call => call[0] === "close"), "the surface Core just created is closed");
  assert.ok(!r.calls.some(call => call[0] === "stream"));
});

test("focus stream client sends the hello line and decodes what the pipe returns", async () => {
  const pipeName = "office-in-editor-v1-" + randomBytes(8).toString("hex") + "-focus-" + randomBytes(16).toString("hex");
  const token = randomBytes(32).toString("hex");
  const surfaceToken = "oie-surface-" + randomBytes(16).toString("hex");
  let hello = "";
  const image = randomBytes(200_000);
  const server = net.createServer(socket => {
    socket.on("data", chunk => {
      hello += chunk.toString("utf8");
      if (!hello.endsWith("\n")) return;
      const bytes = Buffer.concat([sdk.encodeFocusEvent({ type: "state", state: "parked", cause: "open", activationId: 0, parkId: 1 }),
        sdk.encodeFocusFrame({ frameSeq: 1, parkId: 1, width: 10, height: 10, format: "png" }, image)]);
      // Dribble the bytes so the frame arrives in many chunks.
      let offset = 0;
      const step = () => { if (offset >= bytes.length) { socket.end(); return; } socket.write(bytes.subarray(offset, offset += 7_919)); setImmediate(step); };
      step();
    });
  });
  await new Promise(resolve => server.listen("\\\\.\\pipe\\" + pipeName, resolve));
  try {
    const events = [], frames = [];
    let ended = false;
    const client = new streamModule.FocusStreamClient(pipeName, token, surfaceToken, {
      event: event => events.push(event), frame: frame => frames.push(frame), ended: () => { ended = true; } });
    await client.connect();
    await until(() => ended, 5000);
    assert.deepEqual(JSON.parse(hello), { v: 1, token, surfaceToken });
    assert.equal(events[0].state, "parked");
    assert.equal(frames.length, 1);
    assert.ok(Buffer.from(frames[0].image).equals(image));
    client.dispose();
    assert.throws(() => new streamModule.FocusStreamClient("\\\\.\\pipe\\elsewhere", token, surfaceToken, {}), { code: "INVALID_PARAMS" });
  } finally { server.close(); }
});

test("a failed open tells an outdated add-in apart from a missing window", () => {
  assert.equal(uiModule.openFailureText("office_addin_outdated"), uiModule.FOCUS_TEXT.addinOutdated);
  assert.equal(uiModule.openFailureText("office_window_unavailable"), uiModule.FOCUS_TEXT.windowUnavailable);
  assert.equal(uiModule.openFailureText("request_failed"), uiModule.FOCUS_TEXT.openFailed);
});

test("focus page plants the surface token for UI Automation and keeps a strict CSP", () => {
  const surfaceToken = "oie-surface-" + "c".repeat(32);
  const html = uiModule.focusHtml(surfaceToken);
  assert.ok(html.includes(`aria-label="${surfaceToken}"`));
  for (const corner of ["tl", "tr", "bl", "br"]) assert.ok(html.includes(`aria-label="${surfaceToken}:${corner}"`));
  assert.match(html, /default-src 'none'; style-src 'nonce-[^']+'; script-src 'nonce-[^']+'/u);
  assert.ok(!html.includes("unsafe-"));
  assert.ok(html.includes("event.isTrusted"));
  // Frames are drawn 1:1; the canvas must not be CSS-stretched to the tab, or a resize stretches the old frame.
  assert.ok(!/#frame\{[^}]*width:100%/u.test(html));
  assert.ok(html.includes("context.drawImage(bitmap,0,0);"));
  // Wheel forwarding only for trusted events, and the page must be able to cancel its own scrolling.
  assert.ok(html.includes("root.addEventListener('wheel'") && html.includes("{passive:false}") && /wheel',event=>\{\s*if\(!event\.isTrusted/u.test(html));
  assert.ok(editorModule.FOCUS_EXTENSIONS.has(".xlsx") && !editorModule.FOCUS_EXTENSIONS.has(".vsdx"));
});
