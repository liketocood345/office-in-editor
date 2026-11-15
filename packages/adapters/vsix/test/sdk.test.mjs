import assert from "node:assert/strict";
import { test } from "node:test";
import * as net from "node:net";
import { randomBytes } from "node:crypto";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import { createContext, runInContext } from "node:vm";

const root = fileURLToPath(new URL("../../../../", import.meta.url));
const compiled = await build({
  entryPoints: [fileURLToPath(new URL("../../../adapter-sdk/src/index.ts", import.meta.url))],
  bundle: true, platform: "node", format: "esm", target: "node18", write: false,
  absWorkingDir: root,
});
const sdk = await import("data:text/javascript;base64," + Buffer.from(compiled.outputFiles[0].text).toString("base64"));
const uiBuild = await build({ entryPoints: [fileURLToPath(new URL("../src/ui.ts", import.meta.url))],
  bundle: true, platform: "node", format: "esm", target: "node18", write: false, absWorkingDir: root });
const ui = await import("data:text/javascript;base64," + Buffer.from(uiBuild.outputFiles[0].text).toString("base64"));

test("webview boot reports ready and geometry without global collisions; ACK stops retries", () => {
  const messages = [], timers = [], listeners = {};
  const nodes = Object.fromEntries(["status", "surface", "ole-root"].map(id => [id, { hidden: true, addEventListener() {} }]));
  const context = createContext({
    acquireVsCodeApi: () => ({ postMessage: value => messages.push(value) }),
    document: { getElementById: id => nodes[id], body: {}, title: "" },
    window: { innerWidth: 900, innerHeight: 600, addEventListener: (name, handler) => { listeners[name] = handler; } },
    ResizeObserver: class { observe() {} }, setTimeout: callback => timers.push(callback),
  });
  runInContext("const status = 'existing browser binding';", context);
  const surfaceToken = "oie-surface-" + "a".repeat(32);
  const html = ui.hostHtml("Test", "Loading", surfaceToken);
  runInContext(html.match(/<script nonce="[^"]+">([\s\S]*?)<\/script>/u)[1], context);
  assert.equal(messages[0].action, "ready");
  assert.equal(messages[1].action, "layout");
  assert.equal(messages[1].width, 900);
  assert.equal(messages[1].surfaceToken, surfaceToken);
  assert.equal(messages[1].clientOnly, true);
  assert.equal(html.includes("ole-root"), true);
  assert.equal(html.includes("oie-beacon"), true);
  assert.equal(html.includes("width:1px"), true);
  assert.equal(html.includes("background:#000"), true);
  assert.equal(html.includes(surfaceToken + ":tl"), true);
  assert.equal(html.includes("oie-surface-label"), false);
  assert.equal(html.includes("open-external"), false);
  assert.equal(timers.length, 4);
  listeners.message({ data: { kind: "layout-ack" } });
  for (const callback of timers) callback();
  assert.equal(messages.length, 2);
  listeners.message({ data: { kind: "layout-request" } });
  assert.equal(messages.length, 3);
  listeners.message({ data: { kind: "attached" } });
  assert.equal(nodes.status.hidden, true);
  assert.equal(html.includes("open-external"), false);
});

test("editor document open uses an explicit bounded pipe budget", async () => {
  assert.equal(sdk.OPEN_DOCUMENT_TIMEOUT_MS, 60_000);
  assert.equal(sdk.MAX_PIPE_TIMEOUT_MS, 60_000);
  let call;
  const broker = new sdk.OfficeBroker({ request: async (...args) => { call = args; return {}; } });
  await broker.openInOffice("fixture.docx", "embedded", { fallback: false });
  assert.equal(call[0], "office.open");
  assert.equal(call[2].timeoutMs, 60_000);
});

test("endpoint accepts only the versioned user pipe namespace", () => {
  assert.deepEqual(sdk.parseEndpoint(JSON.stringify({
    pipeName: "office-in-editor-v1-" + "a".repeat(16), protocolVersion: 1, version: "0.1.0",
  })), { pipeName: "office-in-editor-v1-" + "a".repeat(16), protocolVersion: 1, version: "0.1.0" });
  for (const value of [
    { pipeName: "arbitrary-pipe", protocolVersion: 1, version: "0.1.0" },
    { pipeName: "office-in-editor-v1-" + "a".repeat(16), protocolVersion: 2, version: "0.1.0" },
  ]) assert.throws(() => sdk.parseEndpoint(JSON.stringify(value)), { code: "INVALID_ENDPOINT" });
});

test("response correlation failures remain unknown rather than successful writes", () => {
  assert.throws(
    () => sdk.parseResponse(JSON.stringify({ id: "another-operation", ok: true, result: {} }), "expected"),
    error => error.code === "INVALID_RESPONSE" && error.outcomeUnknown,
  );
});

test("worker error messages do not cross the diagnostic boundary", () => {
  const privateText = "PRIVATE DOCUMENT BODY AND LOCAL PATH";
  assert.throws(
    () => sdk.parseResponse(JSON.stringify({
      id: "request", ok: false, error: { code: "REFERENCE_NOT_FOUND", message: privateText },
    }), "request"),
    error => error.code === "REFERENCE_NOT_FOUND" && !String(error.stack).includes(privateText),
  );
});

test("outbound frames enforce the wire byte budget", () => {
  const frame = JSON.parse(sdk.encodeRequest("request", "reference.get", { refId: "example" }));
  assert.equal(frame.v, 1);
  assert.equal(frame.method, "reference.get");
  assert.throws(
    () => sdk.encodeRequest("request", "reference.get", { refId: "字".repeat(sdk.MAX_LINE_BYTES) }),
    { code: "REQUEST_TOO_LARGE" },
  );
});

test("explicit exports exclude private paths and unrelated runtime fields", () => {
  const reference = {
    refId: "test-ref", application: "word", documentName: "Document", kind: "word-range",
    createdAt: "2026-01-01T00:00:00Z", text: "selected text", locator: { start: 5, end: 18 },
    sha256: "abc", dirty: true, sourcePath: "PRIVATE_PATH", token: "PRIVATE_TOKEN", ownerId: "PRIVATE_OWNER",
  };
  const json = sdk.referenceAsJson(reference);
  assert.equal(JSON.parse(json).text, "selected text");
  assert.ok(!json.includes("PRIVATE_"));
  assert.ok(!sdk.referenceAsMarkdown(reference).includes("PRIVATE_"));
});

test("untrusted document labels are HTML text, never markup", () => {
  const value = sdk.escapeHtml('<img src=x onerror="alert(1)">&\'');
  assert.ok(!value.includes("<img"));
  assert.ok(value.includes("&lt;img"));
  assert.ok(value.includes("&quot;"));
  assert.ok(value.includes("&#39;"));
});

async function mockPipe(handler) {
  const name = "office-in-editor-v1-" + randomBytes(8).toString("hex");
  const server = net.createServer(socket => {
    socket.on("error", () => {});
    let input = "";
    socket.on("data", chunk => {
      input += chunk.toString("utf8");
      const newline = input.indexOf("\n");
      if (newline >= 0) handler(socket, JSON.parse(input.slice(0, newline)));
    });
  });
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen("\\\\.\\pipe\\" + name, resolve);
  });
  return { name, close: () => new Promise(resolve => server.close(resolve)) };
}

test("named pipe accepts split response frames without starting a real worker", { skip: process.platform !== "win32" }, async () => {
  const pipe = await mockPipe((socket, request) => {
    assert.equal(request.method, "health");
    const response = JSON.stringify({ id: request.id, ok: true, result: { connected: true } }) + "\n";
    socket.write(response.slice(0, 13));
    socket.end(response.slice(13));
  });
  try { assert.deepEqual(await sdk.requestPipe(pipe.name, "health", {}), { connected: true }); }
  finally { await pipe.close(); }
});

test("a lost pipe after request admission never reports a safe retry", { skip: process.platform !== "win32" }, async () => {
  const pipe = await mockPipe(socket => socket.end());
  try {
    await assert.rejects(
      sdk.requestPipe(pipe.name, "reference.deliver", {}),
      error => error.outcomeUnknown === true,
    );
  } finally { await pipe.close(); }
});

const compiledNotices = await build({
  entryPoints: [fileURLToPath(new URL("../src/notices.ts", import.meta.url))],
  bundle: true, platform: "node", format: "esm", target: "node18", write: false,
});
const { ErrorNoticeBudget } = await import("data:text/javascript;base64,"
  + Buffer.from(compiledNotices.outputFiles[0].text).toString("base64"));

test("background failures never consume a toast and repeated codes log once", () => {
  const budget = new ErrorNoticeBudget();
  assert.deepEqual(budget.observe("BROKER_UNAVAILABLE", false), { log: true, notify: false });
  assert.deepEqual(budget.observe("BROKER_UNAVAILABLE", false), { log: false, notify: false });
  assert.deepEqual(budget.observe("BROKER_UNAVAILABLE", true), { log: false, notify: true });
  assert.deepEqual(budget.observe("BROKER_UNAVAILABLE", true), { log: false, notify: false });
  assert.deepEqual(budget.observe("OTHER_ERROR", true), { log: true, notify: false });
});

test("failure logs have a finite per-session code budget", () => {
  const budget = new ErrorNoticeBudget();
  let logged = 0;
  let notices = 0;
  for (let index = 0; index < 100; index += 1) {
    const result = budget.observe("ERROR_" + index, true);
    if (result.log) logged += 1;
    if (result.notify) notices += 1;
  }
  assert.equal(logged, 32);
  assert.equal(notices, 1);
});
