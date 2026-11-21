import { createServer } from 'node:http';
import { createHash, randomBytes, timingSafeEqual } from 'node:crypto';
import { spawn } from 'node:child_process';
import { readFile, writeFile } from 'node:fs/promises';
import { createInterface } from 'node:readline';
import path from 'node:path';
import { DocumentPorts } from './document-ports.mjs';

const roots = process.argv.slice(2).map(root => path.resolve(root));
if (roots.length < 1 || roots.length > 2) throw new Error('Expected one or two fixture directories');
const assetRoot = roots[0];
const files = new Map();
for (const [route, file, type] of [
  ['/', 'index.html', 'text/html'],
  ['/app.js', 'app.js', 'text/javascript'],
  ['/style.css', 'style.css', 'text/css'],
]) {
  files.set(route, { body: await readFile(path.join(assetRoot, file)), type });
}

async function startDocument(document) {
  const root = path.dirname(document);
  const owner = JSON.parse(await readFile(path.join(root, 'ownership.json'), 'utf8'));
  if (owner.project !== 'office-in-editor' || owner.purpose !== 'word-browser-local' ||
      !/^[a-f0-9]{32}$/.test(owner.run) || path.basename(root).toLowerCase() !== 'word-web-' + owner.run) {
    throw new Error('Invalid experiment directory');
  }

  const expectedSecret = randomBytes(32);
  const worker = spawn(path.join(assetRoot, 'WordWorker.exe'), [root], {
    cwd: root, windowsHide: true, stdio: ['pipe', 'pipe', 'ignore'],
  });
  const lines = createInterface({ input: worker.stdout });
  let pending;
  let initial;
  let lastSave;
  let workerFailed = false;
  let closing = false;
  let proofWritten = false;
  let origin;

  function awaitWorker(timeoutMs) {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        pending = undefined;
        workerFailed = true;
        reject(new Error('word_timeout'));
      }, timeoutMs);
      pending = value => {
        clearTimeout(timeout);
        pending = undefined;
        resolve(value);
      };
    });
  }

  lines.on('line', line => {
    if (!pending) return;
    try { pending(JSON.parse(line)); }
    catch { workerFailed = true; pending({ ok: false, error: 'worker_protocol' }); }
  });
  function unavailable() {
    workerFailed = true;
    pending?.({ ok: false, error: 'word_unavailable' });
  }
  worker.on('error', unavailable);
  worker.stdin.on('error', unavailable);
  worker.on('exit', unavailable);

  try {
    initial = await awaitWorker(25_000);
    if (!initial?.ok || initial.engine !== 'Microsoft Word' || initial.wordWindowHidden !== true)
      throw new Error('word_unavailable');
  } catch (error) {
    await writeFile(path.join(root, 'web-error.json'), JSON.stringify({ reason: error.message }), { flag: 'wx' });
    worker.stdin.end(); // No termination. The worker owns normal cleanup or retains its document.
    throw error;
  }

  async function command(message) {
    if (workerFailed || closing || pending) return { ok: false, error: 'word_busy' };
    const result = awaitWorker(15_000);
    worker.stdin.write(JSON.stringify(message) + '\n');
    return result;
  }

  function reply(res, code, value, type = 'application/json') {
    res.writeHead(code, {
      'Content-Type': type + '; charset=utf-8',
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
      'Referrer-Policy': 'no-referrer',
      'Content-Security-Policy': "default-src 'none'; script-src 'self'; style-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'",
    });
    res.end(type === 'application/json' ? JSON.stringify(value) : value);
  }

  function authorized(req) {
    if (req.headers['sec-fetch-site'] === 'cross-site') return false;
    if (req.headers.origin && req.headers.origin !== origin) return false;
    const match = /^Bearer ([0-9a-f]{64})$/.exec(req.headers.authorization ?? '');
    return !!match && timingSafeEqual(Buffer.from(match[1], 'hex'), expectedSecret);
  }

  async function readBody(req) {
    if (req.headers['content-type'] !== 'application/json') throw new Error('invalid_request');
    const chunks = [];
    let size = 0;
    for await (const chunk of req) {
      size += chunk.length;
      if (size > 24_576) throw new Error('request_too_large');
      chunks.push(chunk);
    }
    const value = JSON.parse(Buffer.concat(chunks).toString('utf8'));
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('invalid_request');
    return value;
  }

  async function endCleanly() {
    if (closing) return { ok: true };
    const value = await command({ op: 'close' });
    if (value.ok) {
      closing = true;
      worker.stdin.end();
      server.close();
    }
    return value;
  }

  const server = createServer({ requestTimeout: 5000, headersTimeout: 4500, maxHeaderSize: 8192 }, async (req, res) => {
    if (!origin || req.headers.host !== new URL(origin).host) return reply(res, 403, { error: 'invalid_host' });
    const file = files.get(req.url);
    if (file && req.method === 'GET') return reply(res, 200, file.body, file.type);
    if (!authorized(req)) return reply(res, 403, { error: 'unauthorized' });
    try {
      if (req.method === 'GET' && req.url === '/api/document') {
        const value = await command({ op: 'read' });
        return reply(res, value.ok ? 200 : 503, value);
      }
      if (req.method !== 'POST' || req.headers.origin !== origin) return reply(res, 405, { error: 'invalid_request' });
      const body = await readBody(req);
      if (req.url === '/api/save') {
        if (typeof body.text !== 'string' || body.text.length > 4096 || !Number.isSafeInteger(body.revision)) {
          return reply(res, 400, { error: 'invalid_request' });
        }
        const value = await command({ op: 'save', text: body.text, revision: body.revision });
        if (value.ok && value.wordSaved) lastSave = value;
        return reply(res, value.ok ? 200 : 409, value);
      }
      const view = body.vscode === true ? 'VS Code' : body.edge === true ? 'Edge' : undefined;
      const proofMarker = view === 'VS Code' ? ' · 已通过 VS Code 标签页保存' : ' · 已通过普通 Edge 标签页保存';
      if (req.url === '/api/proof' && !proofWritten && view && lastSave?.revision === 1 && lastSave.text === initial.text + proofMarker &&
          body.revision === lastSave.revision) {
        proofWritten = true;
        await writeFile(path.join(root, 'browser-proof.json'), JSON.stringify({
          browser: view, normalTab: view === 'Edge', webview: view === 'VS Code', editedViaWebForm: true, wordSaved: true, revision: lastSave.revision,
          wordWindowHidden: true,
          textSha256: createHash('sha256').update(lastSave.text, 'utf8').digest('hex'),
        }), { flag: 'wx' });
        return reply(res, 200, { ok: true });
      }
      if (req.url === '/api/proof-failed' && !proofWritten && view &&
          typeof body.error === 'string' && /^[a-z_]{1,40}$/.test(body.error)) {
        proofWritten = true;
        await writeFile(path.join(root, 'browser-proof.json'), JSON.stringify({ browser: view, error: body.error }), { flag: 'wx' });
        return reply(res, 200, { ok: true });
      }
      if (req.url === '/api/close') {
        const value = await endCleanly();
        return reply(res, value.ok ? 200 : 409, value);
      }
      return reply(res, 404, { error: 'not_found' });
    } catch {
      if (!res.headersSent && !res.destroyed) reply(res, 503, { error: workerFailed ? 'save_unconfirmed' : 'request_failed' });
    }
  });
  server.maxConnections = 8;
  server.keepAliveTimeout = 1500;
  server.on('error', () => worker.stdin.end());
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      server.removeListener('error', reject);
      resolve();
    });
  });
  origin = 'http://localhost:' + server.address().port;
  const closed = new Promise(resolve => server.once('close', resolve));
  const url = origin + '/#key=' + expectedSecret.toString('hex');
  try {
    await writeFile(path.join(root, 'web-ready.json'), JSON.stringify({
      url, run: owner.run, serverPid: process.pid,
    }), { flag: 'wx' });
  } catch (error) {
    await endCleanly();
    throw error;
  }
  return { port: server.address().port, closed, close: endCleanly };
}

const ports = new DocumentPorts(2);
const fixtures = roots.map(root => path.join(root, 'synthetic.docx'));
const open = file => ports.acquire(file, startDocument);
const [first, duplicate] = await Promise.all([open(fixtures[0]), open(fixtures[0])]);
const sessions = [first];
let secondaryFailed = false;
for (const fixture of fixtures.slice(1)) {
  try { sessions.push(await open(fixture)); }
  catch { secondaryFailed = true; } // Keep the first document's listener alive.
}
const values = sessions.map(session => session.port);
await writeFile(path.join(assetRoot, 'port-verification.json'), JSON.stringify({
  activeFiles: sessions.length, ports: values, sameFileReused: first === duplicate,
  distinctPorts: new Set(values).size === sessions.length, secondaryFailed,
}), { flag: 'wx' });

if (process.env.OIE_WORD_WEB_SUPERVISED === '1') {
  // Keep the backend available until its IDE supervisor requests shutdown,
  // even when every document listener has already been released.
  const control = createInterface({ input: process.stdin });
  let stopping = false;
  control.on('line', async line => {
    if (line !== 'shutdown' || stopping) return;
    stopping = true;
    const outcomes = await Promise.allSettled(sessions.map(session => session.close()));
    const closed = outcomes.every(outcome => outcome.status === 'fulfilled' && outcome.value.ok);
    await writeFile(path.join(assetRoot, 'backend-shutdown.json'), JSON.stringify({ closed }), { flag: 'wx' });
    if (closed) control.close();
  });
}
