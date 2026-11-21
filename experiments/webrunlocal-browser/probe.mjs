import { spawn } from 'node:child_process';
import { randomBytes, randomInt } from 'node:crypto';
import { createServer } from 'node:http';
import { createConnection } from 'node:net';
import { setTimeout as delay } from 'node:timers/promises';

const [edge, profile] = process.argv.slice(2);
if (!edge || !profile) throw new Error('Expected isolated Edge executable and profile');
const endpoints = [83, 84]; // PluginOK's documented default and fallback.
const phases = ['websocket_error', 'websocket_timeout', 'version_timeout', 'version_rejected', 'version_received'];

function browserProbe(ports, sid) {
  const results = [];
  function check(port) {
    return new Promise(resolve => {
      const socket = new WebSocket(`ws://127.0.0.1:${port}?sid=${sid}&flag=0`);
      let opened = false;
      let finished = false;
      const finish = (phase, details = {}) => {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        socket.close();
        resolve({ port, opened, phase, ...details });
      };
      const timer = setTimeout(() => finish(opened ? 'version_timeout' : 'websocket_timeout'), 2200);
      socket.addEventListener('open', () => {
        opened = true;
        socket.send(JSON.stringify({ req: 'Wrl_Version', rid: 1, para: { Mac: 0, More: 0, Config: 0, Router: 0, IP: 0 } }));
      }, { once: true });
      socket.addEventListener('error', () => finish('websocket_error'), { once: true });
      socket.addEventListener('close', () => finish(opened ? 'version_rejected' : 'websocket_error'), { once: true });
      socket.addEventListener('message', ({ data }) => {
        if (typeof data !== 'string' || data.length > 16_384) return finish('version_rejected');
        try {
          const reply = JSON.parse(data);
          if (reply.req !== 'Wrl_Version' || reply.rid !== 1) return;
          if (reply.ret !== 0 || !reply.data) return finish('version_rejected');
          finish('version_received', {
            version: /^\d+(?:\.\d+){1,3}$/.test(reply.data.Version ?? '') ? reply.data.Version : undefined,
            validDay: /^\d{4}-\d{2}-\d{2}$/.test(reply.data.ValidDay ?? '') ? reply.data.ValidDay : undefined,
            advanced: reply.data.Adv === 1 || reply.data.Adv === true
          });
        } catch { finish('version_rejected'); }
      });
    });
  }
  async function run() {
    for (const port of ports) {
      results.push(await check(port));
      if (results.at(-1).phase === 'version_received') break;
    }
    const edgeVersion = /\bEdg\/(\d+(?:\.\d+){1,3})\b/.exec(navigator.userAgent)?.[1];
    const status = document.getElementById('status');
    status.textContent = results.some(row => row.phase === 'version_received') ? '已收到 PluginOK 版本回复' : 'PluginOK 握手未完成';
    await fetch('/report', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ edgeVersion, visible: !document.hidden, results }) });
  }
  void run();
}

function checkListener(port) {
  return new Promise(resolve => {
    const socket = createConnection({ host: '127.0.0.1', port });
    const finish = reachable => { socket.destroy(); resolve({ port, reachable }); };
    socket.once('connect', () => finish(true));
    socket.once('error', () => finish(false));
    socket.setTimeout(600, () => finish(false));
  });
}

const events = [];
const begin = performance.now();
const event = name => { if (events.length < 16) events.push({ name, elapsedMs: Math.round(performance.now() - begin) }); };
const listeners = await Promise.all(endpoints.map(checkListener));
event('runtime_ports_checked');
const nonce = randomBytes(18).toString('base64');
const html = `<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>PluginOK localhost 检查</title>
    <main><h1>PluginOK localhost 检查</h1><p id="status" role="status">正在连接本机承载器</p></main>
    <script nonce="${nonce}">(${browserProbe.toString()})(${JSON.stringify(endpoints)},${randomInt(1, 1_000_000_000)});</script></html>`;
let authority, received;
const report = new Promise(resolve => { received = resolve; });
let accepted = false;
const server = createServer({ requestTimeout: 3000, headersTimeout: 3000 }, async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  if (req.headers.host !== authority) { res.writeHead(403).end(); return; }
  if (req.method === 'GET' && req.url === '/') {
    event('localhost_page_served');
    const wsSources = endpoints.map(port => `ws://127.0.0.1:${port}`).join(' ');
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8',
      'Content-Security-Policy': `default-src 'none'; script-src 'nonce-${nonce}'; connect-src 'self' ${wsSources}; base-uri 'none'; frame-ancestors 'none'` }).end(html);
    return;
  }
  if (req.method !== 'POST' || req.url !== '/report' || req.headers.origin !== 'http://' + authority || accepted) {
    res.writeHead(404).end(); return;
  }
  try {
    const chunks = [];
    let length = 0;
    for await (const chunk of req) {
      length += chunk.length;
      if (length > 2048) throw new Error('report_too_large');
      chunks.push(chunk);
    }
    const raw = JSON.parse(Buffer.concat(chunks));
    if (accepted || !/^\d+(?:\.\d+){1,3}$/.test(raw.edgeVersion ?? '') || typeof raw.visible !== 'boolean' ||
        !Array.isArray(raw.results) || !raw.results.length || raw.results.length > endpoints.length) throw new Error('invalid_report');
    const results = raw.results.map((row, index) => {
      if (row.port !== endpoints[index] || typeof row.opened !== 'boolean' || !phases.includes(row.phase)) throw new Error('invalid_probe_row');
      const { port, opened, phase } = row;
      return { port, opened, phase,
        version: typeof row.version === 'string' && /^\d+(?:\.\d+){1,3}$/.test(row.version) ? row.version : undefined,
        validDay: typeof row.validDay === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(row.validDay) ? row.validDay : undefined,
        advanced: typeof row.advanced === 'boolean' ? row.advanced : undefined };
    });
    accepted = true;
    event('browser_report_received');
    res.writeHead(204).end();
    received({ edgeVersion: raw.edgeVersion, visible: raw.visible, results });
  } catch { if (!res.destroyed) res.writeHead(400).end(); }
});
server.maxConnections = 8;
await new Promise((resolve, reject) => {
  server.once('error', reject);
  server.listen(0, '127.0.0.1', resolve);
});
authority = 'localhost:' + server.address().port;
let browser, result, outcome;
try {
  browser = spawn(edge, ['--user-data-dir=' + profile, '--no-first-run', '--no-default-browser-check', '--new-window', 'http://' + authority + '/'], { stdio: 'ignore' });
  const started = new Promise(resolve => {
    browser.once('spawn', () => resolve(true));
    browser.once('error', () => resolve(false));
  });
  if (await started) {
    event('ordinary_edge_started');
    result = await Promise.race([report, delay(12_000, undefined, { ref: false })]);
    outcome = !result ? 'browser_report_missing' : result.results.some(row => row.phase === 'version_received') ? 'version_received' : 'pluginok_handshake_unavailable';
  } else outcome = 'edge_start_failed';
} finally {
  const closed = new Promise(resolve => server.close(resolve));
  server.closeAllConnections();
  await closed;
  event('localhost_listener_closed');
  browser?.unref();
}
console.log(JSON.stringify({ outcome, listeners, browser: result, browserPid: browser?.pid, port: Number(authority.split(':')[1]), events,
  officeEditorChecks: { word: 'not_run', excel: 'not_run', powerpoint: 'not_run' } }));
