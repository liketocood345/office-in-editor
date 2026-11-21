const vscode = require('vscode');
const { randomBytes } = require('node:crypto');
const { readFile, rename, writeFile } = require('node:fs/promises');
const { createServer } = require('node:http');
const path = require('node:path');
const { setTimeout: delay } = require('node:timers/promises');

const browserCommand = 'workbench.action.browser.open';
const formats = [
  ['word', 'Word', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
  ['excel', 'Excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
  ['powerpoint', 'PowerPoint', 'application/vnd.openxmlformats-officedocument.presentationml.presentation']
];
const signals = ['activeX', 'nativeControl', 'nativeControlReadFailed', 'mimeHandler', 'fallbackVisible', 'fallbackHit', 'visible',
  'objectLoadFailed', 'objectCspBlocked', 'documentComplete', 'reload'];

async function readReport(request) {
  const chunks = [];
  let bytes = 0;
  for await (const chunk of request) {
    bytes += chunk.length;
    if (bytes > 1024) throw new Error('report_too_large');
    chunks.push(chunk);
  }
  const raw = JSON.parse(Buffer.concat(chunks));
  if (!raw || signals.some(name => typeof raw[name] !== 'boolean')) throw new Error('invalid_signals');
  return Object.fromEntries(signals.map(name => [name, raw[name]]));
}

async function listen({ title, mime, clsid, registered }, log) {
  const reports = [];
  const nonce = randomBytes(18).toString('base64');
  const home = `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Office probe · ${title}</title>
      <script nonce="${nonce}" src="/page.js"></script></head>
      <body><h1>${title} browser probe</h1><object id="office" type="${mime}" classid="clsid:${clsid}" width="420" height="120">
      <p id="fallback">The browser has not activated the Office document object.</p></object>
      </body></html>`;
  const client = await readFile(path.join(__dirname, 'page.js'));
  let origin, authority, finished;
  const completed = new Promise(resolve => { finished = resolve; });
  const server = createServer({ requestTimeout: 3000, headersTimeout: 3000 }, async (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    if (req.headers.host !== authority) { res.writeHead(403).end(); return; }
    if (req.method === 'GET' && req.url === '/') {
      log('html_served');
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8',
        'Content-Security-Policy': `default-src 'none'; object-src 'self'; script-src 'nonce-${nonce}'; connect-src 'self'; base-uri 'none'` }).end(home);
      return;
    }
    if (req.method === 'GET' && req.url === '/page.js') {
      log('script_served');
      res.writeHead(200, { 'Content-Type': 'application/javascript' }).end(client);
      return;
    }
    if (req.method !== 'POST' || req.url !== '/report') { res.writeHead(404).end(); return; }
    if (req.headers.origin !== origin || reports.length >= 2) { res.writeHead(403).end(); return; }
    try {
      const report = await readReport(req);
      if (reports.length >= 2 || report.reload !== (reports.length === 1)) {
        res.writeHead(409).end(); return;
      }
      reports.push(report);
      log(report.reload ? 'reload_reported' : 'page_reported', report);
      res.writeHead(204).end();
      if (reports.length === 2) finished();
    } catch {
      log('report_rejected');
      if (!res.destroyed) res.writeHead(400).end();
    }
  });
  server.maxConnections = 8;
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const { port } = server.address();
  authority = '127.0.0.1:' + port;
  origin = 'http://' + authority;
  log('listener_ready', { port, registered });
  return { port, origin, reports, completed,
    async close() {
      const done = new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
      server.closeAllConnections();
      await done;
    } };
}

function classify(registered, reports) {
  if (!registered) return 'office_class_not_registered';
  if (reports.length === 0) return 'page_report_missing';
  if (reports.length === 1) return 'reload_report_missing';
  if (reports.some(page => page.objectCspBlocked)) return 'object_blocked_by_page_policy';
  if (reports.some(page => page.activeX || page.nativeControl)) return 'activation_signal_requires_edit_verification';
  if (reports.some(page => page.nativeControlReadFailed || !page.documentComplete)) return 'object_state_inconclusive';
  if (reports.every(page => !page.mimeHandler && page.fallbackVisible && page.fallbackHit)) return 'browser_did_not_activate_office';
  return 'no_activation_signal_fallback_unconfirmed';
}

async function activate(context) {
  const root = process.env.OIE_BROWSER_VSCODE_LAB;
  const base = path.join(process.env.LOCALAPPDATA ?? '', 'OfficeInEditorLab');
  if (!root || path.dirname(root).toLowerCase() !== base.toLowerCase() ||
      !/^office-browser-[0-9a-f]{32}$/.test(path.basename(root)) || process.platform !== 'win32') return;
  const inputs = JSON.parse(await readFile(path.join(root, 'inputs.json'), 'utf8'));
  if (path.basename(root) !== 'office-browser-' + inputs.run) return;
  const output = vscode.window.createOutputChannel('Office Browser Probe');
  context.subscriptions.push(output);
  const events = [];
  const started = performance.now();
  function log(stage, data = {}) {
    if (events.length >= 96) return;
    const event = { stage, elapsedMs: Math.round(performance.now() - started), ...data };
    events.push(event);
    output.appendLine(JSON.stringify(event));
  }
  async function snapshot(stage, data) {
    // One small file per transition, never per event or browser request.
    const file = path.join(root, `${stage}.json`);
    const temporary = file + '.partial';
    await writeFile(temporary, JSON.stringify({ stage, run: inputs.run, vscodeVersion: vscode.version, ...data, events }), { flag: 'wx' });
    await rename(temporary, file);
  }
  const sessions = [];
  const results = [];
  let error;
  log('extension_activated');
  await snapshot('activated', {});
  try {
    if (!(await vscode.commands.getCommands()).includes(browserCommand)) throw new Error('integrated_browser_unavailable');
    for (const [application, title, mime] of formats) {
      const registered = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(inputs.classes[application] ?? '');
      const clsid = registered ? inputs.classes[application] : '00000000-0000-0000-0000-000000000000';
      const record = (stage, data) => log(stage, { application, ...data });
      const session = await listen({ title, mime, clsid, registered }, record);
      sessions.push(session);
      record('browser_open_requested');
      await vscode.commands.executeCommand(browserCommand, session.origin + '/');
      record('browser_open_returned');
      await Promise.race([session.completed, delay(6500, undefined, { ref: false })]);
      const outcome = classify(registered, session.reports);
      results.push({ application, port: session.port, registered, outcome, reports: session.reports });
      record(outcome);
      await snapshot(application, { results });
    }
  } catch (failure) {
    error = failure.message === 'integrated_browser_unavailable' ? failure.message : 'probe_failed';
    log(error);
  } finally {
    const tabs = vscode.window.tabGroups.all.flatMap(group => group.tabs);
    const probeTitles = formats.map(([, title]) => `Office probe · ${title}`);
    const matched = tabs.filter(tab => probeTitles.includes(tab.label));
    const tabCount = matched.length;
    const closed = await Promise.allSettled(sessions.map(session => session.close()));
    const listenersClosed = closed.every(result => result.status === 'fulfilled');
    log('listeners_closed', { listenersClosed });
    await snapshot('result', { error, results, distinctPorts: new Set(results.map(result => result.port)).size === formats.length,
      integratedBrowserTabs: tabCount, listenersClosed, editSaveVerified: false });
    // This command is scoped to the isolated window started by the lab.
    await vscode.commands.executeCommand('workbench.action.closeWindow');
  }
}

exports.activate = activate;
