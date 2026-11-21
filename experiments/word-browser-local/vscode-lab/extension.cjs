const vscode = require('vscode');
const { randomBytes } = require('node:crypto');
const { watch } = require('node:fs');
const { lstat, readFile, writeFile } = require('node:fs/promises');
const http = require('node:http');
const path = require('node:path');

const viewType = 'oieLocal.wordWebviewLab';
const limit = 32_768;
const readJson = async file => JSON.parse(await readFile(file, 'utf8'));
const key = file => path.normalize(file).toLowerCase();

function waitForJson(file) {
  return new Promise((resolve, reject) => {
    let done = false;
    const finish = (error, value) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      watcher.close();
      if (error) reject(error);
      else resolve(value);
    };
    const read = () => void readJson(file).then(value => finish(undefined, value), () => {});
    const watcher = watch(path.dirname(file), (_kind, name) => {
      if (!name || name.toString() === path.basename(file)) read();
    });
    watcher.on('error', error => finish(error));
    const timer = setTimeout(() => finish(new Error('backend_start_timeout')), 28_000);
    read();
  });
}

function wordRequest(endpoint, route, body) {
  const payload = body === undefined ? undefined : Buffer.from(JSON.stringify(body));
  if (payload && payload.length > 24_576) return Promise.reject(new Error('invalid_request'));
  const headers = { Host: endpoint.host, Authorization: 'Bearer ' + endpoint.token };
  if (payload) Object.assign(headers, { Origin: endpoint.origin, 'Content-Type': 'application/json', 'Content-Length': payload.length });
  return new Promise((resolve, reject) => {
    const request = http.request({ hostname: '127.0.0.1', port: endpoint.port, path: '/api/' + route,
      method: payload ? 'POST' : 'GET', headers, agent: false, signal: AbortSignal.timeout(16_000) }, response => {
      const chunks = [];
      let size = 0;
      response.on('data', chunk => {
        size += chunk.length;
        if (size > limit) request.destroy(new Error('invalid_response'));
        else chunks.push(chunk);
      });
      response.on('error', reject);
      response.on('end', () => {
        try {
          const value = JSON.parse(Buffer.concat(chunks).toString('utf8'));
          if (response.statusCode >= 300 && value?.ok) throw new Error('invalid_response');
          resolve(value);
        } catch (error) { reject(error); }
      });
    });
    request.on('error', reject);
    request.end(payload);
  });
}

async function allowedFixtures() {
  const roots = JSON.parse(process.env.OIE_VSCODE_WEB_LAB_ROOTS ?? '[]');
  const base = path.join(process.env.LOCALAPPDATA ?? '', 'OfficeInEditorLab');
  if (roots.length !== 2 || new Set(roots.map(key)).size !== 2) throw new Error('invalid_fixtures');
  const fixtures = [];
  for (const root of roots) {
    const run = /^word-web-([a-f0-9]{32})$/.exec(path.basename(root))?.[1];
    if (!run || key(path.dirname(root)) !== key(base) || (await lstat(root)).isSymbolicLink()) throw new Error('invalid_fixture_root');
    const owner = await readJson(path.join(root, 'ownership.json'));
    if (owner.run !== run || owner.project !== 'office-in-editor' || owner.purpose !== 'word-browser-local')
      throw new Error('invalid_fixture_owner');
    const file = path.join(root, 'synthetic.docx');
    if (!(await lstat(file)).isFile()) throw new Error('invalid_fixture');
    fixtures.push({ root, run, file });
  }
  return fixtures;
}

async function activate(context) {
  if (process.platform !== 'win32' || vscode.env.remoteName || !vscode.workspace.isTrusted) return;
  const fixtures = await allowedFixtures();
  const registered = new Map(fixtures.map(fixture => [key(fixture.file), fixture]));
  const first = fixtures[0];
  const asset = vscode.Uri.joinPath(context.extensionUri, 'media');
  const html = await readFile(vscode.Uri.joinPath(asset, 'index.html').fsPath, 'utf8');
  const requestFile = path.join(first.root, 'vscode-start.json');
  const evidence = path.join(first.root, 'vscode-verification.json');
  let started;
  let firstReloaded = false;
  let initialRead;
  let reloadRead;
  let secondaryRead;
  let proof = false;
  let finished = false;
  let startWasAbsent = false;

  async function start() {
    try {
      await writeFile(requestFile, JSON.stringify({ run: first.run, reason: 'first_office_open' }), { flag: 'wx' });
    } catch (error) {
      if (error.code !== 'EEXIST') throw error;
      const existing = await readJson(requestFile);
      if (existing.run !== first.run || existing.reason !== 'first_office_open') throw new Error('invalid_start_request');
    }
    const ports = await waitForJson(path.join(first.root, 'port-verification.json'));
    if (!ports.distinctPorts || !ports.sameFileReused || ports.activeFiles !== 2) throw new Error('invalid_port_allocation');
    for (const fixture of fixtures) {
      const ready = await readJson(path.join(fixture.root, 'web-ready.json'));
      const parsed = new URL(ready.url);
      const secret = /^#key=([a-f0-9]{64})$/.exec(parsed.hash)?.[1];
      if (ready.run !== fixture.run || parsed.protocol !== 'http:' || parsed.hostname !== 'localhost' ||
          !parsed.port || parsed.username || parsed.password || parsed.pathname !== '/' || parsed.search || !secret)
        throw new Error('invalid_endpoint');
      fixture.endpoint = { origin: parsed.origin, port: Number(parsed.port), host: parsed.host, token: secret };
    }
    if (fixtures[0].endpoint.port === fixtures[1].endpoint.port) throw new Error('port_conflict');
  }

  function render(panel, smoke) {
    const nonce = randomBytes(16).toString('hex');
    const webview = panel.webview;
    const csp = "default-src 'none'; script-src 'nonce-" + nonce + "'; style-src " + webview.cspSource + "; form-action 'none'; base-uri 'none'";
    const css = webview.asWebviewUri(vscode.Uri.joinPath(asset, 'style.css')).toString();
    const js = webview.asWebviewUri(vscode.Uri.joinPath(asset, 'app.js')).toString();
    webview.html = html
      .replace('<html lang="zh-CN">', '<html lang="zh-CN"' + (smoke ? ' data-smoke="vscode"' : '') + '>')
      .replace('<meta charset="utf-8">', '<meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="' + csp + '">')
      .replace('href="/style.css"', 'href="' + css + '"')
      .replace('defer src="/app.js"', 'defer nonce="' + nonce + '" src="' + js + '"');
  }

  async function verify() {
    if (finished || !proof || !initialRead || !reloadRead || !secondaryRead) return;
    finished = true;
    const tabs = vscode.window.tabGroups.all.flatMap(group => group.tabs);
    const found = new Set(tabs.filter(tab => tab.input instanceof vscode.TabInputCustom && tab.input.viewType === viewType)
      .map(tab => key(tab.input.uri.fsPath)));
    const passed = startWasAbsent && fixtures.every(fixture => found.has(key(fixture.file))) &&
      initialRead.revision === 0 && reloadRead.revision === 1 && secondaryRead.revision === 0 &&
      initialRead.port === reloadRead.port && initialRead.port !== secondaryRead.port;
    await writeFile(evidence, JSON.stringify({ passed, lazyStart: startWasAbsent, twoRealTabs: found.size === 2,
      initialRead, reloadRead, secondaryRead, wordSaved: proof }), { flag: 'wx' });
  }

  context.subscriptions.push(vscode.window.registerCustomEditorProvider(viewType, {
    async openCustomDocument(uri) {
      const fixture = uri.scheme === 'file' && !uri.authority ? registered.get(key(uri.fsPath)) : undefined;
      if (!fixture) throw new Error('Only the two synthetic lab fixtures are allowed');
      started ??= start();
      await started;
      return { uri, fixture, dispose() {} };
    },
    resolveCustomEditor(document, panel) {
      const fixture = document.fixture;
      const primary = fixture === first;
      let pending = false;
      let disposed = false;
      panel.webview.options = { enableScripts: true, localResourceRoots: [asset] };
      panel.title = 'Word ' + (primary ? 'A' : 'B') + ' · ' + fixture.endpoint.port;
      const subscription = panel.webview.onDidReceiveMessage(async message => {
        if (disposed || message?.kind !== 'word-request' || !Number.isSafeInteger(message.id)) return;
        const respond = value => panel.webview.postMessage({ kind: 'word-response', id: message.id, value });
        if (pending) { await respond({ ok: false, error: 'word_busy' }); return; }
        pending = true;
        try {
          const { route, body } = message;
          let payload;
          if (route === 'save' && typeof body?.text === 'string' && body.text.length <= 4096 && Number.isSafeInteger(body.revision)) {
            payload = { text: body.text, revision: body.revision };
          } else if (primary && route === 'proof' && body?.revision === 1) {
            payload = { vscode: true, revision: 1 };
          } else if (primary && route === 'proof-failed' && /^[a-z_]{1,40}$/.test(body?.error ?? '')) {
            payload = { vscode: true, error: body.error };
          } else if (route === 'close') payload = {};
          else if (route !== 'document') { await respond({ ok: false, error: 'invalid_request' }); return; }
          const value = await wordRequest(fixture.endpoint, route, payload);
          if (!disposed) await respond(value);
          if (value.ok && route === 'document') {
            const observed = { port: fixture.endpoint.port, revision: value.revision };
            if (!primary) secondaryRead ??= observed;
            else if (firstReloaded) reloadRead ??= observed;
            else initialRead ??= observed;
          }
          if (primary && value.ok && route === 'proof' && !firstReloaded && !disposed) {
            proof = true;
            firstReloaded = true;
            render(panel, false);
          }
          await verify();
        } catch {
          if (!disposed) await respond({ ok: false, error: 'request_failed' });
        } finally { pending = false; }
      });
      panel.onDidDispose(() => { disposed = true; subscription.dispose(); });
      render(panel, primary && !firstReloaded);
    },
  }, { webviewOptions: { retainContextWhenHidden: true } }));

  try {
    await lstat(requestFile);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
    startWasAbsent = true;
  }
  for (const [index, fixture] of fixtures.entries()) {
    await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(fixture.file), viewType,
      { viewColumn: index ? vscode.ViewColumn.Two : vscode.ViewColumn.One, preview: false });
  }
}

exports.activate = activate;
