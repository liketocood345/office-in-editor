import { app, BrowserWindow } from 'electron';
import { randomBytes, timingSafeEqual } from 'node:crypto';
import { createRequire } from 'node:module';
import { createServer } from 'node:net';
import { setTimeout as delay } from 'node:timers/promises';

// Loaded by the main entry of an editor build with the native host adapter.
const addon = createRequire(import.meta.url)('./oie-native-slot.node');
if (typeof addon.operate !== 'function') throw new Error('Office native host addon unavailable');
const secret = randomBytes(32);
const pipe = '\\\\.\\pipe\\office-in-editor-slot-' + randomBytes(16).toString('hex');
const tokenPattern = /^oie-surface-[0-9a-f]{32}$/;
const connections = new Set();
const slots = new Map();
const watchedWindows = new WeakSet();

function failure(code) {
  return Object.assign(new Error(code), { code });
}

function outerOrigin(frame) {
  try {
    const url = new URL(frame.url);
    if (url.protocol === 'vscode-webview:' && /^[0-9a-z-]{1,128}$/.test(url.hostname)) {
      return url.protocol + '//' + url.hostname;
    }
  } catch { /* Detached or navigating frame. */ }
  return '';
}

async function locate(token) {
  const selector = '#ole-root[data-oie-surface="' + token + '"]';
  for (const win of BrowserWindow.getAllWindows().slice(0, 16)) {
    const main = win.webContents.mainFrame;
    if (!main?.url.startsWith('vscode-file:') && !main?.url.startsWith('file:')) continue;
    const candidates = main.frames.slice(0, 64).filter(frame => outerOrigin(frame));
    const matches = await Promise.all(candidates.map(async outer => {
      try {
        const origin = outerOrigin(outer);
        if (!origin) return undefined;
        const frames = outer.framesInSubtree.slice(0, 8).filter(frame =>
          outerOrigin(frame) === origin || frame.url === 'about:blank' || frame.url === 'about:srcdoc');
        const found = await Promise.all(frames.map(frame =>
          frame.executeJavaScript('Boolean(document.querySelector(' + JSON.stringify(selector) + '))')
            .catch(() => false)));
        return found.includes(true) ? origin : undefined;
      } catch { return undefined; }
    }));
    const origins = matches.filter(Boolean);
    if (origins.length === 1) return { win, origin: origins[0] };
    if (origins.length > 1) throw failure('native_surface_ambiguous');
  }
  throw failure('native_surface_frame_missing');
}

function readEditorRect(origin) {
  const frames = [...document.querySelectorAll('.monaco-workbench iframe')].filter(frame =>
    frame.src.startsWith(origin + '/'));
  if (!frames.length) return { reason: 'iframe_missing' };
  const visible = frames.filter(frame => {
    const style = getComputedStyle(frame);
    return style.visibility === 'visible' && style.display !== 'none';
  });
  if (!visible.length) return { reason: 'iframe_hidden' };
  const rects = [];
  for (const frame of visible) {
    const rect = frame.getBoundingClientRect();
    const left = Math.max(0, rect.left);
    const top = Math.max(0, rect.top);
    const right = Math.min(innerWidth, rect.right);
    const bottom = Math.min(innerHeight, rect.bottom);
    if (right - left >= 32 && bottom - top >= 32)
      rects.push([left, top, right - left, bottom - top].map(n => Math.round(n * devicePixelRatio)));
  }
  if (rects.length !== 1) return { reason: rects.length ? 'iframe_ambiguous' : 'iframe_unlaid' };
  return { rect: rects[0] };
}

async function measure({ win, origin }) {
  try {
    const result = await win.webContents.mainFrame.executeJavaScript('(' + readEditorRect.toString() + ')(' + JSON.stringify(origin) + ')');
    const rect = result?.rect;
    if (Array.isArray(rect) && rect.length === 4 && rect.every(n => Number.isInteger(n) && n >= 0 && n <= 32768)) return { rect };
    if (['iframe_missing', 'iframe_hidden', 'iframe_ambiguous', 'iframe_unlaid'].includes(result?.reason))
      return { reason: result.reason };
  } catch { /* The webview can be detached when moved or closed. */ }
  return { reason: 'measurement_failed' };
}

async function awaitInitialGeometry(socket, token) {
  let location;
  let reason = 'frame_missing';
  let attempts = 0;
  // The inner webview can post its ready message while the workbench's
  // parent iframe is still hidden during an editor transition.
  for (const pause of [0, 40, 100, 200, 350, 500]) {
    if (pause) await delay(pause);
    if (socket.destroyed) throw failure('native_surface_disconnected');
    attempts++;
    if (!location) {
      try { location = await locate(token); }
      catch (error) {
        if (error?.code !== 'native_surface_frame_missing') throw error;
        continue;
      }
    }
    const measured = await measure(location);
    if (measured.rect) return { ...location, rect: measured.rect, geometryAttempts: attempts };
    reason = measured.reason;
  }
  throw failure('native_surface_' + reason);
}

function native(op, slot, token, rect = slot.rect, visible = false) {
  const result = JSON.parse(addon.operate(op, slot.win.getNativeWindowHandle(), token, ...rect, Number(visible)));
  if (result.error) throw failure(result.error);
  return result;
}

function send(socket, message) {
  if (!socket.destroyed && socket.writableLength < 32_768) socket.write(JSON.stringify(message) + '\n');
}

function watchWindow(win) {
  if (watchedWindows.has(win)) return;
  watchedWindows.add(win);
  let waiting = false;
  win.on('close', event => {
    const owned = [...slots].filter(([, slot]) => slot.win === win);
    if (!owned.length) return;
    event.preventDefault();
    if (waiting) return;
    waiting = true;
    const groups = new Map();
    for (const [token, slot] of owned) {
      const tokens = groups.get(slot.socket) ?? [];
      tokens.push(token);
      groups.set(slot.socket, tokens);
    }
    for (const [socket, tokens] of groups) send(socket, { close: tokens });
    setTimeout(() => {
      waiting = false;
      if (!win.isDestroyed() && ![...slots.values()].some(slot => slot.win === win)) win.close();
    }, 2500).unref();
  });
}

async function dispatch(socket, message) {
  const { op, token, visible } = message;
  if (typeof token !== 'string' || !tokenPattern.test(token)) throw failure('invalid_native_surface');
  let slot = slots.get(token);
  if (slot && slot.socket !== socket) throw failure('native_surface_owned');
  if (op === 'create') {
    if (slot) throw failure('native_surface_exists');
    if (slots.size >= 32) throw failure('native_surface_limit');
    const location = await awaitInitialGeometry(socket, token);
    if (socket.destroyed) throw failure('native_surface_disconnected');
    slot = { ...location, visible: visible === true, socket };
    if (slots.has(token) || slots.size >= 32) throw failure('native_surface_conflict');
    const identity = native('create', slot, token, slot.rect, slot.visible);
    slots.set(token, slot);
    watchWindow(slot.win);
    return identity;
  }
  if (!slot) throw failure('native_surface_missing');
  if (op === 'release') {
    native('release', slot, token);
    slots.delete(token);
    return { released: true };
  }
  if (op === 'inspect') return { ...native('inspect', slot, token), viewport: slot.rect, geometryAttempts: slot.geometryAttempts };
  if (op !== 'layout') throw failure('invalid_native_operation');
  const measured = visible === true ? (await measure(slot)).rect : undefined;
  const rect = measured ?? slot.rect;
  const show = Boolean(measured);
  if (socket.destroyed) throw failure('native_surface_disconnected');
  if (show !== slot.visible || rect.some((n, index) => slot.rect[index] !== n)) {
    if (show) native('layout', slot, token, rect, true);
    else native('hide', slot, token);
    slot.rect = rect;
    slot.visible = show;
  }
  return { visible: show };
}

const server = createServer(socket => {
  if (connections.size >= 16) return socket.destroy();
  connections.add(socket);
  let buffer = '';
  let authenticated = false;
  let queue = Promise.resolve();
  let pending = 0;
  socket.setEncoding('utf8');
  socket.setTimeout(3000, () => { if (!authenticated) socket.destroy(); });
  socket.on('error', () => socket.destroy());
  socket.on('close', () => {
    connections.delete(socket);
    for (const [token, slot] of slots) {
      if (slot.socket !== socket) continue;
      try {
        native('release', slot, token);
        slots.delete(token);
      } catch {
        // A live Core child vetoes destruction. Preserve its session.
        try { native('hide', slot, token); } catch { /* Parent gone. */ }
      }
    }
  });
  socket.on('data', chunk => {
    buffer += chunk;
    if (buffer.length > 16_384) return socket.destroy();
    let newline;
    while ((newline = buffer.indexOf('\n')) !== -1) {
      const line = buffer.slice(0, newline);
      buffer = buffer.slice(newline + 1);
      if (line.length > 1024 || pending >= 16) return socket.destroy();
      let message;
      try { message = JSON.parse(line); } catch { return socket.destroy(); }
      if (!message || !Number.isSafeInteger(message.id)) return socket.destroy();
      if (!authenticated) {
        if (typeof message.secret !== 'string' || !/^[0-9a-f]{64}$/.test(message.secret) ||
            !timingSafeEqual(Buffer.from(message.secret, 'hex'), secret)) return socket.destroy();
        authenticated = true;
        socket.setTimeout(0);
        send(socket, { id: message.id, value: { protocol: 1 } });
        continue;
      }
      pending++;
      queue = queue.then(async () => {
        if (socket.destroyed) return;
        try { send(socket, { id: message.id, value: await dispatch(socket, message) }); }
        catch (error) {
          send(socket, { id: message.id, error: /^native_[a-z0-9_]+$/.test(error?.code) ? error.code : 'native_host_failed' });
        }
      }).finally(() => { pending--; });
    }
  });
});
server.on('error', () => { delete process.env.OIE_NATIVE_PIPE; delete process.env.OIE_NATIVE_TOKEN; });
server.listen(pipe);
process.env.OIE_NATIVE_PIPE = pipe;
process.env.OIE_NATIVE_TOKEN = secret.toString('hex');
app.on('quit', () => {
  for (const socket of connections) socket.destroy();
  server.close();
});
