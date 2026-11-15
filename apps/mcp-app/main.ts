import { App } from '@modelcontextprotocol/ext-apps';
import { prepareContextUpdate } from './context.mjs';

const app = new App({ name: 'Office in Editor', version: '0.1.4' }, {}, { autoResize: true });
const status = document.getElementById('status')!;
const label = document.getElementById('label')!;
const list = document.getElementById('references')!;
const reconnect = document.getElementById('reconnect') as HTMLButtonElement;
type Configuration = { targetId: string; token: string; appOrigin: string; epoch?: string; label?: string };
let configuration: Configuration | undefined;
let controller: AbortController | undefined;
let hostConnected = false;
let listening: Promise<void> | undefined;
let generation = 0;
let items: any[] = [];
let deliveryChain = Promise.resolve();

function setStatus(message: string) { status.textContent = message; }
function render() {
  list.replaceChildren(...items.map(item => {
    const row = document.createElement('li');
    row.textContent = `${item.documentName} · ${item.refId.slice(0, 8)}`;
    return row;
  }));
}
function safeOrigin(value: string): string {
  const url = new URL(value);
  if (url.protocol !== 'http:' || url.hostname !== '127.0.0.1' || !url.port || url.username || url.password || url.pathname !== '/') {
    throw new Error('invalid_local_endpoint');
  }
  return url.origin;
}
async function acknowledge(config: Configuration, delivery: any, state: string) {
  const response = await fetch(`${safeOrigin(config.appOrigin)}/app/ack`, {
    method: 'POST', headers: { Authorization: `Bearer ${config.token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ targetId: config.targetId, requestId: delivery.requestId, refId: delivery.refId, status: state }),
    cache: 'no-store', credentials: 'omit'
  });
  if (!response.ok) throw new Error('acknowledgement_failed');
}
async function receive(config: Configuration, current: number, delivery: any) {
  if (current !== generation) return;
  let attempted = false;
  try {
    if (!app.getHostCapabilities()?.updateModelContext) {
      await acknowledge(config, delivery, 'unsupported');
      setStatus('Context updates are unavailable in this chat host. The reference remains queued.');
      return;
    }
    const update = prepareContextUpdate(items, delivery);
    // Mark in flight before crossing the host boundary; unknown outcomes are never replayed.
    await acknowledge(config, delivery, 'updating');
    if (current !== generation) { await acknowledge(config, delivery, 'unknown'); return; }
    attempted = true;
    await app.updateModelContext(update.payload);
    if (current !== generation) { await acknowledge(config, delivery, 'unknown'); return; }
    await acknowledge(config, delivery, 'attached');
    items = update.items;
    render();
    setStatus(`${items.length} reference(s) added to this chat context. Nothing was sent.`);
  } catch {
    if (attempted) {
      try { await acknowledge(config, delivery, 'unknown'); } catch { /* Disconnect preserves the in-flight outcome. */ }
    }
    if (current === generation) setStatus('Context delivery was not confirmed. Inspect Office References before retrying.');
  }
}
async function listen(config: Configuration, current: number) {
  const ownController = new AbortController();
  controller = ownController;
  reconnect.disabled = true;
  try {
    const response = await fetch(`${safeOrigin(config.appOrigin)}/app/events?targetId=${encodeURIComponent(config.targetId)}`, {
      headers: { Authorization: `Bearer ${config.token}`, Accept: 'text/event-stream' },
      signal: ownController.signal, cache: 'no-store', credentials: 'omit'
    });
    if (!response.ok || !response.body) {
      if (response.status === 401) throw new Error('registration_expired');
      throw new Error('connection_unavailable');
    }
    if (current === generation) setStatus('Ready to receive Office references in this chat context.');
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    while (current === generation) {
      const chunk = await reader.read();
      if (chunk.done) break;
      buffer += decoder.decode(chunk.value, { stream: true });
      if (buffer.length > 2 * 1024 * 1024) throw new Error('event_too_large');
      let boundary: number;
      while ((boundary = buffer.indexOf('\n\n')) >= 0) {
        const frame = buffer.slice(0, boundary).replace(/\r/g, '');
        buffer = buffer.slice(boundary + 2);
        const data = frame.split('\n').filter(line => line.startsWith('data:')).map(line => line.slice(5).trimStart()).join('\n');
        if (!data) continue;
        const delivery = JSON.parse(data);
        if (delivery.refId && delivery.requestId && typeof delivery.text === 'string')
          deliveryChain = deliveryChain.then(() => receive(config, current, delivery));
      }
    }
    if (current === generation) setStatus('Disconnected. Reconnect explicitly when this chat is ready.');
  } catch (error) {
    if (current === generation && (error as Error)?.name !== 'AbortError') {
      setStatus((error as Error)?.message === 'registration_expired'
        ? 'Registration expired. Register this chat again; preserved references are available in Office References.'
        : 'Connection unavailable. References are preserved; no automatic retry is running.');
    }
  } finally {
    if (controller === ownController) controller = undefined;
    if (current === generation) reconnect.disabled = false;
  }
}
function startListening() {
  if (!configuration || !hostConnected || listening) return;
  const current = generation;
  const task = listen(configuration, current);
  listening = task;
  void task.finally(() => {
    if (listening === task) listening = undefined;
    if (current !== generation) startListening();
  });
}

app.ontoolresult = result => {
  const meta = (result as any)._meta?.officeInEditor;
  if (!meta || typeof meta.targetId !== 'string' || typeof meta.token !== 'string' || typeof meta.appOrigin !== 'string') {
    setStatus('This result does not contain a registered Office chat.'); return;
  }
  try { safeOrigin(meta.appOrigin); } catch { setStatus('Invalid local bridge endpoint.'); return; }
  const changed = !configuration || configuration.targetId !== meta.targetId || configuration.token !== meta.token || configuration.appOrigin !== meta.appOrigin;
  if (changed) {
    ++generation;
    controller?.abort();
    items = [];
    render();
  }
  configuration = meta;
  label.textContent = String(meta.label ?? 'Registered Office chat');
  startListening();
};
reconnect.addEventListener('click', startListening);
window.addEventListener('pagehide', () => { ++generation; hostConnected = false; controller?.abort(); }, { once: true });
void app.connect().then(() => {
  hostConnected = true;
  if (!app.getHostCapabilities()?.updateModelContext) {
    setStatus('This chat host cannot accept Office context updates.'); return;
  }
  if (configuration) startListening();
  else setStatus('Waiting for this chat registration.');
}).catch(() => setStatus('The chat host did not initialize this Office bridge.'));

