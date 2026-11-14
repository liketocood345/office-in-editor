import { App } from '@modelcontextprotocol/ext-apps';
import { prepareContextUpdate } from './context.mjs';

const app = new App({ name: 'Office in Editor', version: '0.1.0' }, {}, { autoResize: true });
const status = document.getElementById('status')!;
const label = document.getElementById('label')!;
const list = document.getElementById('references')!;
const reconnect = document.getElementById('reconnect') as HTMLButtonElement;
let configuration: { targetId: string; token: string; appOrigin: string; label?: string } | undefined;
let controller: AbortController | undefined;
let connected = false;
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
async function acknowledge(delivery: any, state: string) {
  if (!configuration) return;
  const response = await fetch(`${safeOrigin(configuration.appOrigin)}/app/ack`, {
    method: 'POST', headers: { Authorization: `Bearer ${configuration.token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ targetId: configuration.targetId, requestId: delivery.requestId, refId: delivery.refId, status: state }),
    cache: 'no-store', credentials: 'omit'
  });
  if (!response.ok) throw new Error('acknowledgement_failed');
}
async function receive(delivery: any) {
  try {
    if (!app.getHostCapabilities()?.updateModelContext) {
      await acknowledge(delivery, 'unsupported');
      setStatus('This chat host cannot accept context updates. The reference remains available in Office References.');
      return;
    }
    const update = prepareContextUpdate(items, delivery);
    // This updates this App's hosting chat context. It never submits a message.
    await app.updateModelContext(update.payload);
    items = update.items;
    render();
    await acknowledge(delivery, 'attached');
    setStatus(`${items.length} reference(s) added to this chat. Nothing was sent automatically.`);
  } catch {
    setStatus('The reference was not confirmed. Open Office References to inspect or retry it.');
    // An unknown host outcome must not be automatically retried or labelled attached.
  }
}
async function listen() {
  if (!configuration || !connected) return;
  controller?.abort();
  controller = new AbortController();
  reconnect.disabled = true;
  try {
    const response = await fetch(`${safeOrigin(configuration.appOrigin)}/app/events?targetId=${encodeURIComponent(configuration.targetId)}`, {
      headers: { Authorization: `Bearer ${configuration.token}`, Accept: 'text/event-stream' },
      signal: controller.signal, cache: 'no-store', credentials: 'omit'
    });
    if (!response.ok || !response.body) throw new Error('connection_unavailable');
    setStatus('Ready. Select this registered chat in the Office ribbon.');
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    while (true) {
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
        if (delivery.refId && delivery.text) deliveryChain = deliveryChain.then(() => receive(delivery));
      }
    }
    setStatus('Disconnected. Reconnect explicitly when this chat is ready.');
  } catch (error) {
    if ((error as Error)?.name !== 'AbortError') setStatus('Connection unavailable. References are preserved; no automatic retry is running.');
  } finally { reconnect.disabled = false; }
}

// Register one-shot handlers before the initialization handshake.
app.ontoolresult = result => {
  const meta = (result as any)._meta?.officeInEditor;
  if (!meta || typeof meta.targetId !== 'string' || typeof meta.token !== 'string' || typeof meta.appOrigin !== 'string') {
    setStatus('This tool result does not contain a registered Office chat.');
    return;
  }
  try { safeOrigin(meta.appOrigin); } catch { setStatus('Invalid local bridge endpoint.'); return; }
  if (configuration && configuration.targetId !== meta.targetId) {
    controller?.abort();
    items = [];
    render();
  }
  configuration = meta;
  label.textContent = String(meta.label ?? 'Registered Office chat');
  if (connected) void listen();
};
reconnect.addEventListener('click', () => { void listen(); });
window.addEventListener('pagehide', () => controller?.abort(), { once: true });
void app.connect().then(() => {
  connected = true;
  if (!app.getHostCapabilities()?.updateModelContext) {
    setStatus('Context updates are unavailable in this chat host. Office chat attachment is unsupported here.');
    return;
  }
  if (configuration) void listen();
  else setStatus('Waiting for this chat registration.');
}).catch(() => setStatus('The chat host did not initialize this Office bridge.'));
