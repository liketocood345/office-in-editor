const host = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : undefined;
const parameters = new URLSearchParams(location.hash.slice(1));
const secret = host ? undefined : parameters.get('key') || sessionStorage.getItem('oie-lab-key');
const verify = host ? document.documentElement.dataset.smoke === 'vscode'
  : parameters.get('verify') === '1' && /Edg\//.test(navigator.userAgent) && navigator.webdriver === false;
const proofClient = host ? { vscode: true } : { edge: true };
const form = document.querySelector('#editor');
const field = document.querySelector('#text');
const save = document.querySelector('#save');
const end = document.querySelector('#end');
const status = document.querySelector('#status');
let revision;
let committed;
let loadError = false;
let busy = false;
let sequence = 0;
const responses = new Map();

if (host) window.addEventListener('message', event => {
  const message = event.data;
  if (message?.kind !== 'word-response') return;
  responses.get(message.id)?.(message.value);
});

if (/^[0-9a-f]{64}$/.test(secret ?? '')) {
  sessionStorage.setItem('oie-lab-key', secret);
  history.replaceState(null, '', location.pathname);
}

async function request(route, body) {
  if (host) {
    const id = ++sequence;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        responses.delete(id);
        reject(new Error('host_timeout'));
      }, 18_000);
      responses.set(id, result => {
        clearTimeout(timer);
        responses.delete(id);
        if (result?.ok) resolve(result);
        else reject(new Error(result?.error ?? 'request_failed'));
      });
      host.postMessage({ kind: 'word-request', id, route, body });
    });
  }
  const headers = { Authorization: 'Bearer ' + secret };
  const options = { method: 'GET', headers, signal: AbortSignal.timeout(18_000) };
  if (body !== undefined) {
    options.method = 'POST';
    options.body = JSON.stringify(body);
    headers['Content-Type'] = 'application/json';
  }
  const response = await fetch('/api/' + route, options);
  const result = await response.json();
  if (!response.ok || !result.ok) throw new Error(result.error ?? 'request_failed');
  return result;
}

function refreshButtons() {
  save.disabled = loadError || busy || field.value === committed;
  end.disabled = loadError || busy || field.value !== committed;
}

field.addEventListener('input', refreshButtons);
form.addEventListener('submit', async event => {
  event.preventDefault();
  if (save.disabled) return;
  busy = true;
  refreshButtons();
  field.readOnly = true;
  status.value = 'Word 正在保存…';
  try {
    const result = await request('save', { text: field.value, revision });
    revision = result.revision;
    committed = result.text;
    status.value = 'Word 已保存到本机 synthetic.docx。';
    if (verify && revision === 1) {
      await request('proof', { ...proofClient, revision });
      status.value = host ? 'VS Code 标签页编辑与 Word 保存验证通过。' : '普通 Edge 标签页编辑与 Word 保存验证通过。';
    }
  } catch (error) {
    loadError = true;
    if (verify) void request('proof-failed', { ...proofClient, error: error.message }).catch(() => {});
    status.value = error.message === 'edit_conflict' ? '文件版本已变化。新文字仍保留在此页。'
      : '保存未确认。新文字仍保留在此页。';
  } finally {
    busy = false;
    field.readOnly = false;
    refreshButtons();
  }
});

end.addEventListener('click', async () => {
  busy = true;
  refreshButtons();
  try {
    await request('close', {});
    loadError = true;
    field.disabled = true;
    status.value = '实验已结束。';
  } catch { status.value = '本地会话仍保留。'; }
  finally { busy = false; refreshButtons(); }
});

request('document').then(result => {
  committed = field.value = result.text;
  revision = result.revision;
  field.disabled = false;
  status.value = '已连接本机 Microsoft Word。';
  refreshButtons();
  if (verify && revision === 0) {
    field.value += host ? ' · 已通过 VS Code 标签页保存' : ' · 已通过普通 Edge 标签页保存';
    field.dispatchEvent(new Event('input'));
    form.requestSubmit(save);
  }
}).catch(error => {
  loadError = true;
  status.value = '未连接本机 Word。请用本轮实验链接打开。';
  if (verify) void request('proof-failed', { ...proofClient, error: error.message }).catch(() => {});
});
