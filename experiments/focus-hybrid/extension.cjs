const vscode = require('vscode');
const { execFile, spawn } = require('node:child_process');
const { randomBytes } = require('node:crypto');
const { appendFileSync } = require('node:fs');
const { readFile, rename, writeFile } = require('node:fs/promises');
const path = require('node:path');
const { createInterface } = require('node:readline');
const { setTimeout: delay } = require('node:timers/promises');
const { promisify } = require('node:util');
const execute = promisify(execFile);
const apps = Object.freeze({ word: 'Word', excel: 'Excel', powerpoint: 'PowerPoint' });
const policies = new Set(['always', 'switch', 'mouse']);
const sessions = new Map();
const opening = new Map();
const isOffice = app => typeof app === 'string' && Object.hasOwn(apps, app);
const powershell = () => path.join(process.env.SystemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');

function page(app) {
  const nonce = randomBytes(18).toString('base64');
  return `<!doctype html><html lang="zh-CN"><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}'">
<style nonce="${nonce}">
html,body{width:100%;height:100%;margin:0;padding:0;overflow:hidden}body{display:flex;flex-direction:column}
header{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:6px 10px;background:var(--vscode-editorGroupHeader-tabsBackground);color:var(--vscode-foreground)}
nav{display:flex;align-items:center;gap:6px}header button{padding:3px 6px;font:inherit;cursor:pointer;color:var(--vscode-foreground);background:var(--vscode-dropdown-background);border:1px solid var(--vscode-dropdown-border)}
main{position:relative;width:100%;min-height:0;flex:1;overflow:hidden}#edit,img{display:block;width:100%;height:100%;border:0;padding:0;margin:0;box-shadow:none}
#edit{background:none;cursor:pointer}img{object-fit:fill}p{position:absolute;top:0;left:0;margin:12px;pointer-events:none}[hidden]{display:none!important}
</style>
<header><nav aria-label="打开 Office 标签页">${Object.entries(apps).map(([id, name]) => `<button data-office="${id}"${id === app ? ' aria-current="page"' : ''}>${name}</button>`).join('')}</nav></header>
<main aria-label="Office Focus Preview Surface ${app}"><button id="edit" aria-label="进入 ${apps[app]} 编辑" hidden>
<img id="preview" alt="本机 ${apps[app]} 预览"></button><p id="status" role="status">等待 ${apps[app]} 预览</p></main>
<script nonce="${nonce}">
const api=acquireVsCodeApi(), edit=document.getElementById('edit'), image=document.getElementById('preview');
const status=document.getElementById('status'), surface=document.querySelector('main');
edit.onclick=event=>api.postMessage({action:'activate',trusted:event.isTrusted});
for(const button of document.querySelectorAll('[data-office]'))button.onclick=()=>api.postMessage({action:'open',app:button.dataset.office});
let layoutTimer;
new ResizeObserver(()=>{clearTimeout(layoutTimer);layoutTimer=setTimeout(()=>api.postMessage({action:'layout'}),100);}).observe(surface);
window.addEventListener('message',({data})=>{
 if(data.kind==='frame'){
   const received=performance.now();
   image.onload=()=>{
     const decodeMs=Math.round(performance.now()-received);
     edit.hidden=false;status.hidden=true;
     requestAnimationFrame(()=>requestAnimationFrame(()=>{
       const rect=image.getBoundingClientRect();
       api.postMessage({action:'paint',sequence:data.sequence,decodeMs,dpr:devicePixelRatio,
         width:Math.round(rect.width*devicePixelRatio),height:Math.round(rect.height*devicePixelRatio)});
     }));
   };
   image.src='data:image/png;base64,'+data.png;
 }
 if(data.kind==='status'){status.textContent=data.text;status.hidden=!data.text;}
});
api.postMessage({action:'ready'});
</script></html>`;
}

exports.activate = async context => {
  const root = process.env.OIE_FOCUS_LAB;
  if (process.platform !== 'win32' || !root || !/^focus-hybrid-[a-f0-9]{32}$/.test(path.basename(root)) ||
      path.dirname(root).toLowerCase() !== path.join(process.env.LOCALAPPDATA ?? '', 'OfficeInEditorLab').toLowerCase()) return;
  let input;
  for (let attempt = 0; attempt < 20 && !input; attempt++) {
    try { input = JSON.parse(await readFile(path.join(root, 'inputs.json'), 'utf8')); }
    catch (error) { if (error.code !== 'ENOENT') throw error; await delay(100); }
  }
  if (!input || 'focus-hybrid-' + input.run !== path.basename(root) || !isOffice(input.app) || !Number.isSafeInteger(input.officePid) ||
      !Number.isSafeInteger(input.codePid) || input.stem !== 'synthetic-' + input.run) return;
  const { McpRegistration } = require('./mcp.cjs');
  const mcp = new McpRegistration(context.extensionPath);
  context.subscriptions.push(mcp);
  await mcp.refresh();
  function showDocument(doc) {
    const session = openDocument(context, root, doc, open, mcp.status.mode);
    sessions.set(doc.app, session);
    void session.completed.finally(() => { if (sessions.get(doc.app) === session) sessions.delete(doc.app); }).catch(() => {});
  }
  async function open(application) {
    if (!isOffice(application)) return;
    const existing = sessions.get(application);
    if (existing) { existing.reveal(); return; }
    if (opening.has(application)) return opening.get(application);
    const task = execute(powershell(), ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
      path.join(root, 'open-fixture.ps1'), '-Run', input.run, '-App', application, '-AsJson'],
    { windowsHide: true, timeout: 22000, maxBuffer: 16384 }).then(({ stdout }) => {
      const doc = JSON.parse(stdout);
      if (doc.app !== application || doc.stem !== input.stem || !Number.isSafeInteger(doc.officePid)) return;
      showDocument({ ...input, ...doc, smoke: false, cycles: 0, primary: false });
    }).catch(() => vscode.window.showErrorMessage(`无法打开 ${apps[application]} 实验标签页。`))
      .finally(() => opening.delete(application));
    opening.set(application, task);
    return task;
  }
  showDocument({ ...input, primary: true });
};

function openDocument(context, root, input, openApp, mcpStatus) {
  const appName = apps[input.app];
  const panel = vscode.window.createWebviewPanel('officeFocusHybridLab', `Office in Editor · ${appName}`, vscode.ViewColumn.One,
    { enableScripts: true, retainContextWhenHidden: true, localResourceRoots: [] });
  const output = vscode.window.createOutputChannel(`Office Focus Hybrid · ${appName}`);
  context.subscriptions.push(panel, output);
  let ready = false, closed = false, frame, painted = 0, paintedSize, paintedWhileParked = false;
  let smokeRequested = false, saveRequested = false, editing = false, keyboardFocused = false;
  let resizedFrom, resizeActivationPending = false, resizeSmokeStarted = false;
  const cycles = input.smoke && Number.isInteger(input.cycles) ? input.cycles : 0;
  // 'click' drives preview clicks; 'raise' pulls the parked Office forward externally; 'raise-baseline' does the same with adoption disabled.
  const scenario = cycles && ['raise', 'raise-baseline'].includes(input.scenario) ? input.scenario : 'click';
  let round = 0, sweepPassed = 0, sweepPhase = 'first-park', sweepTimer, sweepDeadline;
  let enteredOffice = false, heldKeyboard = false, returnedToPreview = false;
  let lastPaintSequence = 0, lastPaintPark = -1, activationId = 0, parkId = 0, seq = 0, clockSent;
  const clockSampleTarget = 5;
  let clockSamples = 0, clockBest;
  const events = [];
  const started = performance.now();
  const suffix = input.primary ? '' : `-${input.app}`;
  const journal = path.join(root, `events${suffix}.jsonl`);
  let journalAvailable = true;
  function trace(stage, value, layer = 'host', frameSeq) {
    const tMs = Math.round(performance.now() - started);
    const corr = { activationId, parkId, ...(frameSeq === undefined ? {} : { frameSeq }) };
    const row = { run: input.run, app: input.app, layer, stage, seq: ++seq, tMs, ms: tMs, corr, ...value };
    if (events.length >= 72) events.shift();
    events.push(row);
    const line = JSON.stringify(row);
    if (journalAvailable) {
      try { appendFileSync(journal, line + '\n'); }
      catch { journalAvailable = false; output.appendLine('{"stage":"journal_write_failed"}'); }
    }
    if (!closed) output.appendLine(line);
  }
  const args = [String(input.officePid), String(input.codePid), input.stem, input.app];
  if (input.smoke) args.push('--smoke');
  const native = spawn(path.join(root, 'native.exe'), args, { stdio: ['pipe', 'pipe', 'ignore'], windowsHide: true });
  const nativeExited = new Promise(resolve => native.once('close', resolve));
  function tellNative(command) { if (native.stdin.writable && !native.stdin.destroyed) native.stdin.write(command + '\n'); }
  if (input.smoke && policies.has(input.pixelPolicy)) tellNative(`policy ${input.pixelPolicy}`);
  if (scenario === 'raise-baseline') tellNative('adopt off');
  function verifyOfficeSave(returnToEditor = () => tellNative('return-to-editor')) {
    if (saveRequested) return;
    saveRequested = true;
    if (!keyboardFocused) {
      trace('smoke_save_skipped_no_keyboard_focus');
      returnToEditor();
      return;
    }
    trace('smoke_save_requested');
    const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
      path.join(root, 'save-probe.ps1'), '-Run', input.run, '-App', input.app, '-OfficePid', String(input.officePid)];
    if (input.pixelPolicy === 'mouse') args.push('-ExercisePointer');
    const save = spawn(powershell(), args,
    { windowsHide: true, stdio: 'ignore', timeout: 12000 });
    save.on('error', () => trace('smoke_save_worker_error'));
    save.once('close', code => {
      trace('smoke_save_finished', { code });
      returnToEditor();
    });
  }
  native.stdin.on('error', () => {});
  native.on('error', () => trace('native_start_failed'));
  native.on('exit', code => trace('native_exited', { code }));
  const nativeLines = createInterface({ input: native.stdout });
  function showFrame() {
    if (!closed && ready && panel.visible && frame) void panel.webview.postMessage({ kind: 'frame', ...frame });
  }
  let finish;
  const ended = new Promise(resolve => {
    finish = () => {
      if (closed) return;
      closed = true;
      resolve();
    };
    panel.onDidDispose(finish);
    panel.onDidChangeViewState(() => {
      if (panel.visible) tellNative('layout');
      showFrame();
    });
    panel.webview.onDidReceiveMessage(message => {
      if (message?.action === 'ready') { ready = true; trace('webview_ready'); showFrame(); }
      if (message?.action === 'layout' && panel.visible) tellNative('layout');
      if (message?.action === 'open' && panel.active && panel.visible) void openApp(message.app);
      if (message?.action === 'paint' && Number.isSafeInteger(message.sequence) && message.sequence === frame?.sequence &&
          message.sequence > lastPaintSequence && Number.isInteger(message.width) && Number.isInteger(message.height)) {
        lastPaintSequence = message.sequence;
        painted++;
        if (frame.parked) paintedWhileParked = true;
        const dimensionsChanged = message.width !== paintedSize?.width || message.height !== paintedSize?.height;
        const firstForPark = frame.parked && frame.parkId !== lastPaintPark;
        paintedSize = { width: message.width, height: message.height };
        if (dimensionsChanged || firstForPark) {
          const display = {};
          if (Number.isFinite(message.decodeMs)) display.decodeMs = message.decodeMs;
          if (Number.isFinite(message.dpr)) display.dpr = message.dpr;
          trace('preview_painted', { parked: frame.parked, ...paintedSize, ...display }, 'webview', message.sequence);
        }
        if (firstForPark) {
          lastPaintPark = frame.parkId;
          if (parkId > 0 && enteredOffice) returnedToPreview = true;
          sweepPreviewReady();
        }
        if (resizedFrom && (paintedSize.width !== resizedFrom.width || paintedSize.height !== resizedFrom.height)) {
          resizedFrom = undefined;
          resizeActivationPending = true;
          trace('smoke_resized_activation_request', paintedSize);
          void delay(400, undefined, { ref: false }).then(() => tellNative('activate'));
        }
      }
      if (message?.action === 'activate' && panel.active && panel.visible && frame && !editing) {
        activationId++;
        trace('preview_user_click', { trusted: message.trusted === true }, 'webview');
        tellNative('activate');
      }
    });
    nativeLines.on('line', line => {
      if (line.length > 1_050_000) return;
      let event;
      try { event = JSON.parse(line); } catch { return; }
      if (event?.stage && /^[a-z][a-z0-9_]{0,47}$/.test(event.stage)) {
        const details = {};
        for (const field of ['hresult', 'nativeUs', 'width', 'height', 'alignmentErrorPx', 'cropWidthPx', 'cropHeightPx',
          'intervalMs', 'framesRead', 'framesSent', 'unchanged', 'throttled', 'oversize', 'bytes', 'helperCpuMs',
          'displayCount', 'editorDpi', 'officeBits', 'raised', 'recovered', 'offscreenAtDeadline', 'elapsedMs',
          'editorRestored', 'docked', 'direct', 'watchdog']) {
          if (Number.isInteger(event[field])) details[field] = event[field];
        }
        if (['office', 'editor', 'other', 'none'].includes(event.owner)) details.owner = event.owner;
        if (event.stage === 'office_foreground_released') parkId++;
        if (event.stage === 'office_foreground_adopted') activationId++;
        let calibrated;
        if (event.stage === 'native_clock' && Number.isInteger(event.nativeUs) && clockSent !== undefined) {
          const received = performance.now();
          const sample = { roundTripMs: Math.round(received - clockSent),
            hostOffsetMs: Math.round((received + clockSent) / 2 - started - event.nativeUs / 1000) };
          Object.assign(details, sample, { sample: ++clockSamples });
          clockSent = undefined;
          if (!clockBest || sample.roundTripMs < clockBest.roundTripMs) clockBest = sample;
          if (clockSamples < clockSampleTarget) sendClock();
          else calibrated = { samples: clockSamples, ...clockBest };
        }
        trace(event.stage, details, 'native');
        if (calibrated) trace('native_clock_calibrated', calibrated);
        // Calibrate only after the helper has bound the Office window, so startup work does not inflate the round trip.
        if (event.stage === 'office_window_verified' && !clockSamples && clockSent === undefined) sendClock();
        if (cycles && ['smoke_preview_click_blocked', 'smoke_editor_click_blocked', 'smoke_return_denied'].includes(event.stage)) {
          trace('sweep_input_blocked', { round, phase: sweepPhase, reason: event.stage });
          finish();
          return;
        }
        if (event.stage === 'office_keyboard_focus_confirmed') heldKeyboard = keyboardFocused = true;
        if (event.stage === 'office_foreground_adopted') {
          editing = enteredOffice = true;
          void panel.webview.postMessage({ kind: 'status', text: '' });
        }
        if (cycles && event.stage === 'smoke_parked_raise_result') sweepRaiseObserved(event);
        if (cycles && event.stage === 'smoke_raise_precondition_failed') {
          trace('sweep_input_blocked', { round, phase: sweepPhase, reason: event.stage });
          finish();
          return;
        }
        if (event.stage === 'native_foreground_confirmed') {
          editing = enteredOffice = true;
          if (cycles && scenario === 'click') sweepActivated();
          else if (input.smoke && resizeActivationPending) {
            resizeActivationPending = false;
            void delay(650, undefined, { ref: false }).then(() => tellNative('return-to-editor'));
          } else if (input.smoke) verifyOfficeSave();
        }
        if (event.stage === 'office_foreground_released') {
          editing = false;
          keyboardFocused = false;
        }
        if (!cycles && event.stage === 'wgc_parked_frame_sampled' && !smokeRequested && input.smoke) {
          smokeRequested = true;
          void delay(1500, undefined, { ref: false }).then(() => {
            activationId++;
            trace('smoke_activation_request');
            tellNative('activate');
          });
        } else if (!cycles && event.stage === 'wgc_parked_frame_sampled' && input.smoke && saveRequested && !resizeSmokeStarted && paintedSize) {
          resizeSmokeStarted = true;
          resizedFrom = paintedSize;
          trace('smoke_resize_sidebar');
          void vscode.commands.executeCommand('workbench.action.toggleSidebarVisibility').then(undefined, () => trace('smoke_resize_failed'));
        }
        if (event.stage === 'native_activation_denied' || event.stage === 'activation_precondition_failed') {
          void panel.webview.postMessage({ kind: 'status', text: `请在 VS Code 前台点击预览以激活 ${appName}` });
        }
        if (event.stage === 'office_modal_blocking') {
          void panel.webview.postMessage({ kind: 'status', text: `${appName} 有未处理的对话框，请先关闭对话框` });
        }
        if (event.stage === 'office_modal_cleared') {
          void panel.webview.postMessage({ kind: 'status', text: '' });
        }
        if (event.stage === 'native_alignment_unconfirmed') {
          void panel.webview.postMessage({ kind: 'status', text: `当前内容区无法与 ${appName} 窗口对齐` });
        }
      } else if (Number.isSafeInteger(event?.sequence) && typeof event.png === 'string' && /^[a-zA-Z0-9+/=]+$/.test(event.png)) {
        if (!frame) trace('first_frame_received', { parked: event.parked === true, base64Bytes: event.png.length });
        frame = { sequence: event.sequence, png: event.png, parked: event.parked === true, parkId };
        showFrame();
      }
    });
    panel.webview.html = page(input.app);
    if (cycles) {
      sweepDeadline = setTimeout(() => {
        trace('sweep_timeout', { round, phase: sweepPhase });
        finish();
      }, 30_000 + cycles * 3600);
      sweepDeadline.unref();
    } else if (input.smoke) void delay(25_000, undefined, { ref: false }).then(finish);
  });

  function sweepPreviewReady() {
    if (!cycles || (sweepPhase !== 'first-park' && sweepPhase !== 'returning')) return;
    if (sweepPhase === 'returning') {
      trace('sweep_round_passed', { round });
      sweepPassed++;
      if (sweepPassed === cycles) {
        sweepPhase = 'complete';
        finish();
        return;
      }
    }
    const initial = round === 0;
    const idleAt = performance.now();
    const idleCpu = process.cpuUsage();
    sweepPhase = 'preview-ready';
    sweepTimer = setTimeout(() => {
      if (closed || sweepPhase !== 'preview-ready') return;
      if (initial) {
        const cpu = process.cpuUsage(idleCpu);
        trace('sweep_idle', { intervalMs: Math.round(performance.now() - idleAt),
          hostCpuMs: Math.round((cpu.user + cpu.system) / 1000) });
      }
      round++;
      const returnMethod = round % 2 ? 'click' : 'foreground-request';
      if (scenario !== 'click') {
        sweepPhase = 'raising';
        trace('sweep_round_started', { round, scenario, returnMethod });
        tellNative('smoke-raise-parked');
        return;
      }
      sweepPhase = 'activating';
      trace('sweep_round_started', { round, returnMethod });
      tellNative('smoke-click-preview');
    }, initial ? 2600 : 350);
    sweepTimer.unref();
  }

  function sweepRaiseObserved(event) {
    if (sweepPhase !== 'raising') return;
    trace('sweep_raise_observed', { round, raised: event.raised, recovered: event.recovered,
      offscreenAtDeadline: event.offscreenAtDeadline });
    if (event.recovered === 1) {
      // Fixed variant: Office is now editing in the tab; return the same way as click rounds.
      sweepPhase = 'returning';
      sweepTimer = setTimeout(() => {
        if (!closed && sweepPhase === 'returning') tellNative(round % 2 ? 'smoke-click-editor' : 'return-to-editor');
      }, 330);
      sweepTimer.unref();
    } else {
      // Baseline variant: Office never left the parked position, so the current preview is still valid.
      sweepPhase = 'returning';
      sweepPreviewReady();
    }
  }

  function sweepActivated() {
    if (sweepPhase !== 'activating' || !keyboardFocused) {
      trace('sweep_focus_unconfirmed', { round, phase: sweepPhase });
      finish();
      return;
    }
    sweepPhase = 'editing';
    function returnToEditor() {
      if (closed || sweepPhase !== 'editing') return;
      sweepPhase = 'returning';
      tellNative(round % 2 ? 'smoke-click-editor' : 'return-to-editor');
    }
    if (round === 1) verifyOfficeSave(returnToEditor);
    else {
      if (input.pixelPolicy === 'mouse' && round % 2) tellNative('smoke-move-pointer');
      sweepTimer = setTimeout(returnToEditor, 330);
      sweepTimer.unref();
    }
  }

  function sendClock() {
    clockSent = performance.now();
    tellNative('clock');
  }
  const completed = ended.then(async () => {
    clearTimeout(sweepTimer);
    clearTimeout(sweepDeadline);
    tellNative('exit');
    await Promise.race([nativeExited, delay(2500, undefined, { ref: false })]);
    const result = { run: input.run, app: input.app, vscodeVersion: vscode.version, painted, paintedWhileParked, captureSource: `wgc-${input.app}`,
      confirmedForeground: enteredOffice, confirmedKeyboardFocus: heldKeyboard, returnedToPreview,
      recordedNativeExit: events.some(row => row.stage === 'probe_exited'),
      ...(cycles ? { sweep: { requested: cycles, passed: sweepPassed, phase: sweepPhase, scenario } } : {}), events };
    await writeFile(path.join(root, `result${suffix}.partial`), JSON.stringify(result));
    await rename(path.join(root, `result${suffix}.partial`), path.join(root, `result${suffix}.json`));
  });
  void completed.then(() => {
    if (input.smoke) return vscode.commands.executeCommand('workbench.action.closeWindow');
  }).catch(() => trace('session_finalization_failed'));
  trace('extension_ready', { mcpStatus });
  return { finish, completed, reveal: () => panel.reveal() };
}

exports.deactivate = () => {
  for (const session of sessions.values()) session.finish();
  return Promise.allSettled(Array.from(sessions.values(), session => session.completed));
};
