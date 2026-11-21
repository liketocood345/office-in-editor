import { app } from 'electron';
import { createRequire } from 'node:module';
import { writeFile } from 'node:fs/promises';
import { setTimeout as wait } from 'node:timers/promises';

const require = createRequire(import.meta.url);
const resultPath = process.env.OIE_CODEOSS_LAB_RESULT;
let started = false;

async function until(read, label) {
    for (let attempt = 0; attempt < 48; attempt++) {
        const value = await read();
        if (value) return value;
        await wait(250);
    }
    throw new Error(`${label}: timed out`);
}

async function test(win) {
    const result = { stage: 'loading', usedOffice: false, capturedFrames: false };
    let call;
    try {
        const addon = require('./oie-codeoss-slot.node');
        if (typeof addon.operation !== 'function') throw new Error('Node-API entry points unavailable');
        const handle = win.getNativeWindowHandle();
        call = (name, ...values) => addon.operation(name, handle, ...[0, 1, 2, 3].map(index => values[index] ?? 0));
        const read = (expression) => win.webContents.executeJavaScript(expression);
        result.stage = 'workbench-layout';
        const group = await until(() => read(`(() => {
            const body = document.querySelector('.part.editor .editor-group-container.active .editor-container')
                || document.querySelector('.part.editor .editor-group-container .editor-container');
            if (!body || !document.querySelector('.monaco-workbench')) return null;
            const r = body.getBoundingClientRect();
            if (r.width < 320 || r.height < 200) return null;
            return { left: r.left, top: r.top, right: r.right, bottom: r.bottom, scale: devicePixelRatio };
        })()`), 'workbench editor content');
        if (!win.isVisible()) throw new Error('lab window is not visible');
        const scale = group.scale;
        const slot = [group.left, group.top, group.right, group.bottom].map(n => Math.round(n * scale));

        result.stage = 'stock-command-palette';
        const palette = await until(async () => {
            const measured = await read(`(() => {
            const widget = document.querySelector('.quick-input-widget');
            if (!widget || getComputedStyle(widget).display === 'none') return null;
            const r = widget.getBoundingClientRect();
            return r.width > 0 && r.height > 0
                ? { left: r.left, top: r.top, right: r.right, bottom: r.bottom } : null;
            })()`);
            result.layout = { group, palette: measured };
            if (!measured || Math.min(group.right, measured.right) - Math.max(group.left, measured.left) < 60 ||
                Math.min(group.bottom, measured.bottom) - Math.max(group.top, measured.top) < 32) return null;
            return measured;
        }, 'genuine command palette overlapping editor');
        const left = Math.max(group.left, palette.left);
        const top = Math.max(group.top, palette.top);
        const right = Math.min(group.right, palette.right);
        const bottom = Math.min(group.bottom, palette.bottom);
        if (right - left < 60 || bottom - top < 32) throw new Error('editor and command palette did not substantially overlap');
        const x = Math.floor((left + right) / 2);
        const y = Math.floor((top + bottom) / 2);
        const px = Math.round(x * scale);
        const py = Math.round(y * scale);
        const domHitsPalette = () => read(`Boolean(document.elementFromPoint(${x}, ${y})?.closest('.quick-input-widget'))`);
        result.domBeforeAttach = await domHitsPalette();
        if (!result.domBeforeAttach) throw new Error('the workbench hit test did not identify the stock palette before attachment');

        result.stage = 'native-slot';
        result.attach = call('attach', ...slot);
        if (result.attach !== 'attached') throw new Error(result.attach);
        await wait(150);
        result.domAfterAttach = await domHitsPalette();
        result.stockPaletteWin32Hit = call('sample', px, py);
        result.stockPaletteConflictObserved = result.domAfterAttach && result.stockPaletteWin32Hit === 'native-document';

        result.stage = 'native-overlay';
        const bounds = [left + 4, top + 3, right - 4, bottom - 3].map(n => Math.round(n * scale));
        result.overlay = call('overlay', ...bounds);
        if (result.overlay !== 'overlay-shown') throw new Error(result.overlay);
        await wait(150);
        result.nativeOverlayWin32Hit = call('sample', px, py);
        result.nativeOverlayPaintCallback = call('paint-state');
        result.nativeOverlayMarker = call('paint', px, bounds[1] + 6);
        const outsideX = Math.round(((group.right + Math.min(group.right, palette.right)) / 2) * scale);
        const outsideY = Math.round(((group.bottom + Math.min(group.bottom, palette.bottom)) / 2) * scale);
        result.documentHitOutsideOverlay = call('sample', outsideX, outsideY);
        result.nativeOverlayAction = call('activate', px, py);
        result.afterActionWin32Hit = call('sample', px, py);
        result.focusAfterAction = call('focus');
        result.nativeOverlayPassed = result.nativeOverlayWin32Hit === 'palette-button' &&
            result.nativeOverlayMarker === 'marker-painted' && result.documentHitOutsideOverlay === 'native-document' &&
            result.nativeOverlayAction === 'action-fired-focus-restored' && result.afterActionWin32Hit === 'native-document';

        // A host-owned, non-topmost popup calibrates desktop readback against the same marker and button.
        // The document stays in the same child slot during both comparisons.
        result.stage = 'owned-popup-control';
        call('clear-overlay');
        const popupControl = { creation: call('popup-control', ...bounds) };
        if (popupControl.creation === 'overlay-shown') {
            await wait(150);
            popupControl.hit = call('sample', px, py);
            popupControl.paintCallback = call('paint-state');
            popupControl.marker = call('paint', px, bounds[1] + 6);
            popupControl.documentOutside = call('sample', outsideX, outsideY);
            popupControl.action = call('activate', px, py);
            popupControl.hitAfterAction = call('sample', px, py);
        }
        result.ownedPopupControl = popupControl;
        result.stage = 'completed';
    } catch (error) {
        result.error = String(error.message).slice(0, 400);
    } finally {
        if (call && result.attach === 'attached') result.cleanup = call('detach');
        await writeFile(resultPath, JSON.stringify(result, null, 2), { flag: 'wx' });
        app.quit();
    }
}

if (resultPath) {
    app.on('browser-window-created', (_event, win) => {
        win.webContents.on('did-finish-load', () => {
            const url = win.webContents.getURL();
            if (started || !url.startsWith('vscode-file://vscode-app/') || !url.includes('workbench')) return;
            started = true;
            void test(win).catch(error => { console.error(error); app.exit(1); });
        });
    });
}
