import { randomBytes } from "node:crypto";
import { escapeHtml } from "../../../adapter-sdk/src";

export { escapeHtml };

export function actionButton(action: string, label: string): string {
  return '<button type="button" data-action="' + escapeHtml(action) + '">' + escapeHtml(label) + "</button>";
}

/** No runtime value enters script source, CSS, or an executable URL. */
export function panelHtml(title: string, body: string): string {
  const nonce = randomBytes(18).toString("base64");
  return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}'; base-uri 'none'; form-action 'none'">
<title>${escapeHtml(title)}</title>
<style nonce="${nonce}">
html,body{margin:0;height:100%;width:100%;overflow:hidden;background:var(--vscode-editor-background);color:var(--vscode-foreground)}
body{font-family:var(--vscode-font-family);padding:24px;max-width:960px;line-height:1.65}
body.host{padding:0;max-width:none}
.host-status{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;padding:24px;text-align:center}
.host-ready{position:absolute;inset:0;background:transparent}
h1{font-size:1.6em}h2{font-size:1.1em;margin-top:24px}
.notice{border-left:3px solid var(--vscode-focusBorder);padding:10px 16px;background:var(--vscode-textBlockQuote-background)}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin:20px 0}
button{cursor:pointer;font:inherit;padding:7px 12px;background:var(--vscode-button-background);color:var(--vscode-button-foreground);border:1px solid transparent;border-radius:3px}
button:focus-visible{outline:2px solid var(--vscode-focusBorder);outline-offset:3px}
button:hover{background:var(--vscode-button-hoverBackground)}
pre{white-space:pre-wrap;overflow-wrap:anywhere;background:var(--vscode-textCodeBlock-background);padding:14px}
dt{font-weight:600}dd{margin:0 0 10px;overflow-wrap:anywhere}.muted{color:var(--vscode-descriptionForeground)}
</style></head><body><h1>${escapeHtml(title)}</h1>${body}
<script nonce="${nonce}">
const api=acquireVsCodeApi();
for(const button of document.querySelectorAll('button[data-action]')){
  button.addEventListener('click',()=>api.postMessage({action:button.getAttribute('data-action')}));
}
</script></body></html>`;
}

export function hostHtml(
  title: string,
  message: string,
  surfaceToken: string,
  appearance?: { loaderSrc?: string; cspSource?: string; officePage?: { url: string; origin: string } },
): string {
  void title;
  const nonce = randomBytes(18).toString("base64");
  const token = escapeHtml(surfaceToken);
  const loaderSrc = appearance?.loaderSrc ? escapeHtml(appearance.loaderSrc) : "";
  const imgCsp = appearance?.cspSource ? ` img-src ${appearance.cspSource};` : "";
  const officePage = appearance?.officePage;
  const pageCsp = officePage ? ` frame-src ${escapeHtml(officePage.origin)};` : "";
  const frame = officePage ? `<iframe class="office-page" title="本机 Office" sandbox="allow-scripts allow-same-origin" src="${escapeHtml(officePage.url)}"></iframe>` : "";
  const loaderImg = loaderSrc
    ? `<img id="status-loader" class="host-loader-gif" width="90" height="90" alt="" src="${loaderSrc}">`
    : "";
  return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';${imgCsp}${pageCsp} base-uri 'none'; form-action 'none'">
<title>${token}</title>
<style nonce="${nonce}">
html,body,#ole-root{margin:0;height:100%;width:100%;min-height:100%;min-width:100%;overflow:hidden;background:var(--vscode-editor-background);color:var(--vscode-foreground)}
#ole-root{position:absolute;inset:0}
.office-page{position:absolute;inset:0;width:100%;height:100%;border:0}
.host-status{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;padding:24px;text-align:center;color:var(--vscode-foreground);font-family:var(--vscode-font-family)}
.host-loader{display:flex;flex-direction:column;align-items:center;justify-content:center}
.host-loader-gif{width:90px;height:90px;object-fit:contain;display:block;flex:none}
.host-loader-line{margin:16px 0 0;padding:0;font:inherit;color:inherit;line-height:1.4}
.host-ready{position:absolute;inset:0}
.oie-beacon{position:absolute;width:1px;height:1px;padding:0;margin:0;border:0;border-radius:0;pointer-events:none;z-index:3;background:#000}
.oie-beacon[data-oie-beacon="tl"]{left:0;top:0}
.oie-beacon[data-oie-beacon="tr"]{right:0;top:0}
.oie-beacon[data-oie-beacon="bl"]{left:0;bottom:0}
.oie-beacon[data-oie-beacon="br"]{right:0;bottom:0}
[hidden]{display:none!important}
</style></head><body>
<div id="ole-root" class="ole-root" role="application" aria-label="${token}" title="${token}" data-oie-surface="${token}">
${frame}
<div class="oie-beacon" data-oie-beacon="tl" role="img" aria-label="${token}:tl"></div>
<div class="oie-beacon" data-oie-beacon="tr" role="img" aria-label="${token}:tr"></div>
<div class="oie-beacon" data-oie-beacon="bl" role="img" aria-label="${token}:bl"></div>
<div class="oie-beacon" data-oie-beacon="br" role="img" aria-label="${token}:br"></div>
<div id="status" class="host-status"${message ? "" : " hidden"}><div class="host-loader">${loaderImg}<p class="host-loader-line"><span id="status-label">${escapeHtml(message)}</span><span id="status-dots">.</span></p></div></div>
<div id="surface" class="host-ready" hidden></div>
</div>
<script nonce="${nonce}">
(()=>{
const api=acquireVsCodeApi();
const surfaceToken=${JSON.stringify(surfaceToken)};
document.title=surfaceToken;
const root=document.getElementById('ole-root');
const statusNode=document.getElementById('status');
const surface=document.getElementById('surface');
const statusLabel=document.getElementById('status-label');
const statusDots=document.getElementById('status-dots');
const statusLoader=document.getElementById('status-loader');
let acknowledged=false;
let marks=1;
let ellipsis=0;
if(typeof setInterval==='function'&&statusDots){
  ellipsis=setInterval(()=>{
    if(!statusDots||statusNode.hidden) return;
    marks=marks%3+1;
    statusDots.textContent='.'.repeat(marks);
  },400);
}
function stopLoader(){
  if(ellipsis&&typeof clearInterval==='function') clearInterval(ellipsis);
  ellipsis=0;
}
window.addEventListener('error',()=>api.postMessage({action:'script-error'}));
api.postMessage({action:'ready'});
window.addEventListener('message',event=>{
  const data=event.data||{};
  if(data.kind==='layout-ack') acknowledged=true;
  if(data.kind==='layout-request') report();
  if(data.kind==='status'&&typeof data.text==='string'){
    stopLoader();
    if(statusLoader) statusLoader.hidden=true;
    if(statusDots) statusDots.textContent='';
    if(statusLabel) statusLabel.textContent=data.text;
    else statusNode.textContent=data.text;
    statusNode.hidden=false;
    surface.hidden=true;
  }
  if(data.kind==='attached'){
    stopLoader();
    statusNode.hidden=true;
    surface.hidden=false;
  }
});
let reportQueued=false;
function report(){
  if(reportQueued) return;
  reportQueued=true;
  requestAnimationFrame(()=>{
    reportQueued=false;
    const box=root&&root.getBoundingClientRect?root.getBoundingClientRect():null;
    const width=Math.round((box&&box.width)||window.innerWidth||0);
    const height=Math.round((box&&box.height)||window.innerHeight||0);
    const dpr=window.devicePixelRatio||1;
    const viewport=window.visualViewport;
    api.postMessage({action:'layout',surfaceToken,width,height,x:0,y:0,dpr,clientOnly:true,innerW:Math.round(window.innerWidth||0),innerH:Math.round(window.innerHeight||0),vvW:viewport?Math.round(viewport.width):0,vvH:viewport?Math.round(viewport.height):0,vvScale:viewport?viewport.scale:1});
  });
}
window.addEventListener('resize',report);
new ResizeObserver(report).observe(document.body);
if(root) new ResizeObserver(report).observe(root);
for(const delay of [100,300,750,1500,3000,6000]) setTimeout(()=>{if(!acknowledged)report();},delay);
report();
})();
</script></body></html>`;
}
