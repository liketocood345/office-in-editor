import { randomBytes } from "node:crypto";
import { escapeHtml } from "../../../adapter-sdk/src";

/** User-facing text for focus-hybrid notices and states. Kept here so the page script holds no copy. */
export const FOCUS_TEXT = {
  loading: "正在加载 Office",
  modal: "Office 正在显示对话框，点击此处查看",
  captureUnavailable: "无法显示预览，点击此处进入 Office 编辑",
  activationDenied: "Office 没有切到前台，请再点击一次",
  lost: "Office 已关闭此文档",
  openFailed: "打开未完成。",
  addinOutdated: "Office 加载项是旧版本。请关闭所有 Office 窗口后重新打开此标签。",
  windowUnavailable: "Office 没有交出文档窗口。请确认 Office 加载项已启用，然后重新打开此标签。",
} as const;

/** Text for a failed focus.open, by the Core error code. */
export function openFailureText(code: string): string {
  if (code === "office_addin_outdated") return FOCUS_TEXT.addinOutdated;
  if (code === "office_window_unavailable") return FOCUS_TEXT.windowUnavailable;
  return FOCUS_TEXT.openFailed;
}

/**
 * Focus-hybrid tab page. The root element's accessible name is the surface token (plus four corner beacons), which
 * is how Core finds the tab's physical bounds; the page never learns screen coordinates or window handles.
 * Frames arrive as JPEG or PNG bytes and are drawn 1:1 on a canvas; a trusted click asks Core to bring Office over the tab.
 * No runtime value enters script source except the JSON-encoded token and fixed text.
 */
export function focusHtml(surfaceToken: string, appearance?: { loaderSrc?: string; cspSource?: string }): string {
  const nonce = randomBytes(18).toString("base64");
  const token = escapeHtml(surfaceToken);
  const imgCsp = appearance?.cspSource ? ` img-src ${appearance.cspSource};` : "";
  const loader = appearance?.loaderSrc
    ? `<img id="status-loader" class="loader" width="90" height="90" alt="" src="${escapeHtml(appearance.loaderSrc)}">` : "";
  return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';${imgCsp} base-uri 'none'; form-action 'none'">
<title>${token}</title>
<style nonce="${nonce}">
html,body,#ole-root{margin:0;height:100%;width:100%;overflow:hidden;background:var(--vscode-editor-background);color:var(--vscode-foreground);font-family:var(--vscode-font-family)}
#ole-root{position:absolute;inset:0;cursor:text}
#frame{position:absolute;left:0;top:0;display:block}
.status{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px;text-align:center;background:var(--vscode-editor-background)}
.status.over{background:color-mix(in srgb,var(--vscode-editor-background) 70%,transparent)}
.loader{width:90px;height:90px;object-fit:contain}
.status p{margin:16px 0 0;line-height:1.4}
.oie-beacon{position:absolute;width:1px;height:1px;pointer-events:none;z-index:3;background:#000}
.oie-beacon[data-oie-beacon="tl"]{left:0;top:0}.oie-beacon[data-oie-beacon="tr"]{right:0;top:0}
.oie-beacon[data-oie-beacon="bl"]{left:0;bottom:0}.oie-beacon[data-oie-beacon="br"]{right:0;bottom:0}
[hidden]{display:none!important}
</style></head><body>
<div id="ole-root" role="application" tabindex="0" aria-label="${token}" title="${token}" data-oie-surface="${token}">
<canvas id="frame" aria-hidden="true"></canvas>
<div class="oie-beacon" data-oie-beacon="tl" role="img" aria-label="${token}:tl"></div>
<div class="oie-beacon" data-oie-beacon="tr" role="img" aria-label="${token}:tr"></div>
<div class="oie-beacon" data-oie-beacon="bl" role="img" aria-label="${token}:bl"></div>
<div class="oie-beacon" data-oie-beacon="br" role="img" aria-label="${token}:br"></div>
<div id="status" class="status">${loader}<p id="status-text">${escapeHtml(FOCUS_TEXT.loading)}</p></div>
</div>
<script nonce="${nonce}">
(()=>{
const api=acquireVsCodeApi();
const surfaceToken=${JSON.stringify(surfaceToken)};
const root=document.getElementById('ole-root');
const canvas=document.getElementById('frame');
const context=canvas.getContext('2d');
const statusNode=document.getElementById('status');
const statusText=document.getElementById('status-text');
const loader=document.getElementById('status-loader');
let parkId=0, frameSeq=0, drawn=false, drawing=false, queued=null, lost=false;
window.addEventListener('error',()=>api.postMessage({action:'script-error'}));
function show(text,over){
  statusText.textContent=text;
  if(loader) loader.hidden=true;
  statusNode.classList.toggle('over',!!over);
  statusNode.hidden=false;
}
function hide(){ statusNode.hidden=true; }
async function draw(frame){
  if(drawing){ queued=frame; return; }
  drawing=true;
  try{
    const bitmap=await createImageBitmap(new Blob([frame.image],{type:frame.format==='png'?'image/png':'image/jpeg'}));
    // Draw the frame 1:1 in device pixels. The parked window already has the tab's physical size, so this is
    // pixel-exact; while a resize is in flight the old frame is shown unscaled (clipped or with a margin) instead
    // of being stretched, until Core sends a frame at the new size.
    const ratio=window.devicePixelRatio||1;
    if(canvas.width!==bitmap.width||canvas.height!==bitmap.height){ canvas.width=bitmap.width; canvas.height=bitmap.height; }
    canvas.style.width=(bitmap.width/ratio)+'px';
    canvas.style.height=(bitmap.height/ratio)+'px';
    context.drawImage(bitmap,0,0);
    bitmap.close();
    drawn=true;
    if(!lost&&statusNode.dataset.sticky!=='1') hide();
  }catch(_){ api.postMessage({action:'frame-error'}); }
  finally{
    drawing=false;
    const next=queued; queued=null;
    if(next&&next.frameSeq>frameSeq) draw(next);
  }
}
window.addEventListener('message',event=>{
  const data=event.data||{};
  if(data.kind==='layout-request') report();
  if(data.kind==='frame'&&data.parkId>=parkId&&data.frameSeq>frameSeq){
    parkId=data.parkId; frameSeq=data.frameSeq;
    draw({image:data.image,frameSeq:data.frameSeq,format:data.format});
  }
  if(data.kind==='state'){
    if(data.parkId>parkId) parkId=data.parkId;
    if(data.state==='editing'&&statusNode.dataset.sticky!=='1') hide();
    if(data.state==='lost'){ lost=true; }
  }
  if(data.kind==='status'&&typeof data.text==='string'){
    statusNode.dataset.sticky=data.sticky?'1':'';
    show(data.text,drawn&&!data.opaque);
  }
  if(data.kind==='clear-status'){ statusNode.dataset.sticky=''; if(drawn||data.force) hide(); }
});
function activate(event){
  if(!event.isTrusted||lost) return;
  const box=root.getBoundingClientRect();
  const point=event.clientX!==undefined&&box.width>0&&box.height>0
    ?{x:Math.min(1,Math.max(0,(event.clientX-box.left)/box.width)),y:Math.min(1,Math.max(0,(event.clientY-box.top)/box.height))}
    :undefined;
  api.postMessage({action:'activate',surfaceToken,trusted:true,point});
}
root.addEventListener('pointerdown',event=>{ if(event.button===0) activate(event); });
// Wheel over the parked preview scrolls the parked Office window in place (it is not activated). Browser deltas are
// converted to Win32 wheel units (120 = one notch; up/away is positive) and merged per animation frame.
let wheelY=0, wheelX=0, wheelPoint=null, wheelKeys={ctrl:false,shift:false}, wheelQueued=false;
function wheelUnits(delta, mode){ return mode===1 ? delta*40 : mode===2 ? delta*360 : delta*1.2; }
root.addEventListener('wheel',event=>{
  if(!event.isTrusted||lost||!drawn) return;
  event.preventDefault();
  // The frame is drawn 1:1 from the parked window, so map the point onto the drawn frame, not the whole tab.
  const box=canvas.getBoundingClientRect();
  if(box.width<=0||box.height<=0) return;
  const x=(event.clientX-box.left)/box.width, y=(event.clientY-box.top)/box.height;
  if(x<0||x>1||y<0||y>1) return;
  wheelY-=wheelUnits(event.deltaY,event.deltaMode);
  wheelX+=wheelUnits(event.deltaX,event.deltaMode);
  wheelPoint={x,y}; wheelKeys={ctrl:event.ctrlKey,shift:event.shiftKey};
  if(wheelQueued) return;
  wheelQueued=true;
  requestAnimationFrame(()=>{
    wheelQueued=false;
    const deltaY=Math.trunc(wheelY), deltaX=Math.trunc(wheelX);
    wheelY-=deltaY; wheelX-=deltaX;
    if(deltaY===0&&deltaX===0) return;
    api.postMessage({action:'wheel',surfaceToken,trusted:true,point:wheelPoint,deltaY,deltaX,ctrl:wheelKeys.ctrl,shift:wheelKeys.shift});
  });
},{passive:false});
root.addEventListener('keydown',event=>{ if(event.key==='Enter'||event.key===' '){ event.preventDefault(); activate(event); } });
let reportQueued=false;
function report(){
  if(reportQueued) return;
  reportQueued=true;
  requestAnimationFrame(()=>{
    reportQueued=false;
    const box=root.getBoundingClientRect();
    api.postMessage({action:'layout',surfaceToken,width:Math.round(box.width||window.innerWidth||0),height:Math.round(box.height||window.innerHeight||0),dpr:window.devicePixelRatio||1});
  });
}
window.addEventListener('resize',report);
new ResizeObserver(report).observe(root);
api.postMessage({action:'ready'});
report();
})();
</script></body></html>`;
}
