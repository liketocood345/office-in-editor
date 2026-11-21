let objectLoadFailed = false;
let objectCspBlocked = false;
window.addEventListener('error', event => { if (event.target?.id === 'office') objectLoadFailed = true; }, true);
window.addEventListener('securitypolicyviolation', event => {
  if (event.effectiveDirective === 'object-src') objectCspBlocked = true;
});

window.addEventListener('load', () => {
  // OBJECT fallback layout can settle after deferred scripts have executed.
  // Defer the hit test until the target has had a chance to paint it.
  setTimeout(reportAfterLayout, 180);
}, { once: true });

function reportAfterLayout() {
  const object = document.getElementById('office');
  const fallback = document.getElementById('fallback');
  const box = fallback.getBoundingClientRect();
  const hit = box.width && box.height ? document.elementFromPoint(box.left + 3, box.top + 3) : null;
  let nativeControl = false;
  let nativeControlReadFailed = false;
  try { nativeControl = object.object != null; }
  catch { nativeControlReadFailed = true; }
  fetch('/report', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      activeX: typeof window.ActiveXObject !== 'undefined',
      nativeControl, nativeControlReadFailed,
      mimeHandler: Boolean(navigator.mimeTypes.namedItem(object.type)),
      fallbackVisible: Boolean(box.width && box.height),
      fallbackHit: fallback.contains(hit), visible: !document.hidden,
      objectLoadFailed, objectCspBlocked, documentComplete: document.readyState === 'complete',
      reload: performance.getEntriesByType('navigation')[0]?.type === 'reload'
    })
  }).then(response => {
    if (response.ok && !sessionStorage.getItem('office-browser-reloaded')) {
      sessionStorage.setItem('office-browser-reloaded', '1');
      location.reload();
    }
  }).catch(() => {});
}
