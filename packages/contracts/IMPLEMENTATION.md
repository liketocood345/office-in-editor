# Implementation contract v1

Review snapshot: 2026-11-15, source version 0.1.4. This file records existing v1 wire shapes and intended behavior; it is not an end-to-end compatibility certificate. Current implementation gaps and the additional normative lifecycle requirements are documented in [the architecture audit](../../docs/ARCHITECTURE_AUDIT.md) and [the application matrix](../../docs/OFFICE_COMPATIBILITY.md). New lease/generation/routing fields described there are not yet implemented by v1 callers.

This source checkout is separate from private planning. Never copy local absolute paths, user identity, chat transcripts, credentials, generated state, binaries or logs into tracked files. DSH is deferred. Preserve the existing BSD-3-Clause license.

## Shared runtime

Core is a .NET 8 Windows private worker. Publish executable name: OfficeInEditor.Core.exe.
- `--endpoint`: print one JSON object `{pipeName, protocolVersion:1, version}` and exit with no mutations.
- `--ensure-broker`: ensure one per-user broker via named mutex and return endpoint JSON after bounded readiness.
- `--broker`: one current-user-only named pipe; no console/UI/auto-start, no idle disk polling.
- `--mcp`: passive newline JSON-RPC MCP stdio bridge. Initialize/listing must never launch Office or create references.

Pipe name: `office-in-editor-v1-` + first 16 lowercase SHA256 hex chars of current Windows user SID UTF-8. Every pipe request is a single UTF-8 JSON line: `{v:1,id,method,params}`. Response `{id,ok:true,result}` or `{id,ok:false,error:{code,message}}`. One request per connection. Maximum line 2 MiB, bounded timeouts.

Core owned data root is resolved at runtime under the current user's LocalApplicationData / OfficeInEditor. Installed components are under LocalApplicationData / Programs / OfficeInEditor / current / core or office. No source machine paths.

## Core methods

- `health`: version, advertised capabilities (`nativeOfficeEmbedding:true`, `nativeAutoSaveLocalRedirect:false`), bound loopback app port if active. The first flag currently means the experimental path is enabled, not that the installed application/format has passed embedding or source-file persistence. Future capability reports must distinguish candidate, supported, and verified per application/format/mode. The tab does not locate document content.
- `host.register`: `{hostId,kind:'vscode'|'cursor',label}` -> lease. `host.unregister`, `host.list`.
- `chat.register`: `{ownerId,hostId?,label}` -> target `{id,label,hostId,epoch,token}`; token stays in MCP `_meta`/app-only data, never user prompt/log. Marks that chat as the owner's current target and the latest-activity default. `chat.list` returns no tokens. `chat.unregister` requires owner/token. Only explicitly registered chats, never pretend all existing chats.
- `chat.defaultTarget`: no params. Returns the most recently active registered chat `{id,label,hostId,epoch,ownerId,online}` or `{}`. Never errors for an empty directory.
- `routing.touch`: `{ownerId,targetId?}`. Records MCP/chat activity. If `targetId` is omitted, uses that ownerId's current chat. Pushes `bindChatTarget` to Office instances whose pid is the active OLE in-place server. Push failure does not fail the MCP tool. Missing target: no prompt, no bind.
- `office.register`: `{officeId,application:'word'|'excel'|'powerpoint',pipeName,pid}`; register one Office Add-in instance, no document paths. `office.unregister`, `office.list`.
- `office.capture`: `{officeId}` -> request `captureSelection` on Office's pipe and capture immutable ref. Office reply `{application,documentId,documentName,sourcePath?,kind,text,locator,dirty,autoSaveAvailable,autoSaveEnabled}`. sourcePath is local runtime data only, never log it or ship it. Core creates UUID + SHA256 and bounded immutable reference storage.
- `reference.capture`: receives the same selection object (used for native right-click); -> reference metadata and refId.
- `reference.list` / `reference.get:{refId}` / `reference.release:{refId,confirm:true}` (only explicit user release, no original files).
- `reference.deliver`: `{refId,targetId,requestId}` -> bounded queue, no auto submit.
- `reference.deliverSelection`: `{selection,targetId,requestId}` -> deduplicated capture+queue used by Office right-click.
- `delivery.list` / `delivery.ack` only authorized app token; result never pretends a queued reference is attached.
- `office.open`: `{path,mode,fallback?,width?,height?,sessionId?,surfaceToken?}`. Editor-tab opens send `fallback:false`, the surface token (`oie-surface-` + 32 hex), and webview size. Embedded mode requires the token and at least 32×32. NativeOfficeLauncher passes `bindSource:false` and `OleCreateFromFile` on a **staging copy** under `%TEMP%\OfficeInEditor\embed` so the user's file is not truncated during create. The Core container is an owned popup whose screen rect is the token/beacon union (not `WS_CHILD` of Chromium; that attaches input queues). Place stays hidden until `open_ok`. `OleSite.SaveObject` must `IPersistFile.Save` the original `DocumentPath` even when unbound. Unbound close uses `OLECLOSE_NOSAVE`. `inplace:true` still does not certify CustomDocument dirty/cancel-close. Failure aborts this session's HWND; NOSAVE only for proven unmodified create failures. Embedded mode must never Process.Start a standalone application, including on failure. `mode=external` is only the explicit independent-window command. `office.route.get` may include a short lastOleFailure tag without paths. An active session prevents idle Core shutdown; each kind STA must pump messages **including during OleCreate/DoVerb** (`IMessageFilter`). Chromium Webview is not an OLE site. HideNewFrames must skip dialogs. OneNote without IOleDocument fails fast. Verified Excel/PPT embed, viewport borders, and cancel-close remain release gates.
- `office.detach_window`: tears down the OLE session (OleClose + DestroyWindow). Hiding while switching tabs uses `office.layout_window` with `visible:false`.
- `office.describe_window` / `office.attach_window` / `office.layout_window` / `office.detach_window`: when an OLE session is active, these place or hide Core's owned container over the custom-editor `#ole-root` token rect; otherwise they forward to the Office add-in. Attach/layout take `surfaceToken` plus inner `width,height`. Geometry comes from four-corner beacons / UIA token; never `window.screenX/Y` and never a workbench-sized Chromium HWND. Only the Core OLE container HWND is ours; never `SetParent` `OpusApp` / `WINWORD`. Switch-away is hide-only (`visible:false`); do not UIDeactivate. The tab does not capture selection or edit the document.
- Document writes are an explicit whitelist: `word.insert_text`, `excel.set_range_values`, `excel.set_range_formulas`, `powerpoint.add_slide`, and `powerpoint.add_textbox`, plus explicit `office.save_document`. Each edit requires a caller operation id, an Office capability report, and durable recovery recorded before dispatch. Unknown outcomes are not replayed. This is not a native AutoSave local-backend redirect and not an automatic save engine. Agent edits use MCP, not the editor tab.

## Office callback pipe

Office .NET Framework COM add-in hosts current-user-only `office-in-editor-office-<pid>-<instanceId>`. Same line protocol. `captureSelection`, `inspectDocument`, `prepareEdit`, `applyEdit`, `revertEdit`, `saveDocument`, and `bindChatTarget` run on the Office STA/UI thread with a bounded dispatcher. `bindChatTarget` `{id,label,epoch}` silently selects that chat for Add to Chat (embedded MCP rebind); empty/unknown id is a no-op. No global input hooks, no background selection polling, no forced application quit. Ribbon/menus may refresh core registration on explicit interaction. Broker restart must be tolerated. Edit callbacks omit absolute source paths. Ribbon and the editor tab do not show product-explanation copy.

## MCP + Apps

Tools: `office_list_documents`, `office_capture_selection`, `office_list_references`, `office_get_reference`, `office_inspect_document`, `office_apply_edit`, `office_get_operation`, `office_revert_edit`, `office_save_document`, `office_release_operation`, `office_capabilities`, `office_register_chat`, plus app-only acknowledgement. Read-only hints except explicitly creating references/targets or invoking a whitelisted write. Write tools are idempotent for one operation id and must not be retried under a new id after an unknown outcome. `office_register_chat` creates a target for this tool invocation's specific chat App, not a guessed active window. Resource `ui://office-in-editor/chat` HTML app, content type `text/html;profile=mcp-app`, tool metadata `_meta.ui.resourceUri`. App UI is built at apps/mcp-app and copied to core publish output `mcp-app.html`.

Every `tools/call` (including failed ones once a chat is attributable) invokes `routing.touch` so an in-IDE OLE Office instance rebinds Add to Chat to that chat. Attribution: `_meta.officeInEditor.targetId`, else this MCP stdio ownerId's current chat. Standalone Office does not receive the push; it uses `chat.defaultTarget` only when the user has not picked a chat by hand.

The HTML App receives `_meta` containing target id/token and loopback app URL via tool-result, uses supported MCP Apps `updateModelContext`, and never `sendMessage`/submit. Core loopback endpoint is current-machine only and token authenticated. Streaming connection `/app/events?targetId=...` with Bearer token via fetch, server-sent newline events `{refId,text,documentName,kind,locator}`; bounded pending references. Acknowledge via app-only tool or token-protected POST `/app/ack`. `/app/events` is push streaming, no periodic disk/network polling. Public app resource CSP permits only the exact loopback origin. Only mark accepted context when App host acknowledges updateModelContext; unknown/unsupported stays queued.

## VSIX expectations

One root VSIX for VS Code/Cursor, with the application candidates declared in package.json. The tab provides a full-size `#ole-root`, an `oie-surface-` token, and four 1px black corner beacons named `{token}:tl|tr|bl|br`. Layout messages send CSS width/height only (`clientOnly`); they must not send `window.screenX/Y` as an overlay origin. Core locates the tab by those beacons (UIA Name union of at least two corners), maps the token onto the editor webview HWND, and `SetParent`s its own container as a `WS_CHILD` of that surface. It must reject a workbench-sized Chromium HWND. The OLE host fills the tab content box; Core does not reserve a Word status-bar strip or paint a themed tab chrome. Core owns the OLE container and document viewport. The add-in must not position an independent Office main window as a tab fallback. Selection, Add to Chat and agent edits remain in the add-in/MCP. Use stable public extension APIs; the native surface experiment is not a public HWND slot supplied by Custom Editor API and needs separate identity/clipping/input compatibility validation. No third-party renderer masquerades as Office.

Register MCP through stable VS Code definition provider if available and official Cursor `vscode.cursor.mcp.registerServer` if present. Core executable is discovered from the stable installed user product directory, fallback packaged payload only where appropriate. Feature-check APIs; no private chat commands, IDE DB reads, UI injection or file config rewrites. Native target selection requires a chat explicitly registered via the MCP App. Show reference inbox/status, registered targets, selection capture, queue reference, explicit release/export. Existing settings and draft input preserved. Workspace trust required for local Office control; remote hosts unsupported.

## Ownership and privacy

Runtime references are bounded: 64 MiB / 512 refs, each <=1 MiB. Total private data under one finite user budget, logs minimal and capped; no document bodies or paths in diagnostic logs. Writes only on actual requests; rollback/journals bounded. Last host/Office/app leases release worker with bounded unknown-outcome recovery, never kill Office. Installation may leave Word registration pending next launch; never close an existing Office process.

## Focus-hybrid route

Contract added 2026-11-21. Core serves these methods since 2026-11-21 (see "Core implementation" below). `health.capabilities.focusHybrid` is true when the capture helper `oie-focus-capture.exe` sits next to Core, which the payload build puts there. Hosts use the OLE route above while the flag is false. Sources of truth are `packages/shared-core/FocusContract.cs` and `packages/adapter-sdk/src/focus.ts`, kept in sync by the unit tests on both sides.

The native Office main window is visible over the tab only while it holds the foreground (`editing`). Otherwise it is parked and the tab shows captured frames. Core owns the state machine, the window movement and the original-position recovery. The tab only reports clicks, size and visibility, and draws frames. The tab never holds an HWND and never starts native processes.

**States and causes.** States: `parked | activating | editing | releasing | lost`. Causes: `open | click | foreground | layout | timer | policy | dialog | close | recovery`.
- Any time the parked Office window (or its own dialog) owns the foreground, Core moves it to `editing` with cause `foreground`. Core never leaves it focused off-screen (D1).
- `lost` means the Office window or process is gone, and the session only accepts `focus.close`.

**Control methods** (v1 request pipe, host lease fields as for `office.open`):
- `focus.open` `{path, surfaceToken, editorPid, cssWidth, cssHeight, pixelRatio, pixelPolicy?}` → `{sessionId, application, shared, state, stream:{pipeName, token}}`.
  - `surfaceToken` is `oie-surface-` + 32 hex. It is also the UIA Name of the preview element, which is how Core finds the physical bounds.
  - `editorPid` is recorded with the surface; Core finds the IDE top-level window from the surface token. Sizes are CSS px from 32 to 16384, and `pixelRatio` is 0.5–8.
  - `pixelPolicy` is `always | switch | mouse` and defaults to `always`.
  - Sessions are deduplicated by file identity: a second tab on the same file joins with `shared:true`. The path is never echoed, logged or put in events.
- `focus.activate` `{sessionId, surfaceToken, activationId, trusted:true, point?:{x,y}}` → `{state, activationId}`.
  - Missing or false `trusted` fails with `untrusted_activation`.
  - `point` is normalized to the surface (0–1). Core does not replay the click: S05 forbids input injection, so the next click lands in Office. Placing the caret through the object model is an open P5 option in ADR 0003.
- `focus.layout` `{sessionId, surfaceToken, visible, cssWidth, cssHeight, pixelRatio}` → `{state}`.
  - Core re-reads the UIA bounds itself. The tab never sends screen coordinates.
  - `visible:false` while editing returns to `parked`.
- `focus.policy` `{sessionId, pixelPolicy}` → `{pixelPolicy}`.
- `focus.close` `{sessionId, surfaceToken, decision?: save|discard}` → `{closed, dirty, remainingSurfaces}`.
  - Closing a surface that is not the last one detaches it.
  - Closing the last surface over unsaved changes without a decision returns `{closed:false, dirty:true}` and keeps everything as is. The host asks the user, then calls again with a decision, or keeps the tab on cancel.
  - Core restores the original window position before closing. It never posts `WM_CLOSE` over unsaved changes.
- `focus.wheel` `{sessionId, surfaceToken, trusted:true, point:{x,y}, deltaY, deltaX, ctrl?, shift?}` → `{forwarded}`.
  - Wheel over a parked, visible tab scrolls the parked Office window in place; Office is not activated and stays parked, and the new pixels come back as frames. This is the one user-approved S05 exception (ADR 0003): wheel only, never keyboard or clicks.
  - Win32 conventions: `deltaY` > 0 scrolls up, `deltaX` > 0 scrolls right, 120 is one notch, each bounded to ±2400; both 0 is rejected. Missing or false `trusted` fails with `untrusted_activation`.
  - `point` is normalized to the drawn frame, which maps 1:1 onto the parked window. Core posts `WM_MOUSEWHEEL` / `WM_MOUSEHWHEEL` to the deepest content child at that point (Word `_WwG`, Excel `EXCEL7`, PowerPoint `mdiClass`); the frame window itself ignores them. `forwarded` is false while editing, for a hidden tab, or for a tab Office is not parked for.
  - The tab merges steps that arrive while one is in flight and drops steps made before Office is parked.
- `focus.status` `{sessionId}` → `{state, application, surfaces, dirty, activationId, parkId, pixelPolicy}`.

**Stream pipe.** `stream.pipeName` is the broker pipe name + `-focus-` + sessionId. It has the same current-user-only ACL and is one connection per surface.
- The client first writes one line: `{"v":1,"token":<stream.token>,"surfaceToken":...}`. A bad token closes the connection.
- Server messages are `u32 LE length of (kind + body)`, then a `u8 kind`, then the body:
  - kind 1 is an event, encoded as UTF-8 JSON.
  - kind 2 is a frame: `u16 LE header length`, then header JSON `{frameSeq, parkId, width, height, format}`, then the image bytes. `format` is `jpeg` (the default stream) or `png`.
- A message may not exceed 32 MiB and a frame header 4 KiB. Core drops an oversize frame and sends a later one; it never truncates.
- Events:
  - `{type:"state", state, cause, activationId, parkId}`
  - `{type:"notice", code}`, where `code` is one of `modal_blocking | modal_cleared | alignment_unconfirmed | activation_denied | office_exited | capture_unavailable`
  - `{type:"closed", reason: closed|office_exited|superseded}`
- `frameSeq` increases per session. The first frame after a park carries the new `parkId`. Frames contain pixels only: no titles, paths or text.

**Diagnostics.** Same correlation fields as the experiment (`activationId`, `parkId`, `frameSeq`). They hold state, timings and sizes only. Latency up to 500 ms (P95) is within target and is not an optimization goal.

**Core implementation.** The state machine is `FocusStateMachine.cs`, ported from the earlier focus-hybrid lab prototype. The other parts are:
- `FocusDispatcher.cs`: one STA thread for all sessions, with the foreground hook and a 100 ms tick.
- `Win32FocusWindowOps.cs`: window moves, focus and UIA bounds, plus `FocusRecovery`.
- `FocusStream.cs`: the stream pipe.
- `CoreFocus.cs`: the `focus.*` registry.

Behaviour:
- **Before the first park** the state is `releasing` with cause `open`. Core waits 1.3 s for the tab layout and any Office startup dialog. Then it parks Office, or docks it with cause `open` when Office already holds the foreground.
- **Watchdog.** The 100 ms tick catches foreground changes that sent no event: a missed raise is adopted (D1) and a missed loss is parked (cause `timer`). An unconfirmed park reports `releasing` with `alignment_unconfirmed`, stops capture and retries every 500 ms.
- **Modal dialogs (D3).** A modal dialog while parked is reported as `modal_blocking`. A click on the tab then shows the dialog over the tab (state `editing`, cause `dialog`) instead of off-screen.
- **Several tabs on one file.** Office docks into the tab clicked last. The other tabs only draw frames.
- **Host goes away.** Its surfaces are dropped. A session that no tab shows any more puts Office back where it was and leaves it open.
- **Crash recovery (D2).** Each moved window's original placement (handle and rectangle, no path) is kept in `focus-recovery.json` in the Core data root. The file is written only when a session starts or ends. On start, Core moves back any window a crashed Core left parked.
- **Close.** If the add-in cannot be reached and there is no decision, Core restores the window and leaves Office open instead of guessing whether there are unsaved changes. In that case `focus.status` reports `dirty: null`.
- **Lab switch.** `OIE_FOCUS_HYBRID=1` forces `focusHybrid` on, to test window control without frames. `OIE_FOCUS_CAPTURE=<full path>` points Core at a helper built elsewhere.
- **Capture helper.** `packages/focus-capture` builds `oie-focus-capture.exe`: Windows.Graphics.Capture plus PNG encoding only, one process per session, started on the first park. `FocusCaptureHost.cs` drives it.
  - Arguments are `<hwnd> <pid>`. stdin takes `prepare`, `publish <parkId>`, `stop` and `exit`. Closing stdin (including Core dying) ends the helper.
  - stdout uses the stream envelope. Kind 1 is a status `{stage}`. Kind 2 is a frame with header `{parkId, width, height, format:"png"}`.
  - The helper sends frames only while the window is off-screen and only when pixels change, at the parked window's visible size, JPEG quality 0.9 (about 8–12 ms to encode a 1200×760 frame). The first frame after each `publish`, or after resuming from 0 fps, is always sent.
  - Frame rate: Core sends `fps 30` while the tab is visible and the editor is not minimized, and `fps 0` (paused, capture kept warm) otherwise, so only the tab the user can see is streamed at 30 fps. stdin also takes `fps <n>` and `format jpeg|png`.
  - Office pauses its own animations (animated GIFs) whenever it is not the foreground app, wherever the window is: measured on the lab, PowerPoint updated about 11 times/s while active (on screen, off screen or 1 px on screen) and never while inactive, even fully visible on screen. Synthetic `WM_ACTIVATEAPP`, `WM_NCACTIVATE` and `WM_ACTIVATE` did not resume them. A parked preview is therefore still for such content; this matches what Office shows behind another window and is not a forwarding fault.
  - Core assigns `frameSeq` and forwards a frame only while the session is `parked` with the same `parkId`.
  - A helper that fails to start, reports a capture failure, or crashes 3 times within a minute is given up on, and the session reports `capture_unavailable` once.

**Host (VSIX).** Code is in `packages/adapters/vsix/src`: `focus-editor.ts` runs each tab, `focus-stream.ts` is the pipe client and `focus-ui.ts` is the page.
- **Route.** `officeInEditor.hostRoute` is `auto` (default), `focus` or `ole`. With `auto`, Word, Excel and PowerPoint files use this route when `health.capabilities.focusHybrid` is true. Other formats always use the OLE route. `officeInEditor.pixelPolicy` sets the pixel policy.
- **Page.** The page reuses the OLE page's markers: the root's accessible name is the surface token, plus the four corner beacons. Frames are drawn on a canvas.
  - A left click, Enter or Space with `event.isTrusted` sends `activate` with the click point normalized to the tab.
  - Size is sent from a `ResizeObserver`. Visibility comes from the panel.
- **Frames.** The extension forwards only the newest frame, and drops frames from an older park.
- **Status text.** Notices map to text on the tab: modal dialog, capture unavailable, activation denied, Office exited.
- **Close.** VS Code can't cancel closing a webview tab, so a dirty last tab shows a modal Save / Don't Save / Cancel prompt.
  - Save or Don't Save calls `focus.close` again with the decision.
  - Cancel reopens the file in a new tab, which joins the session. The old surface then detaches on retry.
- **Shutdown.** The extension does not close sessions itself. Unregistering the host makes Core put Office back where it was and leave it open.
- **Also changed.** The SDK stream decoder now keeps the chunks of a large frame apart until the frame is complete, so a multi-MB frame is copied once, not once per chunk.

Office add-in methods Core calls (current-user pipe, verified Office pid), implemented in `packages/office-addin/FocusDocuments.cs`:
- `resolveDocumentWindow` `{path}` → `{found:true, hwnd, documentId}` or `{found:false, reason?}`.
  - The path must be local and absolute. It matches an open, non-embedded document by full path, or by volume serial and file index. It is never echoed back.
  - `hwnd` comes from the object model (D4): Word `ActiveWindow.Hwnd`, Excel `Windows(1).Hwnd`, PowerPoint `Windows(1).HWND` when the build has it.
  - Otherwise, for PowerPoint, the add-in uses the one `PPTFrameClass` frame of this process whose title is the window caption. Two matches give `found:false`.
  - The add-in normalizes the handle to the top-level frame and checks the process and class. Core checks them again.
  - If no add-in has the file, Core starts Office once and polls for up to 45 s.
- `documentState` `{documentId}` → `{documentId, dirty}`. `dirty` is `!Saved`, and is true when it can't be read.
- `closeDocument` `{documentId, save}` → `{documentId, closed, saved}`.
  - Word calls `Close(wdSaveChanges | wdDoNotSaveChanges)` and Excel calls `Close(save)`.
  - PowerPoint saves first when asked, then calls `Close()`, which never prompts. `Saved` is never written (S06).
  - A failure returns `SAVE_FAILED` or `CLOSE_FAILED`, and Office keeps the document open.
  - The Office application itself is not quit when its last document closes. That is a follow-up decision.

## Required lifecycle extension (not implemented v1 wire fields)

- Authenticate host lease/epoch, editorId/generation, sessionId and document/source identity at queue admission and STA execution. Expired owners cannot layout/save/close another host's session. Bound cancel-before-open tombstones and resource counts.
- Resolve routes per document/session with owner/token/epoch checks and a monotonic routeVersion. Broadcast PID-wide routing and successful-cache-push assumptions are insufficient. Resolve the authoritative target again at Add to Chat submission.
- Use one document SaveCoordinator for manual/MCP/edit-triggered/timed/close saves. The two plugin autosave switches remain independent, default off; one bounded in-memory dirty heartbeat is allowed only while enabled. No dirty change means no Save, no close/reopen, no source-file polling loop. This does not redirect native Office AutoSave.
- Implement editable CustomDocument dirty/save/revert/backup integration and cancelable close ownership. Preserve a recovery path on canceled/failed/unknown close before releasing the document UI. Embedded copies cannot be advertised as safe original-file editing without verified persistence.
- Use native frame client coordinates, committed OLE border widths, document viewport and clipping as separate layout state. Account for NULL versus zero border widths and DPI units; do not assign the entire frame to the document after accepting toolbar space.
- Prove token-to-host/native-surface ownership; size similarity is not identity. Hide/close only windows proven to belong to the specific document/session, never global newly appeared windows or unknown dialogs.
- Default diagnostics to bounded memory; export only on explicit request. Remove production fixed-machine debug paths and fixed ingest endpoints. Set global kind/session/STA/queue budgets and safe idle reclamation.
- Measure 30-second editable readiness separately from the required 60-second tab terminal deadline; the current source has not demonstrated that all paths meet either limit. Start at tab/openId creation, including connection, handshake and queue time; transport timeouts cannot silently extend the deadline. An unresponsive COM call yields an unknown outcome, not permission to kill Office or replay writes. Retain recoverable dirty/unknown objects rather than unconditionally aborting them at the UI deadline.
