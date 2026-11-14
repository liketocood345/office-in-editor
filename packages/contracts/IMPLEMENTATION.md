# Implementation contract v1

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

- `health`: version, capabilities (`nativeOfficeEmbedding:false`, `nativeAutoSaveLocalRedirect:false`), bound loopback app port if active.
- `host.register`: `{hostId,kind:'vscode'|'cursor',label}` -> lease. `host.unregister`, `host.list`.
- `chat.register`: `{ownerId,hostId?,label}` -> target `{id,label,hostId,epoch,token}`; token stays in MCP `_meta`/app-only data, never user prompt/log. `chat.list` returns no tokens. `chat.unregister` requires owner/token. Only explicitly registered chats, never pretend all existing chats.
- `office.register`: `{officeId,application:'word'|'excel'|'powerpoint',pipeName,pid}`; register one Office Add-in instance, no document paths. `office.unregister`, `office.list`.
- `office.capture`: `{officeId}` -> request `captureSelection` on Office's pipe and capture immutable ref. Office reply `{application,documentId,documentName,sourcePath?,kind,text,locator,dirty,autoSaveAvailable,autoSaveEnabled}`. sourcePath is local runtime data only, never log it or ship it. Core creates UUID + SHA256 and bounded immutable reference storage.
- `reference.capture`: receives the same selection object (used for native right-click); -> reference metadata and refId.
- `reference.list` / `reference.get:{refId}` / `reference.release:{refId,confirm:true}` (only explicit user release, no original files).
- `reference.deliver`: `{refId,targetId,requestId}` -> bounded queue, no auto submit.
- `reference.deliverSelection`: `{selection,targetId,requestId}` -> deduplicated capture+queue used by Office right-click.
- `delivery.list` / `delivery.ack` only authorized app token; result never pretends a queued reference is attached.
- `office.open`: `{path}` only explicit user action; only existing local .docx/.xlsx/.pptx, use installed Microsoft app executable, no fake embedding, no global Office option changes.
- Document writes are an explicit whitelist: `word.insert_text`, `excel.set_range_values`, `excel.set_range_formulas`, `powerpoint.add_slide`, and `powerpoint.add_textbox`, plus explicit `office.save_document`. Each edit requires a caller operation id, an Office capability report, and durable recovery recorded before dispatch. Unknown outcomes are not replayed. This is not a native AutoSave local-backend redirect and not an automatic save engine. Native embedding remains unsupported.

## Office callback pipe

Office .NET Framework COM add-in hosts current-user-only `office-in-editor-office-<pid>-<instanceId>`. Same line protocol. `captureSelection`, `inspectDocument`, `prepareEdit`, `applyEdit`, `revertEdit`, and `saveDocument` run on the Office STA/UI thread with a bounded dispatcher. No global input hooks, no background selection polling, no forced application quit. Ribbon/menus may refresh core registration on explicit interaction. Broker restart must be tolerated. Edit callbacks omit absolute source paths.

## MCP + Apps

Tools: `office_list_documents`, `office_capture_selection`, `office_list_references`, `office_get_reference`, `office_inspect_document`, `office_apply_edit`, `office_get_operation`, `office_revert_edit`, `office_save_document`, `office_release_operation`, `office_capabilities`, `office_register_chat`, plus app-only acknowledgement. Read-only hints except explicitly creating references/targets or invoking a whitelisted write. Write tools are idempotent for one operation id and must not be retried under a new id after an unknown outcome. `office_register_chat` creates a target for this tool invocation's specific chat App, not a guessed active window. Resource `ui://office-in-editor/chat` HTML app, content type `text/html;profile=mcp-app`, tool metadata `_meta.ui.resourceUri`. App UI is built at apps/mcp-app and copied to core publish output `mcp-app.html`.

The HTML App receives `_meta` containing target id/token and loopback app URL via tool-result, uses supported MCP Apps `updateModelContext`, and never `sendMessage`/submit. Core loopback endpoint is current-machine only and token authenticated. Streaming connection `/app/events?targetId=...` with Bearer token via fetch, server-sent newline events `{refId,text,documentName,kind,locator}`; bounded pending references. Acknowledge via app-only tool or token-protected POST `/app/ack`. `/app/events` is push streaming, no periodic disk/network polling. Public app resource CSP permits only the exact loopback origin. Only mark accepted context when App host acknowledges updateModelContext; unknown/unsupported stays queued.

## VSIX expectations

One root VSIX for VS Code/Cursor. Standard public API only. `Open Office document` creates a truthful session/control tab, native editing inside editor is explicitly unavailable. External Microsoft Office opening is an explicit separate action. No third-party renderer masquerades as Office.

Register MCP through stable VS Code definition provider if available and official Cursor `vscode.cursor.mcp.registerServer` if present. Core executable is discovered from the stable installed user product directory, fallback packaged payload only where appropriate. Feature-check APIs; no private chat commands, IDE DB reads, UI injection or file config rewrites. Native target selection requires a chat explicitly registered via the MCP App. Show reference inbox/status, registered targets, selection capture, queue reference, explicit release/export. Existing settings and draft input preserved. Workspace trust required for local Office control; remote hosts unsupported.

## Ownership and privacy

Runtime references are bounded: 64 MiB / 512 refs, each <=1 MiB. Total private data under one finite user budget, logs minimal and capped; no document bodies or paths in diagnostic logs. Writes only on actual requests; rollback/journals bounded. Last host/Office/app leases release worker with bounded unknown-outcome recovery, never kill Office. Installation may leave Word registration pending next launch; never close an existing Office process.
