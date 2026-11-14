# Private shared core

This package implements the Windows-only .NET 8 broker and MCP stdio bridge. It is an internal plugin component, with no product window, tray, service, scheduled task, automatic startup, or Office installer. It never kills Office.

## Build and verify

```powershell
dotnet build packages/shared-core/OfficeInEditor.Core.csproj
dotnet run --project packages/shared-core/Tests/OfficeInEditor.Core.Tests.csproj
pwsh -NoProfile -File packages/shared-core/Tests/Smoke.ps1
```

The integration executable uses its own verified temporary directory and a fake Office callback pipe. It does not launch or modify Microsoft Office. The smoke script exercises the real per-user broker and MCP process, creates an ephemeral chat target, then releases it normally; it does not create a reference. Other valid plugin owners are retained.

Publishing uses `OfficeInEditor.Core.exe`. The parent build copies its bundled MCP App to the publish directory as `mcp-app.html`. A self-contained Windows publish must include the ASP.NET Core framework used by the loopback server; do not copy just the apphost executable from a framework-dependent build.

## Commands

| Command | Behavior |
| --- | --- |
| `--endpoint` | Print `{pipeName,protocolVersion,version}` and exit without creating application files or starting a broker. |
| `--ensure-broker` | Connect to the existing per-user instance or start the private worker, with bounded readiness. |
| `--broker` | Hold the per-user mutex, serve the current-user-only pipe and ephemeral loopback App endpoint. |
| `--mcp` | Run newline JSON-RPC MCP on inherited stdio. Initialization acquires a process-owned broker lease; initialization and discovery never launch Office or capture references. |
| `--shutdown-if-idle` | Contact an existing worker only. Exit 0 when absent or safely shutting down; exit 3 when valid owners or work remain. Never force process termination. |

Pipe identity is derived from a hash of the Windows SID. A named mutex prevents multiple mutable authorities. Pipe clients and servers use `CurrentUserOnly`; registration owners are observed by PID plus process creation identity and exit events. Every connection carries one bounded UTF-8 JSON line using the shared contract, with a 2 MiB message cap and bounded request deadlines. No disk polling is used for discovery or ownership.

## Implemented protocol

`health` reports version, explicit capabilities, `appOrigin`, `appPort`, bounded storage status and aggregate owner counts. Lists are direct arrays. Capture results are reference metadata with `refId` (and `id` alias), application/document identity, display name, kind, creation time, text length, SHA-256 and native save observations. Unknown `autoSaveAvailable` remains null; `autoSaveState` is preserved when supplied.

The broker implements host/chat/Office registration and lists; Office callback capture; immutable reference capture/list/get/explicit release; deduplicated delivery and selection-plus-delivery; token-authorized delivery listing/acknowledgement; and explicit external opening of existing local DOCX/XLSX/PPTX through the registered Microsoft Office application.

`office.route.get` always reports standalone mode with `native_embedding_unavailable`. Native editor embedding and native AutoSave redirection remain unsupported. Explicit document writes are limited to the declared recoverable whitelist and a manual Save; there is no automatic save engine, macro execution, arbitrary COM reflection, or file deletion tool.

Host registration returns a lease required by `host.unregister`. Office registration must originate from the declared live Office process, and its callback pipe must match `office-in-editor-office-<pid>-<instanceId>`. COM objects stay in the Office add-in; capture is delegated to `captureSelection` through its pipe. The core does not implement or emulate Office STA dispatch.

Runtime MCP ownership uses `runtime.acquireOwner`/`runtime.releaseOwner`; the stdio proxy manages this automatically. `runtime.shutdownIfIdle` returns `busy` or `shutting_down`. The final owner release starts a single idle deadline, not an I/O polling loop. Releasing one host never ends unrelated Office/MCP ownership.

## MCP App

Tools are `office_list_documents`, `office_capture_selection`, `office_list_references`, `office_get_reference`, `office_capabilities`, `office_register_chat`, and app-only `office_app_ack`. Resources expose `ui://office-in-editor/chat` as `text/html;profile=mcp-app`, with CSP permitting only the active loopback origin. Tool and legacy App resource metadata are both included.

`office_register_chat` registers this particular tool invocation's App target. It does not enumerate or guess an IDE's other chats. Its bearer token and App origin appear under `_meta.officeInEditor` only, never model-facing content or structured content.

The App connects using authenticated `fetch` to `/app/events?targetId=...`. Responses are push SSE `data:` records containing request/ref identity, captured text, display name, kind and locator. Tokens are forbidden in URL query parameters. Exact Host and loopback address checks, bounded bodies, controlled CORS preflight and private-network headers are enforced. Authorization remains required even for an opaque/null App origin. No request logging or periodic stream/disk heartbeats are enabled.

POST `/app/ack` accepts bearer authentication and `{targetId,requestId,refId,status}`. Only an authenticated `attached` acknowledgement may mark accepted context; `unsupported` remains queued and `failed` remains a failure. A disconnected App invalidates that target; no automatic target substitution or retry is performed. No endpoint submits or sends a chat message.

## Storage and limitations

References resolve under the user's private runtime data root. They are UUID-named immutable files, each at most 1 MiB, with at most 512 references / 64 MiB in total. Disk writes happen only on capture or confirmed release. Runtime diagnostic logs are not written. Source paths and captured content are never diagnostics; MCP reference reads and App events omit absolute source paths.

Startup reads only the bounded reference directory. Unexpected entries, linked directories/files, corrupt records or excess capacity block new capture instead of triggering speculative cleanup. Explicit release removes only the recorded reference, never its source document. In-use or uncertain references are never deleted by a background collector.

Delivery queues, targets, process leases and request tombstones are bounded memory state: 32 host/Office/chat/MCP-owner registrations, 512 deliveries, 32 queued items per target, and 1024 deduplication entries. References survive a broker restart; ephemeral chat routes and delivery receipts do not. After restart clients must explicitly register targets again and inspect the reference inbox; a lost receipt must not be reported as attached. No document write request exists in this version, so worker shutdown cannot abandon a core-owned COM write.

The main capability gaps remain explicit. Successful MCP selection capture or a truthful external Office launch does not establish native embedding or native AutoSave redirection.
