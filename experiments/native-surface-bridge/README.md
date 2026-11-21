# Native surface bridge experiment

This is an original C++20 DLL prototype, with a C ABI and a synthetic User32 test executable. It is a development experiment, not a standalone Office application. It is not included in the production VSIX or enabled by the current editor provider.

## What the component implements

- Creates real `WS_CHILD` slots on a native parent's owning UI thread. A broker in a different process cannot claim to own that parent.
- Binds each slot to an opaque tab identity and a generation. Duplicate identities, stale tokens and out-of-bounds layout requests are rejected.
- Uses client coordinates and native parenting. Parent movement needs no screen-position polling. Identical layouts do not issue another native move/show operation.
- Preserves content Z order on layout and visibility changes. An explicitly supplied, host-owned native overlay can be ordered above the slots; content cannot label itself an overlay.
- Refuses normal slot release while a child document container remains attached. Host close requires explicit release of its slots.
- Bounds the number of hosts and slots. No timers, worker threads, network, file access, registry writes, capture, input injection, global hooks or Office automation exist in the DLL.

The prototype uses `SetWindowPos` to lay out its own slots and to order an explicitly registered host-owned overlay. It does not reparent an Office main frame, change foreign window styles, search by application title, or raise Office windows globally.

## The unresolved browser boundary

The supported overlay mode is **native child windows**. This implementation does not import a Windows window into the DOM, and does not turn an existing HTML menu into a native overlay. Passing the mode flag is a caller contract, not independent proof of browser composition support.

A real browser/IDE adapter must supply the native parent on its UI thread and integrate its menus, command palette, focus, DPI and shutdown behavior. Loading a DLL in the VS Code extension-host process does not put it in Electron's UI process. The current standard VSIX has no implemented provider for this contract.

Consequently, the synthetic tests prove native ownership and lifecycle behavior. They do **not** prove Edge/Chrome/Cursor composition, Office input/IME, Word editing/save, process-crash recovery or an end-to-end latency below 40 ms. Unexpected host destruction can still destroy its child HWNDs; Office session recovery and close arbitration must remain in the independent Core before any production attachment.

## Build and verify

Requires Windows and installed Visual C++ x64 Build Tools. From the repository:

```powershell
./experiments/native-surface-bridge/build.ps1
./.local/build/native-surface-bridge/native-surface-smoke.exe
```

Outputs are ignored local build files: `oie-native-surface.dll`, its import library and `native-surface-smoke.exe`. The executable creates only synthetic windows, never opens a browser or Office, and never touches documents. The tests check native HWND state and sibling hit testing; they do not inspect compositor pixels.

Keep the DLL loaded until every host has been explicitly closed. The prototype is not a hostile-input boundary or a production browser SDK. The C ABI is defined in `bridge.h`; it must not be exposed as an unauthenticated HWND-taking HTTP API.

## Intended integration shape

```text
VSIX (TypeScript): document tab lifecycle
       |
Browser/IDE native provider (host-side integration still required)
       |
Native bridge (C++ / C ABI): slot + native overlay ownership
       |
Shared Core (C#): own OLE container + document/save/MCP state
       |
Installed Office + the existing Office add-in (C#)
```

The DLL is a concrete prototype of the native part, not evidence that the missing browser/IDE provider already exists. A localhost server can carry control messages; it cannot supply that provider by itself.
