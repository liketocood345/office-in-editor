# Native Office add-in

This package builds one .NET Framework 4.8 managed COM add-in for installed Windows Word, Excel and PowerPoint. It contains no Microsoft Office binaries, renderer or PIAs. Office APIs are called through IDispatch on Office's existing UI/STA thread.

## Implemented behavior

- A native **Office in Editor** Ribbon tab has **Refresh chats**, a menu of explicitly registered chat targets, and **Add to Chat**. Target labels show a short target ID and, when present, a short host ID to distinguish identical titles; routing always uses the full target ID, never the title, and no private token is displayed. Refresh is an explicit action that can start the installed private Core worker. Normal Office startup only tries to register with an already running worker and never starts one or writes a document.
- No chat is selected initially. The native context-menu command stays visible and disabled until a registered target is selected. A stale target is rejected by Core; refreshing clears a target whose ID/epoch is no longer registered.
- Clicking Add to Chat captures a bounded immutable selection before dispatching IPC, then calls `reference.deliverSelection` with one request ID. It queues context; it never submits a chat message or claims that a host has accepted context. Failures and unknown outcomes are shown in the Ribbon status, with no automatic replay or modal dialog.
- A current-user ACL-protected, network-denied callback pipe accepts `captureSelection` and the explicit edit callbacks `inspectDocument`, `prepareEdit`, `applyEdit`, `revertEdit`, and `saveDocument`. It dispatches them to the Office STA. Requests are bounded to 12 seconds and 2 MiB. It has no global hooks, SendKeys, clipboard access, selection polling, private IDE APIs, or Office termination.
- `autoSaveEnabled` is a read-only observation of `AutoSaveOn`. `autoSaveAvailable` is `true` only when enabled, otherwise `null`: a false getter cannot distinguish a disabled toggle from an available toggle that is off. `autoSaveState` distinguishes `enabled`, `off_or_unavailable` and `unsupported`. No setter is used to probe support.
- Native Office-in-editor embedding and redirecting native AutoSave to ordinary local files remain **unsupported**. This add-in provides no timer, debounce, AutoRecover replacement or save polyfill. The initial version provides standalone Office capture and collaboration only; it does not invent an embedded/Cursor origin or silently route to an active window.

## Selection and menu coverage

| Application | Implemented selection | Implemented native context IDs |
| --- | --- | --- |
| Word | Nonempty text/range, including selected table text; story/start/end locator | `ContextMenuText`, `ContextMenuHeading`, `ContextMenuList`, `ContextMenuTableCell`, `ContextMenuTextTable` |
| Excel | Worksheet Range selections; at most 5,000 cells and 32 areas, values/formulas/addresses | `ContextMenuCell`, `ContextMenuListRange` |
| PowerPoint | Selected text, selected shapes, or up to ten selected slides / 100 shapes; slide/shape IDs | `ContextMenuTextEdit`, `ContextMenuShape`, `ContextMenuPicture`, `ContextMenuTable`, `ContextMenuThumbnail`, `ContextMenuSlideSorter` |

The menu IDs were checked against [Microsoft's Office Fluent command identifier lists](https://github.com/OfficeDev/office-fluent-ui-command-identifiers/tree/main/Microsoft%20365/Current%20Channel). Each command requests `insertBeforeMso="Cut"`. This is relative ordering, not a guarantee of absolute first position against every Office version or other add-in. Unsupported object selections are rejected rather than substituted with ActiveCell or an entire document. Non-text PowerPoint shapes have a semantic ID/name/type representation; no screenshot/export or temporary file is generated. Group contents, charts, equations and full table structure are not expanded by this version.

Selections larger than 1 MiB of serialized UTF-8 JSON are rejected without truncation. Word/PPT length guards run before requesting long BSTR text. Excel cells are fetched individually to avoid a whole-range BSTR array allocation. Collection traversal stops at a 2.5-second cooperative budget; a synchronous Office COM call itself cannot be forcibly interrupted. The callback request has a twelve-second timeout, and no Office process is killed on timeout. Agent edits are the same whitelist Core already journals: Word text insertion, Excel value or formula ranges up to 200 cells, a new PowerPoint slide, and a new text box. Prepare reads the prior state and returns compensation; Apply checks that hash before changing anything; Revert refuses when the post-edit hash no longer matches. Manual Save calls the document `Save` only for an existing local `.docx`, `.xlsx`, or `.pptx` and never changes `AutoSaveOn`. Unsaved documents are not given a Save As dialog. Excel writes can clear Excel's own undo stack; the stored recovery record is the supported reversal. Native embedding and native AutoSave redirection stay unsupported.

## Build and verify without Office

```powershell
dotnet build packages/office-addin/OfficeInEditor.OfficeAddin.csproj -c Release
dotnet build packages/office-addin/tests/OfficeInEditor.Addin.Smoke.csproj -c Release
& packages/office-addin/tests/bin/Release/net48/OfficeInEditor.Addin.Smoke.exe
```

The output DLL is `bin/Release/net48/OfficeInEditor.Addin.dll`. The smoke executable checks Ribbon XML/ordering, selection immutability and rejection, nullable AutoSave observations, pipe ACLs, size limits, callback method restrictions and real STA dispatch using test doubles. It does not activate, register, open, close or alter Office. A successful build/smoke does **not** establish live Word/Excel/PPT compatibility; live acceptance remains a separate user-controlled step.

## Registration

The product installer copies the DLL to the stable current-user product `office` directory, then calls the repository's `scripts/install-office.ps1` with `-PayloadRoot` and `-InstallRoot`. The script reads assembly metadata only and writes current-user COM/add-in keys in the detected Office registry view. It uses `mscoree.dll` with a stable CodeBase, with no RegAsm/HKLM writes or COM activation. `-ValidateOnly` checks existing registrations without activating Office. `-WhatIf` uses PowerShell's standard dry-run support after payload preflight.

Uninstall validates ownership and the installation's CodeBase, then removes only this add-in's bounded HKCU registration keys. It does not delete documents, user references, Office temporary/recovery files, product payload files or arbitrary registry branches. Existing Office processes are never restarted or unloaded; registration changes take effect on a later normal launch. User or enterprise Office policy may independently disable COM add-ins.

No document body or source path is written to diagnostics. A local source path is included only in the private selection protocol as allowed by the shared contract; unsaved names, URLs and network paths are not treated as local files. No absolute machine path, credential or chat transcript is present in source.
