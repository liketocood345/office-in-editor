# Native Office add-in

Word has one **periodic save** Ribbon switch, off by default, with a 60-second interval (5–3600). Standalone Word executes the native `FileSave` command on its STA only for the current named writable local DOCX. It never activates a background document, closes/reopens Office, or falls back to a different save command after failure. Embedded Word delegates configuration to Core's timer for the exact session/document; Core owns the OLE source-file persistence, including background sessions. The add-in never waits synchronously for Core from Office's STA. These code paths have fake coverage; live acceptance is separate. Excel/PowerPoint retain their existing native Save foundation.

This package builds one .NET Framework 4.8 managed COM add-in for installed Windows Word, Excel and PowerPoint. It contains no Microsoft Office binaries, renderer or PIAs. Office APIs are called through IDispatch on Office's existing UI/STA thread.

## Implemented behavior

- A native **Office in Editor** Ribbon tab has **Refresh chats**, a menu of registered chats, and **Add to Chat**. Routing uses the full target ID. The tab does not show product-explanation copy. Refresh can start the installed private Core worker. Normal Office startup only registers with an already running worker.
- Standalone Office: without a manual Chat-menu choice, Add to Chat uses Core `chat.defaultTarget` (latest activity among registered chats). If none, no prompt and no delivery. A manual choose/clear sticks until the user changes it.
- Embedded OLE Office: Core `bindChatTarget` after each attributable MCP call selects that chat. Missing attribution is silent.
- Add to Chat sends `reference.captureAndDeliver` with the exact document ID and manual target ID/epoch when applicable. Core resolves the current destination before calling `captureSelection(documentId)` and rechecks it before enqueueing. Missing or stale targets are silent; no alternate chat is substituted. It queues context and never submits a chat message.
- Embedded document metadata comes from authenticated `office.source.resolve`, verified against each Word native window using `OBJID_NATIVEOM`, then registered with `office.document.bind`. Several sessions in the same process remain distinct. `refreshDocumentBindings` supports a direct Core snapshot or asynchronous refresh. Route callbacks are per document/session and reject old route versions and retired broker epochs. Dynamic menu content is invalidated on each drop, with target epochs included in menu tags.
- A current-user ACL-protected, network-denied callback pipe accepts `captureSelection`, `bindChatTarget`, the explicit edit callbacks `inspectDocument`, `prepareEdit`, `applyEdit`, `revertEdit`, and `saveDocument`, and the window callbacks `describeWindow`, `attachWindow`, `layoutWindow`, and `detachWindow`. It dispatches them to the Office STA. Requests are bounded to 12 seconds and 2 MiB. It has no global hooks, SendKeys, clipboard access, selection polling, private IDE APIs, or Office termination.
- `autoSaveEnabled` is a read-only observation of `AutoSaveOn`. `autoSaveAvailable` is `true` only when enabled, otherwise `null`: a false getter cannot distinguish a disabled toggle from an available toggle that is off. `autoSaveState` distinguishes `enabled`, `off_or_unavailable` and `unsupported`. No setter is used to probe support.
- The add-in does not create WinForms windows or a managed `HWND_MESSAGE` WndProc (both crashed Word: `clr.dll` `0xc0000005` / `80131506`). STA dispatch uses `QueueUserAPC` onto the Office UI thread. Pipe registration starts on `OnStartupComplete`. Editor-tab in-place activation is Core's OLE container; this add-in still accepts `describeWindow`/`attachWindow` when OLE did not take the session. It does not change Office window styles. Selection, Add to Chat, and MCP edits stay here. Redirecting native AutoSave to ordinary local files remains **unsupported**.

## Selection and menu coverage

| Application | Implemented selection | Implemented native context IDs |
| --- | --- | --- |
| Word | Nonempty text/range, including selected table text; story/start/end locator | `ContextMenuText` only (LTSC-safe; extra heading/table IDs can disable the whole add-in) |
| Excel | Worksheet Range selections; at most 5,000 cells and 32 areas, values/formulas/addresses | `ContextMenuCell`, `ContextMenuListRange` |
| PowerPoint | Selected text, selected shapes, or up to ten selected slides / 100 shapes; slide/shape IDs | `ContextMenuTextEdit`, `ContextMenuShape`, `ContextMenuPicture`, `ContextMenuTable`, `ContextMenuThumbnail`, `ContextMenuSlideSorter` |

`GetCustomUI` chooses Word, Excel or PowerPoint from the `ribbonId` Office passes (for example `Microsoft.Word.Document`), not from whether `OnConnection` has finished. Word currently attaches only `ContextMenuText` so a rejected heading/table ID cannot drop the whole customUI. The menu IDs were checked against [Microsoft's Office Fluent command identifier lists](https://github.com/OfficeDev/office-fluent-ui-command-identifiers/tree/main/Microsoft%20365/Current%20Channel). Each command requests `insertBeforeMso="Cut"`. This is relative ordering, not a guarantee of absolute first position against every Office version or other add-in. Unsupported object selections are rejected rather than substituted with ActiveCell or an entire document. Non-text PowerPoint shapes have a semantic ID/name/type representation; no screenshot/export or temporary file is generated. Group contents, charts, equations and full table structure are not expanded by this version.

Selections larger than 1 MiB of serialized UTF-8 JSON are rejected without truncation. Word/PPT length guards run before requesting long BSTR text. Excel cells are fetched individually to avoid a whole-range BSTR array allocation. Collection traversal stops at a 2.5-second cooperative budget; a synchronous Office COM call itself cannot be forcibly interrupted. The callback request has a twelve-second timeout, and no Office process is killed on timeout. Agent edits are the same whitelist Core already journals: Word text insertion, Excel value or formula ranges up to 200 cells, a new PowerPoint slide, and a new text box. Prepare reads the prior state and returns compensation; Apply checks that hash before changing anything; Revert refuses when the post-edit hash no longer matches. Manual Save calls the document `Save` only for an existing local `.docx`, `.xlsx`, or `.pptx` and never changes `AutoSaveOn`. Unsaved documents are not given a Save As dialog. Excel writes can clear Excel's own undo stack; the stored recovery record is the supported reversal. Native embedding and native AutoSave redirection stay unsupported.

## Build and verify without Office

```powershell
dotnet build packages/office-addin/OfficeInEditor.OfficeAddin.csproj -c Release
dotnet build packages/office-addin/tests/OfficeInEditor.Addin.Smoke.csproj -c Release
& packages/office-addin/tests/bin/Release/net48/OfficeInEditor.Addin.Smoke.exe
```

The output DLL is `bin/Release/net48/OfficeInEditor.Addin.dll`. The smoke executable checks Ribbon XML/ordering, selection immutability and rejection, nullable AutoSave observations, pipe ACLs, size limits, callback method restrictions and real STA dispatch using test doubles. It does not activate, register, open, close or alter Office. A successful build/smoke does **not** establish live Word/Excel/PPT compatibility; live acceptance remains a separate user-controlled step.

## Registration

The product installer copies the DLL to the stable current-user product `office` directory, then calls the repository's `scripts/install-office.ps1` with `-PayloadRoot` and `-InstallRoot`. The script reads assembly metadata only and writes current-user COM/add-in keys in the detected Office registry view. It uses `mscoree.dll` with a stable CodeBase, with no RegAsm/HKLM writes or COM activation. When Word/Excel/PowerPoint are not running, it also deletes only `Resiliency\DisabledItems` values whose decoded bytes mention `officeineditor.addin`; other disabled-document recovery values are left alone. `-ValidateOnly` checks existing registrations without activating Office. `-WhatIf` uses PowerShell's standard dry-run support after payload preflight.

Uninstall validates ownership and the installation's CodeBase, then removes only this add-in's bounded HKCU registration keys. It does not delete documents, user references, Office temporary/recovery files, product payload files or arbitrary registry branches. Existing Office processes are never restarted or unloaded; registration changes take effect on a later normal launch. User or enterprise Office policy may independently disable COM add-ins.

No document body or source path is written to diagnostics. A local source path is included only in the private selection protocol as allowed by the shared contract; unsaved names, URLs and network paths are not treated as local files. No absolute machine path, credential or chat transcript is present in source.
