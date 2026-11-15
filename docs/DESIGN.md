# Office in Editor 设计手册

**设计目标（优先于实现细节）：标签嵌套覆盖具有合格本地文档服务器的 Office 文件编辑器，含 Word、Excel、PowerPoint，以及满足条件的 Visio、OneNote、Publisher、Project。支持多标签、多应用；30 秒内可交互为性能目标，60 秒为当前交接规定的标签终态硬上限。适配走同一套容器契约，不按 Office 主窗类名逐个打补丁。不含 Outlook、OneDrive 等非文档编辑成员。**

**0.1.4 规范补全（2026-11-15 活测）：** 选择器已扩至七款候选；多会话、每 kind 独立 STA。加载项选区/MCP 仍只分派 Word/Excel/PowerPoint。标签嵌套以 `#ole-root` 令牌矩形为唯一内容区；**打开成功前容器必须保持隐藏**，失败不得留下挡住 IDE 其它区域的 HWND。保存只有一条路径：与功能区保存按钮相同的一次原件 Save；插件只保留**定时保存**，实现为周期静默触发该按钮，取消「编辑后保存」。原生 AutoSave 后端重定向仍未实现。Excel/PPT 的 OLE 创建必须在 STA 上泵消息；无 `IOleDocument`/`IOleInPlaceObject` 的格式（常见于桌面 OneNote）短状态失败，不抢主窗。全系列可交互与关闭取消仍待验收。

嵌套几何、静默加载、保存按钮与仅定时保存以本文为准；身份、路由、关闭取消和修复关卡见 [架构审查](ARCHITECTURE_AUDIT.md)；逐应用差异见 [Office 兼容契约](OFFICE_COMPATIBILITY.md)。[本地保存与诊断](LOCAL_AUTOSAVE_AND_DIAGNOSTICS.md) 是 0.1.3 双开关历史，不能替代本文。

本手册描述当前仓库实现的 VS Code / Cursor 分支，以及尚未写进代码的设计目标。DeepSeek Harness 适配不在本版范围内。这里不记录本机路径、账号或聊天原文。

## 产品形态

产品是两层插件，不是独立软件、托盘、系统服务或开机启动项。

```text
              一份 Office 加载项
                      │
             共享 Core（私有执行器）
                      │
               统一 Host Adapter
             ┌────────┴────────┐
     同一份 VSIX          DSH Cordis 适配
     VS Code / Cursor     当前暂缓，预留分支
```

- Office 加载项负责 Ribbon、右键 Add to Chat、选区捕获、受白名单约束的原生对象模型读写。独立模式使用相同加载项；标签失败不自动回退到独立窗口。
- Core 负责身份、引用、聊天路由、MCP、配额、去重、可恢复写事务，以及编辑器模式里的 x64 OLE 文档容器。
- VSIX 只做宿主薄适配：打开文档、把自定义编辑器 Webview 表面令牌交给 Core、注册 MCP。标签页是 Webview，提供 SuperDoc 式 100% 根节点（`#ole-root`）和表面令牌，不定位正文、不捕获选区、不代替 Office 编辑，也不是 OLE 站点。

直接打开 Office 与从编辑器打开，使用同一份加载项和同一个 Core。

公共状态、保存仲裁、事务恢复、引用与清理只能有一套。宿主表面适配、文档容器适配、应用对象模型适配分开；共用容器不等于共用 Word 的选区/保存对象模型。DSH 未来只增加 Cordis 宿主适配，不能复制 Core 或 Office 加载项。

## 两种打开方式

独立模式目标：用户在 IDE 外直接启动范围内文件编辑器时，相同加载项仍可使用。当前只实现 Word / Excel / PowerPoint，Visio / OneNote / Publisher / Project 需先补应用适配。Add to Chat **不猜未登记的聊天**，只在 Core 已显式登记的目标里选。Outlook 等范围外应用不装这条加载项路径。

- **未手动干预**（没有在 Ribbon 里点选或清除具体聊天）：每次 Add to Chat 只投给当前 **最新活动** 的已登记聊天。活动指该聊天最近一次 `chat.register`、经归因的 MCP 工具调用、向该聊天的投递或 App 连接。不弹窗、不状态长文、不改投到别的聊天、不自动发送消息。
- **获取不到**最新活动目标：无提示、无行为。Add to Chat 可以保持禁用，或点击后直接返回；不得出现错误条、MessageBox 或 toast。
- **手动干预**：用户在 Chat 菜单里点选或清除之后，以这次选择为准，直到再次手动改选。手动清除后不再自动填回最新活动。

编辑器模式：VS Code 与 Cursor 共用一个 VSIX。用户在 Cursor 里打开 **范围内的本地 Office 文件** 时，默认自定义编辑器是 `officeInEditor.document`（选择器是扩展名表，不是写死三种）。标签页 Webview **不能**变成 `IOleInPlaceSite`。写法对齐 SuperDoc 自定义编辑器：HTML/CSS 把 `#ole-root` 铺满 webview（`retainContextWhenHidden`），**内容只占该标签的编辑器表面**，不得盖住侧栏、聊天、标签栏或其它标签。Cursor 没有可用的 `workbench.editor.hibernation`；防休眠靠 `retainContextWhenHidden`、打开成功后钉住标签、以及切走时只 Hide 容器。Core 只 `SetParent` **自己的容器** HWND，永不 `SetParent` 各应用主框。跨进程把容器做成 Chromium 的 `WS_CHILD` 会把输入队列接到 IDE 线程，键鼠会溢出到标签外，因此当前嵌套是 **owned `WS_POPUP`，屏幕矩形等于令牌并集**，这是输入隔离下的嵌套，不是按 `screenX/Y` 乱摆的独立工具窗。用 webview 种下的表面令牌（及四角打点）找到 **我们自己的** `#ole-root`，禁止按宽度高度去匹配其它窗口，禁止把 IDE 主客户区或 workbench 级 `Chrome_RenderWidgetHostHWND` 当编辑器，也禁止按 `OpusApp` / `XLMAIN` / `PPTFrameClass` / OneNote / Visio 主窗类名做适配表。**创建与激活期间容器保持隐藏**；`ShowObject` 与 `DoVerb(SHOW)` 也不得在 `open_ok` 前把容器亮出来。打开失败必须拆掉本次 HWND，不得留下挡住非标签区域输入的残窗。容器实现 `IOleClientSite`、`IAdviseSink`、`IOleInPlaceSite`、`IOleInPlaceFrame`、`IOleDocumentSite`（以及目标应用作为文档服务器实际需要的同一组接口），对已有本地文件按 **文件的 OLE 文档类** 创建对象，再原位激活。不分发 `dsoframer.ocx`。关标签拆除会话（Close + DestroyWindow），不是只 Hide；切走标签只 Hide，且不得 `UIDeactivate` 拆掉功能区。从编辑器打开时原位失败 **不得** `Process.Start` 再开独立 Office，失败页也不得自动唤起独立窗口；标签只留短状态。独立窗口只走命令 `Office: Open in Microsoft Office`。选区、Add to Chat 仍在加载项里；Agent 改文档走 MCP。

编辑器模式里，只有 **已原位激活、且其进程被 Core OLE 会话认作服务进程** 的 Office 实例算「IDE 窗口内」。`inplace=false` 的独立窗口按独立模式路由。

## Office 全系列：谁进标签，谁不进

嵌套的对象是 **「用户打开一份本地文件、在文档界面里编辑」** 的 Office 应用，不是整个 Microsoft 365 套件。

### 范围内（设计目标）

凡是 Windows 上把某种 **本地文件** 登记为 OLE 文档服务器（或等价文档类）的 Office 应用，都走 **同一条** 标签嵌套路径：

| 种类 | 典型本地文件（示例，以本机登记为准） | 说明 |
| --- | --- | --- |
| Word | `.docx` `.doc` `.docm` `.rtf` | 有应用适配；当前内嵌/原件回存待修复验收 |
| Excel | `.xlsx` `.xls` `.xlsm` `.csv`（仅当 Excel 登记为文档服务器） | 有应用适配；真实内嵌/保存待验收 |
| PowerPoint | `.pptx` `.ppt` `.pptm` | 有应用适配；真实内嵌/保存待验收 |
| Visio | `.vsdx` `.vsd` `.vstx` | 与 Word 同类：文件 + 绘图文档服务器 |
| OneNote | 本机可打开的笔记本/分区文件（如 `.one` `.onepkg`） | **仅当** 该版本向 OLE 登记了文档类；UWP/云本若没有文档服务器，记为该格式暂不支持，不改去抢 OneNote 主窗 |
| 同类文件编辑器 | Publisher、Project 等 **以本地文件为工作对象** 且登记了文档类的成员 | 同一套容器，不单开 HWND 特例 |

扩展名只给出候选，ProgID、CLSID 从 **文件 → `GetClassFile` / 本机 OLE 登记** 探测，还需验证 DocObject/持久化/加载项能力。不为每个产品写窗口类适配。Manifest 选择器是静态贡献，不能随本机注册自动消失；未验证格式应保持可选关联并在运行时拒绝。没有合格文档服务器时短状态失败，不 `Process.Start`。Publisher 的产品退役、OneNote 的层级模型及其它特殊约束见兼容契约。

只支持具备原生自动保存功能的目标产品线/版本；按产品、许可、构建号和位数列准入证据，不能用统一的“Office 16”或当前文件 AutoSaveOn 值替代版本政策。自定义自动 Save 不改变这一版本门槛。

### 范围外（明确不做标签嵌套）

不以「打开一份本地文件并在其中编辑」为工作方式的套件成员 **不进** 自定义编辑器，也不做窗口嵌套：

- **Outlook**（邮件、日历、联系人、任务）
- **OneDrive**（同步与存储，不是文档编辑器）
- Teams、Skype for Business、Office Hub、SharePoint 管理端等协作/门户
- 商店版/UWP 外壳若不能作为该文件的 OLE 文档服务器

范围外的应用即使本机已安装，VSIX 也不关联它们的文件类型，Core 也不去找它们的主窗口。

## 一次适配所有窗口（禁止按应用打补丁）

嵌套契约只有一套。新增 Visio / OneNote **不得** 再增加 `VISIO` / `ONENOTE` 主窗类名、按客户区宽高匹配、或 `SetParent` 应用主框。

统一做法：

1. **表面**：每个自定义编辑器标签种下唯一 `oie-surface-` 令牌，`#ole-root` 铺满 webview。定位只认这个令牌及其四角打点（见下节），不认 `window.screenX/Y`，不认「看起来差不多大的窗」，也不把 IDE 主客户区或 workbench 级 Chromium HWND 当编辑器。
2. **容器**：Core 为 **这一次打开** 建一个自己的 HWND，实现文档站点接口。只 `SetParent` 这个容器，永不 `SetParent` Office 主框。当前因跨进程 `WS_CHILD` 会附着 Chromium 输入队列，容器以 owned `WS_POPUP` 铺在令牌矩形上；尺寸必须等于 `#ole-root`，不得铺满 IDE。
3. **文档对象**：按已验证的文档接口创建/绑定。`OleCreateFromFile` 只从 **暂存副本** 创建，禁止直接以写方式打开用户原件。原件回写只走保存按钮那一条 `IPersistFile.Save(原路径)`（见「原件保存」）。rclsid 按规范可传 CLSID_NULL；`GetClassFile` 失败时才用 ProgID 补 CLSID，补不到则短状态失败。容器客户区要再扣除 Office 协商的边界 UI。
4. **主框**：激活时出现的独立主窗不是嵌套目标。仅在证明窗口属于本 session/document 且不影响用户独立文档后，才允许针对性 Hide。**禁止全桌面新窗口差集 Hide/Close**，禁止隐藏 `#32770` / `NUIDialog` / `bosa_sdm*` 等安全、修复、保存提示；无法证明归属就停止控制。类名最多作诊断。
5. **加载项**：每款补齐注册、文档身份、Ribbon/context ID、选区、保护状态和保存适配，再按能力开放 MCP。缺模型不改为抓主窗，也不能宣称已有全系列协作能力。

### 标签表面定位与外观

标签自己提供几何，不向其它窗口「对齐」。

- **打点**：`#ole-root` 四角各一个 **1 像素黑点**（`pointer-events: none`），无障碍名为 `{token}:tl` / `:tr` / `:bl` / `:br`。点只辅助定位，不进入文档、不拦截输入。Core 用 UIA Name 读取至少两个角的屏幕矩形，取并集作为该标签内容区；对角齐全时以并集为准。
- **禁止**：用 IDE 窗口原点、CSS `screenX/Y`、整块 workbench `Chrome_RenderWidgetHostHWND`、或按宽高相似匹配其它窗来摆 OLE 容器。覆盖聊天、侧栏、标签栏即失败，应保持隐藏并短状态，不得将错就错铺满界面。
- **皮肤**：标签 webview 的底色与状态文字跟随 **本机 Microsoft Office 界面主题**（`HKCU\Software\Microsoft\Office\16.0\Common` 的 `UI Theme`；选「使用系统设置」时跟 Windows 应用浅/深色），**不**跟 Cursor / VS Code 编辑器色板。Colorful / White 为白底深字；Dark Gray / Black / 深色为深底浅字，使打开中缝隙与 Word 皮肤一致。

OneNote 现代客户端、仅云笔记本、无 `IOleObject`/`IOleDocument`/`IOleInPlaceObject` 的宿主：创建后若不是文档对象，**立即短状态失败**，不得 `DoVerb` 空等或铺一层失败 HWND。不允许改成「找到 OneNote 主窗再塞进标签」。

### 静默加载、切走标签与输入范围

- **静默加载**：`CreateWindow` 起就是隐藏的。`Place(show=true)`、可见 `ShowWindow` 只允许在 `open_ok` 之后。加载中标签可以显示 Webview 短状态（「正在打开」），那是页内 DOM，不是覆盖整个 IDE 的原生窗。
- **失败**：Abort/Dispose 本次容器；不得把 32×32 或错误矩形的 popup 留在工作区上。失败不得劫持侧栏、聊天、标签栏的键鼠。
- **切走**：只 Hide 该会话容器，保留 OLE 原位 UI（功能区边距不清零）。Cursor 自定义编辑器 `retainContextWhenHidden:true`；打开成功后 `keepEditor` + 钉住，避免预览标签被挤掉而 `dispose`→detach。本机 Cursor 无编辑器休眠设置可关。
- **切回**：用已信任的令牌矩形缓存，或重新打点；禁止用 `window.screenX/Y` 回退。禁止在切回时对整个 IDE 树做无超时 `FindFirst`（会把 STA 卡死数十秒）。
- **溢出即缺陷**：命中矩形相对 IDE 主窗 ≥80% 面积、或明显大于 Webview 上报的编辑区，一律视为未找到表面，保持隐藏。

## 原件保存与定时保存

原生 AutoSave 存储后端重定向 **仍未实现**，能力字段保持 false。插件不得把它改成 true。

本地保存只有 **与保存按钮相同** 的一次提交：

1. **保存按钮**（OLE 原位功能区、Ctrl+S、MCP `office_save_document`）都进入容器 `IOleClientSite.SaveObject`：对用户原路径 `IPersistFile.Save`。不得因 `bindSource:false` 就返回 `E_NOTIMPL`。不得把内存 `IStorage` 当 OOXML 写回 `.docx/.xlsx/.pptx`。
2. **创建**用暂存副本做 `OleCreateFromFile`，避免打开阶段截断原件。未绑定源文件的关闭用 `OLECLOSE_NOSAVE`，禁止 `PROMPTSAVE` 把空嵌入写回原件。
3. **只保留定时保存**，以及并列的 **关标签保存**（默认开）。取消「编辑后保存」开关与 2 秒去抖写盘。定时保存默认开启、间隔默认 60 秒（5–3600），实现为 **周期静默触发保存按钮**（独立模式 `CommandBars.ExecuteMso("FileSave")`，标签嵌入模式直接 `SaveObject`），不另写一套序列化、不 Close/Quit、不另存副本、不改 `AutoSaveOn`。关最后一个同文件标签时，关标签保存开启则 `SaveObject`，关闭则不写盘。
4. 无改动可不写盘；外部改文件、只读、冲突则暂停该文档的定时保存，短状态，不弹窗风暴。干净文档不因心跳读全文。
5. 安装时若 Office 正在运行，不得覆盖已加载的加载项 DLL；Ribbon 变更等到 Office 下次启动。标签内保存不依赖这次 DLL 替换，走 Core。

## 多标签、多应用、按种类 60 秒

性能另有 **30 秒内真实可交互** 目标；60 秒是规定的成功/失败终态硬上限，源码尚未证明所有路径都能兑现，不能用它宣称达成 30 秒。队列和握手均计入标签墙钟，详见架构审查中的测量关卡。

打开时允许：

- 一次打开 **多个** 自定义编辑器标签（同一应用多种文件，或几种应用一起）；
- **多种** 范围内 Office 同时处于打开中或已嵌套。
- **同一原件的多个标签共用一份文档对象**。第二个及以后的标签用 `IOleDocumentView.Clone` / `CreateView`，服务器不支持多视图时共用同一 HWND。任一标签里的编辑就是这份文档；禁止为同文件再 `OleCreateFromFile` 或静默复制工作副本。最后一个标签拆除对象。
- **Office 自定义编辑器不走默认预览「新换旧」**。`vscode.openWith` 用 `{ preview: false }`；标签一出现就 `keepEditor`，并用 `tabGroups.onDidChangeTabs` 把仍为 preview 的 Office 标签转成普通标签。不得改全局 `workbench.editor.enablePreview`。

关闭：

- Ribbon「关标签保存」与「定时保存」并列，默认开。关最后一个该文件标签时，开则走 `SaveObject`，关则 `OLECLOSE_NOSAVE`。先 Hide/Destroy 容器 HWND，避免标签已关表面仍停在 IDE 上。

计时与排队：

- **每个标签** 从 tab-created / openId 创建起计算自己的墙钟 **60 秒**，包含连接、表面握手、排队与 COM 阶段，不能等到发送 `office.open` 才起算。到期必须结束 UI 等待：成功或短状态失败/未知。后台标签同样要写出终态，不能只更新可见标签。
- **每种应用独立 60 秒**。种类键是 **OLE 文档服务器身份**（文件 `GetClassFile` / CLSID / local server），不是每个扩展名一种：`.docx` 与 `.doc` 同属 Word；Word 与 Excel 是两种。Word 的打开不占用 Excel 的 60 秒，Visio 不占用 PowerPoint。不同种类可并行创建/激活。
- 同一种类内部若文档服务器要求单线程/单 STA，只在该种类内排队；排队等待仍计入 **该标签自己的** 60 秒。种类 A 排队不得让种类 B 的计时停住。
- 全局只允许一个 OLE 会话、或一种类打开就 `office_busy` 挡掉其它标签， **不符合** 本目标。
- 60 秒到仍未原位成功：标签转短状态失败/未知，只回收本次已证明未修改的创建资源；已交互、dirty、COM 卡住或未知对象按恢复规则保留接管入口，不能无条件 NOSAVE。**不得** `Process.Start` 独立应用或把独立主窗当成功，不结束用户独立 Office 进程。

失败与并发时仍遵守：不杀 Office、不改文件关联、关标签才拆对应会话。

## 编辑器模式与 OLE

- Webview 不是 OLE 站点；[WebView2 也不提供 ActiveX 容器](https://github.com/MicrosoftEdge/WebView2Feedback/issues/321)。
- 协议参考：KB311765 DsoFramer 样本（hueidou/DsoFramer、stjeong/DotNetSamples 等）与 ActiveQt 的 `IOleInPlaceSite`。只读接口与激活动词，自写窄宿主，不注册旧 OCX。
- 反向项目（Office 里嵌浏览器）不采用。不把 Electron/VS Code 私有 workbench DOM 当几何来源。允许的嵌套只限于 Core OLE **容器**铺在 `#ole-root` 令牌矩形上（当前 owned `WS_POPUP`；禁止再把容器 `SetParent` 进 Chromium 以免输入队列并入 IDE）。
- 加载项仍必须能在独立打开的范围内应用、以及 OLE 容器里的对应实例中加载；不在 Office 进程里创建 WinForms `Control` 或托管窗口过程。`OnConnection` 只保存 Application，管道与 Core 注册放到 `OnStartupComplete`。STA 回调用 APC 切回 Office 线程，不建 `HWND_MESSAGE`。新种类若没有 Application 加载项模型，嵌套仍走容器路径，MCP/Ribbon 能力按缺模型缩小。
- 每个标签一次打开对应 **一条** OLE 会话。会话记住该次原位对象窗口的进程 ID。随后 `office.register` 若 pid 匹配某条嵌入会话，则该实例标为嵌入，接收 MCP 重绑。多种类、多标签可以同时有多条会话；不要用「全局只有一条 HasSession」挡打开。
- 容器 HWND 先隐藏创建。几何只来自令牌/四角打点的屏幕矩形。禁止 `HWND_TOPMOST`、禁止为对齐去 `SetForegroundWindow`、禁止按其它窗口的宽高做匹配、禁止用 `window.screenX/Y` 把 popup 摆到 IDE 原点或铺满工作区。未 `open_ok` 时窗口保持隐藏；`ShowObject` 也不得在 attach 前把容器亮出来。`GetContainer` 返回自身 `IOleContainer`。按暂存副本 `OleCreateFromFile` 之后 `OleRun`，把 `IOleObject` 赋给站点，**表面已解析且仍隐藏** 时再原位激活，成功后再 Show。Word 会在 `DoVerb` 期间回调 `ActivateMe`，此时必须已经能 `CreateView`。关标签走 `office.detach_window` 只拆除 **该标签** 对应会话。
- **每种** Office 文档服务器一条 STA。该线程在空闲时泵消息（`MsgWaitForMultipleObjects` + `PeekMessage`），**在 `OleCreateFromFile` / `DoVerb` 等阻塞 COM 调用期间也必须泵**（`IMessageFilter.MessagePending`），否则 Excel/PPT 会在调用方 STA 上卡到标签 60 秒上限。不能只在每次 `Invoke` 返回后 `PeekMessage`。不同种类的 STA 互不阻塞。
- 存在 **任意** 活动 OLE 会话时 Core 不算空闲，不得按 20 秒空闲把容器进程退出。全部会话拆除后才恢复空闲退出。

## Add to Chat 目标

引用仍然只排队，从不自动提交聊天消息，也不抢焦点。

### IDE 窗口内（嵌入）

每一次可归因的 MCP 工具调用（含 `office_register_chat`、读/写/捕获，不论成败）都必须把该嵌入实例的 Add to Chat 目标 **重新定位** 为「发起刚刚这次 MCP 操作的聊天」。

归因顺序：

1. `tools/call` 的 `_meta.officeInEditor.targetId`（MCP App 回传）。
2. 否则该 MCP 进程 `ownerId` 下最近一次 `office_register_chat` / `routing.touch` 所绑定的聊天。
3. 仍没有：不改当前目标，**无提示、无行为**。

定位只改加载项当前目标与菜单选中态，不打开聊天、不切 Cursor 窗口、不发消息。下一次 MCP 再覆盖。用户在嵌入实例里手动点选只维持到下一次 MCP。

当前 Core 通过加载项回调 `bindChatTarget` 推送 `{id,label,epoch}`，回调失败不得让 MCP 工具失败。目标契约还须按 host/session/document 核验归属与 routeVersion；不能向全部内嵌实例广播或把推送失败当成菜单绑定成功。最终 Add to Chat 以 Core 权威路由为准，推送仅更新缓存。

### IDE 外（独立 Office）

见上文「未手动干预 → 最新活动；获取不到 → 静默」。`chat.defaultTarget` 返回最新活动的公开聊天对象；没有则空对象。加载项在启动注册、Refresh、以及未手动干预的 Add to Chat 时读取它，不向用户解说。

不得把「最近一个 Cursor 窗口」或未登记会话当成目标。

## 标签页与 Ribbon 文案

插件要短。编辑器自定义标签页、Office **Office in Editor** 选项卡、侧栏能力列表都 **不放** 产品说明长句（例如 AutoSave 未重定向、Agent 走 MCP、引用需 App 确认、排队不发送）。

允许留下的：按钮名（Refresh chats、Choose chat、Add to Chat）、当前聊天短标题、失败时的短错误码。缺省状态行为是空白，不是一段教程。

诊断面板可以保留能力字段，仍不在文档标签页里展开说明。

## 明确未实现

- 把 Office 原生 AutoSave 的存储后端改成普通本地文件。本机若没有这项原生能力，不改 `AutoSaveOn`、不用 AutoRecover 或“保存后复制”来冒充原生 AutoSave。插件定时保存是另一条、明确标注的本地 Save，见上文。
- 把未经过 `chat.register` 的 IDE 会话当成 Add to Chat 目标。
- 自动发送聊天消息、后台抢焦点、把 Word / Excel / PowerPoint / Visio / OneNote **主框** `SetParent` 进标签。
- 按窗口类名、宽高、屏幕坐标为每个 Office 产品写一套适配器。
- 把 Outlook、OneDrive 或其它非文件编辑套件成员嵌进标签。
- 没有 OLE 文档类的 OneNote / 商店版外壳：不抢主窗冒充嵌套。
- 任意宏、任意 COM、删除文件、关闭时丢弃修改。

这些限制会在健康检查里保持为否，不因为写操作可用而改口。健康检查的否 **不** 写进标签页正文。

## Add to Chat 选区

右键命令请求排在 Cut 之前，这是相对顺序，不保证在所有 Office 版本里绝对第一。捕获的是当前语义选区：Word 文本范围、Excel 工作表区域、PowerPoint 文本、形状或少量幻灯片。空选区、过大选区和不支持的对象会被拒绝，不用整个文档代替。拒绝时：若来自无目标的静默路径，仍无提示；若用户已选定目标，可用短状态码。

引用是不可变的，有容量上限。释放引用只删除插件自己的引用记录，不删除用户原件。

## Agent 操作 Office

MCP 写操作走 Office 原生对象模型，并由 Core 串行化到同一文档。

| 操作 | 作用 |
| --- | --- |
| `office_inspect_document` | 读取已绑定文档和该文档实际允许的写种类 |
| `office_apply_edit` | 执行一次白名单编辑 |
| `office_revert_edit` | 仅当编辑后的校验仍匹配时撤销 |
| `office_save_document` | 对已经有本地路径的文档调用一次手动 Save |
| `office_get_operation` | 查看结果，不重放 |
| `office_release_operation` | 用户明确确认后释放恢复材料，保留去重墓碑 |

允许的编辑种类（对象模型白名单；**嵌套路径不依赖这份表**）：

- Word：在正文指定字符位置插入有限长度文本。
- Excel：写入不超过 200 个单元格的值，或写入以 `=` 开头的公式。
- PowerPoint：在末尾新增一张空白幻灯片，或在指定幻灯片上新增一个文本框。
- Visio / OneNote / Publisher / Project：当前连 inspect、选区和手动 Save 分派也未实现，能力必须报告不支持。后续先补文档身份与只读能力，再逐项开放已验证的等价 Save/白名单；不能仅因存在通用文件入口就声称 MCP 可用。

调用方必须自带稳定的 `operationId`。相同身份和相同参数不会执行第二次。Core 在真正改文档之前把补偿数据写入恢复记录。若响应丢失，结果记为未知，禁止换一个新编号重做。确认文档没有被改到一半之后，才能显式释放该未知记录。

撤销只覆盖这一次编辑。用户之后又改过同一区域时，撤销会被拒绝。Excel 的 COM 写入可能清掉 Excel 自己的撤销栈，不能把 Ctrl+Z 当成唯一恢复手段。

手动 Save 只接受已经存在本地路径的范围内文件。未保存文档不会弹出另存为对话框。Save 不修改 `AutoSaveOn`，也不把文档标成已保存来假装落盘。

每次上述 MCP 调用，只要能归因到已登记聊天，就按「IDE 窗口内」规则重绑 Add to Chat。

## 风险边界

- 不安装全局键盘或鼠标钩子，不用 SendKeys 或剪贴板监听。
- 不在后台抢焦点，不调用会主动激活窗口的 Select。
- 默认不因空闲反复读全文。仅定时保存：有界心跳看内存 dirty，到期静默点一次保存按钮；无改动不写盘。失败后暂停该文档，不关闭/重开、不弹窗风暴。不装全局键鼠钩子、不用 SendKeys 点按钮（独立模式走 `ExecuteMso("FileSave")`）。
- 相同故障不循环弹窗。缺目标时连 Ribbon 长状态也不写。
- 恢复材料有数量和体积上限。满额后停止新增编辑，不自动删掉尚未确认的恢复记录。
- 超时视为结果未知，不自动重试写操作，不结束 Office 进程。
- 文档身份、宿主租约、generation 与路由归属贯穿 open/layout/save/close；不能仅按 PID 决定文件或聊天。标签必须采用可取消的关闭与 dirty/恢复生命周期，确认关闭前不丢掉所有权。
- 诊断默认内存有界；显式脱敏导出，不默认写固定机器路径或 POST 固定调试服务。kind、STA、会话、引用与恢复目录都有总量和空闲回收预算。
- 卸载编辑器扩展不卸载 Office 加载项。卸载加载项不删除用户文档或引用材料。注册变更等到 Office 下一次正常启动生效，安装器不关闭正在运行的 Office。
