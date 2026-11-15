# 当前实现复核：窗口、标签隔离、保存与聊天（2026-11-15）

嵌套与保存的现行规范以 [DESIGN.md](DESIGN.md) 2026-11-15 活测补全为准。本文部分条目（`WS_CHILD` 才算嵌套、SaveObject `E_NOTIMPL`、双自动保存开关、OpenOnSta 一律 Place true）已被后续实现和手册替换，只作当时证据。不要把本报告当修复完成。

## 1. 官方文档视图与 IDE 表面是两个问题

若“小窗”指 Office 的 OLE/DocObject 原位视图，官方已经提供文档服务器这一侧，但容器仍要实现每个文档/视图的 site、UI 激活、边界空间、菜单、保存与释放。VS Code Custom Editor 对扩展公开的是 Webview 和文档生命周期，没有在该 API 中提供原生 HWND 安装槽。因此“Word 能在自建原生窗口编辑”只证明 Office 一侧，未证明能安全作为 IDE 标签。[IOleDocumentSite](https://learn.microsoft.com/en-us/windows/win32/api/docobj/nn-docobj-ioledocumentsite)、[Custom Editor API](https://code.visualstudio.com/api/extension-guides/custom-editors)

**当前嵌套（DESIGN 已改口）：** OleSite.CreateWindow 创建 owned `WS_POPUP`，Place 调用 NativeWindow.OwnPopup。这是输入隔离下的嵌套（跨进程 `WS_CHILD` 进 Chromium 会把键鼠并入 IDE），不是按 `screenX/Y` 乱摆的独立工具窗。几何必须等于 `#ole-root` 令牌并集；铺满 IDE 或挡住侧栏即缺陷。不能沿用“已经作为 Webview WS_CHILD”的描述，也不再把 owned popup 本身当成违约。

这条路线必须自行同步遮挡、可见性、标签切换、焦点、屏幕坐标与 DPI，天然比真正父子窗口多一套状态；同时仍未落实 OLE SetBorderSpace/active object/ResizeBorder，Word 顶部遮挡缺口没有被 popup 改造解决。官方接口需要完整宿主实现，不能仅返回成功；也不能把项目自身漏接归因于 Office 不支持。

后续实现已在标签内种四角定位点，并用 Office `UI Theme` 作为 webview 底色。手册要求：点为 1 像素黑点；并集不得大于编辑区；禁止 `screenX/Y` 与 workbench 级 HWND。几何关卡仍须活测：命中矩形相对 IDE ≥80% 或大于编辑区 +48 必须拒绝。

## 2. 标签串扰的确定路径

| 当前代码 | 为什么表现为概率性串扰 |
| --- | --- |
| OpenOnSta 曾无论 visible 都 Place(..., true) | **已改为** 创建/激活期间 Place(false)，`open_ok` 后才 Show；须活测失败路径不留挡输入的 HWND |
| HideNewFrames 对全桌面窗口做前后差集隐藏 | 其它文档/软件在快照后出现的窗口可能被误藏；并未验证当前 session/document 归属 |
| FromUia/FromAccessible 的 owner 取 ForegroundIde，扫描却覆盖所有 IDE | token 来自后台 B 窗口，owner 可能被选为前台 A 窗口 |
| popup 缓存只跟随 IDE 顶层移动；layout 调 Place 未传新 x/y | 侧栏、分屏、移标签后原点可能过时；宽高更新也不能修正原点 |
| x/y 乘 dpr，但宽高仍 CSS，SDK 又丢 dpr | 同一矩形混合单位，高 DPI/IDE 缩放容易错位 |
| 所有 kind 共用 OleExclusive；每 kind 一个 STA 队列 | 一个慢 open 拖延其它标签的 hide/layout/save；等全局锁的另一个 kind 也处理不了自己的队列 |
| Office 加载项 selected/manualOverride/embeddedBound 按 Application 实例保存，Core 广播 | 多文档共用进程/加载项时，A 标签的聊天绑定覆盖 B |

VSIX 每次 hostDocument 的 bounds/closed/layoutRunning 等局部闭包和 session token 是独立的，不能笼统说所有变量都全局共享。全局 trace/lastFailure 主要污染诊断；真正功能串扰来自上述窗口归属、显示时序、调度及应用级路由。

改进先做可验证的限制：后台创建不显示/不 UIActivate(TRUE)；停止桌面差集控制；owner 必须属于命中 token 的宿主；统一原生像素矩形与每 session 的 visible/active/generation。激活与布局分开，不能为普通 resize 反复激活 Office。全局锁不得作为身份错误的补救；按服务器要求决定锁范围，不先移除所有串行化。

对严格标签嵌入，先证明指定 IDE 版本的原生表面身份与裁剪是否可可靠取得；不能证明时报告承载能力限制，不能把 popup 覆盖当正式完成。若 COM 阻塞仍需故障隔离，可由共享 Core 管理每 session 私有容器 worker，Office 加载项保持统一；但拆 worker 不保证 Office 服务器也使用独立进程，不能据此承诺进程级隔离。

## 3. 官方接口可以做插件本地自动保存

Word Document.Save、Excel Workbook.Save、PowerPoint Presentation.Save 都能保存已命名的本地文档；插件在适当时机调用即可，不必关闭重开界面。当前 LocalAutoSave.Save 已调用 IPersistFile.Save 或对象 Save，MCP 手动保存也复用它。[Word Save](https://learn.microsoft.com/en-us/office/vba/api/word.document.save)、[Excel Save](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.save)、[PowerPoint Save](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.save)

需要区分原生 AutoSaveOn：它是受位置/格式/状态限制的开关，文档禁用时强设会报错；当前公开文档没有提供把它替换成任意本地存储后端的通用接口。这不妨碍插件自动调用 Save。[AutoSaveOn](https://learn.microsoft.com/en-us/office/vba/api/word.document.autosaveon)

规范已改为：仍 `bindSource:false`，但从暂存副本 `OleCreateFromFile`；`SaveObject` 必须对用户原路径 `IPersistFile.Save`，不得 `E_NOTIMPL`。创建副本仍不等于文件绑定；活测须证明保存按钮与定时点击更新原件、关闭不把原件写成空文件。[OleCreateFromFile](https://learn.microsoft.com/en-us/windows/win32/api/ole2/nf-ole2-olecreatefromfile)

自动触发也有限制：当前每秒只看 ActiveDocument/ActiveWorkbook/ActivePresentation，后台文档不被检查；“编辑模式 2 秒”按距上次保存计算，不是距最后编辑的严格防抖。Excel Change、Word DocumentChange/SelectionChange、PPT 选择及结构事件也不是同一套任意编辑事件。[Excel Change](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.change)、[Word DocumentChange](https://learn.microsoft.com/en-us/office/vba/api/word.application.documentchange)

SaveCoordinator 只调度保存按钮、MCP Save、定时静默点击与关闭保存；取消编辑后保存。IPersistFile.Save 后应核对 SaveCompleted / 文件戳，不能只验证第一次成功。[IPersistFile.Save](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-ipersistfile-save)、[SaveCompleted](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-ipersistfile-savecompleted)

按绑定文档维护 dirty 调度，各应用事件辅助，加用户启用后的有界内存检查；无改动不读写文件。保护未知/冲突/保存取消暂停该文档，不改全局 DisplayAlerts，不标 Saved=true 伪造成功。验收必须反复保存、切换后台、关闭重开并回读原文件。

## 4. 聊天识别不足与绑定缺陷叠加

当前 Core 只记录登记、MCP touch、投递、App 连接，不知道用户切换聊天或普通输入。McpServer.ownerId 是每 stdio 进程随机身份，不是聊天身份；缺少自定义 targetId 时回落进程最后登记目标。同一 MCP 进程被多个聊天共用时会产生歧义；返回一次自定义 _meta 不会自动给未来工具请求补入它。

还有可直接修的断链：

1. Core open 完成发 refreshDocumentBindings，但加载项无该 handler，错误又被吞。若加载项早于 ServerPid 确认注册，Embedded=false 不被及时刷新，首次 MCP 可能根本不推送绑定。EmbeddedDocumentBindings 目前也没有实际使用点。
2. TouchRouteAsync 向所有 Embedded 实例广播，而非目标文档；共用 500ms 顺序推送，失败仍 bound:true。不是可靠绑定确认。
3. Add to Chat 先快照 selected，再异步查询 defaultTarget。查询期间新 MCP 已绑定 C 时，旧查询 B 返回仍可覆盖 C，随后 B 的投递又被标成最新活动，错误会强化自身。
4. embeddedBound 置 true 后没有完整的失效清除；刷新不再更新目标，旧聊天可能永久锁住。
5. App 断流即 RemoveChat；Reconnect 却拿旧 ID/token 重连，不能恢复。Core 重启丢内存聊天，加载项仍保留旧 selected。

### 已核验的官方来源身份路径

Cursor Hooks 提供稳定 conversation_id 与 MCP 前后事件；可把来源身份送到 Core。VS Code 有 Agent Hooks，但字段/预览状态随 harness 和版本而异，需按本机能力验证。当前代码尚未接入。[Cursor Hooks](https://cursor.com/docs/hooks)、[VS Code Hooks](https://code.visualstudio.com/docs/agent-customization/hooks)

建议以宿主事件为来源：用 prompt/tool 事件更新“最近实际操作”的已登记聊天；在本插件 register-chat 结果与该事件的会话身份之间建立映射。只提取必要身份，丢弃 hook 自带的正文、邮箱、转录路径、工具全文；不写事件流水、不改审批决策。Hook 只是自有插件载荷内的短生命周期桥接，不是独立常驻软件。安装/注册遵循宿主机制，不覆盖用户已有 hooks。

**来源识别和引用交付是两条能力。** Hook 知道 conversation 不等于能给该聊天注入上下文；仍需已登记的 MCP App/官方交付适配及收件确认。Hook 的活动事件也不等于聊天焦点变化：仅点击一个没有其它操作的聊天，不应伪称已观察到活动。缺字段或无可交付目标时静默保持未绑定，不能改为猜测。

Core 按 host/session/document 保存权威目标、失效 epoch 与递增 routeVersion；加载项菜单只缓存显示。Add to Chat 在 Core 内原子“选目标 + 入队”，确认后这次投递目标固定；旧异步查询不得覆盖新版本。用 invocation/operation 身份关联工具来源，不把任意 Hook 身份不加授权地绑定到所有 Office。失败的已归因调用同样更新来源；慢旧调用的完成不能覆盖新调用。短断流进入有界 offline 状态，重登记/换 epoch 后才恢复交付，禁止不断新建孤儿聊天。

## 5. 建议的修复批次

1. **停止误控制与串扰**：去全桌面差集隐藏、后台不显示/激活、固定表面所有权、每 session 可见性/矩形 generation；布局边界与 DPI 一并验证。
2. **闭环保存与退出**：原件身份、持久化、重复 Save、dirty 生命周期、取消关闭与恢复入口，之后再宣告内嵌自动保存。
3. **闭环路由**：补缺失 handler/登记刷新，Core 权威路由、版本和原子入队，修断线恢复；接入宿主官方会话事件，移除 PID 广播和进程身份猜聊天。
4. **建立最小回归集**：A/B 标签倒序 open/layout 回包、后台打开、同 Word PID 多文档、旧 defaultTarget 晚到、新 MCP 失败、App 短断流/Core 重启、同一原件多次保存回读。先用 fake 与合成文件，再扩大应用/格式。

这些修复复用现有共享 Core、VSIX 和 Office 加载项；无需再建立另一套通用框架。平台承载的可行性关卡和项目自身确定错误应分别验收。
