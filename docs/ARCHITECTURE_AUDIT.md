# 架构补全与实现审查（2026-11-15）

本文是对 0.1.4 工作区源码、设计手册和交接材料的静态审查。它补充设计约束，**不是修复完成、已安装版本或真实 Office 验收声明**。本次没有启动 Office、操作用户文档、修改安装或注册表。历史文档的测试仅对其注明的构建和模式有效。

后续源码与 [DESIGN.md](DESIGN.md) 已把嵌套定为 owned popup 铺在 `#ole-root` 令牌矩形上（跨进程 `WS_CHILD` 进 Chromium 会并入输入队列）。几何、静默加载、保存按钮与仅定时保存以 DESIGN 为准。OleExclusive、HideNewFrames 对话框例外仍须活测。聊天官方 Hook 见 [后续代码复核](CODE_REVIEW_FOLLOWUP.md)。

## 1. 当前实现与目标的区别

| 能力 | 当前源码证据 | 尚未证明的部分 |
| --- | --- | --- |
| 双层插件、共享 Core | 同一 VSIX + 私有 Core/MCP + COM 加载项 | 原生 HWND 承载是否能跨 IDE 版本稳定工作 |
| 全系列文件入口 | package.json 已声明 Word、Excel、PowerPoint、Visio、OneNote、Publisher、Project 文件；OfficeDocumentKind 做分类 | 选择器不是兼容证明；无服务器时仍猜测扩展名，尚未严格 fail closed |
| 多标签、多种类 | OleDocumentHost 已有多会话、每 kind 独立 STA | 文档归属、同 PID 多文档、关闭取消、资源回收未闭环 |
| Office 协作插件 | 注册、Ribbon、枚举、选区、编辑分派覆盖 Word/Excel/PowerPoint | 其它四款连 inspect、选区、保存分派也未实现 |
| 仅定时保存 | 规范：周期静默点保存按钮，默认 60s；取消编辑后保存 | 独立窗口 Ribbon 与标签内 SaveObject 均须活测写回原件 |
| 原生 AutoSave 后端重定向 | 能力为 false | 未实现；插件 Save 不得冒充原生重定向 |
| 30 秒可交互 | 有阶段诊断和超时机制 | 没有双 IDE、多应用冷/热启动实测数据 |

尤其需要修正旧说明：NativeOfficeLauncher.Open 仍传 `bindSource:false`，但从暂存副本 `OleCreateFromFile`；原件回写走 `SaveObject`/`IPersistFile.Save`，不是 0.1.3 文件绑定路线。源码已有多会话和扩展选择器，旧文档的“三格式/单实例”、双自动保存开关已过时。详细应用矩阵见 [OFFICE_COMPATIBILITY.md](OFFICE_COMPATIBILITY.md)。工作区有未提交实现，代码行号会移动，定位以方法名为准。

## 2. 倒 Y 的职责与扩展点

~~~text
                    一份 Office 加载项
         Ribbon / 右键 / 原生文档身份 / 选区 / 应用适配
                             │
                共享 Core（唯一权威与保存仲裁）
        会话 / 路由 / 引用 / MCP / 事务 / 恢复 / 配额
                             │
                  统一 Host Adapter 接口
                     ┌───────┴────────┐
             一份 VSIX 薄外壳       DSH Cordis 适配
             VS Code / Cursor      仅预留，当前不做
~~~

分支只适配宿主生命周期、文档表面和聊天能力，不复制路由库、保存器、恢复目录、清理器或 Office 加载项。新增分支必须复用同一套权限、去重和配额契约。

共享核心内部仍需分清三种接口，不能用“一次适配”抹掉应用差异：

| 接口 | 必须负责 | 不应负责 |
| --- | --- | --- |
| HostSurfaceAdapter | 宿主租约、标签身份、表面身份、可见性、关闭协调、宿主支持的聊天交付 | Office 正文操作、应用主框嵌套、猜测其它窗口 |
| NativeDocumentAdapter | 文档服务器能力、创建/激活、frame/viewport、文件持久化、可取消关闭 | 聊天目标猜测、全桌面窗口整形 |
| OfficeApplicationAdapter | 文档对象身份、选区语义、保护/忙碌状态、保存、有限 MCP 白名单、菜单 ID | 复制 Core 的权威状态、用 ActiveDocument 替代指定 documentId |

以上是逻辑边界，不要求把产品拆成独立程序。分发仍是一份 VSIX 携带插件自有载荷；Office 加载项需标准用户级注册，必要的安装器可作单文件交付入口，但不能绕过扩展安装与 Office 加载机制。不得捆绑 Office 本体、增加后台服务、托盘、常驻更新器或独立编辑界面。

### 表面承载是单独的发布关卡

VS Code 公共 Custom Editor API 给出的是 Webview/文档生命周期，并未在该接口契约中提供可供插件嵌套原生 HWND 的专用槽。`#ole-root` 和无障碍 token 只能标识网页节点，**不能证明它有独占的原生 HWND**。因此，当前 UIA/MSAA + SetParent 路线是有兼容性风险的实现实验，不能仅因使用 VSIX 就宣称全部依赖均为稳定扩展 API。[VS Code Custom Editor API](https://code.visualstudio.com/api/extension-guides/custom-editors)

必须验证 token 与具体宿主、标签、原生父窗口的一一归属、裁剪和输入范围。发现的是共享 Chromium renderer/workbench 祖先或其它 IDE 的相似窗口时，应拒绝附着。禁止用宽高相似、全屏包含或最近窗口补足身份；不得覆盖聊天、侧栏或其它标签。仅能 SetParent 插件自己的容器，不能 SetParent Office 主框。若无法证明专属表面，保持短状态失败；不自动改成独立应用。

现行定位辅助：标签 `#ole-root` 四角 1 像素黑点，Name 为 `{token}:tl|tr|bl|br`。Core 取角点并集，不得把 IDE 主客户区当编辑器。标签底色跟随 Office `UI Theme`，不跟随 Cursor 编辑器主题。

### 会话与文档身份

后续协议应把下列身份作为一组受验证的能力，而非几个可猜测字符串：

`hostId + hostLease + hostEpoch + editorId + generation + sessionId + documentId + sourceFileId`

- open、layout、hide、save、close、路由和 MCP 操作均核对归属；入队前与 STA 真正执行前都核对 generation/deadline。
- PID、HWND、文件路径仅作辅助。PID 可以有多个文档；HWND 会复用；相同路径可能被替换、链接或从不同别名访问。
- cancel-before-open 建立有界取消墓碑，禁止晚到的 open 重新创建会话。关闭、取消、失联回收都幂等且仅作用于自己的会话。
- 宿主失联进入恢复状态，不把未知或 dirty 文档按 NOSAVE 丢掉。已独立打开的 Office 不随宿主退出。
- kind 使用已确认的服务器身份、CLSID/安装信息与位数归组；不能只用 exe 文件名证明同一安装。每 kind 的 STA 串行化不代替每 document 的编辑与保存锁。

### 能力发现、版本与格式

扩展名只产生候选。必须分别确认本机授权安装、文件类、所需 OLE/DocObject 接口、激活、加载项、选区、持久化与恢复能力。拒绝列表之外还要有正向 Office 服务器与能力检查，不能接受任何非 Outlook 的 LocalServer。

Manifest 的文件选择器是静态贡献，不能声称会随本机 OLE 注册动态消失。未验证格式应使用显式选择/可选关联并在运行时拒绝；不能因扩展名命中就宣告已支持。模板、放映、模具、笔记本导入包与普通文档需要不同打开/保存策略。

保留用户“只支持已有原生自动保存时代的 Office 版本”的门槛；不得把只读到一个 API、文件恰好能打开或旧版测试通过视为版本准入。发布矩阵需写产品线、版本/构建、许可、渠道、Office/加载项位数和 OS。当前文件的 `AutoSaveOn=false` 不等于产品版本过旧；LTSC 与 OneNote 等例外需独立记录能力，不扩大为全版本兼容。

### 路由与 Add to Chat

- 独立模式按既有交互约定：手动选择/清除优先；未手动干预时只取最新活动的已登记聊天；无目标静默。引用仅加入上下文队列，不自动发送。
- 内嵌模式按调用来源重绑，但范围是该宿主/文档会话，不是该 PID 的全部文档，更不是所有已嵌入 Office。
- Core 持有按 session/document 的权威路由与递增 routeVersion。targetId 必须校验 owner、token/epoch 和宿主边界；推送加载项只是更新菜单缓存。
- Add to Chat 最终提交前重新核对当前文档、会话 generation 和 routeVersion。500ms 推送失败不得被描述为菜单已成功绑定，也不得沿用旧目标发送。
- 同一加载项支持两条宿主分支；未来 Cordis 只实现聊天能力适配，不给 Office 增加一套分支专用菜单或写入器。

### 保存、恢复与关闭

权威口径见 [DESIGN.md](DESIGN.md)「原件保存与定时保存」。Core 为每个文档提供一个 SaveCoordinator：统一调度手动 Save、MCP Save、**定时保存**和关闭保存。插件 **取消编辑后保存**；只保留定时保存，且必须与保存按钮同一条提交。Office 原生 AutoSave 不能被全局关闭，需观察其结果并避免和插件 Save 重叠；能力字段不得把原生 AutoSave 标成已重定向。

- 保存按钮、Ctrl+S、MCP `office_save_document`、定时心跳都进入 `IOleClientSite.SaveObject` → `IPersistFile.Save(用户原路径)`。不得因 `bindSource:false` 返回 `E_NOTIMPL`。不得把内存 `IStorage` 当 OOXML 写回原件。
- 创建用 `%TEMP%\OfficeInEditor\embed` 暂存副本做 `OleCreateFromFile`，避免打开阶段截断用户文件。未绑定或未验证源文件的关闭用 `OLECLOSE_NOSAVE`；`PROMPTSAVE` 不得把空嵌入写回 `.docx/.xlsx/.pptx`。
- 独立窗口的定时保存：`CommandBars.ExecuteMso("FileSave")`，失败再回落到文档 `Save`。标签嵌入不依赖加载项 DLL 当次替换，走 Core `SaveObject`。
- 默认开启定时保存，间隔 5–3600 秒、默认 60 秒。有界心跳看内存 dirty；无改动不写盘。失败、冲突、只读后暂停该文档定时保存，短状态，不关文档、不循环弹窗。
- 多标签应按已绑定 dirty 文档调度，不能悄悄把后台文档排除。
- 保存前校验文档 COM 身份、源文件身份、generation、保护/只读/忙碌和外部修改。长度+mtime 仅是低成本冲突提示。
- 关闭必须可取消：先协调 dirty/保存/恢复，再确认 Office closed，最后释放视图/会话与标签所有权。VS Code 可编辑 CustomDocument 的 dirty、save、revert、backup 生命周期也要接入；不能在 panel dispose 后才弹保存并吞异常。[VS Code 文档生命周期](https://code.visualstudio.com/api/extension-guides/custom-editors)
- panel 的 dispose 通知本身不是可异步否决的关闭钩子。若宿主已移除视图而 Office 拒绝关闭，必须提供重新接管入口。
- NOSAVE 仅可回收**已证明尚不可编辑且未被修改**的失败创建对象，以及从未绑定原件的暂存嵌入；已有交互、dirty 或结果未知时必须保留可接管的恢复入口。

### 时间、资源、输入与诊断

保留两个不同指标：**30 秒内真实可交互是性能目标**；现行交接规定 **60 秒标签终态硬上限**，源码虽有预算常量，尚未证明连接等待、COM 挂起等所有路径都能兑现。两者都不是 API 返回或成功创建 HWND。从 tab-created / openId 创建起算，分别记录 queued、surface-bound、object-created、UI-activated、first-content-visible、editable/failed。握手、排队与重试计入同一标签墙钟，不能每阶段重置；传输超时不能再额外延长 UI 截止时刻。COM 卡住时不能承诺安全强杀或同步中止，UI 到期后标记失败/未知并隔离后续调用。

全局设定 session、kind、队列和原生窗口预算；每 kind 空闲时释放 STA/COM，计数清零后 Core 才可退出。不能无限增加永久线程。`open_ok` 前容器保持隐藏。后台切换只 hide 对应自有容器，不 close 文档，不触发 Save，不 UIDeactivate。正常用户输入不得溢出标签令牌矩形；不装全局键鼠钩子、SendKeys 或全局焦点修正循环。STA 在 OleCreate/DoVerb 等待期间必须泵消息。

生产默认不往固定地址 POST，也不持续写调试文件。诊断使用有界内存环形缓冲，用户显式导出一次脱敏快照；不记录正文、选区、真实路径、聊天令牌/原文。错误按文档和错误码合并，状态变化才通知；Office 自身安全/许可/保存提示不能全桌面隐藏。清理只触及插件拥有且身份已验证的资源，不跟随链接目录、不删原文件、不清 Office 全局缓存。

## 3. 风险清单与建议顺序

下表为当前源代码事实及后果分析，文件位置相对仓库；不是已经发生数据损坏的断言。

| 优先级 | 证据 | 风险 / 应收敛的行为 |
| --- | --- | --- |
| P0 | NativeOfficeLauncher.Open；OleDocumentHost.OpenOnSta / OleSite.SaveObject | 规范已改为暂存副本创建 + SaveObject 写原路径；须活测证明保存按钮与 60s 定时点击都更新用户文件，且关闭未保存会话不把原件写成 0 字节。CustomDocument dirty/取消关闭仍缺 |
| P0 | OleDocumentHost.SnapshotVisibleTopLevel / HideNewFrames | 隐藏桌面所有新出现且未被排除的顶层窗。规范：必须跳过 `#32770` / `NUIDialog` / `bosa_sdm*` 等对话框，且须证明本次 session 归属。误藏 Excel 修复/许可窗会把 OleCreate 卡到 60 秒 |
| P0 | extension.ts 的 CustomEditor 注册与 hostDocument.close；OleDocumentHost.CloseOnSta | 标签先 dispose，关闭失败被吞；Office 取消保存后会话还在、标签却没了。实施可取消关闭与恢复接管 |
| P1 | OleSite.GetWindowContext / SetBorderSpace / CompleteActivation / QueueViewLayout | 工具栏边距未落实，文档设回满客户区；Word 顶部遮挡的主要代码嫌疑，详见下一节 |
| P1 | WebviewSurface 令牌/四角打点 | 规范禁止 sizeMatch、`screenX/Y`、workbench 级 HWND；≥80% IDE 面积必须拒绝。须活测切回标签仍用打点并集，不铺满界面 |
| P1 | ui.ts 的 layout 上报；OleSite.Place 缓存分支 | CSS 尺寸写入原生 SetWindowPos，初次与缓存分支单位不同；高 DPI/缩放可能放大错位 |
| P1 | CoreState.cs:400–428；OfficeAddin.cs:303–320 | 显式 targetId 缺 owner/token/epoch 核验、向全部内嵌实例广播，推送失败仍 bound；旧聊天目标或跨文档投递 |
| P1 | OleDocumentHost.SourceForProcess；EmbeddedDocumentBindings.Update | PID 第一条源绑定、仅 Word、每次替换整张绑定表；不满足多文档/多应用。当前 SourceBound=false 时更不会提供此绑定 |
| P1 | SDK api.ts 的 open/layout/detach；CoreState.RemoveHost；OleDocumentHost.Session | 缺宿主 lease/generation；失联宿主会话可能残留，晚到调用可能作用于错误生命周期 |
| P1 | Runtime.cs:184–200；extension.ts:290 等 | 默认追加源机器固定调试文件且无上限，默认 loopback ingest；违反隐私、有限 IO 与垃圾上限。统一显式、有界诊断 |
| P1 | DocumentOperations.cs:274–284、554–570、409–411 | 能力宣告不完整反映保护状态，部分读取保护失败按可写继续；应未知即拒绝写，不把后续 COM 报错当保护策略 |
| P2 | OfficeDocumentKind.TryClassify；OleDocumentHost.TryOpen / StaWorker | 无服务器仍扩展名猜测、接受非拒绝列表服务器、kind 只按 exe 名、STA 无退出/总上限；补正向能力与资源预算 |

表中 extension.ts/ui.ts 位于 packages/adapters/vsix/src；Core 文件位于 packages/shared-core；加载项文件位于 packages/office-addin。优先消除 P0，再验证布局和文档身份，最后逐款扩大兼容范围。不能先批量关联全部文件，再把失败当用户自担风险。

## 4. Word 顶部 UI 覆盖正文的分析

**首要代码缺口：frame 的空间和 document viewport 的空间没有分开。** 当前 GetBorder 返回整个客户区，RequestBorderSpace/SetBorderSpace 恒 S_OK；后者不保存边距、不重排文档。GetWindowContext 的 pos/clip、QueueViewLayout、激活和恢复路径又都给文档整个 ClientRect。工具栏与正文因此可能被分配到同一片空间，下一次 resize 还会覆盖先前布局。

微软规定 SetBorderSpace 实际分配工具栏空间，SDI 的 SetRect viewport 应扣掉工具栏、状态栏等区域。RequestBorderSpace 恒成功本身并不必然有错；问题是批准后没有落实。当前 UIActivate → SetRect → Show 顺序本身符合规范，不应倒序“修复”。[SetBorderSpace](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-ioleinplaceuiwindow-setborderspace)、[SetRect](https://learn.microsoft.com/en-us/windows/win32/api/docobj/nf-docobj-ioledocumentview-setrect)

~~~text
容器客户区 C = (0, 0, W, H)
已提交的工具栏四边宽度 B = (L, T, R, Bottom)
正文视口 V = (L, T, W-R, H-Bottom)

┌────────────── C ──────────────┐
│   Office 请求并分配的顶部 UI   │
├──────────────────────────────┤
│              V               │
│         文档正文视口          │
└──────────────────────────────┘
~~~

这只是协商模型，不是给 Ribbon 硬编码一个高度。BORDERWIDTHS 的字段是四边宽度，不是绝对 RECT 坐标。若 Word 已在内部管理某块 UI，且没有向容器请求非零边距，不能重复扣减；需用日志辨明 UI 的父窗口与实际请求。当前静态证据足以确认协议实现缺口，**没有证明唯一运行时根因，也没有证明文件内容被删除**。

建议统一布局流程：

1. 记录并校验 Request/SetBorderSpace 的非负宽度、NULL/全零、HRESULT、当前 frame、view 与 clip 矩形；确认是哪一层工具栏占位。
2. RequestBorderSpace 只决定能否满足；SetBorderSpace 提交/释放边距。接口要正确表达 NULL，不能仅靠不可空的 ref NativeRect。NULL 与全零结构在 OLE 中语义不同。回调中不泵消息、不弹框、不重入激活。
3. UIActivate 返回后重新读取最新边距再计算 V；首次激活、恢复、WM_SIZE、OnPosRectChange、DPI/父表面变化全部走同一个布局方法。clip 是实际可见裁剪范围，不能把文档重新扩大到工具栏区域。
4. SetActiveObject 当前未保存活动对象，缺少 ResizeBorder 通知。正确持有/释放 IOleInPlaceActiveObject；frame 尺寸变化时通知其重新协商，再使用最新 V。[ResizeBorder](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-ioleinplaceactiveobject-resizeborder)
5. 原生父客户区作为原生布局真值，CSS 消息仅触发布局或附带明确换算。当前缓存分支把 visualViewport 的 CSS 像素直接给 SetWindowPos，需消除混用。记录两端 DPI awareness/窗口 DPI 并检查 API 返回；跨进程 SetParent 对 DPI 有额外约束。[SetParent](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setparent)

不应通过移动页面、加固定 Webview padding、强制隐藏 Ribbon、调整文档页边距或按 Office 主窗类名补偏移解决。它们掩盖容器分配错误，还会改变编辑行为或在其它 DPI/应用上复发。

## 5. 后续修复与验收方案

先用 fake COM/窗口边界夹具验证身份、边距状态机、NULL/全零、超时/取消、失败不跨会话；再用合成文件做真实 Office 验证。不要用用户文档探测崩溃或数据回存。

| 关卡 | 需要的证据 |
| --- | --- |
| 表面 | 两个 IDE 同时开、相同尺寸、多窗口/多标签时仍只附着所属表面；聊天/侧栏无覆盖 |
| Word 布局 | 页首放明显标记；Ribbon 展开/折叠、resize、标签切换、DPI 100/150/200%、IDE 缩放/跨屏后页首可见；记录 C/B/V/clip 及 GetRect |
| 持久化 | 编辑后按源文件身份关闭并回读原文件，标记存在；外部修改、只读、保存事件取消、保存未知均不假报成功 |
| 自动保存 | 两开关各自与同时开启；同一改动只一次 Save；无改动不写盘；后台文档策略明确；不关/重开界面 |
| 关闭恢复 | 保存/不保存/取消、IDE 退出/崩溃、Core 中断、重载扩展、延迟 open 回包；取消后标签/文档仍可接管 |
| 输入 | 文档、Ribbon、聊天往返点击；快捷键仅当前目标处理；无全局键鼠 hook、无循环激活 |
| 容量与诊断 | 达到 kind/session/恢复预算停止新增；空闲无磁盘轮询；日志/引用有界，卸载不删原文件 |
| 性能 | 以真实可编辑为终点，分冷/热启动测量队列与各阶段；30 秒目标、60 秒终态分别报告，不拿 CLI/管道成功替代 |

发布记录需包含 IDE/Office 构建号、位数、文件类别、模式（独立/内嵌）、原件保存回读和失败恢复结果。历史独立 Word 自动保存通过不代表当前内嵌通过；Excel/PowerPoint fake 回归不代表真实 Office 通过；新增四款必须逐款验证后才可声明兼容。
