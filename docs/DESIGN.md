# Office in Editor 设计手册

**2026-11-21 当前实验方向：叠层 + 像素转发焦点切换。** SSH 隔离环境已验证同一 VS Code 窗口的 Word、Excel、PowerPoint 标签、失焦回切、经对象模型 `Save()` 写回原件并独立读回，以及原生 Add to Chat。Trident＋DSOFramer、WebRunLocal、原生槽和纯文本 localhost 保留各自的实验记录。正式 VSIX 与 Cursor 的接入仍待验收。已发现的实现缺陷（D1–D11）与攻坚阶段（P0–P6）记录在项目根目录 `planning/HANDOFF.md` 的「焦点混合实现缺陷」一节（不在本仓库内），下文引用的编号即指该节。

**本手册各节的适用路线。** 条款只约束其所属路线，不同路线之间的禁止项不互相套用：

| 章节 | 适用路线 |
| --- | --- |
| 技术路线总结与路线裁决 | 焦点混合（当前实验方向） |
| 本地 Word 网页与 VS Code 标签页 | 纯文本 localhost 对照 |
| 产品形态、两种打开方式中的「独立模式」、Add to Chat 目标、标签页与 Ribbon 文案、明确未实现、Add to Chat 选区、Agent 操作 Office、风险边界 | 各路线共用；与承载方式相关的句子已就地标注 |
| 两种打开方式中的「编辑器模式」、Office 全系列、一次适配所有窗口、原件保存与定时保存的第 1–2 条、多标签与 60 秒、编辑器模式与 OLE | 原位 OLE 容器（0.1.4 VSIX 当前实现） |

原位 OLE 路线中禁止叠层、整窗隐藏避让、owned popup、截图和 `Process.Start` 独立 Office 的条款，不约束焦点混合实验。焦点混合进入正式 VSIX 前，须以 ADR 写出它自己的承载、保存、关闭与路由契约，并同步修订项目根目录的 `docs/RISK_BOUNDARIES.md`。

## 技术路线总结与路线裁决

原生窗口承载路线的共同堵点是桌面窗口与 IDE 弹层的 z 序合成。当前用户裁决允许在隔离 SSH 实验中使用「叠层 + 像素转发焦点切换」。下方保留的网页对照、OLE 容器与产品契约各自适用于其描述的承载路线。

### 候选路线与堵点

| # | 路线 | 堵点 |
| --- | --- | --- |
| 1 | 原生窗口嵌入（叠窗 / DSOFramer / Trident / WebRunLocal / MFCBind） | z 序遮挡；塞不进 Chromium；键盘孤岛（本次接受为特性） |
| 2 | Chromium 转 IE（Edge IE 模式 / 再嵌套） | IE 模式只存在于 Edge 独立应用，不可嵌入 Electron |
| 3 | 网页 UI + 无头桌面 Word（纯网页 + COM 后台读写） | 界面不是 Ribbon（唯一软肋，无技术堵点） |
| 4 | 微软网页版 Office（WOPI / Office Online） | 需服务器 + 授权；非本地引擎、非本地文件 |
| 5 | 像素转发（抓帧 + 反注输入） | IME / 文字 / Office 弹出层 / 延迟 / 光标 |
| 6 | Chromium 上挖透明洞透原生窗 | Chromium 无"外来 HWND 图层"，必须 fork |
| 7 | 原生叠层上挖洞透 Chromium 弹层 | 拿不到弹层几何（黑盒）+ 实时性 / 圆角 / 焦点抖动 |
| 8 | **叠层 + 像素转发焦点切换（本次选定）** | **无法监听 VS Code 自发 popup（唯一公认硬伤）** |

共同死线：凡"原生窗口 + 与 IDE 弹层和平共处"的方案，都依赖"实时知道 VS Code 弹层在哪里"，而弹层几何对扩展是黑盒。只有路线 3 与路线 5 不需要该信息；路线 3 的代价是界面非 Ribbon，路线 5 的代价是 IME / 文字 / Office 弹出层 / 延迟。路线 8 也是原生窗口方案，但它不和弹层同时占用编辑区：预览态由 Webview 独占，不需要弹层几何；编辑态由 Office 独占，只需要知道「Office 持焦期间 VS Code 是否自发弹层」这一个事件，这就是下文的唯一硬伤。

### 选定方向：叠层 + 像素转发焦点切换

焦点驱动模型（**键盘孤岛是期望行为**，规避 IDE 与 Office 快捷键冲突）：

- **叠层可见 ⟺ 叠层持有焦点**：只要焦点在叠层上，叠层就在最上，键鼠全归 Office，VS Code 在后方。这是设计不变式；当前实现只处理「编辑中失焦」一侧，停放中的 Office 被 Alt+Tab、任务栏或自身对话框拉到前台时会在屏外持焦（交接缺陷 D1）。
- **其余情况一律展示像素转发画面**：焦点在像素转发窗口（标签页/画面）、焦点在 VS Code 其它位置（其它标签/终端/侧栏/命令面板）、或焦点在其它应用，都显示转发画面。
- 选择 Office 标签可显示预览。点击预览由标签发出激活请求，助手请求原生 Office 窗口前台、对齐内容区并核对键盘焦点。
- 焦点离开叠层：点击 VS Code 任意 chrome、切到其它标签/应用 → `EVENT_SYSTEM_FOREGROUND` 侦测到即把叠层移出可视区，标签页先显示上一张有效预览，再更新离屏转发画面。
- 像素预览与原生窗口共用文档内容区的物理像素边界。预览铺满内容区，裁掉采集多出的窗口外缘。边界一致不等于分辨率一致：当前预览最长边至多 960 像素、每 300 ms 至多一帧，显示时拉伸铺满；节流期间到达的帧直接丢弃且不补发，内容停下前的最后一次变化可能一直不显示（交接缺陷 D5）。原生窗口在屏幕上编辑时不得出现采集提示框。首帧尚未更新时用最后一张有效预览占位。

机制：

1. **叠层 z 序**：独立顶层窗口，不设 `WS_EX_TOPMOST`，用焦点跟随（谁有焦点谁在前）。前台切换用 `SetWinEventHook(EVENT_SYSTEM_FOREGROUND)` 驱动；当前根据 Webview 的布局通知与 UIA 内容区重新对齐。IDE 窗口移动时的 `EVENT_OBJECT_LOCATIONCHANGE / MOVESIZEEND` 跟随仍需验收。当前的「移出可视区」是把 Office 主框整窗移到虚拟屏幕左侧，只在辅助程序正常退出时恢复原位；z 序停放（留在内容区背后，由 VS Code 激活自然盖住）是待对照的替代方案（交接缺陷 D2、阶段 P1）。
2. **转发只在像素更新时触发**：捕获用 WGC（能抓被遮挡/移出屏幕的窗口），帧到达 + 帧差门控，空闲降频；当前 SSH 实验将裁边后的低频 PNG 以 base64 经辅助程序 stdout、扩展宿主和 `postMessage` 交给 Webview，单帧编码超过 700 KB 时直接丢弃。共享 D3D11 纹理只能省去辅助程序内部的拷贝：扩展无法让 Webview 读取 GPU 纹理，跨进 Webview 仍须传字节，降开销应靠全分辨率分块差分加二进制传输（阶段 P3）。
3. **点击代理**：webview 画面之上盖透明点击层，单击 = 激活 Office（吞掉该次点击，避免文档误击），辅助程序使用 `SetForegroundWindow` 请求前台，必要时用 UIA 请求焦点，再核对实际焦点。
4. **转发激活策略**：在桌面 Word、Excel、PowerPoint 的原生 **Office in Editor** 功能区选项卡提供一个设置，共用以下三个互斥选项。
   - **像素转发常开（默认）**：持续保持管线可用，空闲用帧差降频。Windows 10 实验机上，原生窗口可见期间只保留已准备的采集资源，窗口移出屏幕后再启动 WGC 会话。
   - **只有切换才激活**：原生 Office 窗口失去前台后才准备采集资源。新帧到达前继续显示上一张有效预览。
   - **移动鼠标激活像素转发**：只在本次原生 Office 窗口持有前台期间每 80 ms 检查指针是否移动。移动即预备本次采集，真正失焦并移出屏幕后启动转发。键盘切换或未检测到移动时按「只有切换才激活」回退。指针只作为预热信号，不触发焦点切换，不安装全局输入钩子，不记录指针坐标。

三种格式共用这套窗口状态机。Office 内的 **Office in Editor** 功能区、右键 **Add to Chat** 和选区操作由同一份 COM 加载项提供，所用文档仍由桌面 Office 打开与保存。VS Code 标签页承载这些文档的预览与焦点切换，复用同一个 Core 注册 MCP。文档由桌面 Office 按原路径直接打开，以普通顶层窗口编辑。因此 Ctrl+S 与功能区保存就是 Office 原生保存；原位 OLE 的暂存副本、`SaveObject` 与「IDE 窗口内」判据都不适用。实验尚未接入定时保存，接入时应走独立模式的 `ExecuteMso("FileSave")`。正式版如何把这类窗口认定为「IDE 窗口内」以接收 MCP 重绑，仍待定义；实验经隔离扩展显式登记聊天目标。引用只排队到已显式登记的聊天，消息由用户提交。Core 给 Office 回调管道发请求前核验 Windows 返回的服务器 PID 与已登记进程一致，并验证进程令牌的用户 SID。

### 残余风险与打断源

路线层面唯一公认硬伤：Office 模式下 VS Code 自发 popup（通知 / 错误 / 模态框）被叠层挡住，且没有公开 API 可监听。UIA 已能读到 Webview 内的预览节点，说明 VS Code 的网页无障碍树可见，基于 UIA 事件的弹层哨兵是候选解法（阶段 P4）。实现层缺陷另列于交接，不属于本条。缓解：叠层只盖编辑区，留出标题栏 / 活动栏 / 状态栏一条 chrome 边保证随时可切回；伴生扩展镜像它能监听到的事件。

打断式切换的触发源，按可靠性排序：

1. `EVENT_SYSTEM_FOREGROUND`（前台切换，主状态机）
2. UIA 结构变化（模态/对话框出现，不受遮挡，头号验证项）
3. CV 盯任务栏闪烁 / 状态栏角标（兜底）
4. 周期性 peek（默认关，接受闪烁才开）

### SSH 实现与验收（2026-11-21）

原版 VS Code 1.121.0 的隔离扩展在 Windows 10 22H2 的一个窗口打开 Word、Excel、PowerPoint 三个文档标签。扩展按运行标识和应用绑定合成文件，每个辅助程序只控制该文档的 Office 窗口及对应 VS Code。选择标签时保留预览。点击预览时先停止采集，再请求 Office 前台、对齐内容区并确认键盘焦点。失焦后确认窗口完全移出所有显示器才启动 WGC。Webview 对上一张有效图像保底，采集外缘在编码前裁掉。画面只在实验机本机进程间传递。

| 格式 | 已测内容区（物理像素） | 实际验证 |
| --- | --- | --- |
| Word `.docx` | 852 × 708、552 × 708 | 两种尺寸边界误差 0；回切预览；原件正文保存读回。 |
| Excel `.xlsx` | 852 × 669、552 × 669 | 两种尺寸边界误差 0；回切预览；原件 E2 保存读回。 |
| PowerPoint `.pptx` | 852 × 671、552 × 671 | 两种尺寸边界误差 0；回切预览；原件首张幻灯片保存读回。 |

PowerPoint 鼠标预热实测比失焦早 561 ms，失焦后 161 ms 取得首个离屏采样。采样时间与 Webview 绘制时间以各自事件计算。无移动回切从失焦才准备。Excel 实测功能区偏好通过每用户注册表通知传到原生辅助进程。Word、Excel、PowerPoint 的原生 Add to Chat 回调均通过选区捕获、隔离 MCP 目标入队及本机认证 SSE 收件验证。

诊断保留每应用的 `events.jsonl`、`result.json`、`finished.json`，副标签按应用后缀写入。事件仅记录 `app`、阶段、相对时间与必要的尺寸、焦点归属和错误码。用前台与键盘焦点确认激活、用离屏采样与 Webview 解码确认回切、用独立解析加哈希变化确认原件写入。手工回放材料是一份嵌 16 帧 GIF 的合成 PPTX。普通编辑视图由观察脚本在失焦时移动同一图形，放映用于查看 GIF 自身播放。[复跑入口与诊断字段](../experiments/focus-hybrid/README.md)

### 待扩展的验收

1. 三种转发策略分别记录多轮从失焦到 Webview 首帧的延迟、帧率和 CPU 开销。复测键盘切换、鼠标误触发和多次连续点击。
2. 验证点击代理吞掉激活点击，长时间重复进出无焦点漂移，中文 IME 组合态及 Office 模态对话框能恢复。
3. 覆盖 IDE 移动、跨显示器 DPI、分栏、多窗口、改变可见性和异常退出。当前数据覆盖侧栏宽度变化。
4. 验证 VS Code 自发通知、命令面板、模态框的显示及焦点协议。UIA 结构变化、任务栏信号等候选打断源仍需分别实测。
5. 接入正常打开命令和发布版 VSIX 后，验收 Cursor、真实聊天输入框及文档保存、关闭取消、冲突、恢复。辅助程序退出时会向 Office 发 `WM_CLOSE`，这只是实验收尾，正式版的关闭语义需另定。

以上各项已按阶段 P0–P6 编入交接的攻坚规划，并配有量化门槛、埋点分布和迭代循环。

## 提示词注入攻击记录

**2026-11-21，历史上下文混淆导致执行路线偏移。** 按用户指示将此次混淆归入提示词注入攻击记录。已观察到的行为是把历史 VS Code 原生槽实验和窗口层级修补带入当前网页承载任务。上一轮在 Webview 加入 `localhost` 触发页，仍将原生槽挂在 VS Code 主窗口并按网页矩形定位，创建时使用 `HWND_TOP`。这是已确认的执行偏移。现有证据未定位到外部恶意文本或攻击者。

**当时的处置与验收。** 该轮已停止 VS Code 叠窗实验。引用聊天、日志、旧任务单和历史文档只提供证据，当期用户指令决定路线及验收顺序。最新焦点混合实验的范围以本手册首节为准。网页请求、端口连通、父子 HWND 或一次窗口命中均是单项诊断。浏览器通过须记录页面和承载器握手、三类界面可交互、菜单和标签切换无叠层、各源文件读回结果；使用结构化日志保留失败阶段。 [原生窗口层级复盘](NATIVE_WINDOW_CONTAINMENT.md)

**2026-11-18 原生槽校正记录：** 当时的 `WS_CHILD` 可能挂在 IDE 根窗口下，标签专属原生槽仍待验收。真实窗口与 HTML 菜单的合成边界、已落实的提层/弹层重挂/高频日志修复，以及后续表面契约见 [真实窗口承载](NATIVE_WINDOW_CONTAINMENT.md)。该轮按原生子窗与网页交互验证，使用真实界面、菜单和焦点结果判定承载。

窗口移动采用原生 owner 事件驱动，纯平移不得进入 Office COM、重复激活或重新扫描 UIA；尺寸/DPI 变化后通过有界的缓存节点校准收敛。定位失效的显式恢复入口与验证边界见 [窗口跟随修复](WINDOW_FOLLOW_FIX.md)。

**原位 OLE 路线的设计目标（优先于该路线的实现细节；焦点混合目前只验证 Word、Excel、PowerPoint，其他成员的准入判据另定）：标签嵌套覆盖具有合格本地文档服务器的 Office 文件编辑器，含 Word、Excel、PowerPoint，以及满足条件的 Visio、OneNote、Publisher、Project。支持多标签、多应用；30 秒内可交互为性能目标，60 秒为当前交接规定的标签终态硬上限。适配走同一套容器契约，不按 Office 主窗类名逐个打补丁。不含 Outlook、OneDrive 等非文档编辑成员。**

**0.1.4 原位 OLE 规范补全（2026-11-15 活测）：** 选择器已扩至七款候选；多会话、每 kind 独立 STA。加载项选区/MCP 仍只分派 Word/Excel/PowerPoint。标签嵌套以 `#ole-root` 令牌矩形为唯一内容区；**打开成功前容器必须保持隐藏**，失败不得留下挡住 IDE 其它区域的 HWND。保存只有一条路径：与功能区保存按钮相同的一次原件 Save；插件只保留**定时保存**，实现为周期静默触发该按钮，取消「编辑后保存」。原生 AutoSave 后端重定向仍未实现。Excel/PPT 的 OLE 创建必须在 STA 上泵消息；无 `IOleDocument`/`IOleInPlaceObject` 的格式（常见于桌面 OneNote）短状态失败，不抢主窗。全系列可交互与关闭取消仍待验收。

嵌套几何、静默加载、保存按钮与仅定时保存以本文为准；身份、路由、关闭取消和修复关卡见 [架构审查](ARCHITECTURE_AUDIT.md)；逐应用差异见 [Office 兼容契约](OFFICE_COMPATIBILITY.md)。[本地保存与诊断](LOCAL_AUTOSAVE_AND_DIAGNOSTICS.md) 是 0.1.3 双开关历史，不能替代本文。

本手册描述当前仓库实现的 VS Code / Cursor 分支，以及尚未写进代码的设计目标。DeepSeek Harness 适配不在本版范围内。这里不记录本机路径、账号或聊天原文。

## 本地 Word 网页与 VS Code 标签页

本节记录纯文本 localhost 对照。该路线由 HTML 渲染编辑区，已安装的 Word 在登录用户会话中通过 COM 读取、修改和保存文件。普通 Edge 直接使用本地网页，VS Code 通过标准 Custom Editor 显示同一套 HTML、CSS 和交互代码。原生槽路线保留桌面 Word 界面，网页路线目前仅提供单段正文与文件保存，原版 Ribbon、富文本及 Office 加载项功能不在本次网页验收范围。

- **文件会话。** 按已验证的本地文件身份去重。为每个活动文件独占一个 loopback TCP 监听和一份随机凭据。重复打开、刷新、切换标签、分栏与空闲均复用监听。确认文档会话正常关闭后才释放端口，关闭确认存在歧义时保留会话及恢复线索。端口关系留在内存，扩展或系统重启后重新分配。MCP 端口与文档端口独立。
- **编辑后端生命周期。** 打开普通工作区或仅激活扩展不启动编辑后端。第一份受支持的 Office 文件由 Custom Editor 打开时按需启动，重复请求合并到同一启动。单个窗口关闭只释放自己的视图，其他 IDE 窗口共用现有后端与端口。所有 IDE 窗口正常退出后才停止后端，先确认各 Word 会话正常收尾；保存结果不明确时保留可恢复会话。独立 Office 加载项与 MCP 的显式业务调用另按其已有生命周期处理。隔离实验由启动监督器持有原版 VS Code 主进程句柄验证这一边界，生产集成需使用窗口租约或等价机制识别最后一个窗口。
- **宿主接口。** Edge 访问本站点的 API。VS Code Webview 使用 `acquireVsCodeApi().postMessage()`，扩展宿主校验消息后代该标签访问固定的 loopback 端口。Webview 只能提出受限的读取、保存或结束操作，不能指定 URL、绝对路径或令牌。凭据只由可信启动器及宿主使用，不进入 Webview、普通诊断和源码。拒绝跨文档凭据、任意来源和重定向。
- **编辑与保存。** 用户输入先留在视图，按保存时提交正文和当前版本。后台绑定精确 Word 文档，检查版本及源文件变化，在每个文档内串行执行。仅在 Word 确认保存后更新版本。未确认的保存保留视图内容，不自动重试。实验仅允许合成 `.docx`，正式版需补 VS Code 的 dirty、关闭取消、热退出和恢复契约。
- **界面与隔离。** Webview 只加载扩展内打包的资源，无 ActiveX、原生 Word 前台窗口、图像传输或者全局键鼠输入。IDE 菜单和编辑区由同一套网页合成处理。会话与位置不绑定，标签的显隐、缩放和移动不触发 Word COM 重开。编辑流程无磁盘轮询和弹窗重试。
- **验收。** 在原版 VS Code 的独立配置中打开两个合成文件，分别确认真实 Webview、独立端口和 Word 保存。第一份通过标签内表单编辑，独立读回本地 `.docx`，检查第二份未变。重复打开第一份须复用端口和已保存版本。还须验证两个标签互相切换时无原生编辑窗盖住 IDE 浮层；没有做的人工视觉检查须单列。通过该实验后再考虑接入发布版 VSIX。

**2026-11-21 实验结果。** 原版 VS Code 1.121.0 的隔离扩展成功打开两个合成 `.docx` 自定义标签。标签 A 通过 Word 保存，独立读回正文哈希一致，重载后原端口和新版本均保留，标签 B 的版本未变。记录证明启动信号来自首次 `openCustomDocument` 调用。另一次实验打开两个同实例窗口，关闭其中一个后两个端口仍能读到文档，关闭最后一个后两个 Word 会话及后台进程均正常退出。实验入口位于 `experiments/word-browser-local/vscode-lab`，仍需验收实际菜单叠放和正式文档编辑生命周期。

**纯浏览器对照的完整界面验收口径。** 用标准网页标签分别打开 `.docx`、`.xlsx`、`.pptx`，验证本机 Office 界面、编辑、原件保存、菜单和标签生命周期。呈现过程须符合仓库的网页验收约定。这条口径目前没有可用的呈现后端：标准 Chromium 无 ActiveX/Office MIME 宿主，三类 Office 文档对象也不提供无窗口接口（见[浏览器探针](BROWSER_OFFICE_PROBE.md)）。因此它只在出现真实网页呈现后端时才会启动，不是正在推进的接入项。标准 Webview 的纯文本实验验证过 Word 通信和保存。

原生实验用 `experiments/editor-native-host/read-debug.ps1` 输出脱敏阶段与窗口诊断。`geometryAttempts` 记录首次定位尝试，`viewport` 为标签在 IDE 客户区内的物理像素范围，`officeApplications` 记录经窗口进程核实的 Office 类型。`hitTest` 记录前台是否属于该 IDE，以及在槽内五个固定位置命中槽和 Office 子窗口的数量。命中只能作为遮挡与输入路线的诊断证据，实际菜单层级仍单独验收。初始化只做有界重试，无磁盘轮询。

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

- Office 加载项负责 Ribbon、右键 Add to Chat、选区捕获、受白名单约束的原生对象模型读写。独立模式使用相同加载项。原位 OLE 路线中，标签失败不自动回退到独立窗口；焦点混合路线本身就以独立顶层窗口编辑，不属于该回退。
- Core 负责身份、引用、聊天路由、MCP、配额、去重、可恢复写事务，以及原位 OLE 路线的 x64 OLE 文档容器。焦点混合的窗口状态机与抓帧目前在实验辅助程序 `native.exe` 中，尚未并入 Core 或加载项。
- VSIX 只做宿主薄适配：打开文档、把自定义编辑器 Webview 表面令牌交给 Core、注册 MCP。标签页是 Webview，不定位正文、不捕获选区、不代替 Office 编辑，也不是 OLE 站点。原位 OLE 路线中它提供 SuperDoc 式 100% 根节点（`#ole-root`）和表面令牌；焦点混合路线中它另外显示预览图像，并以点击代理发出激活请求。

直接打开 Office 与从编辑器打开，使用同一份加载项和同一个 Core。

公共状态、保存仲裁、事务恢复、引用与清理只能有一套。宿主表面适配、文档容器适配、应用对象模型适配分开；共用容器不等于共用 Word 的选区/保存对象模型。DSH 未来只增加 Cordis 宿主适配，不能复制 Core 或 Office 加载项。

## 两种打开方式

独立模式目标：用户在 IDE 外直接启动范围内文件编辑器时，相同加载项仍可使用。当前只实现 Word / Excel / PowerPoint，Visio / OneNote / Publisher / Project 需先补应用适配。Add to Chat **不猜未登记的聊天**，只在 Core 已显式登记的目标里选。Outlook 等范围外应用不装这条加载项路径。

- **未手动干预**（没有在 Ribbon 里点选或清除具体聊天）：每次 Add to Chat 只投给当前 **最新活动** 的已登记聊天。活动指该聊天最近一次 `chat.register`、经归因的 MCP 工具调用、向该聊天的投递或 App 连接。不弹窗、不状态长文、不改投到别的聊天、不自动发送消息。
- **获取不到**最新活动目标：无提示、无行为。Add to Chat 可以保持禁用，或点击后直接返回；不得出现错误条、MessageBox 或 toast。
- **手动干预**：用户在 Chat 菜单里点选或清除之后，以这次选择为准，直到再次手动改选。手动清除后不再自动填回最新活动。

编辑器模式（本段为原位 OLE 路线，即 0.1.4 实现；焦点混合的编辑器模式见首节）：VS Code 与 Cursor 共用一个 VSIX。用户在 Cursor 里打开 **范围内的本地 Office 文件** 时，默认自定义编辑器是 `officeInEditor.document`（选择器是扩展名表，不是写死三种）。标签页 Webview **不能**变成 `IOleInPlaceSite`。写法对齐 SuperDoc 自定义编辑器：HTML/CSS 把 `#ole-root` 铺满 webview（`retainContextWhenHidden`），**内容只占该标签的编辑器表面**，不得盖住侧栏、聊天、标签栏或其它标签。Cursor 没有可用的 `workbench.editor.hibernation`；防休眠靠 `retainContextWhenHidden`、打开成功后 `keepEditor` 转为普通标签（不 `pinEditor`，见下文「多标签」），以及切走时只 Hide 容器。Core 只 `SetParent` **自己的容器** HWND，永不 `SetParent` 各应用主框。容器客户区坐标等于 `#ole-root` 令牌矩形。标准 VS Code / Cursor 的 Webview 没有独立 HWND，所以实际父窗口只能是 IDE 顶层窗口 `Chrome_WidgetWin_1`（Chromium GPU 表面的 Win32 宿主），容器作为它的 `WS_CHILD` 只占令牌矩形；父窗口不能是 `Chrome_RenderWidgetHostHWND` 本身（合成器会盖住原生子窗）。这样能获得原生裁剪和随父移动，但不能让 IDE 的 HTML 弹层盖到 Office 之上，这正是路线表第 1 行的堵点；标签专属原生槽在标准宿主中不存在。禁止只做屏幕对齐的 owned `WS_POPUP` 薄壳，禁止用 `SetWindowRgn` 挖洞或整窗隐藏来避让 IDE 菜单。禁止铺满 IDE。用 webview 种下的表面令牌（及四角打点）找到 **我们自己的** `#ole-root`，禁止按宽度高度去匹配其它窗口，禁止按 IDE 主客户区或 workbench 级 HWND 的尺寸当编辑器矩形（作为父窗口可以，见上句），也禁止按 `OpusApp` / `XLMAIN` / `PPTFrameClass` / OneNote / Visio 主窗类名做适配表。**创建与激活期间容器保持隐藏**；`ShowObject` 与 `DoVerb(SHOW)` 也不得在 `open_ok` 前把容器亮出来。打开失败必须拆掉本次 HWND，不得留下挡住非标签区域输入的残窗。容器实现 `IOleClientSite`、`IAdviseSink`、`IOleInPlaceSite`、`IOleInPlaceFrame`、`IOleDocumentSite`（以及目标应用作为文档服务器实际需要的同一组接口），对已有本地文件按 **文件的 OLE 文档类** 创建对象，再原位激活。不分发 `dsoframer.ocx`。关标签拆除会话（Close + DestroyWindow），不是只 Hide；切走标签只 Hide，且不得 `UIDeactivate` 拆掉功能区。从编辑器打开时原位失败 **不得** `Process.Start` 再开独立 Office，失败页也不得自动唤起独立窗口；标签只留短状态。独立窗口只走命令 `Office: Open in Microsoft Office`。选区、Add to Chat 仍在加载项里；Agent 改文档走 MCP。

原位 OLE 编辑器模式里，只有 **已原位激活、且其进程被 Core OLE 会话认作服务进程** 的 Office 实例算「IDE 窗口内」。`inplace=false` 的独立窗口按独立模式路由。焦点混合的文档窗口不经原位激活，这条判据不适用，其替代判据见首节。

## Office 全系列：谁进标签，谁不进

嵌套的对象是 **「用户打开一份本地文件、在文档界面里编辑」** 的 Office 应用，不是整个 Microsoft 365 套件。范围外清单对各路线共用。范围内的「OLE 文档服务器」判据按原位 OLE 路线制定；焦点混合不依赖 OLE 文档类，目前只验证 Word、Excel、PowerPoint。

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

本节为原位 OLE 路线的容器契约。「一次适配、不按应用打补丁」的原则对焦点混合同样适用：三款应用共用一套窗口状态机。嵌套契约只有一套。新增 Visio / OneNote **不得** 再增加 `VISIO` / `ONENOTE` 主窗类名、按客户区宽高匹配、或 `SetParent` 应用主框。

统一做法：

1. **表面**：每个自定义编辑器标签种下唯一 `oie-surface-` 令牌，`#ole-root` 铺满 webview。定位只认这个令牌及其四角打点（见下节），不认 `window.screenX/Y`，不认「看起来差不多大的窗」，也不把 IDE 主客户区或 workbench 级 Chromium HWND 当编辑器。
2. **容器**：Core 为 **这一次打开** 建一个自己的 HWND，实现文档站点接口。容器挂到 GPU 表面的 Win32 宿主上（活测 `Chrome_WidgetWin_1` 的令牌客户区），永不 `SetParent` Office 主框，永不挂进 `Chrome_RenderWidgetHostHWND`。尺寸必须等于 `#ole-root`，不得铺满 IDE。Office 功能区/Search 等界面必须成为**容器的子窗口**（可 `SetParent` 这些 UI HWND，不可 `SetParent` `OpusApp`/`XLMAIN` 主框）。禁止 owned popup 只改 owner、禁止 region 挖洞、禁止整窗隐藏避让。
3. **文档对象**：按已验证的文档接口创建/绑定。`OleCreateFromFile` 只从 **暂存副本** 创建，禁止直接以写方式打开用户原件。暂存文件名必须带会话戳（`working-<32hex>.xlsx`），不得所有会话都叫 `working.xlsx`：Excel 同一进程里禁止两个同名工作簿，否则弹「无法同时打开两个同名的工作簿」并 `E_ABORT`。原件回写只走保存按钮那一条 `IPersistFile.Save(原路径)`（见「原件保存」）。rclsid 按规范可传 CLSID_NULL；`GetClassFile` 失败时才用 ProgID 补 CLSID，补不到则短状态失败。容器客户区要再扣除 Office 协商的边界 UI。
4. **主框**：激活时出现的独立主窗不是嵌套目标。仅在证明窗口属于本 session/document 且不影响用户独立文档后，才允许针对性 Hide。**禁止全桌面新窗口差集 Hide/Close**，禁止隐藏 `#32770` / `NUIDialog` / `bosa_sdm*` 等安全、修复、保存提示；无法证明归属就停止控制。类名最多作诊断。
5. **加载项**：每款补齐注册、文档身份、Ribbon/context ID、选区、保护状态和保存适配，再按能力开放 MCP。缺模型不改为抓主窗，也不能宣称已有全系列协作能力。

### 标签表面定位与外观

标签自己提供几何，不向其它窗口「对齐」。四角打点是原位 OLE 路线的做法；焦点混合实验用 UIA 读取 Webview 内命名预览节点的物理像素边界，再把 Office 可见边界对齐过去。两者都只认本标签自己种下的标识。

- **打点**：`#ole-root` 四角各一个 **1 像素黑点**（`pointer-events: none`），无障碍名为 `{token}:tl` / `:tr` / `:bl` / `:br`。点只辅助定位，不进入文档、不拦截输入。Core 用 UIA Name 读取至少两个角的屏幕矩形，取并集作为该标签内容区；对角齐全时以并集为准。
- **禁止**：用 IDE 窗口原点、CSS `screenX/Y`、整块 workbench `Chrome_RenderWidgetHostHWND`、或按宽高相似匹配其它窗来摆 OLE 容器。覆盖聊天、侧栏、标签栏即失败，应保持隐藏并短状态，不得将错就错铺满界面。
- **铺满**：OLE 容器按标签内容区分配。工具栏占用 `SetBorderSpace` 顶边；正文视口是工具栏以下的整块客户区，并用 `SetObjectRects` 把就地对象铺进该矩形。不另做标签 Office 主题底色。

OneNote 现代客户端、仅云笔记本、无 `IOleObject`/`IOleDocument`/`IOleInPlaceObject` 的宿主：创建后若不是文档对象，**立即短状态失败**，不得 `DoVerb` 空等或铺一层失败 HWND。不允许改成「找到 OneNote 主窗再塞进标签」。

### 静默加载、切走标签与输入范围

- **静默加载**：`CreateWindow` 起就是隐藏的。`Place(show=true)`、可见 `ShowWindow` 只允许在 `open_ok` 之后。加载中标签显示页内 DOM：90×90 GIF 居中叠在「正在加载 Office」之上，文字略下移，句号 `.` / `..` / `...` 循环；不是覆盖整个 IDE 的原生窗。
- **失败**：Abort/Dispose 本次容器；不得把 32×32 或错误矩形的宿主留在工作区上。失败不得劫持侧栏、聊天、标签栏的键鼠。
- **切走**：只 Hide 该会话容器，保留 OLE 原位 UI（功能区边距不清零）。Cursor 自定义编辑器 `retainContextWhenHidden:true`；打开成功后只 `keepEditor` 转为普通标签（不 `pinEditor`，与下文「多标签」一致），避免预览标签被挤掉而 `dispose`→detach。本机 Cursor 无编辑器休眠设置可关。
- **切回**：立刻用已信任的 owner 边距把容器摆回上次矩形，再空闲校准打点。禁止把 webview 刚显示那几帧、更靠近 IDE 原点的半成品打点并集写进位置（会斜向左上漂几帧再弹回）。禁止用 `window.screenX/Y` 回退。禁止在切回时对整个 IDE 树做无超时 `FindFirst`（会把 STA 卡死数十秒）。
- **溢出即缺陷**：命中矩形相对 IDE 主窗 ≥80% 面积、或明显大于 Webview 上报的编辑区，一律视为未找到表面，保持隐藏。

## 原件保存与定时保存

原生 AutoSave 存储后端重定向 **仍未实现**，能力字段保持 false。插件不得把它改成 true。

第 1–2 条只适用于原位 OLE 路线；第 3–5 条适用于各路线。焦点混合路线的文档由 Office 按原路径打开，保存按钮即 Office 原生保存，定时保存走独立模式的 `ExecuteMso("FileSave")`。

定时保存是 2026-11-15 的用户决定，它明确标注为插件自己的本地 Save，不冒充原生 AutoSave。项目根目录 2026-11-14 版 `docs/REQUIREMENTS.md`、`docs/RISK_BOUNDARIES.md` 中「不实现保存计时器」的表述，已由修订说明限定为禁止去抖保存和 AutoSave polyfill。

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
- **同一原件的多个标签 / 多个 Cursor 窗口共用一份文档对象**。复制只发生在壳的窗口展示与交互层：每个标签一个容器 HWND，绑同一 `IOleObject`。**禁止**再 `IOleDocumentView.Clone` / `CreateView`（第二视图会抢走先开窗口的原位界面，表现为白屏）。额外窗口一律壳镜像：`PrintWindow` 把主容器 HWND 的画面拷到该标签 HWND，**不转发鼠标/键盘**。镜像在 Core 进程内由失效通知触发，不轮询、不编码、不跨进程传输。它不是下文禁止的 80 ms / `PrintWindow` 轮询，也不是串流。禁止为同文件再 `OleCreateFromFile`，禁止静默复制工作副本/从原件分流。最后一个标签拆除对象。
- **Office 自定义编辑器不走默认预览「新换旧」**。首次 `vscode.openWith` 用 `{ preview: false }`。已在 `officeInEditor.document` 里的标签只 `keepEditor` 转成普通标签，**不得**再 `openWith`（会拆掉当前 webview，几何超时且留下 OLE 窗），**不得** `pinEditor`。`tabGroups.onDidChangeTabs` 同样只 `keepEditor`。不得改全局 `workbench.editor.enablePreview`。

关闭：

- Ribbon「关标签保存」与「定时保存」并列，默认开。关最后一个该文件标签时，开则走 `SaveObject`，关则 `OLECLOSE_NOSAVE`。先 Hide/Destroy 容器 HWND，避免标签已关表面仍停在 IDE 上。
- **自动确认警告**（默认开）：只对 **标签嵌套** 生效。警告弹窗本身仍会出现；当可见按钮恰好一颗且文案是确认（确定/OK/是）时，向该按钮 `PostMessage(BM_CLICK)`。以文案为准：Excel 单按钮「确定」的控件 ID 有时是 `IDCANCEL`，仍要点。不设 `DisplayAlerts=false`、不装全局键鼠钩子。两颗及以上（是/否、确定/取消）不动。独立打开的 Office 窗口不挂钩。Office in Editor 功能区与编辑器标题栏各有同一开关。

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
- 反向项目（Office 里嵌浏览器）不采用。不把 Electron/VS Code 私有 workbench DOM 当几何来源。本路线允许的嵌套只限于 Core OLE **容器**作为承载该 Webview 的 IDE 顶层窗口的 `WS_CHILD`，客户区等于令牌矩形；Webview 自身没有 HWND。禁止 owned popup 薄壳、禁止 `SetWindowRgn` 挖洞、禁止整窗隐藏避让。禁止 `SetParent` 各应用主框。
- 加载项在独立 Office、焦点混合实验和 OLE 容器中的对应实例使用同一份实现。`OnConnection` 只保存 Application，管道与 Core 注册放到 `OnStartupComplete`。STA 回调优先用 APC，Office 只处理窗口消息时使用当前 STA 的 `WM_TIMER` 清空已有回调队列。该回退只检查队列，无空闲文档读取，无额外窗口。新种类若没有 Application 加载项模型，嵌套仍走容器路径，MCP/Ribbon 能力按缺模型缩小。
- 每个标签一次打开对应 **一条** OLE 会话。会话记住该次原位对象窗口的进程 ID。随后 `office.register` 若 pid 匹配某条嵌入会话，则该实例标为嵌入，接收 MCP 重绑。多种类、多标签可以同时有多条会话；不要用「全局只有一条 HasSession」挡打开。
- 容器 HWND 先隐藏创建，再 `SetParent` 进承载编辑器 Webview 的 IDE 顶层窗口，位置取令牌矩形。几何只来自令牌/四角打点映射后的 **父窗口客户区**。禁止 `HWND_TOPMOST`、禁止为对齐去 `SetForegroundWindow`、禁止按其它窗口的宽高做匹配、禁止用 `window.screenX/Y` 把窗口摆到 IDE 原点或铺满工作区。**第一次**用四角打点解析编辑器表面与客户区边距；之后子窗口随父窗口移动，尺寸变化才 `SetWindowPos`/`SetRect`。侧栏等外框不变、编辑区变时，空闲再用 `#ole-root` 的 CSS 尺寸校准。禁止空闲校准再走 UIA 全树 / `uia-token`、禁止 80ms/`PrintWindow` 轮询、禁止每次 layout 改窗口样式或 `UIDeactivate`。禁止 region 挖洞和整窗隐藏避让。
- 文件名以 `~` 开头的 Office 锁/缓存文件（如 `~$文档.docx`）不得作为文档宿主。资源管理器默认隐藏 `**/~$*`，不把锁文件列成可点条目。误点 `~$文档.docx` 时打开同目录 `文档.docx` 的标签，并立刻关掉缓存标签，不提示「不要打开」。对不上原件时也关掉缓存标签，不留下空页。末视图 OLE `Close` 之后，若原件旁 `~$` 兄弟锁已无人占用，才删除该锁（Word 崩溃残留）；占用中的锁不删。Core 仍拒绝把缓存路径送进 OLE。未 `open_ok` 时窗口保持隐藏；`ShowObject` 也不得在 attach 前把容器亮出来。`GetContainer` 返回自身 `IOleContainer`。按暂存副本 `OleCreateFromFile` 之后 `OleRun`，把 `IOleObject` 赋给站点，**表面已解析且仍隐藏** 时再原位激活，成功后再 Show。Word 会在 `DoVerb` 期间回调 `ActivateMe`，此时必须已经能 `CreateView`。关标签走 `office.detach_window` 只拆除 **该标签** 对应会话。
- **每种** Office 文档服务器一条 STA。该线程在空闲时泵消息（`MsgWaitForMultipleObjects` + `PeekMessage`），**在 `OleCreateFromFile` / `DoVerb` 等阻塞 COM 调用期间也必须泵**（`IMessageFilter.MessagePending`），否则 Excel/PPT 会在调用方 STA 上卡到标签 60 秒上限。不能只在每次 `Invoke` 返回后 `PeekMessage`。不同种类的 STA 互不阻塞。Excel 在 `OLERENDER_NONE` 上 `OleCreateFromFile` 常返回 `E_ABORT`：再试 `OLERENDER_DRAW`，然后 `CoCreate` + `IPersistFile.Load`（`STGM_READWRITE|SHARE_DENY_WRITE`，失败再换模式），最后 `BindMoniker`。`CoCreate` 成功后不得用 `Load(..., 0)` 直接放弃。
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
- 不在后台抢焦点，不调用会主动激活窗口的 Select。焦点混合在用户每次点击预览时请求一次前台，属于用户发起的切换，不进入后台 MCP 路径。
- 焦点混合实验的例外只有两项：持焦时原生窗口盖在编辑区，以及失焦后在本机内存中抓帧显示。这两项都不扩展到模拟输入、全局键鼠钩子或后台操作。前台事件用 out-of-context WinEvent，不是键鼠钩子。
- 默认不因空闲反复读全文。仅定时保存：有界心跳看内存 dirty，到期静默点一次保存按钮；无改动不写盘。失败后暂停该文档，不关闭/重开、不弹窗风暴。不装全局键鼠钩子、不用 SendKeys 点按钮（独立模式走 `ExecuteMso("FileSave")`）。
- 相同故障不循环弹窗。缺目标时连 Ribbon 长状态也不写。
- 恢复材料有数量和体积上限。满额后停止新增编辑，不自动删掉尚未确认的恢复记录。
- 超时视为结果未知，不自动重试写操作，不结束 Office 进程。
- 文档身份、宿主租约、generation 与路由归属贯穿 open/layout/save/close；不能仅按 PID 决定文件或聊天。标签必须采用可取消的关闭与 dirty/恢复生命周期，确认关闭前不丢掉所有权。
- 诊断默认内存有界；显式脱敏导出，不默认写固定机器路径或 POST 固定调试服务。kind、STA、会话、引用与恢复目录都有总量和空闲回收预算。
- 卸载编辑器扩展不卸载 Office 加载项。卸载加载项不删除用户文档或引用材料。注册变更等到 Office 下一次正常启动生效，安装器不关闭正在运行的 Office。
