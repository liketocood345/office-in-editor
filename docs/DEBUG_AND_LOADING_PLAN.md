# 调试与 30 秒加载优化方案

审查日期：2026-11-15。范围：当前工作区源码、交接文档、设计手册、隔离测试与本机载荷指纹。本文是待实施方案，不代表优化已经完成。

后续源码修复、实际通过的检查及仍未完成的阻断项见 [修复状态](REPAIR_STATUS.md)。下文问题表和版本状态保留为修复前审查快照。

## 结论与基线

当前应先解决保存、会话归属和迟到打开问题，再优化耗时。直接修改超时数字不能可靠实现 30 秒加载，反而可能掩盖后台仍在打开 Office 的状态。

交接中的原生验证仅证明临时文档能在 Core 自有 OLE 容器中原位打开；尚未证明经 VSIX 打开、对齐真实标签、切换、保存和关闭的完整链路。Word、Excel、PowerPoint 必须分别验收。

本轮没有启动或结束 Office/Core，没有移动现有窗口、写注册表、替换安装文件或操作真实文档。保留已有未提交改动，只新增本文。隔离测试的构建产物和日志属于忽略目录。

- 仓库基线：HEAD `a0618fd`，工作区存在未提交功能改动，不能把该提交号当成完整运行版本。
- 已安装 Core DLL 与 `payload/core` DLL 指纹不同；已安装 Office 加载项 DLL 与 `payload/office` DLL 指纹也不同。EXE 相同不代表组件一致。指纹亦与旧交接的部分记录不一致，性能验证前需重新冻结一组构建。
- 类型检查通过；JS 测试 17 项、Core 检查 59 项、加载项 smoke 122 项断言通过。这些未覆盖真实标签打开/保存/关闭。
- 隐私扫描未通过：`packages/shared-core/Tests/Program.cs:35–37` 有三处固定 Windows 测试路径；它们不是发现的用户文档路径，但仍触发仓库规则。修正夹具后再发布，不能因此放宽整体扫描。
- 不执行当前 `extension-host.cjs` 的真实打开部分：它声称假文档不会送入 Office，但现有命令已进入 OLE 路径。须先注入假 broker 或换成明确隔离的原生测试。

## 已核实的问题

行号对应本次审查快照；实施时按方法名复核。

| 编号 / 优先级 | 源码证据 | 影响与修正方向 |
| --- | --- | --- |
| D01 / P0 | `OleDocumentHost.cs:107,194–201,209–213,344` | 新开先关闭全局旧会话；关闭用 `OLECLOSE_NOSAVE`；存储只有内存 IStorage，SaveObject 不保存却返回成功。可能丢弃未保存编辑，也未证明编辑写回原文件。先建立真实保存与可取消关闭契约。 |
| D02 / P1 | `extension.ts:279–307` | 打开未返回时关闭标签，因 officeId 为空直接返回，closed 未置位；晚到打开仍 attach。dispose 必须先失效 openId，再处理迟到结果。 |
| D03 / P1 | `api.ts:68–74`、`protocol.ts:5–7`、`OleDocumentHost.cs:235–240`、`Program.cs:106,115` | office.open 未传专用 timeout，实际走默认 10 秒；60 秒常量未使用。Core 55 秒 token 不会中断同步 STA/COM。超时可能丢失会话 ID，后台仍继续打开。 |
| D04 / P1 | `OleDocumentHost.cs:32,84,107,188`、`extension.ts:124–135` | Core 只有全局 active，IDE 可开不同文件标签。旧 detach 执行时又关闭全局新会话。必须按宿主、标签、文档、generation 定位，在 STA 实际执行时复核。 |
| D05 / P1 | `extension.ts:273–305,520–539`、`CoreState.cs:68,199–204` | 打开中隐藏被忽略，完成后仍显示；layout 并发乱序；停用未统一 detach；host 删除不关闭 OLE，HasSession 阻止空闲退出。需要唯一会话所有者与有序销毁。 |
| D06 / P1 | `extension.ts:286–298`、`ui.ts:71–75`、`host-window.ts:8–17` | HTML 先于消息监听注册；几何仅初次和 resize 上报，无确认/重发/重建握手。移动窗口和 DPI/编辑器缩放未处理，screenX/Y 不等于已验证的原生标签内容矩形。 |
| D07 / P1 | `OleDocumentHost.cs:135,155,308,318–324,408–426` | HRESULT 多处忽略，激活后再次 ActivateMe 可能重复建 view，未持有/关闭 view；未激活也可能返回 inplace=true。不能以 attached 响应代替可交互完成。源码只有一次显式 OleRun。 |
| D08 / P1 | `OleDocumentHost.cs:117–137,160–164,209–213,326–332` | ILockBytes 原始接口引用未释放；storage/object 在较晚阶段才转交 site，早期失败可漏释放。要逐项建立 COM 引用/句柄所有权与失败清理。 |
| D09 / P1 | `BridgeProtocol.cs:173–204` | APC 取消仅取消 TCS、不移除 pending；队列无上限。普通 APC 依赖 Office 线程进入 alertable wait，无时限交付保证。还需保护卸载后迟到回调的委托生命周期；这不等于已证明历史崩溃原因。 |
| D10 / P1 | `OfficeAddin.cs:247–255,278–292`、`CoreState.cs:226,385–402`、`McpServer.cs:79–111` | MCP 嵌入绑定随后可能被全局 defaultTarget 覆盖；只按 PID 判嵌入会混淆同一 Word 的独立文档。路由逐实例等待最长 55 秒并处在工具前置链上。改成文档级归属、路由版本与独立短期限。 |
| D11 / P1 | `OfficeOperations.cs:146`、`DocumentOperations.cs:113` | Core applyEdit 未传 recovery，但加载项 Apply 强制要求，真实跨层编辑会在执行前失败。独立 fake 测试未发现契约错位，须补跨层契约测试。 |
| D12 / P1 | `WindowHost.cs:103–155` | 备用窗口路径重新寻找当前 HWND，可能恢复错误窗口；RestoreAll 可能重新显示已 detached 窗口。固定原会话 HWND/PID/TID，成功拆除即移除记录，禁止枚举首窗兜底。 |
| D13 / P2 | `OfficeAddin.cs:367–377` | COM/Ribbon 入口同步检查文件并写 load-trace；虽有约 64 KiB 截断，仍有额外频繁 I/O。默认改内存计时，显式诊断才有限导出。 |

### 两个先行技术关卡

**保存关卡。** `OleCreateFromFile` 从文件内容创建嵌入对象，并不证明它原位编辑源文件。先用合成文档完成“打开→改一个标记→显式保存→正常关闭→独立重新打开源文件比对”。校验实际文档身份及目标路径，但不将它们写入诊断导出。若原文件未更新，应重新选择原文档加载/持久化方案；不能把复合 IStorage 直接当作 docx/xlsx/pptx 覆盖。仅改成 SAVEIFDIRTY 也不能替代 SaveObject 的真实持久化实现。[创建语义](https://learn.microsoft.com/en-us/windows/win32/api/ole2/nf-ole2-olecreatefromfile)、[SaveObject 契约](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-ioleclientsite-saveobject)、[关闭枚举](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/ne-oleidl-oleclose)。

**标签承载关卡。** 当前 `OleDocumentHost.cs:295–297` 创建无 parent/owner 的 Core 顶层工具窗，通过屏幕矩形对齐。它是窗口对齐原型，尚未建立编辑器原生标签承载关系。既有风险边界又禁止悬浮窗覆盖编辑区，这两处需要统一技术结论。必须证明窗口移动、遮挡、最小化、多屏、切标签和宿主退出均受正确生命周期控制，并确认使用公开扩展能力；若做不到，不能仅凭加载快就宣称已完成规范内嵌。禁止转向私有 workbench DOM、跨进程 SetParent、输入钩子或强制置顶。[扩展能力边界](https://code.visualstudio.com/api/extension-capabilities/overview#restrictions)。

## Debug 执行顺序

### 1. 冻结运行版本

记录构建 ID、源码快照、dirty 状态、宿主/Office 版本与架构、已加载 VSIX/Core/add-in 的组件摘要。添加诊断能力上报 buildId，不使用同一个 0.1.0 区分全部构建。源码编译通过不等于已安装文件相同；DLL 已在 Office 中加载时，磁盘上的 DLL 摘要也不能证明内存版本。

需更新时遵循交接：Office 正常退出后才替换加载项/注册；Core 采用自包含发布。不能在本次审查中自动替换安装来消除差异。

### 2. 先用 fake 重现状态错误

注入假 broker、panel、STA 调度器和可控时钟，避免用真实用户文档重现破坏性路径。

- open 延迟 15 秒：第 0.1 秒关闭标签，结果到达后不能显示、挂接或重开第二份文档。
- open 中切换标签：结果到达保持隐藏；恢复可见后只应用最新矩形。
- 连续 A/B 打开与旧 A detach：B 不得被旧请求关闭；A 的未保存内容不得被静默丢弃。
- 无初次 geometry、Webview 隐藏后重建、旧消息迟到：有限握手重发，旧 generation 无效。
- resize/show/hide 乱序：最多一个 layout 在途，保留一个最新待发值，hide/close 优先。
- deactivate/host 失联：只能拆自己的会话；不通过 connect 重新启动 Core 做清理。
- COM 假阻塞超过期限：UI 进入准确的超时/结果未知状态，不能自动重试 open 或重放写入。
- APC 永不交付、取消及 Dispose 后迟到：pending 回归基线，迟到 thunk 安全 no-op。
- 跨 Core/add-in 跑 prepare/apply/revert/save 契约；真实生产 handler 配 fake Office，验证 recovery 等字段实际传递。
- 路由绑定 A 后收到旧刷新 B、同 PID 两份文档：A 不被全局最新聊天覆盖，独立文档不被错误重绑。

### 3. 加一次性阶段计时

每次打开使用同一个随机 openId；IPC 消息另有 requestId，并携带 hostId/editorId/generation。状态记录放有界内存，不记录文件名、绝对路径、正文、选区、聊天原文、token、管道名。

| 计时点 | 解决的疑问 |
| --- | --- |
| 打开命令、activate 进入/完成、provider 进入 | 是否耗在扩展激活，还是已经进入标签加载 |
| HTML 安装、首个 rAF、layout 发送/ACK | 首屏慢、监听竞态或 Webview 重建 |
| endpoint、ensure-broker、host.register | Core 冷启动、连接恢复、注册等待 |
| STA 入队/出队、当前队列长度 | COM 本身慢还是前一个请求堵住 |
| storage、OleCreateFromFile、OleRun | 对象创建和 Office 冷启动耗时 |
| DoVerb、ActivateMe、CreateView、每个 HRESULT | 重复激活、失败被吞、视图是否创建 |
| native view HWND、非零矩形、首次有效绘制 | 容器出现是否真的包含文档视图 |
| visible commit、ready ACK、close 请求/完成 | 真正交互完成、迟到显示和清理 |
| APC 入队/执行、路由更新版本 | 插件调度慢是否被误算成 Office 启动慢 |

跨进程不要直接相减独立的 Stopwatch 时间戳。由 VSIX 用单调时钟测端到端，各进程报告本地 elapsed 和阶段时长；每个接收端只缩短剩余预算，不能重置整段期限。

原生 readiness 至少要求真实 view、正确矩形、激活成功和首次有效绘制；最终用人工输入/保存验证可编辑。不得为了测“可交互”注入键盘输入或循环强制聚焦。容器自身 WM_PAINT、inplace=true 和 postMessage 均不足以单独证明完成。

### 4. 最后进行原生测量

仅使用独立合成的有效 docx/xlsx/pptx，真实用户文档不参与故障注入。Word 先过保存关卡和生命周期，再逐项覆盖 Excel/PowerPoint。已运行的用户 Office 不强杀；所谓冷启动应等用户正常退出后测量，不能为数据干净擅自结束它们。

分别记录 VS Code/Cursor、Core 冷/暖、Office 冷/暖、单窗口/双 IDE、100%/150%/200% DPI、不同缩放、多屏与负坐标。首轮每格至少 20 次，保留样本数、P50/P95/max 与失败率；20 次只能作为初始基线，发布前扩充。保护视图、密码、恢复对话框、外链与 Office 忙状态另列异常组，不混入健康小文档指标。

## 30 秒优化设计

### 指标定义

1. 标签首屏目标不超过 1 秒：出现简短状态，其他内容维持简洁。不等 Core、聊天列表或 Office。
2. 健康本地验收样本的原生可交互目标：端到端 P95 不超过 30 秒；超时/失败不能从成功率统计中删除。暖启动进一步以 8 秒以内为初始目标，均待实测。
3. 所有打开尝试在 30 秒内得到确定的 UI 反馈：ready、failed、cancelled 或 outcome_unknown。**超时反馈不等于 Office 工作已停止，也不算成功加载。** 对任意损坏文件、模态对话框或无响应 Office，不能承诺成功加载硬上限。

初始预算如下，累计时间从打开请求发起计算；应以计时结果调整瓶颈分配。

| 阶段 | 累计目标 | 实施要点 |
| --- | ---: | --- |
| 轻量激活、标签首屏 | 1 秒 | 注册命令/provider 后返回；dashboard/聊天列表移出打开关键路径 |
| 有效几何与 Core 连接、登记 | 3 秒 | 通过信任/文件检查后并行；监听先于 HTML；ready/layout ACK |
| OLE 创建、必要的 OleRun | 20 秒 | 隐藏创建；区分排队与执行；检查 HRESULT；不重复启动 Office |
| 一次视图激活、定位、首绘检查 | 26 秒 | 复用唯一 view；杜绝无条件第二次 CreateView；验证实际显示 |
| 可见性提交与 ready ACK | 28 秒 | 检查 openId/generation/visibility 后提交；去掉冗余 attach RPC |
| UI 终态与调度余量 | 30 秒 | 给出短状态，停止新动作；不把强制关闭/清理当成保证完成 |

### 跨层协议与状态机

建议将耗时同步 office.open 拆为 begin/status/cancel 与完成事件，方法名实施时统一进入 contracts/SDK；查询是事件丢失后的有限补偿，不建立永久轮询。

```text
created → connecting → queued → opening → activating → ready
                 ↘ failed / cancelled / outcome_unknown
ready → hidden / ready → closing → closed
```

- VSIX 在请求发出前生成 openId，并持有文档会话表；closed 立即禁止后续显示。
- Core 的会话键包含 host/editor/document/generation；可先采取明确单会话策略，拒绝第二次打开，绝不能静默替换旧 dirty 文档。
- Core 保留 openId 的有界终态记录，重复请求查询同一结果。丢响应不能触发第二次 open。
- 每阶段检查剩余 deadline；未执行工作可取消并移出队列。只有 Core 已确认取消且未执行时，才报告安全 cancelled。
- IPC 接收不能同步堵在 STA Invoke 上；采用异步完成源并保留持续 STA 消息泵。Task.Run 或 WaitAsync 只改变等待方式，不使 COM 获得强制取消能力。
- 在途同步 COM 取消只能作为受测增强：取消不证明服务端回滚。状态未知时封锁重放/再次写入，回返后重新验证 generation 并处理迟到资源。[COM 取消机制](https://learn.microsoft.com/en-us/windows/win32/com/canceling-method-calls)。
- 超时不能杀共享 Word、TerminateThread、清 AutoRecover、强制丢弃 dirty 对象或无限创建新 worker。若尚不能防止在途调用的迟到窗口，相关路径仍视为未达生命周期门槛。
- 保存/关闭不是 30 秒打开预算下可以牺牲的数据。关闭失败需保留恢复能力及准确 unknown 状态，不能先删记录再宣称 closed。

### 减少可避免的等待

- HTML/几何与 Core 连接并行；MCP 注册/dashboard 加载不应阻塞标签首屏，但实际 Office 打开仍须完成信任检查和宿主所有权登记。
- Webview 使用带 generation 的 ready/layout/ACK；监听先注册，首次有限重发，可见性恢复重新握手。hidden 时 iframe 生命周期见 [VS Code Webview 文档](https://code.visualstudio.com/api/extension-guides/webview#visibility-and-moving)。
- 屏幕坐标必须带坐标空间、DPI/缩放和来源元信息；先测真实内容矩形，不直接将 screenX/Y 乘 devicePixelRatio 当作修复。[screenX 使用 CSS 像素](https://developer.mozilla.org/en-US/docs/Web/API/Window/screenX)。不通过访问宿主私有 DOM 获取几何。
- layout 最多一个在途、一个最新待发送值；重复矩形不发，拖动时限频，close/hide 优先。空闲不运行定时布局轮询。
- 保存唯一 IOleDocumentView，按 HRESULT 决定是否需要备用激活；保持 Object 在 DoVerb 回调前可见。释放顺序覆盖 view、object、storage、ILockBytes、HWND、窗口类。
- MCP 路由在 Core 先同步提交版本化的权威目标；加载项推送单独短期限（初始 500 毫秒）并合并最新值，不等待所有实例。Add to Chat 必须按该文档当前权威模式解析，不能因为推送慢就发给旧目标。
- APC pending 设硬上限，取消立即移除 callback 工作；限制未交付 wake APC，不断路重试。固定进程生命周期的 native callback thunk，旧 token 只 no-op。[普通 APC 的 alertable 条件](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-queueuserapc)。
- 不通过全局输入钩子、模拟按键、提高优先级或强制抢焦点加速加载。加载项 STA 调度方案必须保留交接的崩溃防线，不重引入 WinForms Control 或托管 HWND_MESSAGE。

## 无积累与无额外频繁 I/O 的验收

- 初始诊断预算：内存环形缓冲最多 100 次打开、每次最多 64 个阶段事件、总计最多 1 MiB；满时覆盖旧计时，不能覆盖恢复材料。显式导出单份且有大小上限，默认无持续写日志。
- Core STA/APC 工作队列必须有界；取消 token、委托、超时 tombstone 都有 owner 和回收条件。未知 dirty 文档不能按普通缓存到期清理；达到上限应拒绝新会话，保留已有数据。
- 100 次开关合成文档后，插件窗口、COM 引用、句柄、pending、会话表与自有临时文件回归基线/固定缓存预算。进程内存允许缓存波动，但不能随次数线性增长。
- 空闲 10 分钟额外周期磁盘写入为零；不将 Office 自身正常锁文件/AutoRecover 误算成插件垃圾，也不删除这些文件。
- 故障 100 次不出现弹窗风暴；相同 openId 只保留一次终态提示，不自动重试。使用标签短状态，正常加载不弹提示框。
- 文件写入只发生于用户保存或已授权的 MCP 操作；失败、标签布局和聊天路由不能触发反复保存。恢复写入沿用现有事务账本及容量上限。

## 建议实施批次与通过条件

| 批次 | 修改范围 | 通过条件 |
| --- | --- | --- |
| A：数据/生命周期修复 | D01–D05、D08；明确承载关卡 | 保存回读一致；关闭可取消；旧请求不碰新会话；晚到不遗留；无静默 NOSAVE |
| B：定位与计时 | 构建 ID、阶段计时、D06/D07/D13 | 能对每次慢加载指出具体阶段；真实 view 与几何验证通过；默认无持续文件日志 |
| C：加载链路优化 | 并行启动、绝对 deadline、异步打开、合并布局 | 首屏 ≤1 秒目标；健康样本原生交互 P95 ≤30 秒；异常 30 秒内反馈，未知不重放 |
| D：MCP/加载项闭环 | D09–D12、生产跨层契约 | prepare/apply/revert/save 一致；路由不串；APC 取消无积累；不引发焦点/窗口复活 |
| E：发布验证 | 双 IDE × 三 Office、故障/资源矩阵、隐私 | 组件版本一致，安装/卸载生命周期完整，所有发布红线与隐私扫描通过 |

各批次先实现与测量，再更新设计手册中的能力声明。未完成原生端到端验收前，保留“原型 / 未验收”标记；AutoSave 本地重定向继续报告不支持。
