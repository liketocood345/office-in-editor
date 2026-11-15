# 修复状态与 AutoSave 边界

下文保留为早期修复快照。当前 0.1.4 源码状态、持久化风险与 Word 顶部布局分析见 [架构审查](ARCHITECTURE_AUDIT.md)，逐应用范围见 [兼容矩阵](OFFICE_COMPATIBILITY.md)。历史独立保存或合成探针结果不能替代当前内嵌验收。

更新：2026-11-15。本文件覆盖本次源码修改，不能作为原生端到端验收或已安装版本声明。

## 本次已修改

- MCP applyEdit 实际传递 recovery；Core 假 Office 测试现在要求该字段内容匹配，避免两个独立 fake 同时假绿。
- VSIX 在打开前分配会话 ID，关闭立即标记失效；后台打开中可请求取消，晚到结果补偿 detach。停用时收集并请求关闭所属会话，清理不再启动新 Core。
- 打开和连接/几何并行；标签打开预算 60 秒；layout/detach 的传输等待为 1.5 秒。未完成时报告未知，不自动重放。失败不 `Process.Start` 独立 Office。
- layout 使用单在途、最新值合并；监听先于 HTML 安装，可见性变化请求重新报告矩形。
- Core 同步 OLE 在后台请求工作线程等待，避免堵住 IPC 接收；新打开并发只允许一个，队列上限 32。
- 原型暂限单会话，第二次打开明确失败，不再静默关闭旧文档。STA 执行时复核会话身份。
- 不再使用 NOSAVE 自动关闭；Office 取消/保存失败时保留会话。SaveObject 不再虚报 S_OK，而是明确失败。**这只是阻止静默丢失，并不表示已实现源文件保存。**
- OLE 存储提前转交所有权，释放原始 ILockBytes；复用/关闭 document view，检查激活与布局错误；窗口布局保留 Z 序。
- APC 工作最多 32 项，每实例最多一个未交付 wake；取消移除 pending；进程级静态 thunk 避免卸载后调用已释放委托。
- 嵌入聊天绑定不再被默认聊天刷新覆盖；路由推送有统一 500 毫秒预算。默认关闭文件加载日志，显式诊断才写。
- 移除测试中的固定绝对路径，隐私扫描通过。
- 移除 extension-host smoke 中真实打开伪 DOCX 的部分；该 smoke 只检查注册与脱敏诊断，原生文档打开另行隔离验收。

## 明确未完成

1. **原生 AutoSave 本地重定向未实现。** 微软原生 AutoSave 依赖 OneDrive/SharePoint，受支持对象模型未提供本地保存提供器替换接口。不能通过定时 Save、AutoRecover、注册表或修改 Office 来伪装实现。能力字段继续为 false。[Microsoft AutoSave 说明](https://support.microsoft.com/en-us/office/collab-files/what-you-should-know-about-autosave)、[AutoSaveOn 状态限制](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autosaveon)。
2. 当前 OleCreateFromFile + 内存 IStorage 的保存契约仍需重做；不能将该原型作为可写生产编辑器发布。不能把 SAVEIFDIRTY 或返回 S_OK 当成保存实现。
3. dirty/save/backup 尚未接入可编辑 CustomDocument；原型仍为 readonly provider，无法在 dispose 后取消关标签。保存失败可能需要在保留的 Office 会话手动处理。
4. 会话所有权尚未绑定 Core host lease；宿主崩溃、取消先于 Core 接收 open、同步 COM 挂起等仍须协议级 begin/status/cancel 与失联回收。不能声称完全消除了孤窗。
5. HWND 仍为 Core 顶层容器；DPI、窗口移动、遮挡和原生标签承载合规关卡未通过。按 PID 识别嵌入的路由归属也仍需文档级映射。
6. 30 秒是请求预算和后续验收目标，未完成双 IDE × 三 Office 的真实加载测量；成功率、首绘和可交互耗时未知。

以上问题解决前不替换本机已加载 DLL、不发布新安装包；已有用户 Office 不结束、不移动，不用真实用户文件做破坏性验证。

## 本次隔离验证

- TypeScript 类型检查通过。
- JS 17 项测试通过，新增验证 openInOffice 的实际 timeout 参数。
- Core 59 项检查通过，含恢复参数一致性、持久化幂等与未知结果禁止重放。
- 加载项 323 项断言通过，新增 100 次取消后 pending 归零及迟到回调不执行工作。
- 源码隐私扫描零发现。

这些检查不启动 Office、不修改注册表，不替代原生集成测试。保留所有既有未提交改动。
