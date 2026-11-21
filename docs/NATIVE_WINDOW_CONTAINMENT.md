# 真实窗口承载与实现边界（2026-11-18）

本轮目标是转接真实 Office 窗口：保留原生绘制、OLE、保存与对象模型；主编辑表面不使用截图、视频、Canvas 像素播放或窗口画面串流。这是原位 OLE 路线在该轮的约束；2026-11-21 起焦点混合实验另有例外，见[设计手册首节](DESIGN.md#技术路线总结与路线裁决)。同文件额外标签的壳镜像是 Core 进程内由失效触发的 `PrintWindow` 复制，不编码、不传输，不属于这里的串流。本地 HTTP 只能承载控制协议，不能把 HWND 转换成 HTML 元素。现有命名管道已能传递控制命令，无需为改变传输协议另起常驻服务。

## 当前代码实际实现了什么

`WebviewSurface` 用令牌和四角节点找到标签内容区的物理矩形。`TryEditorContainer` 会选择一个可容纳这个矩形的 Win32 父窗口；找不到紧贴内容区的 HWND 时，会落到整个 IDE 的 `Chrome_WidgetWin_1`。`OleSite.HostInEditor` 再将 Core 容器设为这个窗口的 `WS_CHILD`。

这证明了 Win32 父子关系，不证明取得了标签专属原生插槽。DOM 令牌证明位置，不能证明 HWND 或合成层的所有权。当前根窗口下的子窗仍属于实验承载，菜单前后层没有验收通过；不得把 `GetParent` 成功、子窗随父窗移动或 OLE 可编辑，写成完整内嵌验收。

## 为什么再套一个容器不足以完成

VS Code 自定义编辑器的视图由 HTML/CSS/JavaScript Webview 实现；公开扩展 API 没有提供本文所需的原生 HWND 插槽或工作台合成层入口。普通扩展可以注册编辑器、菜单和 Webview，也可以启动私有本机执行器；这些操作均属于插件正常范围，不需要另行确认。

Win32 子窗口的父子树，与 Chromium 画布中的 DOM 树是不同的层级。把 Core 容器加到 IDE 根窗口下，可以得到原生裁剪和随父移动；不能让画布中的命令面板、HTML 菜单自动插到 Office 原生子窗口之上。按事件抢 `HWND_TOP`、按类名重挂 Office 飞出层，只是在两个层级之间制造新的竞争。

此结论依据 [VS Code Custom Editor](https://code.visualstudio.com/api/extension-guides/custom-editors)、[扩展边界](https://code.visualstudio.com/api/extension-capabilities/overview)及当前公开类型定义作出；不是声称 Windows 无法承载 OLE。Electron 的 [WebContentsView](https://www.electronjs.org/docs/latest/tutorial/web-embeds)由宿主主进程管理，也不等于扩展能获得任意原生 HWND 的混合合成能力。

## 正确的宿主表面契约

完整路线必须有可验证的表面提供者，同时满足以下条件。当前标准 VSIX 尚未提供该能力，以下是待实现契约，不是虚构的现成 API。

1. 一个标签实例绑定一个槽，带 `hostId/tabId/lease/generation`；提供准确 HWND、所属 PID/TID、进程创建身份及 DPI 上下文。尺寸相近、同 PID 或同类名都不能代替归属证明。
2. 槽的客户区就是标签内容区；由宿主控制布局、裁剪、可见性和销毁。Core 仅附着自己的容器，Office 通过 OLE 站点协商内容窗口与边界，不搬动桌面 Office 主框。
3. 宿主必须同时说明覆盖层的合成机制：IDE 菜单、命令面板、通知、拖动分组预览如何位于原生内容之上。只有一个 HWND 而没有覆盖层契约，仍不足以完成目标。不得用挖洞、整窗隐藏或争抢 Z 序代替。
4. 表面失效、标签迁移或宿主退出时，撤销旧代际命令；保存、恢复和文档生命周期由共享 Core 保留，不因 HWND 失效自动丢弃修改。
5. 焦点与加速键走明确的标签/OLE 边界；无全局键鼠钩子、输入劫持或画面输入转发。Office 弹层不能仅凭类名或服务器 PID 就改父级，必须保留其原有归属。

如需实现宿主协作模块，必须先给出具体接口、涉及的宿主版本及更新恢复策略。修改 IDE 安装代码或注入内部主进程接口，不得冒充标准 VSIX 已具备的能力，也不在本轮默默执行。

## SuperDoc 等特殊格式编辑器的区别

SuperDoc 官方扩展的 [provider](https://github.com/superdoc/docx-editor/blob/main/apps/vscode-ext/src/superdocEditorProvider.ts)读取 DOCX 字节、创建 Webview，并通过消息交给网页；[网页入口](https://github.com/superdoc/docx-editor/blob/main/apps/vscode-ext/webview/main.js)在 DOM 挂载 SuperDoc，导出 DOCX 后交回扩展写盘。编辑界面由浏览器绘制，没有启动 Word，也没有 Office 原生窗口需要参与叠加。

[当前 v2 文档](https://docs.superdoc.dev/editor/)说明其引擎直接编辑 DOCX 包，不能简单概括为 DOCX 转 HTML 再转回。微软 [Hex Editor](https://github.com/microsoft/vscode-hexeditor/blob/main/src/hexEditorProvider.ts)也通过 Webview 绘制自己的编辑界面。因此可参考其标签生命周期、消息和保存接口；不能把 DOM 挂载当作桌面 Office 的原生窗口插槽。引入这些引擎会成为兼容文件格式的自有编辑器路线，不满足本项目继续使用真实 Office 的当前要求。

## 微软官方 Web 版是否可以转接

微软提供 Microsoft 365 for the web，并通过 [WOPI/CSPP](https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/overview)让符合条件的云存储 ISV 接入 Word、Excel、PowerPoint 的网页查看和编辑。该计划不直接向普通 Microsoft 365 客户开放。它运行的是微软 Web 服务，并非把已安装的桌面 Office 暴露为本机网页组件。

编辑集成需要文件信息、读写、锁与解锁等 [WOPI 接口](https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/wopi-requirements)以及宿主页面和身份认证。微软 [PnP-WOPI 示例](https://learn.microsoft.com/en-us/samples/officedev/pnp-wopi/pnp-wopi/)指出，云服务必须能访问 WOPI 宿主端点，因此仅监听 localhost 的桥接不能直接完成微软云端编辑。Web 保存可经存储服务设计回写，但不是原生桌面 AutoSave 指向本机文件。

OneDrive 的 [Word 嵌入](https://support.microsoft.com/en-us/word/embed-a-word-document-in-your-blog)可提供文档嵌入查看，不能当作任意本地文件的完整编辑入口。企业本地部署还有 [Office Online Server](https://learn.microsoft.com/en-us/officeonlineserver/office-online-server-overview)，但微软已宣布 [2026 年底退役](https://techcommunity.microsoft.com/blog/officeeos/announcing-the-retirement-for-office-online-server/4462402)，不作为本项目的新默认路线。

若未来改用 Web 版，加载项需采用 [Office.js/Web Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)接口；当前 Windows COM 加载项与对象模型操作不能原封不动搬入浏览器。当前需求继续使用真实桌面 Office，此处仅记录替代路线，不自动上传文档或改变部署方式。

## 本轮已经实施的安全修复

- 可见原生子窗的 owner 移动不再调用 `SetWindowPos` 提层；子窗由 Windows 随父移动。恢复和布局保留原有兄弟 Z 序。
- 父客户区尺寸或 DPI 失效后，隐藏过期矩形并合并有限校准；后台标签不会因迟到事件显示。
- 移除 `OfficePopupZOrder` 的 PID/类名驱动弹层重挂和提层钩子。瞬态 Office UI 仅从旧帧清理中排除，不改父级、样式或位置。
- `AdoptChild` 只接受本进程私有容器，拒绝空句柄、自挂和父子成环；先验证 DPI 上下文，附着后核验真实父级及 `WS_CHILD/!WS_POPUP`，失败恢复私有容器原有状态。微软 [SetParent 文档](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setparent)说明了样式和跨进程 DPI 风险。
- `DebugNdjson.Probe` 统一为默认关闭的有界内存诊断；开启后最多 64 条固定阶段信息，不保存任意载荷，不启动文件写入或网络上传。

这些修复减少异常行为，不宣称已解决 DOM 与原生窗口的合成边界。本轮合成测试不操作用户文档；真实 Word 菜单、输入法、跨屏及多标签仍需实际验收。

## 读写与延迟口径

拖动、定位与窗口事件不能触发逐事件磁盘写入。Core 周期保存默认 60 秒、可设 5–3600 秒，仅 dirty 时执行；独立加载项的 1 秒心跳主要读取 COM Saved 状态，嵌入文档跳过以避免两套保存器。引用、编辑恢复记录及工作副本是显式操作引起的有界持久化，不得移入布局循环。

40ms 按用户口径作为交互验收阈值。需测量输入到可见反馈的端到端分布，并单独记录控制 IPC 和 Office 响应；只报本地 HTTP/管道往返不能证明整条交互达标。原位 OLE 路线不引入图像采集、编码或传输预算；焦点混合的预览预算另列于设计手册首节。

另有保存风险仍待专门修复：当前 SaveObject 用长度与修改时间同时变化判断外部冲突，同长度外部修改可能漏检；NativeSessionFiles 的保存事务方法尚未接入该入口。本轮未重写保存策略，不能据此承诺数据安全闭环。
