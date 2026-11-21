# 本地 Office → localhost → 标准浏览器：一次实际验证

## 2026-11-21 无窗口承载复核

该轮的交付目标是在网页标签内操作本机安装的 Office 界面并保存原文件。允许增加承载扩展，排除嵌入或叠加 Windows 原生窗口、传送窗口画面、重写编辑器和云端 Office 编辑。这些排除项只是网页承载路线在本轮的范围；同日之后的焦点混合实验恰好使用了叠加原生窗口和本机画面，两者不冲突，见[设计手册首节](DESIGN.md#技术路线总结与路线裁决)。

复跑标准 Edge 157.0.4341.60。localhost 页成功加载，三种已经注册的 Office 文档类均显示 `<object>` 的备用内容，网页未取得 ActiveX 或 Office MIME 处理器。测试配置和监听在结束时清理。

新增本机 COM 探针，分别创建三款 Office 的文档 OLE 对象并查询接口。`S_OK` 为 `0x00000000`，`E_NOINTERFACE` 为 `0x80004002`。

| 本机文档类 | `IOleObject` / `IOleDocument` / `IOleInPlaceObject` | `IOleInPlaceObjectWindowless` / `IViewObjectEx` |
| --- | --- | --- |
| Word.Document.12 | `S_OK` | `E_NOINTERFACE` |
| Excel.Sheet.12 | `S_OK` | `E_NOINTERFACE` |
| PowerPoint.Show.12 | `S_OK` | `E_NOINTERFACE` |

此探针只检查当前 Office 安装的文档类接口，同一应用已有运行进程时跳过。它未加载文件或激活文档视图。微软的无窗口 OLE 承载需要对象与站点两侧都提供对应接口。本次结果关闭了通过这三种文档类直接做无窗口 OLE 承载的路径，测试范围仅限这些公开接口。 [Win32 无窗口 OLE 文档](https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nn-ocidl-ioleinplaceobjectwindowless)

查询会短暂启动对应的 Office COM 服务器。探针释放自身的 COM 对象，服务器何时退出由 Office 决定。本次 Excel、PowerPoint 自行退出，Word 在确认是本轮启动的空后台实例后，通过正常 COM 退出收尾。

承载扩展能通过 Native Messaging 与本机进程通信，VS Code Webview 也可给扩展宿主发消息。两套公开 API 都没有把桌面 Office 编辑界面交给网页渲染的接口。Edge 的 ActiveX 兼容入口使用 Trident。本项目成功的 Trident 样本创建了实在的 Office 子 HWND，仍不符合当前的网页验收条件。 [Edge Native Messaging](https://learn.microsoft.com/en-us/microsoft-edge/extensions/developer-guide/native-messaging) · [VS Code Webview](https://code.visualstudio.com/api/extension-guides/webview) · [Edge IE 模式](https://learn.microsoft.com/en-us/deployedge/edge-ie-mode)

```powershell
$project = 'packages/shared-core/Tests/OfficeInEditor.Core.Tests.csproj'
foreach ($application in 'word', 'excel', 'powerpoint') {
    dotnet run --project $project -- --office-windowless-probe $application
}
```

返回码 `0` 表示成功完成接口查询，需读取 JSON 的 `outcome` 来判断无窗口接口是否可用。

## 2026-11-21 SSH 上的 VS Code 集成浏览器

用原版 VS Code 1.121.0 的 `workbench.action.browser.open` 打开三页，每种文档类一个回环端口。实验扩展从隔离配置启动。页面按目标机已注册的 Word、Excel、PowerPoint OLE 文档类生成 `<object>`，未向其提供任何用户文件。集成浏览器是 VS Code 内的 Chromium 标签。 [VS Code 集成浏览器](https://code.visualstudio.com/docs/debugtest/integrated-browser)

| 检查 | Word | Excel | PowerPoint |
| --- | --- | --- | --- |
| 文档类已注册、网页标签加载 | 是 | 是 | 是 |
| 首次和刷新均执行脚本、回报同一端口 | 是 | 是 | 是 |
| ActiveX、Office 控件接口、Office MIME 处理器 | 均无 | 均无 | 均无 |
| 首次和刷新均显示并命中网页备用内容 | 是 | 是 | 是 |
| 页面 CSP 拒绝对象加载 | 否 | 否 | 否 |

第一轮页脚本先于 `<object>` 备用内容排版，出现过可见性读数不一致。第二轮在页面加载及排版后采样，六次均得到 `browser_did_not_activate_office` 的对应观察。VS Code 标签页和 localhost 请求已通过。三款原生编辑界面均未出现，编辑保存关卡仍未运行。隔离 VS Code 窗口正常退出，三个监听已释放，临时交互任务已注销。复查目标机原有 Word 进程数量没有变化。

```powershell
./experiments/office-browser-vscode/remote.ps1 -Remote '<SSH 别名>'
./experiments/office-browser-vscode/remote.ps1 -Remote '<SSH 别名>' -Run '<首次命令返回的 run>'
```

按 `launch`、`activated`、`word`、`excel`、`powerpoint`、`result`、`finished` 读阶段快照。扩展还将探针事件写入 VS Code 的 `Office Browser Probe` 输出通道。阶段包含标签打开请求和返回、HTML 与脚本响应、首次采样、刷新采样、对象状态和监听退出。发生超时时可以区分启动失败、命令缺失、网页未到达、脚本未回报和控件未激活。输出最多 96 条事件，阶段快照按阶段一次性落盘。日志不收集正文、绝对路径、凭据或截图。

## 2026-11-18 标准浏览器实验

2026-11-18，按约定路线执行一次独立最小实验。它是能力验证，不是已经完成的网页 Office 编辑器，也不改变现有插件的默认打开方式。

## 验证链路

1. 只读查询本机 Word.Document.12、Excel.Sheet.12、PowerPoint.Show.12 官方 OLE 文档类注册。
2. 独立测试进程静默启动仅绑定 loopback、动态端口的本机 HTTP 服务。
3. 服务提供标准 HTML 页面，以 `<object classid="clsid:…" type="…">` 请求浏览器承载已注册的 Office 文档对象；不放截图、视频、自制文档编辑器或屏幕定位窗口。
4. 用独立临时配置启动标准 Microsoft Edge Chromium，检测 ActiveX 主机能力、Office MIME 处理器和实际替代内容是否显示，向服务提交一次结果。
5. 退出本次浏览器实例和测试服务，清理该实验专有临时配置。没有操作现有 Word/Excel/PPT，没有打开或保存用户文档。

页面 CSP 特意允许 `object-src 'self'`，避免因测试自己禁止 OBJECT 而制造失败。页面仅测试本机原生 OLE 对象承载；没有假定 WOPI、Office Online Server 或 Microsoft 365 云端服务存在。

## 本次实测结果

本机为桌面 Office 2021 x64；未检测到 Office Web 服务器。实际浏览器版本为 Edge 157.0.4341.60。

| 项目 | Word | Excel | PowerPoint |
| --- | --- | --- | --- |
| 官方 OLE 文档类已注册 | 是 | 是 | 是 |
| 浏览器 Office MIME 处理器 | 无 | 无 | 无 |
| OBJECT 替代内容可见 | 是 | 是 | 是 |
| 出现可编辑的官方 Office 界面 | 否 | 否 | 否 |

浏览器 `ActiveXObject` 不可用；结果为 `blocked_browser_ole_host_unavailable`。localhost 服务及 HTML 加载成功，但浏览器没有激活 Office 文档对象。

退出核对：`temporaryBrowserProfileRemoved=true`、`serviceStopped=true`；`officeStartedByProbe=false`、`userDocumentsOpened=0`、`pixelTransfer=false`。标准浏览器本身会使用临时配置文件，已在本次结束后清理；这不是插件运行时的逐事件日志或文档写盘。

## 结论及边界

本次实验否定的是**把本机已注册的桌面 Office OLE 对象直接交给标准 Chromium 网页承载**这条具体接口路线。HTTP 服务不能补出缺失的浏览器 OLE/ActiveX 主机。微软对 ActiveX 的浏览器兼容路径是专门的 [Edge IE mode](https://learn.microsoft.com/en-us/deployedge/edge-ie-mode)，不能将其当成 VS Code/Cursor 的标准 Chromium Webview 能力。

这不是 Word 的实际编辑/保存验收，也不证明所有可能架构均不可行。实验没有创建 OLE 会话，没有运行非交互的 Word 服务，没有测试微软 Web 版，更没有实现“桌面 Office 原生窗口变成 HTML”。UI 后端关卡 G0 仍未通过，因此没有继续把这个失败入口接成 IDE 的默认编辑器，也没有安装或覆盖本机组件。

继续完整路线，需要一个实际可用的呈现后端：官方 Office Web/WOPI 部署、宿主原生窗口槽及覆盖层协作，或用户另行选择的图形远程显示。相关差异见 [真实窗口承载](NATIVE_WINDOW_CONTAINMENT.md)。

## VSIX 内分发 ActiveX 宿主的核查

VSIX 可以携带 Windows 原生辅助程序，但 VS Code 的公开 [Webview API](https://code.visualstudio.com/api/extension-guides/webview) 只提供 HTML 界面和消息传递，未开放控件宿主或原生窗口嵌入接口。Edge IE 模式的 ActiveX 能力来自 Trident。带上一个 Trident 宿主会增加一个独立的原生渲染面。

检查固定版本的 DSOFramer 源码，其 COM 接口公开 `IOleInPlaceObject`，没有公开 `IOleInPlaceObjectWindowless`。控件和文档分别创建 `DSOFramerOCXWnd` 与 `DSOFramerDocWnd` 子窗口。先前通过的 Trident HTML 实测在这些窗口下发现了真正的 Word `_WwG` 编辑窗口，并已验证保存。因此把该控件收入 VSIX 不会让 Word 的编辑窗口进入 Webview 的网页合成顺序。它仍需要 IDE 宿主处理原生窗口与网页覆盖层的叠放。

网页路线的另一份实验通过了普通 Edge 的最小闭环。网页与本机 Word 交换文字和保存请求，Word 保持隐藏。已读回实存的合成 `.docx` 并比对正文。实验在 `experiments/word-browser-local`，目前提供单段纯文本编辑。

## 复现

~~~powershell
dotnet run --project packages/shared-core/Tests/OfficeInEditor.Core.Tests.csproj --no-restore -- --browser-office-probe
~~~

退出码0表示实验完成，不表示 Office 内嵌成功；必须读取结果中的 outcome 和 observed。实验最多等待20秒，只管理自己启动的隔离浏览器进程，不按进程名结束浏览器或 Office。入口位于 `Tests/BrowserOfficeProbe.cs`，不会编译进生产 Core。
