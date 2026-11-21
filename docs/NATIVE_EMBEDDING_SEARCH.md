# 原生 Office 承载方案检索

检索日期：2026-11-18。本文记录候选方案及证据，不代表已经完成兼容性验证。

## 当前结论

**后续状态（2026-11-21）：** 下文「最值得继续调查」的 WebRunLocal 经 11-19、11-21 两轮实测，均未完成运行时握手（见 [WebRunLocal 验证](WEBRUNLOCAL_VALIDATION.md)）。当前方向已改为焦点混合（见[设计手册首节](DESIGN.md#技术路线总结与路线裁决)）。本节保留 11-18 的检索判断。

本轮扩展中英文关键词，找到了值得进一步检查的中间件、窗口嵌入示例及官方容器样例。**尚未找到同时满足“标准 VSIX、真实桌面 Office、不传画面、正确参与 IDE 覆盖层、多标签隔离”的现成实现。** 这不是证明所有实现均不可行，也不能把只展示截图的项目算作原生窗口承载成功。

最值得继续调查的是 WebRunLocal / PluginOK / PageHiOffice：它明确声称支持现代浏览器内嵌本机 Office，但核心运行时并没有在本轮审阅中获得可复用源码及标准 VS Code 集成证据。其次是 MFCBind，可用于校正 OLE 容器协议，解决范围与 IDE 承载问题不同。

## 检索覆盖

采用关键词扩展而非仅搜索“Office VS Code”。实际检索词包括：

- `github VSCode native window editor embed HWND extension`
- `github office OLE web browser embedding DocObject WebView2`
- `github Electron embed native HWND child window overlay composition`
- `github webview native child window airspace workaround DirectComposition`
- `github "Office" "DocObject" container MFCBind`
- `github "native window" "browser" "embedding" "ActiveX"`
- `PageOffice Chrome 原理 本地服务 Office 窗口 嵌入`
- `ntko chrome office 嵌入 本地服务 原理`
- `github native office embed browser weboffice local window`
- `github "HwndHost" "airspace" "fix" screenshot`
- `"WebRunLocal" "SetParent"`
- `"201911323165" 浏览器`

先用搜索发现项目，再读原始仓库、关键源文件或厂商文档。没有把聚合文章、搜索摘要或宣传截图作为功能实现的充分证据。

## 候选与判断

### 1. WebRunLocal / PluginOK / PageHiOffice：优先调查的商业运行时线索

[公开仓库](https://github.com/wangzuohuai/WebRunLocal)包含 JS、SDK 示例和运行时包；README 声称可嵌入本机 Office，并明确中间件及小程序有商业授权要求。仓库出现 MIT 文件，不能据此假定所有随附二进制均可按 MIT 分发。

审阅 `Demo/OfficeJS/js/officeApplet.js`：页面经 WebSocket 发送 Office 类型、区域坐标及 OLE/完整嵌入模式。公开 C/C++/C# 文件集中在 SDK 样例；实际运行依赖 WrlEngine/ZbaEngine、MSOfficeApplet 等二进制。本轮未找到这些关键模块的对应完整实现，**不是可直接接入的全开源内核**。[示例接口](https://github.com/wangzuohuai/WebRunLocal/blob/master/Demo/OfficeJS/js/officeApplet.js)

厂商引用的专利说明采用浏览器页面属性调整、设置小程序父窗口、JS 或窗口钩子联动。其描述属于原生窗口整合，不是 Office HTML 引擎。专利文本不能证明当前二进制的具体实现，也没有证明 IDE HTML 菜单能覆盖 Office。[技术说明](https://patents.google.com/patent/CN111176752B/en)

判断：有较高调查价值，但“支持 Electron”尚不能推出“支持 Cursor 标准 Webview”。需要先取得可审计接口、可用测试条件，再做菜单、分屏和多标签验收；不能直接承诺兼容。

### 2. electron-native-windows：有实现的嵌入演示，但不是缺失能力的现成补丁

源码通过创建容器、启动外部应用、枚举 HWND、改窗口样式及 `SetParent` 承载外部窗口。`WindowManager.cc` 还修改父窗口样式，并调用 `HWND_TOPMOST` / `BringWindowToTop`。这解释了演示可以出现“嵌入外观”，却不能作为 IDE 菜单遮挡已解决的证据。[源码](https://github.com/sweetwisdom/electron-native-windows/blob/main/src/WindowManager.cc)

可借鉴：容器边界与进程发现的基本结构。不能直接复制：强制抬高窗口、无 Office 文档身份约束的按进程寻窗、对外部应用生命周期的简单处理。它使用自有 Electron 宿主，不提供标准 VS Code 原生标签槽。本轮未找到明确许可证，复用前需要确认。[仓库](https://github.com/sweetwisdom/electron-native-windows)

### 3. airspace：明确解决另一种需求，当前不采用

该 Tauri 插件有真实 Rust 源码。主方案用 `SetWindowRgn` 从原生内容窗口剪掉 HTML 控件所在区域；另一个方案创建透明置顶覆盖窗口。它不是任意 DOM 与 HWND 的统一合成接口。[项目说明](https://github.com/lukr54/airspace)、[承载实现](https://github.com/lukr54/airspace/blob/main/tauri-plugin-airspace/src/host.rs)、[覆盖窗口实现](https://github.com/lukr54/airspace/blob/main/tauri-plugin-airspace/src/overlay.rs)

判断：适合自有 Tauri 视频播放器等宿主；不能直接装进现有 Electron IDE。其挖洞和独立置顶窗口路线与本项目已明确的要求冲突，不作为默认候选。资料中关于兄弟窗口裁剪的实测可帮助理解故障，但不等于应采用该设计。

### 4. Microsoft MFCBind：值得参考的官方 DocObject 容器样例

官方样例使用 `COleDocObjectItem`，展示 Active Document 激活、视图及容器协作，可对照当前手写 OLE 接口的职责。仓库已归档，样例来自旧版 Visual C++，没有本轮现代 Office 验证。其把嵌入对象存入 Binder 文件的保存方式也不是本项目的本地原文件保存需求。[样例](https://github.com/microsoft/VCSamples/tree/master/VC2010Samples/MFC/ole/MFCBind)、[官方样例目录说明](https://learn.microsoft.com/en-us/cpp/overview/visual-cpp-samples)

判断：可帮助修正 Office 容器内部的布局、激活和菜单协商；不能补上 VS Code 没有暴露的原生标签槽。禁止把“容器修好”与“IDE 混合合成问题解决”混为一项。

### 5. microsoft-ui-reactor 的 Embedded Preview：容易误判的搜索命中

文档开头标注 Draft / design proposal，目标是 **Visual Studio 2022** 的 WPF `HwndHost`。它明确区分 VS Code 的截图预览；附录报告的 WinForms 原型不在仓库内。不能把微软仓库中的设计文档当作已经可用的 VS Code HWND API。[设计文档](https://github.com/microsoft/microsoft-ui-reactor/blob/main/docs/specs/056-visual-studio-embedded-preview.md)

判断：会话代次、尺寸握手等契约有参考价值；其中结束子进程的策略不能直接套到承载未保存 Office 文档的会话。当前无可直接采用的 VSIX 实现。

### 6. 其他线索的适用范围

| 线索 | 本轮核查结果 | 与目标的差距 |
| --- | --- | --- |
| [Darling](https://github.com/Peakk2011/Darling) | 实验性 Win32 宿主，把 Electron 放入原生窗口 | 承载方向相反，需要掌握整个宿主，不能作为现有 VSIX 的替换层 |
| [electron-libghostty](https://github.com/philipp-spiess/electron-libghostty) | README 明确仅 macOS，原生终端视图集成 | 不是 Windows Office 宿主；自有 Electron main process 与标准扩展不同 |
| [OnAir 的嵌窗服务](https://github.com/souz4rafael/onair-native/blob/master/DEVELOPMENT.md) | 文档列出 SetParent 服务及现代 Chromium 的限制 | 不是 Office/VSIX 实现，不把外观演示当作解决证据 |
| [PageOffice](https://www.pageoffice.cn/pages/a7b6ad/) | 官方流程先调用 POBrowser.openWindow，打开自己的浏览器窗口再承载控件 | 已有产品可供行为比较，但这条流程不等于在 Cursor 的标签里承载 |
| [NTKO](https://www.ntko.com/demo/officedemo/) | 官方演示依赖跨浏览器本地连接组件 | 需要另核实实际承载窗口及许可证，不能仅凭“跨浏览器”判断符合 VSIX |

## 官方接口与“无画面”边界

标准 Custom Editor 的公开模型仍是文档模型配合 Webview。官方文档没有提供任意 HWND 导入到标签的 API；这限制了可选实现，并不意味着 VS Code 插件只能做静态网页。[Custom Editors](https://code.visualstudio.com/api/extension-guides/custom-editors)

原生面板请求 [#100616](https://github.com/microsoft/vscode/issues/100616) 已关闭且标记 out-of-scope；不能把 GitHub 的 closed/completed 状态误读为功能完成。较新的 [#322901](https://github.com/microsoft/vscode/issues/322901) 本轮查询为 open，是需求讨论，没有交付实现。

微软的 `WebView2CompositionControl` 确实针对 WPF/WebView2 airspace 提供方案，但其说明明确涉及 GraphicsCaptureSession 和输入转发；方向是把浏览器表面放进 WPF，不是把任意 Office HWND 导入 Chromium DOM。因此既不能当作标准 VSIX 接口，也不能以“Composition”名称宣称完全没有画面转交。[官方设计说明](https://github.com/MicrosoftEdge/WebView2Feedback/blob/main/specs/WPF_WebView2CompositionControl.md)

## 后续验证的优先级

1. **先验证 WebRunLocal 类中间件的承载边界。** 确认是否依赖浏览器私有窗口结构、IDE 进程注入、全局输入拦截、独立置顶窗口或区域挖洞；要求提供标准 VS Code/Cursor Webview 内的证据。公开专利及 JS 只能说明大方向，不能代替这一项。
2. **把 MFCBind 用作 OLE 实现对照。** 梳理激活、正文矩形、菜单协商与保存的对应关系；这一工作能独立改善当前原生容器，但不承诺解决 IDE 覆盖层。
3. **建立筛选门槛后再做运行实验。** 先用合成窗口确认实际父子关系与宿主契约；Office 阶段仅用合成文档。未过关的演示不进入产品默认路径。

任何新候选都须实际通过以下场景：

- IDE 的命令面板、菜单、提示框跨越 Office 区域时可见、可点击；关闭覆盖层后正文正常。
- 两个 Office 标签、分屏、切换、隐藏及窗口移动互不串文档；跨显示器缩放无错位。
- 中文 IME、快捷键、焦点回到 IDE 正常，不使用全局鼠标键盘劫持。
- 保存仍经现有文档身份与仲裁器，关闭/失联不能无提示丢弃修改。
- 空闲时没有新增高频磁盘写入；日志、缓存和重试有上限，清理仅限本会话拥有的文件。
- 首次可编辑时间和输入响应必须实测；“本机窗口”或“无编码传输”都不是低于 40ms 的证据。

本轮只做资料检索与源码阅读，没有运行下载的二进制，没有安装厂商组件，没有调整当前插件路线。结论应表述为“候选待验证”，不能写成“已解决内嵌”。
