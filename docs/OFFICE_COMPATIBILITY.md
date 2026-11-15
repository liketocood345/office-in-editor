# Office 逐应用兼容契约（2026-11-15）

本文区分“范围候选”“有代码”“已真实验收”。它不把文件关联、COM 注册或成功创建窗口当成完整兼容。共同架构、风险与 Word 顶部遮挡分析见 [ARCHITECTURE_AUDIT.md](ARCHITECTURE_AUDIT.md)。

## 当前实现矩阵

| 应用 | package.json 当前候选格式 | 加载项 / MCP 现状 | 发布判断 |
| --- | --- | --- | --- |
| Word | docx/doc/docm/dotx/rtf | 有枚举、文本选区、正文插入、菜单；本地保存只接受 docx | 当前内嵌布局与源文件回存需修复/复验 |
| Excel | xlsx/xls/xlsm/xlsb/csv | 有 Range 选区、有限值/公式写入；本地保存只接受 xlsx | 需真实嵌入/保存/输入验收 |
| PowerPoint | pptx/ppt/pptm | 有文本/形状/幻灯片选择、新幻灯片/文本框；本地保存只接受 pptx | 需真实嵌入/保存/视图验收 |
| Visio | vsdx/vsd/vstx | 无应用识别、注册、Ribbon、枚举、选区、inspect/save 分派 | 条件候选，不能声明完整兼容 |
| OneNote | one/onepkg | 无上述分派；尚无当前版本 DocObject 支持证据 | 条件候选；无合格文档服务器即拒绝 |
| Publisher | pub | 无上述分派 | 条件/遗留候选；另有产品退役限制 |
| Project | mpp | 无上述分派 | 条件候选；事件与保存需专用适配 |

证据：package.json 的 customEditors.selector；OfficeAddin.cs:133–138、RibbonMarkup.cs:21–27、DocumentRegistry.cs:23–28、SelectionCapture.cs:32–37、DocumentOperations.cs:237–248；scripts/install-office.ps1:48 仍只注册前三款。后四款不是简单“缺 MCP 写白名单”，协作基础也未完成。这里的无分派指当前统一加载项，不是否认产品本身有扩展能力。

OfficeDocumentKind 的猜测表与 Manifest 并不完全一致：例如 dotm/xltx/xltm/potx/potm/ppsx/vssx/vstm 在猜测表中、尚未全部贡献选择器。必须按格式列能力，不能由应用名称一概推导。MCP 捕获/编辑未等价采用保存格式限制，因而某旧格式可被捕获，不代表可安全保存。

## Word

- 选区要保留 storyType 与范围。Document.Content 只代表 main story；页眉/页脚/脚注/批注的 start/end 不能直接变为正文插入位置。现有写入只允许正文字符插入，不能宣称完整 Word 编辑。[Document.Content](https://learn.microsoft.com/en-us/office/vba/api/word.document.content)
- 分别处理普通只读、ProtectionType 与 ProtectedViewWindow。Protected View 不是普通 Document 加一个 ReadOnly 标志；保护状态读取失败应拒绝写，不能降为可写。[ProtectionType](https://learn.microsoft.com/en-us/office/vba/api/word.document.protectiontype)、[ProtectedViewWindow](https://learn.microsoft.com/en-us/office/vba/api/word.protectedviewwindow)
- 当前右键仅挂 ContextMenuText。正文、表格、标题、图片等上下文并非一个菜单；“置顶”只能在支持的菜单请求相对位置，并逐构建验收。源码已记录 LTSC 2021 对额外 M365 context ID 的 customUI 兼容问题。
- 重点验收 Ribbon 边界协商、同 PID 多 Document、story locator、修订/保护和原文件持久化。不能用 `_WwG` 或进程匹配替代文档身份。

## Excel

- Selection 可以是 Range、图片或空值；不支持的对象选区应明确拒绝，不以 ActiveCell 替代。现有捕获最多 32 areas/5000 cells，编辑最多单矩形 200 cells。[Application.Selection](https://learn.microsoft.com/en-us/office/vba/api/excel.application.selection)
- Value2 不是显示文本快照，也不保留 Date/Currency 类型。Formula 与动态数组时代的 Formula2 有语义差异；不要把现有 Formula 写入宣称为完整 spill 公式支持。[Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)、[Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula)
- 分别校验工作表保护、合并单元格、数组公式、单元格编辑态与计算忙碌；只检查 workbook.ReadOnly 不够。等待忙碌必须有界，不循环 SendKeys 退出编辑态。[Worksheet.ProtectContents](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protectcontents)
- CSV 另有编码、分隔符、类型推断和保存提示；模板、宏文件和二进制工作簿不能按 xlsx 的保存路径处理。公式栏、工作表标签、任务窗格与 Office Ribbon 都要参与布局验收。
- COM 写入可能影响 Excel 撤销栈，必须保留 Core 的补偿事务；不能依赖 Ctrl+Z 作为唯一恢复手段。[AutoSave 与加载项/宏](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)

## PowerPoint

- 编辑窗口与放映窗口不同。现有 ActiveWindow.Selection 属于编辑窗口模型，未覆盖 SlideShowWindows；播放期间不能把焦点变化解释为编辑文档更换。[DocumentWindow.Selection](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindow.selection)、[SlideShowWindows](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideshowwindows)
- 现有选择支持文本/形状/少量幻灯片；形状捕获主要是文本与身份，不是图像语义提取。稳定 locator 需绑定 presentation/slide/shape 身份，不能只留随着插入删除改变的序号。
- 必测普通视图、排序视图、文本编辑、备注/母版、放映、只读/保护、缩略图和任务窗格。模板与 ppsx 放映文件要另设策略，不把运行放映的顶层窗口抓进标签。

## Visio

- 当前通用 OLE 路线是否支持指定安装必须实测。若将官方 DrawingControl 作为等价文档容器，需单独声明能力与载荷依赖，不能把它的行为当作当前代码的行为。
- DrawingControl.Src 加载的是副本，Document.Save 不能直接回源；官方要求清 Src 后 SaveAs。多个 Control 共享 Application，而每个 Control 自己只有一个 document/window；Control 内 VBA 宏不运行。这直接影响来源绑定、保存、实例隔离和宏能力承诺。[官方 DrawingControl 契约](https://learn.microsoft.com/en-us/office/vba/visio/concepts/about-using-the-visio-drawing-control-in-your-application)
- 绘图、模板、模具、页与 shape 选择分别处理；需要 Visio 专用 Ribbon/context ID、对象保护和稳定 page/shape locator。禁止增加 VISIO 主窗类名来替代文档接口适配。

## OneNote

- 桌面 COM API 的对象是 notebook/section/page 的 ID 与 XML，不是统一的 Documents/FullName/Saved/Save。GetPageContent 可要求选区标记；UpdatePageContent 支持预期最后修改时间，并可能替换匹配 ID 的页级对象；SyncHierarchy 才涉及磁盘来源同步。[Application 接口](https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote)
- 有 OneNote.Application、COMAddIns、文件关联不证明存在可原位激活的 IOleObject/IOleDocument。仅限通过探测的桌面文档服务器；无此能力保持不支持，不抓主窗口。[Win32 API 范围](https://learn.microsoft.com/en-us/office/client-developer/onenote/onenote-home)
- `.onepkg` 是包/导入入口，不等于已就地绑定的单一文件；需要独立证明导入目的地、层级身份与持久化，不允许当作普通 `.one` 直接回写。
- OneNote 的持续保存/同步不能套前三款的本地 Save 心跳；须报告自己的保存能力和冲突语义，不显示无效的“本地定时 Save”。云笔记本、分区跨页修改与托管事件可用性必须单独评估。

## Publisher

- Publisher 使用 SDI；Application.Open 要在新 Application 实例上调用，不能复制 Word 的多文档集合和同 PID 多标签假设。[Application.Open](https://learn.microsoft.com/en-us/office/vba/api/publisher.application.open)
- 选择包含 publication window pane 的文本、形状等，需要自己的 locator、保护/保存与菜单适配。[Selection](https://learn.microsoft.com/en-us/office/vba/api/publisher.selection)
- 当前日期已在产品退役之后：微软公告 M365 从 2026-10-01 起移除 Publisher，永久版支持止于 2026-10-13。因此只能作为确有可用授权安装的遗留候选，不能列成当前 M365 默认可用产品，也不能承诺获得安全更新。[官方退役说明](https://support.microsoft.com/en-us/publisher/microsoft-publisher-will-no-longer-be-supported-after-october-2026)

## Project

- 保存入口是 Application.FileSave，作用于活动项目；必须在 STA 内校验所指项目仍是绑定项目，不能直接套 document.Save 或临时激活其它项目。[Application.FileSave](https://learn.microsoft.com/en-us/office/vba/api/project.application.filesave)
- 微软明确嵌入其它应用/文档时不发生 Project events，因而不能把独立窗口的 BeforeSave/事件驱动自动保存测试直接外推到 OLE。要有嵌入模式独立能力和结果确认方案。[Project.BeforeSave](https://learn.microsoft.com/en-us/office/vba/api/project.project.beforesave)
- task/resource/cell/view 选区有不同语义，需稳定 ID、日历/资源/链接影响范围与事务恢复；先做明确只读能力，再引入有限写入。甘特图、资源视图和编辑网格都须验证布局与输入。

## 共同的准入与发布关卡

1. 只支持本地 Windows 桌面宿主，Office 本体由用户安装授权。Core 为 x64 不代表所有 x86/ARM64 Office 已兼容；进程内 COM 加载项及 ActiveX 二进制必须匹配宿主位数。安装和测试矩阵分别记录 Core 与 Office/加载项位数。[微软位数兼容说明](https://learn.microsoft.com/en-us/office/client-developer/shared/compatibility-between-the-32-bit-and-64-bit-versions-of-office)
2. GetClassFile 只是得到 CLSID；仍须证明完整文档接口、可原位编辑、UI 合并、加载项出现与持久化。[GetClassFile](https://learn.microsoft.com/en-us/windows/win32/api/objbase/nf-objbase-getclassfile)
3. OleCreateFromFile 创建嵌入对象，不是原文件编辑绑定；rclsid 是保留参数，按规范应传 CLSID_NULL，不能把传入候选 CLSID 当成服务器锁定。现代码的内存 IStorage 路线必须另做原件持久化契约。[OleCreateFromFile](https://learn.microsoft.com/en-us/windows/win32/api/ole2/nf-ole2-olecreatefromfile)
4. 保护视图、密码/IRM、宏/外部链接、安全/许可提示、保存事件取消均不可绕过。插件不改全局 AutomationSecurity/DisplayAlerts，不自动信任位置，不隐藏不明对话框。
5. 原生 AutoSave、本插件**定时保存**（周期静默点保存按钮）、AutoRecover、OneDrive 同步分别报告。插件已取消「编辑后保存」。云同步目录有盘符并不等于纯本地文件；不能因 LocalPath 检查通过就宣称隔离了云同步。[AutoSave 与加载项/宏](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
6. 每款至少验收独立/内嵌、右键、选区、准确聊天路由、MCP 能力、原件保存回读、dirty 关闭取消、异常恢复、DPI/多标签和 30 秒可交互。没有实现的项目标 unsupported/pending，不能用空白或“同类兼容”代替。

Outlook、OneDrive、Teams、门户及不能提供文档服务器的 UWP/商店外壳明确不在标签嵌套范围。Access 当前没有选择器或适配，不因同属 Office 或打开本地数据库而自动纳入；未来若扩展数据库类产品，需另评估事务、锁、恢复和文档容器能力。
