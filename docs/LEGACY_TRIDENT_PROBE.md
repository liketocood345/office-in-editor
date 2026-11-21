# Trident + DSOFramer：桌面 Office 网页承载实验

2026-11-20：合成文档已通过 HTML OBJECT 内的真实 Word 打开、编辑、回源保存、两次周期保存和正常关闭。原生 AxHost 打开及新建对照也通过。页面使用 MSHTML/Trident，控件来自固定版本的 DSO 样本源码；不包含 Office 本体。

## 已验证的调用链

```text
自产 HTML OBJECT → 私有 DSO 控件 → 已安装 Word 的 OLE DocObject
                    ↑
可信宿主 IOleContainer → 原生 HRESULT / 文档 HWND / 保存与会话控制
```

HTML 页面没有获得 Office 自动化脚本权限。可信宿主经 IOleContainer 枚举自己页面内的私有控件，再调用固定接口。没有修改 IObjectSafety、kill-bit、浏览器安全区、Office 本体或全局 COM 注册。

## 关键修复

- 宿主使用 PerMonitorV2 DPI。对照实验中，DPI unaware 宿主只能得到未真正激活的文档视图；匹配 Word 的 DPI 后出现 SetActiveObject 回调和真实 Word 子窗。
- 从 IOleInPlaceActiveObject.GetWindow 取得嵌入文档 HWND。嵌入 Word 的 Windows 集合可能返回 0x800A1735，不能沿用独立文档的取窗方法。
- 通过 PreserveSig 记录真实 HRESULT，消除动态调用返回和异常吞掉造成的误判。样本的 ActiveDocument 访问补上对已加载 IOleObject 的回退。
- 视图激活由标准 DoVerb 发起，传播 SetRect/Show 的错误，保留有界内存阶段诊断。
- 打开文件的 Save 回写原地址，新建文档首次保存指定自有目标；随后两次定时保存都要求同一文档 HWND，没有关闭或重开窗口。
- 关闭文档与退出应用分别判断。共享 Word 实例不执行 Quit；自有实例必须核对 HWND 对应 PID、启动身份及文档集合为空。

## 实测判据

| 用例 | 结果 |
| --- | --- |
| 独立 Word，单参数 SaveAs2 | 通过 |
| 可见窗口，两次官方 FileSave 命令 | 通过 |
| 隐藏窗口，两次官方 Document.Save | 通过 |
| 原生控件打开、新建与回源保存 | 通过 |
| HTML OBJECT 打开、编辑及两次周期保存 | 通过 |
| 无效所有权、篡改种子、重复准备及不安全恢复拒绝 | 12 项检查通过 |
| 超时合成进程跨 SSH 断线存活并正常退出 | 通过，临时任务清除，Office PID 不变 |

网页用例的一次监督器记录为约 4.45 秒，覆盖启动、打开、编辑、两次保存和关闭；测试使用已运行的 Word 与小型合成文档，不作为冷启动或复杂文档的基准。没有进行输入至呈现的 40ms 测量。

## 复现

先构建全部实现，再运行测试：

```powershell
./experiments/word-save-probe/build.ps1
./experiments/legacy-trident/build.ps1 -BuildDsoSample
./experiments/office-probe-common/verify-contracts.ps1
./experiments/office-probe-common/deploy.ps1 -Remote <ssh-alias> -Mode G1HtmlOpen
# 对照模式：P0、P0Visible、G1Open、G1Create
./experiments/office-probe-common/verify-isolation.ps1 -Remote <ssh-alias>
```

`-UseCachedSource` 只读取按 Git blob 哈希重新校验的原始缓存，每次重建独立补丁副本，避免补丁反复累积。结果、机器地址、运行目录和二进制位于忽略的 `.local` 中。

## 超时、数据和诊断

交互式临时任务承载测试进程，避免 SSH 作业回收带走未保存文档。监督器等进程句柄，发布结果采用原子文件移动；启动器等文件系统通知，不反复读取日志。正常退出后删除临时任务；需要恢复时保留已记录任务及进程，并拒绝同时启动另一轮 Office 探针。

每个运行目录有所有权、固定合成种子及摘要、产物摘要、阶段检查点和最终结果。检查路径及重解析点，只接受该运行的合成文件。失败始终返回非零；恢复成功也不能把失败测试改成通过。没有强杀 Office、全局键鼠钩子或图像转发。

## 固定来源与分发

DSO 样本镜像固定为 [shines77/DsoFramer 的提交 224163f5](https://github.com/shines77/DsoFramer/tree/224163f5a5acb692118a5be95e5e49fe750bfd50)。镜像声明来源为 Microsoft KB311765；构建保留源码通知、Microsoft 样本 EULA 及镜像 LICENSE。不能仅凭镜像 MIT 标签把全部微软样本代码称为 MIT。分发审查仍需依据实际文件许可。

私有实验移除自动类型库登记和全局注册入口，使用独立 CLSID/ProgID 和免注册清单；原始及私有 CLSID 均做只读 kill-bit 预检。当前产物是开发实验，不混入生产 VSIX。
## 2026-11-21 三格式 localhost 接续验证

按当前任务先在 SSH 的独立 Trident 宿主测试。Trident 从回环 HTTP 页面加载私有 DSOFramer `<object>`，控件由该网页宿主管理。每轮只向浏览器提供静态 HTML，Office 原件留在本机实验目录。页面不接收 COM 自动化接口。日志记录实际页面响应、原位激活 HRESULT、Office 窗口所属进程、子窗口关系、可见性、三个命中点、修改、保存和独立读回。

| 文件 | localhost 页面内原生界面与命中 | 原件保存核验 |
| --- | --- | --- |
| Word `.docx` | 可见，3/3 命中 | 两次周期保存，正文标记从源文件读回 |
| Excel `.xlsx` | 可见，3/3 命中 | 两次周期保存，第一张表 A4 的标记从源文件读回，独立对照确认 D2 公式保留 |
| PowerPoint `.pptx` | 可见，3/3 命中 | 两次周期保存，从源文件读回第一页首个文本框 |

Excel 的 `DoVerb` 返回 `OLEOBJ_S_INVALIDVERB`，该值为非失败 HRESULT。激活逻辑现在也检查是否取得实际活动对象，并在缺失时走 `IOleDocumentSite::ActivateMe`。每轮 Excel 夹具名称包含运行标识，以避免不同目录的同名工作簿触发模态提示。Excel 关闭控件后仍可能持有工作簿，探针只关闭路径与本轮原件精确匹配且已保存的工作簿。最终复测确认 Excel 和 Trident 都已退出。已有 Word 实例保留。 [DoVerb 接口](https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nf-oleidl-ioleobject-doverb)

```powershell
./experiments/legacy-trident/build.ps1 -BuildDsoSample -UseCachedSource
./experiments/word-save-probe/build.ps1
foreach ($mode in 'G1HtmlOpen','G1HtmlExcel','G1HtmlPowerPoint') {
    ./experiments/office-probe-common/deploy.ps1 -Remote '<SSH 别名>' -Mode $mode -Loopback
}
```

验收范围是独立 Trident 宿主和受控 Office 自动化修改。Edge IE 模式、标签切换和完整菜单交互仍待验证。普通 Chromium Edge 和 VS Code 不属于同类待验项：Chromium 不加载 ActiveX（见[浏览器探针](BROWSER_OFFICE_PROBE.md)）。要放进 VS Code，只能把 Trident 宿主作为另一个原生窗口摆在标签上，这会落回设计手册路线表第 1 行的 z 序合成墙，与本实验的「网页内承载」结论无关。SSH 未检测到 Edge IE 模式的企业站点列表策略。私有 OCX 用宿主清单加载，没有全局 COM 注册。实验没有叠层、画面捕获或图像传输。

