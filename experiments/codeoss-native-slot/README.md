# Code-OSS 原生标签槽实验

2026-11-19 在 Windows 10 22H2 的 SSH 实验机上运行。基于 VSCodium Windows x64 便携版 `1.143.17232`，来自 [VSCodium 官方发行页](https://github.com/VSCodium/vscodium/releases/tag/1.143.17232)。实验只修改专用副本的 Electron 入口。主体复用现有 C++ 原生桥，开发扩展仅通过公开命令打开真实命令面板。

| 同一重叠区域 | 远程结果 |
| --- | --- |
| 原版工作台命令面板 | DOM 命中命令面板，Win32 命中下面的原生文档 |
| 宿主拥有的 `WS_CHILD` 菜单层 | Win32 命中按钮，按钮动作和焦点恢复成功。绘制函数完成，桌面上的标记像素未通过 |
| 宿主拥有的临时 `WS_POPUP` 菜单层 | 标记可见，Win32 命中按钮，按钮动作成功。菜单外仍命中原生文档，关闭菜单后原位置重新命中文档并恢复焦点 |

三组测试使用同一个 `EDIT` 原生子窗口模拟文档。宿主只在初次挂载时建立同级窗口顺序，之后没有反复调整文档层级或在菜单打开时隐藏文档。临时弹层只用于菜单，由 IDE 主窗口拥有，未设置系统置顶。

实验分别记录 DOM、Win32 命中和桌面可见性。可见性只读取实验菜单上的一个已知色块并记录是否匹配，且仅在命中自己菜单窗口时读取。按钮通过 `BM_CLICK` 触发，使用真实的 Win32 按钮通知，未发送全局鼠标输入。结果确认了合成窗口、单一菜单和单窗口焦点转移，尚未覆盖完整命令面板、跨屏、缩放、输入法、多标签迁移和实际 Office 编辑。

## 复现

1. 在项目根目录执行 `./experiments/codeoss-native-slot/build.ps1`，需要已经安装 MSVC x64 工具链。输出位于 `.local/build/codeoss-native-slot`。
2. 准备专用的 `%LOCALAPPDATA%/OfficeInEditorLab/codeoss-<运行 ID 前 10 位>` 目录及带 `run` 和 `experiment=codeoss-native-slot` 的 `ownership.json`。完整运行 ID 由实验任务生成。
3. 下载固定的 VSCodium ZIP，核验 SHA-256 `a1090868af8ac0ac04c22d79fdd02246fa115a78daf38a741d8db06df1753c1d`，解压到实验目录内的 `codium`。把本目录的 `run-lab.ps1`、`host-probe.mjs`、`lab-extension` 和构建的 `oie-codeoss-slot.node` 复制到实验根目录。
4. 使用登录用户的 Windows 桌面执行 `run-lab.ps1 -LabRoot <实验目录> -Run <运行 ID>`。它使用独立的用户数据、共享数据与扩展目录，写入一份 `result-*.json`，退出实验进程并还原便携副本的入口文件。

初次远程运行发现版本新增了独立的共享数据目录，`--user-data-dir` 不会重定向这部分。随后加入 `--shared-data-dir`，日志确认共享数据库位于实验目录中。对照创建时间和文件名核实初次运行生成的默认目录后，将它移入实验目录，清除了默认位置的残留。[上游独立启动说明](https://github.com/microsoft/vscode/blob/main/.agents/skills/launch/SKILL.md)

## 对宿主接口的含义

Code-OSS 的内置浏览器已经有布局、可见性和焦点 IPC，可以参考。它的上游实现使用隐藏原生浏览器视图和占位图来处理覆盖。本轮验证了一条不隐藏原生文档的候选方式，保留标签原生槽，由宿主将横跨文档的弹层放在独立的窗口层。[上游内置浏览器架构](https://github.com/microsoft/vscode/blob/main/.github/skills/integrated-browser/SKILL.md)

实验按钮验证了这一层的绘制和动作，仍需让正式的 Quick Input 菜单在这一层渲染，并由工作台传输搜索结果、选择、Esc、焦点与关闭事件。可以先为 Code-OSS 的一个受控文档标签做窄接口，覆盖完整 Quick Input 后再接 Office。当前实验主进程读取自己的测试工作台 DOM 来测量交集，正式接口应由工作台主动提供独立标签的布局和覆盖事件。Cursor 宿主适配待实现。
