# Office in Editor

**类别：高效工作（Productivity）**

Office in Editor 是一款 VS Code / Cursor 插件，用来在编辑器标签页里直接打开 Microsoft Office 的文档（Word）、表格（Excel）和讲稿（PowerPoint）。文档仍由本机的桌面版 Office 打开和保存；插件让它们像代码文件一样出现在编辑器里，并让编辑器中的 AI 可以读取和修改它们。

> 预览版，仅支持 Windows。需要本机已安装并授权的 Microsoft Office，本项目不包含 Office 本体。

## 组成

插件由两部分组成，二者共用一个本机后台（Core）：

| 部分 | 运行位置 | 作用 |
| --- | --- | --- |
| **Office 扩展**（COM 加载项） | Word / Excel / PowerPoint 内 | 提供功能区和右键菜单里的 **Add to Chat**，读取选区，代 AI 执行文档修改 |
| **VS Code 插件**（VSIX） | VS Code / Cursor 内 | 把 Office 文件作为自定义编辑器打开在标签页里，按需启动 Core，并向编辑器注册 Office MCP |

Core 负责文档身份、MCP 服务、修改记录和恢复，并负责调度标签与 Office 窗口。

## 功能

### 在标签页里打开 Office 文件

在资源管理器里打开 `.docx`、`.xlsx`、`.pptx` 等文件，会直接进入插件的标签页，而不是弹出一个独立的 Office 窗口。

- 标签页有焦点时，原生 Office 窗口贴在标签内容区上，键盘和鼠标都交给 Office，所有原生功能照常可用。
- 切到编辑器的其他地方后，Office 窗口移出屏幕，标签里显示它的实时画面。在画面上滚动滚轮会转发给 Office，点一下画面就回到编辑状态。
- 打开文件时 Office 窗口不会单独出现，也不会出现任务栏图标。关闭编辑器时，未修改的文档会在后台关闭；有未保存修改的文档会还原到桌面，让你自己决定是否保存。

路由由设置项 `officeInEditor.hostRoute` 控制：

- `auto`（默认）：Word / Excel / PowerPoint 走上面这条焦点切换路线，其余格式走原位 OLE 路线；
- `focus`：Word / Excel / PowerPoint 一律走焦点切换路线；
- `ole`：一律走原位 OLE 路线。

### Office MCP：让 AI 修改文档

插件向 VS Code / Cursor 注册一个名为 `office-in-editor` 的 MCP 服务器，编辑器里的 AI 可以通过它操作正在打开的 Office 文档：

| 工具 | 作用 |
| --- | --- |
| `office_list_documents` / `office_inspect_document` | 列出已打开的文档，查看每个文档支持哪些修改 |
| `office_capture_selection` / `office_list_references` / `office_get_reference` | 把当前选区捕获成只读引用，并读取引用内容和位置 |
| `office_apply_edit` | 执行一次经过能力检查的修改；执行前先保存恢复记录 |
| `office_get_operation` / `office_revert_edit` | 查询修改结果；在内容未被你再次改动时撤回这次修改 |
| `office_save_document` | 显式保存，走的是与 Office 保存按钮相同的路径 |
| `office_release_operation` | 你确认不再需要后，释放恢复记录（不会删除原文件） |
| `office_capabilities` / `office_register_chat` | 查询支持的操作；登记接收引用的聊天 |

修改都通过 Office 自己的对象模型完成，结果直接体现在打开的文档里。每次修改都有恢复记录，不确定是否生效时只能先查询，不会重复写入。

### Add to Chat：引导 AI 修改哪里

在 Word / Excel / PowerPoint 里选中一段内容，点功能区或右键菜单的 **Add to Chat**，这段选区就会作为引用放进编辑器的聊天输入框。你可以接着说明要怎么改，AI 再按这个引用的位置调用 Office MCP 去修改。

- 引用只放进输入框，不会自动发送；
- 只投递给已经登记过的聊天，找不到目标时不做任何事；
- 引用是只读快照，不会修改或保存文档。

## 安装

### 从 Release 安装

1. 在 [Releases](https://github.com/liketocood345/office-in-editor/releases) 下载 `office-in-editor-<版本>-dist.zip` 并解压。
2. 保存并退出所有 Office 程序。
3. 在解压目录里运行（`-Hosts` 可填 `cursor`、`vscode`，或两者都填）：

   ```powershell
   .\scripts\install.ps1 -VsixPath .\office-in-editor-<版本>.vsix -Hosts cursor
   ```

4. 在编辑器里执行 **Developer: Reload Window**。Office 下次启动时会加载新的加载项。

安装器只为当前用户安装：组件复制到 `%LOCALAPPDATA%\Programs\OfficeInEditor`，加载项注册在当前用户名下。它不会关闭 Office，也不会修改默认文件关联或系统级注册。卸载请运行 `scripts\uninstall.ps1`。

### 发行包结构

```
office-in-editor-<版本>/
├─ README.md、LICENSE
├─ office-in-editor-<版本>.vsix   VS Code 插件
├─ payload/
│  ├─ core/     Core（自带 .NET 8 运行时）、MCP 页面、抓帧程序
│  ├─ host/     插件使用的原生模块
│  └─ office/   Office 扩展（COM 加载项）
└─ scripts/     安装、卸载和加载项注册脚本
```

### 从源码构建

需要 Node.js、.NET 8 SDK 和 Visual Studio C++ 生成工具（x64）。

```powershell
npm ci
npm run package:vsix
.\scripts\install.ps1 -VsixPath .\artifacts\office-in-editor-0.1.4.vsix -Hosts cursor
```

## 源码目录

| 目录 | 内容 |
| --- | --- |
| `packages/adapters/vsix` | VS Code 插件（标签页编辑器、Core 启动、MCP 注册） |
| `packages/adapter-sdk` | 插件与 Core 之间的通信封装 |
| `packages/office-addin` | Office 扩展（COM 加载项、Add to Chat、文档操作） |
| `packages/shared-core` | Core：MCP 服务、文档身份、修改与恢复、窗口调度 |
| `packages/focus-capture` | Office 窗口抓帧程序 |
| `packages/native-host` | 插件使用的原生窗口模块 |
| `packages/contracts` | 插件、Core 与加载项之间的消息约定 |
| `apps/mcp-app` | 在聊天里展示引用的 MCP 页面 |
| `scripts` | 构建、打包、安装与卸载脚本 |

## 隐私

插件只在本机运行，不向外部服务发送文档内容。诊断信息默认只保存在内存里；可选的调试记录只记阶段名和数字，不记录文件路径、正文或聊天内容。

## 当前限制

- Excel 和 PowerPoint 打开时，Office 窗口可能短暂出现；
- PowerPoint 的滚轮转发，以及关闭编辑器时处理未保存文档的流程，尚未完整验证；
- Visio、OneNote、Publisher、Project 只走原位 OLE 路线，属于实验性支持。

## 许可

许可证类型：**混合**，详见 [LICENSE](LICENSE)。

- 来自上游项目的组件沿用上游许可证，许可证原文已复制进 LICENSE 和随组件分发的文件：
  - Core 自带的 .NET 与 ASP.NET Core 运行时：MIT；
  - MCP 页面中打包的 JavaScript 依赖：各自的许可证，原文见 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)。
- 没有上游的部分和本项目自己编写的代码：BSD 3-Clause License。

组件清单见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
