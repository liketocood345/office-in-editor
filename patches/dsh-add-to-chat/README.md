# 补丁：Add to Chat 投递到 DeepSeek Harness

可选补丁，不在主线源码里。打上后，Office 里的 **Add to Chat** 会把选区放进 DeepSeek Harness（dsh）的 VS Code 聊天框，而不是 Cursor composer / VS Code 自带聊天。

## 适用扩展

| 扩展 | 商店 ID | 投递方式 |
| --- | --- | --- |
| [skymecode/deepseek-harness-for-vscode](https://github.com/skymecode/deepseek-harness-for-vscode) | `skymecode.deepseek-harness-for-vscode` | 选区附件（chip） |
| [pwhMass/deepseek-harness-vscode](https://github.com/pwhMass/deepseek-harness-vscode) | `lucasliang.harness-connector-deepseek` | `@` 文件提及 |

两个扩展的命令都以 `deepseekHarness.` 开头并有同名命令，不能同时启用，只启用其中一个。两个都装时补丁优先投递给 skymecode 版。

## 原理

两个扩展都没有给其他扩展提供「往输入框放内容」的接口，补丁只借用它们给自己用户提供的功能：

- **skymecode 版**：它会把当前文本编辑器里的选中文字自动作为附件挂到输入框（设置 `deepseekHarness.autoAttachSelection`，默认开）。补丁把引用（文档名 + 选区文字）放进一个只读虚拟文档 `<文档名> 选区.md`，在编辑器里打开并全选，再聚焦 `deepseekHarness.chatView`。不写磁盘；附件可以在发送前删掉。
- **pwhMass 版**：它有公开命令 `deepseekHarness.addFileToChat(uri)`，会往输入框插入 `@路径`。补丁把引用写成插件存储目录下的 `chat-refs/*.md`（只保留最近 12 个），再调用这条命令。dsh 的系统提示约定 `@` 后的绝对路径指本机文件，发送时由 AI 读取。它的页面在加载完成前会丢消息，所以首次投递前先聚焦 `deepseekHarness.sessionView` 并等 1.5 秒；之后页面常驻，不再等待。

两种方式都只放进输入框，不自动发送。投递失败时回落到原来的 Cursor / VS Code 聊天路径。

新增设置 `officeInEditor.chatTarget`：

- `auto`（默认）：装了上述扩展就投递给 DeepSeek Harness，否则用编辑器自带聊天；
- `builtin`：始终用编辑器自带聊天。

## 使用

在仓库根目录：

```powershell
git apply patches/dsh-add-to-chat/0001-add-to-chat-deepseek-harness.patch
npm run build
npx vsce package --no-dependencies --allow-missing-repository --out artifacts/office-in-editor-dsh.vsix
.\scripts\install.ps1 -VsixPath .\artifacts\office-in-editor-dsh.vsix -Hosts vscode
```

撤销补丁：`git apply -R patches/dsh-add-to-chat/0001-add-to-chat-deepseek-harness.patch`。

补丁基于提交 `187751b` 生成，改动三个文件：`package.json`（新设置）、`packages/adapters/vsix/src/extension.ts`（投递入口）、`packages/adapters/vsix/src/harness-chat.ts`（新增）。

## 已验证与未验证

- 已在隔离实验机（Windows、VS Code 1.121、dsh 0.2.2-rc.5）验证：skymecode 0.6.0 输入框出现选区附件，内容为 Word 选区；connector 0.4.1 输入框出现 `@…/chat-refs/<文档名>-<编号>.md`。
- 未验证：配置 API Key 后 AI 实际读取引用并回复。
- 不含 Office MCP：dsh 扩展不读 VS Code 的 MCP 注册，要让 dsh 的 AI 修改文档，需要另外把 Core（`--mcp`，stdio）配置为 dsh 的 MCP 服务器。
