# Office in Editor

面向 Windows VS Code/Cursor 的 Office 协作插件：共享 VSIX 薄外壳、私有 Core/MCP 与同一份 Office COM 加载项。不包含 Microsoft Office 本体，DSH 适配暂缓。

**当前实验方向是焦点混合**：Office 原生窗口持焦时盖在标签内容区上，失焦后标签显示本机离屏预览。它只在 `experiments/focus-hybrid` 中实现，SSH 实验已通过 Word、Excel、PowerPoint，尚未进入 VSIX。见[设计手册首节](docs/DESIGN.md#技术路线总结与路线裁决)；缺陷清单与攻坚规划在项目根目录的 `planning/HANDOFF.md`（不在本仓库内）。

**0.1.4 VSIX 仍是原位 OLE 原型，尚未通过全系列内嵌编辑验收。** 该路线规范见 [设计手册](docs/DESIGN.md)：标签内容只占 `#ole-root` 令牌矩形；打开成功前容器隐藏；保存只有保存按钮同一条原件 `SaveObject`，插件只保留定时保存（周期静默点该按钮）。Excel/PPT 待活测；无文档服务器的 OneNote 短状态失败。关闭取消、对话框归属仍是缺口。加载项选区/MCP 仍只分派 Word/Excel/PowerPoint。

- [当前架构审查、Word 布局原因与修复关卡](docs/ARCHITECTURE_AUDIT.md)
- [逐应用兼容要求与实现矩阵](docs/OFFICE_COMPATIBILITY.md)
- [历史本地自动保存与诊断验收记录](docs/LOCAL_AUTOSAVE_AND_DIAGNOSTICS.md)
- [设计手册](docs/DESIGN.md)
- [调试方案与历史审查](docs/DEBUG_AND_LOADING_PLAN.md)

```powershell
npm ci
npm run package:vsix
./scripts/install.ps1 -VsixPath ./artifacts/office-in-editor-0.1.4.vsix -Hosts cursor
```

安装前正常保存并退出 Office；安装完成后重新加载编辑器窗口，Office 在下次启动时加载新版。安装器不结束 Office，不更改默认文件关联或系统级注册。运行需要本机已安装并授权的 Microsoft Office。
