# Office 焦点切换实验

此实验在 Windows 交互桌面用独立配置的 VS Code 承载合成 DOCX、XLSX 和 PPTX。标签内的 Word、Excel、PowerPoint 按钮在同一窗口打开对应文档标签。每个原生辅助程序只关联本次文档窗口和该 VS Code。点击预览会先停止采集，再把 Office 的可见窗口边界对齐到内容区。失焦后原生窗口移出屏幕并恢复采集。

## 运行

在仓库根目录执行。需要 MSVC、Windows SDK，以及测试机上已登录同一 SSH 用户的 Windows 桌面、桌面 Word、Excel、PowerPoint 和 VS Code 1.121。

```powershell
& ./experiments/focus-hybrid/build.ps1
& ./experiments/focus-hybrid/remote.ps1 -Remote ollama-relay
& ./experiments/focus-hybrid/remote.ps1 -Remote ollama-relay -Run <run>
```

默认启动长驻的交互会话，从 `Office in Editor · Word` 标签进入。`-App excel` 或 `-App powerpoint` 可选择首个打开的应用。`-Smoke -PixelPolicy mouse` 运行单应用自动验收并检查鼠标预热。已开启对应交互会话时，再次启动会返回同一个运行编号。

Office 加载项、右键 Add to Chat 和 Core 沿用仓库的 `npm run build` 与 `scripts/install.ps1 -SkipEditors`。用隔离扩展随标签注册同一个 MCP 服务。已有 Office 进程按自身的加载周期处理加载项。内置管理员且 UAC 关闭的专用 SSH 测试机需要先用 `register-lab-com.ps1` 注册加载项的机器 COM 类，Office 的每用户加载设置仍由原安装脚本管理。

像素转发策略在 Word、Excel 和 PowerPoint 的 **Office in Editor** 原生功能区中设置。常开策略在编辑期间保留已准备的资源。切换策略等到失焦才准备。鼠标策略在当前 Office 文档处于前台时用小幅移动信号提前准备。所有策略均在确认原生窗口离屏后才让 WGC 出帧。

每次远端启动使用独立目录。`events.jsonl` 持续记录应用、尺寸、采集启停和焦点，`result.json` 汇总整轮实验，`finished.json` 记录应用退出和原文件读回。VS Code 的 `Office Focus Hybrid` 输出面板也显示同一批事件。`-Run` 查询运行中会话时会把最新事件保存为 `.local/focus-hybrid-runs/<run>/ssh-live.json`，结束后会写入 `ssh-evidence.json`。

`preview_painted` 在 Webview 的图像解码完成后记录。`wgc_parked_frame_sampled` 表示辅助程序在窗口离屏时读到了非纯色帧。`paintedWhileParked` 表示 Webview 解码了标记为离屏采样的帧。相同像素会跳过更新，离屏时可以继续显示上一帧。`smoke_activation_request` 来自自动探针，`preview_user_click` 来自 Webview 点击。

在 SSH 交互桌面点击预览即可进入对应的 Office 编辑界面。要手动验证保存，可以输入 `OIE_FOCUS_SAVED` 并按 Ctrl+S，再点回 VS Code 内容区之外。检查 `native_foreground_confirmed`、`office_keyboard_focus_confirmed`、`office_foreground_released` 和后续 `native_parked`。如果出现 `office_modal_blocking`，探针会继续等待对应的 Office 对话框关闭。

无人值守实验使用已验证的合成文档窗口调用对应 Office 对象模型，写入保存标记并执行 `Save()`。之后独立读回原包内的 Word 正文、Excel E2 或 PowerPoint 首张幻灯片，同时校验文件哈希已变化。焦点结果来自 Windows 前台窗口和键盘焦点的实际状态。

对齐采用 UI Automation 提供的 Webview 内容区物理像素边界，以及 DWM 提供的原生 Office 可见窗口边界。网页预览不加边框和留白，采集多出的窗口外缘在编码前裁掉。Windows 10 的采集会话只在原生窗口移出所有显示器后启动，编辑区内不会出现采集框。这个版本使用内存中的低频 PNG。日志只记录状态、时间和尺寸。跨显示器 DPI、中文 IME 和其他 Office 应用仍需单独验收。

动图观察使用合成 PowerPoint 首张幻灯片。该 GIF 在放映时循环播放。普通编辑视图中的观察助手只在演示文稿失焦后移动同一图形，记录处于完全离屏状态时的移动次数。保持 PowerPoint 标签可见，点击 VS Code 的标题栏或侧栏即可观察。

SSH 的 `ad97332a6889416388148297400bbb1e` 实验中，切换侧栏让内容区从 `852 × 708` 变为 `552 × 708`，两种布局下 `native_edit_alignment` 的最大边界误差都是 0 像素。两次进入编辑前均已停止 WGC，回切并确认 Word 离屏后重新启动。首轮保存后的原文件同时通过哈希变化和内容标记读回，合成文档的锁已释放。

## P0 阶段扫测

短轮先覆盖三款应用，再运行整张矩阵。以下命令需要当前实验机的授权 SSH 别名。每次只运行一个自动验收用例，交互桌面中的其他实验窗口可继续保留。

```powershell
& ./experiments/focus-hybrid/build.ps1
& ./experiments/focus-hybrid/sweep.ps1 -Remote '<ssh-alias>' -Policies mouse -Cycles 3 -Repeat 1 -FailFast
& ./experiments/focus-hybrid/sweep.ps1 -Remote '<ssh-alias>' -Cycles 10 -Repeat 2
& ./experiments/focus-hybrid/remote.ps1 -Remote '<ssh-alias>' -Run '<run-id>' -Collect
node ./experiments/focus-hybrid/analyze.mjs .local/focus-hybrid-runs/<run-id>
node ./experiments/focus-hybrid/analyze.mjs --suite .local/focus-hybrid-runs/sweeps/<sweep-id>
```

`-Scenarios raise` 扫测停放中被外部拉到前台的情形（缺陷 D1）：每轮先让隔离 IDE 取得前台，再把停放的 Office 拉到前台，要求它在 850 ms 内转入编辑态并对齐标签，`metrics.json` 记录 `recovered`、`offscreenAtDeadline` 和接管耗时。`-Scenarios raise-baseline` 用同一构建关闭接管逻辑，用来复现缺陷做 A/B。两者可与默认的 `click` 组合，例如 `sweep.ps1 -Remote '<ssh-alias>' -Policies switch -Scenarios 'raise-baseline','raise' -Cycles 10 -Repeat 2`。

也可用 `remote.ps1 -Smoke -Cycles 3 -App word -PixelPolicy mouse` 启动单个短轮。省略 `-Cycles` 保留原来的保存与侧栏缩放冒烟。扫测向精确隔离窗口发送 Windows 鼠标点击，Webview 必须收到受信任的点击事件，随后验证 Office 持有真正的前台与键盘焦点。回切在点击 IDE 状态栏与请求 IDE 前台之间交替。第一轮保存合成文件，最后从 Office 文件包独立读回。鼠标策略的奇数轮提供预热动作。此模式会取得实验桌面的前台，手工编辑应与扫测分开。

每轮以 `corr.activationId` 串联点击和键盘焦点，以 `corr.parkId` 串联失焦、停放、采集启动和该轮首帧，以 `corr.frameSeq` 标记具体预览帧。即使内容尺寸相同，仍记录每次停放后首次完成图像解码并经过浏览器帧回调的预览。`tMs` 是宿主单调时钟，原生 `nativeUs` 在 `office_window_verified` 之后连续做 5 次时钟往返，取往返最短的一次写入 `native_clock_calibrated`，`metrics.json` 的 `clock` 汇总样本与往返分布；旧运行只有一次启动往返，标为 `legacyStartupSample`。运行日志可用 `-Collect` 取回。文件包含阶段、时间、尺寸和统计，预览帧本身只经内存传输。

`environment.json` 记录 Windows、Office、VS Code、UAC 和构建指纹。辅助程序补记显示器数量、当前 IDE DPI 和 Office 位数。扫测按秒记录采集读到的帧、送出的帧、相同帧、节流、编码上限和辅助程序 CPU。每组 `metrics.json` 记录激活与首帧的 P50/P95/最大值、静止采样和原件验证；矩阵摘要另检查同构建复跑的 P95 偏差。异常时先看 `sweep_input_blocked`、`sweep_timeout` 或 `sweep_focus_unconfirmed`，再看第一条缺失的关联事件。首次前台准备的结果见 `smoke_editor_focus_prepared`。
