# 窗口移动跟随修复（2026-11-15）

## 原因

原生 Office 容器当前是 IDE 所有的 popup。IDE 平移不改变 Webview 尺寸，resize/ResizeObserver 不会因此发出布局消息；原有 owner-insets 跟随只有收到布局请求才执行。旧闲时校准也只比较宽高，漏掉左上角位置改变，所以拖动后容易留在旧位置。

## 实现

- NativeOwnerTracker 在一个共享消息线程上，按已经验证的 owner PID/TID 订阅窗口位置、移动结束、最小化/恢复及销毁事件；只接受该 HWND 的窗口自身事件。没有鼠标/键盘钩子、DLL 注入或输入转发。
- 每个容器最多保留一个待处理位置更新。按“已验证的 owner/容器绝对锚点 + owner 当前位置”计算目标，不累计上次移动量。纯平移仅异步 SetWindowPos，保留尺寸、Z 序和焦点；不进入 Office COM、不改样式、不做 UIA 搜索。
- 尺寸或 DPI 改变时先使锚点失效并隐藏旧矩形。交互拖动结束后合并校准，只读取首次定位时缓存的四角/精确节点；不重新枚举桌面、遍历无障碍树。渲染器短暂滞后最多补两次延迟重试，不启动常驻轮询。
- WebviewSurface 的每 STA 锚点缓存上限为 16，独立持有并释放 COM 引用；UIA、MSAA 与精确 HWND 路径均能定点刷新。关闭标签或更换 token 会释放引用与订阅；最后一个订阅关闭时退出共享消息线程。
- 隐藏标签只隐藏，不把原生窗口搬到桌面原点或按 CSS 尺寸缩小。迟到事件、最小化/恢复及 Office ShowObject 回调都不能把失效矩形重新显示。
- Word 正文布局按 client/view/active-object 状态去重；同值位置回调不再反复自排队，队列满会清除 queued 标志。ResizeBorder 调用期间活动对象改变时，完成后仍会补新对象的一轮。

所用 Windows 接口及线程约束见 [SetWinEventHook](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwineventhook)、[SetWindowPos](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos)。

## 恢复入口

若 IDE 重建无障碍节点导致锚点永久失效，选中相应 Office 标签，在命令面板执行 **Office: Re-align Embedded Window**。它通过 office.attach_window 显式重建定位；不会重新打开文档、创建第二个 OLE 视图或丢弃修改。重新发现只发生在显式恢复/首次附着，不放进移动热路径。

## 验证边界

新增测试覆盖纯平移、反复往返无累计漂移、其它窗口/子对象事件隔离、双容器独立跟随、隐藏/最小化/恢复、锚点失效后的迟到事件，以及释放最后订阅后线程退出。真实事件测试仅使用临时 User32 合成窗口，不启动 Office 或操作用户文档。

~~~powershell
dotnet run --project packages/shared-core/Tests/OfficeInEditor.Core.Tests.csproj --no-restore -- --native-owner-smoke
npm run typecheck
npm test
~~~

合成窗口通过不等于已经完成真实 Word × Cursor/VS Code × 多屏 DPI 的视觉验收；更新后仍需检查实际编辑界面、最大化/Snap、跨屏和分屏。此修复保留当前 owned-popup 路线，不将其宣称为 IDE 提供的原生子窗口槽。
