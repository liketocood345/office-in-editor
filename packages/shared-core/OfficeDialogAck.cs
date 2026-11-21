using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeInEditor.Core;

/// <summary>
/// Nested OLE only: when a warning dialog has a single confirm button, post a click.
/// Does not set DisplayAlerts, does not run for standalone Office opens, and is not a keyboard hook.
/// </summary>
internal static class OfficeDialogAck
{
    private const uint BmClick = 0x00F5;
    private const uint WmCommand = 0x0111;
    private const int IdOk = 1;
    private const int IdCancel = 2;
    private const int IdYes = 6;
    private const int IdNo = 7;
    private const int GwlStyle = -16;
    private const int BsTypeMask = 0x000F;
    private const int WsVisible = 0x10000000;
    private const int MaxClicked = 64;
    private static readonly object Gate = new();
    private static readonly Dictionary<string, int> Names = new(StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<uint> Watches = [];
    private static readonly ConcurrentDictionary<long, byte> Clicked = new();
    private static readonly ConcurrentDictionary<uint, Process> WatchedProcesses = new();
    private static int enabled = 1;
    private static long lastPulse;

    internal static bool Enabled
    {
        get => Volatile.Read(ref enabled) != 0;
        set => Volatile.Write(ref enabled, value ? 1 : 0);
    }

    internal static IDisposable Scope(string serverFile, bool allow)
    {
        if (!allow || !Enabled) return Noop.Instance;
        var name = Path.GetFileNameWithoutExtension(serverFile);
        if (string.IsNullOrEmpty(name)) return Noop.Instance;
        lock (Gate) Names[name] = Names.TryGetValue(name, out var count) ? count + 1 : 1;
        Pulse("scope");
        return new KindScope(name);
    }

    internal static void Watch(uint pid, bool allow)
    {
        if (pid == 0 || !allow || !Enabled) return;
        lock (Gate) Watches.Add(pid);
        // #region agent log
        try
        {
            if (WatchedProcesses.ContainsKey(pid)) return;
            var process = Process.GetProcessById((int)pid);
            process.EnableRaisingEvents = true;
            process.Exited += (_, _) =>
            {
                var code = -1;
                try { code = process.ExitCode; } catch { }
                DebugNdjson.Probe("J", "OfficeDialogAck.Watch", "server_exited", new { pid, code });
            };
            if (!WatchedProcesses.TryAdd(pid, process)) process.Dispose();
        }
        catch { }
        // #endregion
    }

    internal static void Unwatch(uint pid)
    {
        if (pid == 0) return;
        lock (Gate) Watches.Remove(pid);
    }

    internal static void Pulse(string origin = "pending")
    {
        if (!Enabled) return;
        var now = Stopwatch.GetTimestamp();
        if (origin == "pending")
        {
            var previous = Interlocked.Read(ref lastPulse);
            if (previous != 0 && Stopwatch.GetElapsedTime(previous, now).TotalMilliseconds < 50)
                return;
        }
        Interlocked.Exchange(ref lastPulse, now);
        HashSet<uint> pids;
        lock (Gate)
        {
            if (Names.Count == 0 && Watches.Count == 0) return;
            pids = [.. Watches];
            if (pids.Count == 0)
            {
                foreach (var name in Names.Keys)
                {
                    try
                    {
                        foreach (var process in Process.GetProcessesByName(name))
                        {
                            try { pids.Add((uint)process.Id); }
                            finally { process.Dispose(); }
                        }
                    }
                    catch { }
                }
            }
        }
        if (pids.Count == 0) return;
        var clicked = 0;
        var seen = 0;
        NativeWindow.EnumProc callback = (window, lp) =>
        {
            if (clicked >= 4) return false;
            if (!NativeWindow.IsWindowVisible(window)) return true;
            NativeWindow.GetWindowThreadProcessId(window, out var pid);
            if (pid == 0 || !pids.Contains(pid)) return true;
            var className = WebviewSurface.ClassName(window);
            if (!IsWarningClass(className)) return true;
            seen++;
            if (TryPostConfirm(window, out _)) clicked++;
            return true;
        };
        NativeWindow.EnumWindows(callback, IntPtr.Zero);
        GC.KeepAlive(callback);
        if (clicked > 0) NativeWindow.PumpWait(30);
        // #region agent log
        if (clicked > 0)
            DebugNdjson.Probe("B", "OfficeDialogAck.Pulse", "clicked",
                new { origin, pids = pids.Count, seen, clicked });
        // #endregion
    }

    private static bool IsWarningClass(string name) =>
        name is "#32770" or "NUIDialog"
        || name.StartsWith("bosa_sdm", StringComparison.OrdinalIgnoreCase);

    private static bool TryPostConfirm(IntPtr dialog, out string skip)
    {
        var key = dialog.ToInt64();
        if (Clicked.ContainsKey(key))
        {
            skip = "already";
            return false;
        }
        var buttons = new List<(IntPtr Hwnd, int Id, string Text, bool Default)>();
        NativeWindow.EnumChildProc child = (window, _) =>
        {
            if (!NativeWindow.IsWindowVisible(window)) return true;
            var className = new StringBuilder(32);
            NativeWindow.GetClassName(window, className, className.Capacity);
            if (!className.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase)) return true;
            var style = NativeWindow.GetWindowLong(window, GwlStyle);
            if ((style & WsVisible) == 0) return true;
            var type = style & BsTypeMask;
            if (type is not (0 or 1)) return true;
            var caption = new StringBuilder(64);
            NativeWindow.GetWindowText(window, caption, caption.Capacity);
            buttons.Add((window, GetDlgCtrlID(window), caption.ToString(), type == 1));
            return buttons.Count < 8;
        };
        NativeWindow.EnumChildWindows(dialog, child, IntPtr.Zero);
        GC.KeepAlive(child);
        if (buttons.Count != 1)
        {
            skip = "count:" + buttons.Count;
            return false;
        }
        var button = buttons[0];
        skip = button.Id + ":" + Normalize(button.Text);
        if (IsCancel(button.Id, button.Text) || !IsConfirm(button.Id, button.Text, button.Default)) return false;
        Remember(key);
        PostMessage(button.Hwnd, BmClick, IntPtr.Zero, IntPtr.Zero);
        PostMessage(dialog, WmCommand, new IntPtr(button.Id == 0 ? IdOk : button.Id), button.Hwnd);
        skip = "clicked:" + skip;
        return true;
    }

    private static bool IsConfirm(int id, string text, bool isDefault)
    {
        if (id is IdOk or IdYes) return true;
        var label = Normalize(text);
        if (IsConfirmLabel(label)) return true;
        return isDefault && label.Length == 0;
    }

    private static bool IsCancel(int id, string text)
    {
        var label = Normalize(text);
        if (IsConfirmLabel(label)) return false;
        if (label is "Cancel" or "取消" or "No" or "否" or "Don't Save" or "不保存") return true;
        return label.Length == 0 && id is IdCancel or IdNo;
    }

    private static bool IsConfirmLabel(string label) =>
        label is "OK" or "确定" or "确认" or "是" or "Yes" or "Continue" or "继续" or "Close" or "关闭";

    private static string Normalize(string text) => text.Replace("&", "", StringComparison.Ordinal).Trim().TrimEnd('.', '。');

    private static void Remember(long hwnd)
    {
        if (Clicked.Count >= MaxClicked) Clicked.Clear();
        Clicked[hwnd] = 0;
    }

    private sealed class KindScope : IDisposable
    {
        private readonly string name;
        private int disposed;
        internal KindScope(string name) => this.name = name;
        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) != 0) return;
            lock (Gate)
            {
                if (!Names.TryGetValue(name, out var count)) return;
                if (count <= 1) Names.Remove(name);
                else Names[name] = count - 1;
            }
        }
    }

    private sealed class Noop : IDisposable
    {
        internal static readonly Noop Instance = new();
        public void Dispose() { }
    }

    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
}
