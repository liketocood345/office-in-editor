using System.Diagnostics;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

/// <summary>
/// Runs packages/focus-capture (oie-focus-capture.exe) for one focus session. Commands go over stdin; status and
/// PNG frames come back on stdout in the stream envelope. The helper exits when Core closes stdin, including when
/// Core dies. A helper that keeps crashing is given up on and reported as <c>capture_unavailable</c>.
/// </summary>
internal sealed class FocusCaptureHost : IFocusCapture, IDisposable
{
    public const string ExecutableName = "oie-focus-capture.exe";
    const int MaxStartsPerMinute = 3;
    static readonly HashSet<string> Failures = new(StringComparer.Ordinal)
        { "capture_unsupported", "capture_device_failed", "capture_item_failed", "capture_session_failed", "capture_start_failed", "capture_target_invalid" };
    readonly object gate = new();
    readonly long hwnd;
    readonly int pid;
    readonly Action<int, int, int, string, byte[]> frame;
    readonly Action<string> notice;
    readonly Queue<long> starts = new();
    Process? process;
    StreamWriter? input;
    bool givenUp, disposed;
    int frameRate = FocusStateMachine.ForegroundFps;

    /// <param name="frame">(parkId, width, height, format, image) for every frame the helper sends.</param>
    /// <param name="notice">Receives <c>capture_unavailable</c> once when capture cannot run.</param>
    public FocusCaptureHost(long hwnd, int pid, Action<int, int, int, string, byte[]> frame, Action<string> notice)
    {
        this.hwnd = hwnd; this.pid = pid; this.frame = frame; this.notice = notice;
    }

    /// <summary>The helper next to Core, or OIE_FOCUS_CAPTURE for lab builds.</summary>
    public static string? Locate()
    {
        var overridden = Environment.GetEnvironmentVariable("OIE_FOCUS_CAPTURE");
        var path = string.IsNullOrEmpty(overridden) ? Path.Combine(AppContext.BaseDirectory, ExecutableName) : overridden;
        return Path.IsPathFullyQualified(path) && File.Exists(path) ? path : null;
    }

    public bool Available { get { lock (gate) return !givenUp && !disposed && Locate() is not null; } }
    public void Prepare(string reason) => Send("prepare");
    public void Publish(int parkId) => Send("publish " + parkId);
    public void Stop() { lock (gate) if (process is not null) Write("stop"); }

    public void SetFrameRate(int fps)
    {
        lock (gate)
        {
            frameRate = Math.Clamp(fps, 0, 60);
            // A helper that is not running picks the rate up when it starts.
            if (process is not null) Write("fps " + frameRate);
        }
    }

    void Send(string command)
    {
        lock (gate)
        {
            if (disposed || givenUp) return;
            if (process is null && !Start()) return;
            Write(command);
        }
    }

    void Write(string command)
    {
        try { input!.Write(command + "\n"); input.Flush(); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { /* The reader notices the exit. */ }
    }

    bool Start()
    {
        var now = Environment.TickCount64;
        while (starts.Count > 0 && now - starts.Peek() > 60_000) starts.Dequeue();
        var path = Locate();
        if (path is null || starts.Count >= MaxStartsPerMinute) { GiveUp(path is null ? "missing" : "restarts"); return false; }
        starts.Enqueue(now);
        var start = new ProcessStartInfo(path) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true };
        start.ArgumentList.Add(hwnd.ToString(System.Globalization.CultureInfo.InvariantCulture));
        start.ArgumentList.Add(pid.ToString(System.Globalization.CultureInfo.InvariantCulture));
        try { process = Process.Start(start); }
        catch (Exception error) when (error is System.ComponentModel.Win32Exception or IOException) { process = null; }
        if (process is null) { GiveUp("start_failed"); return false; }
        input = process.StandardInput;
        input.AutoFlush = false;
        Write("fps " + frameRate);
        var running = process;
        _ = Task.Run(() => ReadAsync(running));
        return true;
    }

    void GiveUp(string why)
    {
        if (givenUp) return;
        givenUp = true;
        DebugNdjson.Write("F", "FocusCapture", "given_up", new { why });
        notice("capture_unavailable");
    }

    async Task ReadAsync(Process running)
    {
        var stream = running.StandardOutput.BaseStream;
        var buffer = new byte[1 << 16];
        var pending = new MemoryStream();
        try
        {
            int read;
            while ((read = await stream.ReadAsync(buffer).ConfigureAwait(false)) > 0)
            {
                pending.Write(buffer, 0, read);
                var data = pending.GetBuffer().AsMemory(0, (int)pending.Length);
                var offset = 0;
                while (true)
                {
                    var used = FocusContract.TryDecode(data.Span[offset..], out var message);
                    if (used == 0) break;
                    offset += used;
                    Handle(message);
                }
                if (offset > 0)
                {
                    var rest = data[offset..].ToArray();
                    pending.SetLength(0);
                    pending.Write(rest);
                }
            }
        }
        catch (Exception error) when (error is IOException or CoreFault or ObjectDisposedException or System.Text.Json.JsonException or InvalidOperationException)
        {
            DebugNdjson.Write("F", "FocusCapture", "read_failed", new { type = error.GetType().Name });
            try { running.Kill(); } catch { }
        }
        finally
        {
            lock (gate) if (process == running) { process = null; input = null; }
            running.Dispose();
        }
    }

    void Handle(FocusContract.StreamMessage? message)
    {
        switch (message)
        {
            case FocusContract.FrameMessage { Header: var header } item when header.ParkId > 0 && FocusContract.FrameFormats.Contains(header.Format):
                frame(header.ParkId, header.Width, header.Height, header.Format, item.Image);
                break;
            case FocusContract.EventMessage { Value: JsonObject value }:
                var stage = value["stage"] is JsonValue text && text.TryGetValue<string>(out var name) ? name : "";
                DebugNdjson.Write("F", "FocusCapture", stage, null);
                if (Failures.Contains(stage)) lock (gate) GiveUp(stage);
                break;
        }
    }

    public void Dispose()
    {
        Process? running;
        lock (gate)
        {
            if (disposed) return;
            disposed = true;
            running = process;
            try { input?.Write("exit\n"); input?.Flush(); input?.Dispose(); } catch (Exception error) when (error is IOException or ObjectDisposedException) { }
        }
        if (running is null) return;
        try { if (!running.WaitForExit(1000)) running.Kill(); }
        catch (Exception error) when (error is InvalidOperationException or System.ComponentModel.Win32Exception) { }
    }
}
