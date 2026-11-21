using OfficeInEditor.Core;

/// <summary>Capture helper host. The spawn check runs only when OIE_FOCUS_CAPTURE points at a built helper.</summary>
static class FocusCaptureHostTests
{
    public static void Run(Action<bool, string> check)
    {
        var configured = Environment.GetEnvironmentVariable("OIE_FOCUS_CAPTURE");
        var notices = new List<string>();
        Environment.SetEnvironmentVariable("OIE_FOCUS_CAPTURE", Path.Combine(Path.GetTempPath(), "missing-" + Guid.NewGuid().ToString("N"), FocusCaptureHost.ExecutableName));
        using (var missing = new FocusCaptureHost(1, 1, (_, _, _, _, _) => { }, notices.Add))
        {
            check(!missing.Available, "focus capture is unavailable without the helper");
            missing.Prepare("switch");
            missing.Publish(1);
            check(notices.SequenceEqual(["capture_unavailable"]), "focus capture reports a missing helper once");
        }
        Environment.SetEnvironmentVariable("OIE_FOCUS_CAPTURE", configured);
        if (string.IsNullOrEmpty(configured) || !File.Exists(configured)) return;

        // A window handle that does not exist: the helper must refuse it and Core must give up quietly.
        var spawned = new List<string>();
        using var host = new FocusCaptureHost(0x7ffffff0, Environment.ProcessId, (_, _, _, _, _) => { }, code => { lock (spawned) spawned.Add(code); });
        check(host.Available, "focus capture helper is found from OIE_FOCUS_CAPTURE");
        host.Prepare("switch");
        for (var i = 0; i < 50; i++) { lock (spawned) if (spawned.Count > 0) break; Thread.Sleep(100); }
        lock (spawned) check(spawned.SequenceEqual(["capture_unavailable"]) && !host.Available, "focus capture helper rejects an invalid window and Core gives up");
    }
}
