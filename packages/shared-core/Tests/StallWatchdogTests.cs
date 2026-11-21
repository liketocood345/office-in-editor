using System.Text.Json.Nodes;
using OfficeInEditor.Core;

/// <summary>The stall watchdog writes one snapshot per stall, only when something is stuck, with no paths in it.</summary>
static class StallWatchdogTests
{
    public static void Run(Action<bool, string> check)
    {
        var root = Path.Combine(Path.GetTempPath(), "oie-stall-" + Guid.NewGuid().ToString("N"));
        var dir = Path.Combine(root, "diagnostics");
        try
        {
            var lockFree = true;
            (string?, long) work = (null, 0);
            var watchdog = new StallWatchdog(root, _ => lockFree, () => work);
            check(watchdog.Check().Count == 0 && !Directory.Exists(dir), "stall watchdog writes nothing while Core is healthy");

            CoreTrace.Mark("test", "before-stall");
            using (CoreTrace.Begin("req:focus.layout"))
            {
                lockFree = false;
                work = ("EndFocusSessionAsync:316", Environment.TickCount64 - 6_000);
                var found = watchdog.Check();
                check(found.Contains("lock") && found.Contains("dispatcher"), "stall watchdog detects a held lock and a long dispatcher call");
            }
            var files = Directory.GetFiles(dir, "stall-*.json");
            check(files.Length == 1, "stall watchdog writes one snapshot for the stall");
            var snapshot = JsonNode.Parse(File.ReadAllText(files[0]))!;
            check(snapshot["dispatcher"]!["work"]!.GetValue<string>() == "EndFocusSessionAsync:316" &&
                snapshot["inFlight"]!.AsArray().Any(item => item!["name"]!.GetValue<string>() == "req:focus.layout") &&
                snapshot["recent"]!.AsArray().Any(item => item!["stage"]!.GetValue<string>() == "before-stall"),
                "stall snapshot names the stuck work, the requests in flight and the recent trail");
            check(!File.ReadAllText(files[0]).Contains(Path.GetTempPath().TrimEnd('\\').Replace("\\", "\\\\"), StringComparison.OrdinalIgnoreCase),
                "stall snapshot carries no machine paths");

            watchdog.Check();
            check(Directory.GetFiles(dir, "stall-*.json").Length == 1, "stall watchdog does not repeat a snapshot for the same stall");
            lockFree = true; work = (null, 0);
            watchdog.Check();
            Thread.Sleep(1100);
            lockFree = false;
            watchdog.Check();
            check(Directory.GetFiles(dir, "stall-*.json").Length == 2, "stall watchdog writes again after recovery and a new stall");
        }
        finally { try { Directory.Delete(root, true); } catch { } }
    }
}
