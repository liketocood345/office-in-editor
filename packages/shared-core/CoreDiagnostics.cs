using System.Collections.Concurrent;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

/// <summary>
/// In-memory trail of what Core was doing: fixed stage names, timestamps and managed thread ids only, never paths,
/// parameters or document content. It is read only when the stall watchdog fires.
/// </summary>
internal static class CoreTrace
{
    const int Capacity = 256;
    static readonly (long Ticks, string Area, string Stage, int Thread)[] Ring = new (long, string, string, int)[Capacity];
    static readonly object Gate = new();
    static long written;
    static long nextOperation;
    static readonly ConcurrentDictionary<long, (string Name, long Start, int Thread)> Active = new();

    /// <summary>Set when the broker starts shutting down; the watchdog reports a shutdown that does not finish.</summary>
    public static long ShutdownStartedAt;

    public static void Mark(string area, string stage)
    {
        lock (Gate)
        {
            Ring[written % Capacity] = (Environment.TickCount64, area, stage, Environment.CurrentManagedThreadId);
            written++;
        }
        DebugTrail.Write(area, stage);
    }

    /// <summary>Tracks an operation in flight (a broker request, a dispatcher call); dispose to end it.</summary>
    public static Operation Begin(string name)
    {
        var id = Interlocked.Increment(ref nextOperation);
        Active[id] = (Safe(name), Environment.TickCount64, Environment.CurrentManagedThreadId);
        return new Operation(id);
    }

    public readonly struct Operation(long id) : IDisposable
    {
        public void Dispose() => Active.TryRemove(id, out _);
    }

    public static IEnumerable<(string Name, long AgeMs, int Thread)> InFlight(long now) =>
        Active.Values.Select(item => (item.Name, now - item.Start, item.Thread)).OrderByDescending(item => item.Item2);

    public static JsonArray Recent(long now)
    {
        lock (Gate)
        {
            var count = (int)Math.Min(written, Capacity);
            var items = new JsonArray();
            for (var i = count; i >= 1; i--)
            {
                var entry = Ring[(written - i) % Capacity];
                items.Add(new JsonObject { ["agoMs"] = now - entry.Ticks, ["area"] = entry.Area, ["stage"] = entry.Stage, ["thread"] = entry.Thread });
            }
            return items;
        }
    }

    /// <summary>Operation names come from protocol method names and code labels; anything else is reduced to a marker.</summary>
    static string Safe(string name) =>
        name.Length is > 0 and <= 80 && name.All(c => char.IsAsciiLetterOrDigit(c) || c is '.' or '_' or '-' or ':') ? name : "unnamed";
}

/// <summary>
/// Watches the broker for stalls: the CoreState lock not obtainable, the thread pool not running work, a dispatcher
/// call or broker request running too long, or a shutdown that does not finish. On the first detection of each stall
/// it writes one snapshot (trace and counters, no content) to diagnostics/stall-*.json; nothing is written otherwise.
/// </summary>
internal sealed class StallWatchdog
{
    const int MaxFiles = 10;
    static readonly TimeSpan Probe = TimeSpan.FromSeconds(3);
    const long DispatcherLimitMs = 5_000, RequestLimitMs = 30_000, ShutdownLimitMs = 15_000;
    readonly string directory;
    readonly Func<TimeSpan, bool> tryLock;
    readonly Func<(string? Work, long StartedAt)> dispatcherWork;
    readonly HashSet<string> reported = new(StringComparer.Ordinal);
    readonly Thread thread;

    public StallWatchdog(string dataRoot, Func<TimeSpan, bool> tryLock, Func<(string? Work, long StartedAt)> dispatcherWork)
    {
        directory = Path.Combine(Path.GetFullPath(dataRoot), "diagnostics");
        this.tryLock = tryLock; this.dispatcherWork = dispatcherWork;
        thread = new Thread(Run) { IsBackground = true, Name = "stall-watchdog" };
    }

    public void Start() => thread.Start();

    void Run()
    {
        while (true)
        {
            Thread.Sleep(1000);
            try { Check(); } catch (Exception) { /* Diagnostics never take Core down. */ }
        }
    }

    /// <summary>One pass; public so tests can drive it without waiting.</summary>
    internal IReadOnlyList<string> Check()
    {
        var found = new List<string>();
        var lockStart = Environment.TickCount64;
        if (!tryLock(Probe)) found.Add("lock");
        var lockWaitMs = Environment.TickCount64 - lockStart;
        using (var ran = new ManualResetEventSlim(false))
        {
            ThreadPool.UnsafeQueueUserWorkItem(_ => { try { ran.Set(); } catch (ObjectDisposedException) { } }, null);
            if (!ran.Wait(Probe)) found.Add("threadpool");
        }
        var now = Environment.TickCount64;
        var (work, startedAt) = dispatcherWork();
        if (work is not null && now - startedAt > DispatcherLimitMs) found.Add("dispatcher");
        if (CoreTrace.InFlight(now).Any(item => item.AgeMs > RequestLimitMs)) found.Add("request");
        var shutdown = Volatile.Read(ref CoreTrace.ShutdownStartedAt);
        if (shutdown > 0 && now - shutdown > ShutdownLimitMs) found.Add("shutdown");
        foreach (var kind in reported.Except(found).ToArray()) { reported.Remove(kind); CoreTrace.Mark("watchdog", "recovered:" + kind); }
        var fresh = found.Where(kind => reported.Add(kind)).ToArray();
        foreach (var kind in fresh)
        {
            ThreadPool.GetAvailableThreads(out var freeWorkers, out _);
            ThreadPool.GetMaxThreads(out var maxWorkers, out _);
            DebugTrail.Write("watchdog", "stall:" + kind, new { threads = ThreadPool.ThreadCount, pending = ThreadPool.PendingWorkItemCount, busyWorkers = maxWorkers - freeWorkers });
        }
        if (fresh.Length > 0) Write(fresh, found, lockWaitMs, now, work, startedAt, shutdown);
        return found;
    }

    void Write(string[] fresh, List<string> all, long lockWaitMs, long now, string? work, long startedAt, long shutdown)
    {
        ThreadPool.GetAvailableThreads(out var availableWorkers, out var availableIo);
        ThreadPool.GetMaxThreads(out var maxWorkers, out var maxIo);
        var snapshot = new JsonObject
        {
            ["version"] = RuntimeIdentity.Version,
            ["detectedAt"] = DateTimeOffset.UtcNow.ToString("O"),
            ["newStalls"] = new JsonArray(fresh.Select(kind => (JsonNode)kind).ToArray()),
            ["allStalls"] = new JsonArray(all.Select(kind => (JsonNode)kind).ToArray()),
            ["lockWaitMs"] = lockWaitMs,
            ["threadPool"] = new JsonObject
            {
                ["threads"] = ThreadPool.ThreadCount, ["pending"] = ThreadPool.PendingWorkItemCount,
                ["busyWorkers"] = maxWorkers - availableWorkers, ["busyIo"] = maxIo - availableIo,
            },
            ["dispatcher"] = new JsonObject { ["work"] = work, ["runningMs"] = work is null ? 0 : now - startedAt },
            ["shutdownMs"] = shutdown > 0 ? now - shutdown : 0,
            ["inFlight"] = new JsonArray(CoreTrace.InFlight(now).Take(32)
                .Select(item => (JsonNode)new JsonObject { ["name"] = item.Name, ["ageMs"] = item.AgeMs, ["thread"] = item.Thread }).ToArray()),
            ["recent"] = CoreTrace.Recent(now),
        };
        try
        {
            Directory.CreateDirectory(directory);
            var name = "stall-" + string.Join("+", fresh) + "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + ".json";
            File.WriteAllText(Path.Combine(directory, name), snapshot.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
            foreach (var old in new DirectoryInfo(directory).GetFiles("stall-*.json").OrderByDescending(file => file.CreationTimeUtc).Skip(MaxFiles))
                old.Delete();
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { /* Best effort. */ }
    }
}
