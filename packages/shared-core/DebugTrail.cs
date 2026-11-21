using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

/// <summary>
/// Opt-in debug trail on disk, for staged debugging on a lab machine. It is off unless the file
/// <c>diagnostics\trace.on</c> exists in Core's data root (checked every few seconds, so it can be switched on and off
/// while Core runs). Each event is one NDJSON line in <c>diagnostics\trace.ndjson</c>: time, process id, area, stage, and
/// only the numeric and true/false fields of its data. Text is never written, so paths, titles and document content
/// cannot reach the file. Lines are written on a background thread; callers never wait on the disk.
/// </summary>
internal static class DebugTrail
{
    const long MaxBytes = 4 * 1024 * 1024;
    const long RecheckMs = 2000;
    static readonly BlockingCollection<string> Pending = new(new ConcurrentQueue<string>(), 4096);
    static readonly object Gate = new();
    static string? directory;
    static volatile bool enabled;
    static long nextCheck;
    static Thread? writer;
    static int written, flushed;

    /// <summary>Sets the data root; until then nothing is written.</summary>
    internal static void Configure(string dataRoot)
    {
        lock (Gate)
        {
            directory = Path.Combine(Path.GetFullPath(dataRoot), "diagnostics");
            nextCheck = 0;
        }
        Refresh();
    }

    internal static bool Enabled { get { Refresh(); return enabled; } }

    static void Refresh()
    {
        var now = Environment.TickCount64;
        if (now < Volatile.Read(ref nextCheck)) return;
        lock (Gate)
        {
            if (now < nextCheck) return;
            nextCheck = now + RecheckMs;
            try { enabled = directory is not null && File.Exists(Path.Combine(directory, "trace.on")); }
            catch (Exception) { enabled = false; }
            if (enabled && writer is null)
            {
                writer = new Thread(Run) { IsBackground = true, Name = "oie-debug-trail" };
                writer.Start();
            }
        }
    }

    internal static void Write(string area, string stage, object? data = null)
    {
        if (!Enabled) return;
        try
        {
            var line = new JsonObject
            {
                ["t"] = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
                ["pid"] = Environment.ProcessId,
                ["area"] = Safe(area),
                ["stage"] = Safe(stage),
            };
            var fields = NumbersOnly(data);
            if (fields is not null) line["data"] = fields;
            if (Pending.TryAdd(line.ToJsonString())) Interlocked.Increment(ref written);
        }
        catch (Exception) { /* Diagnostics never disrupt Core. */ }
    }

    /// <summary>Waits until everything written so far is on disk (tests and shutdown).</summary>
    internal static void Flush(int timeoutMs = 2000)
    {
        var target = Volatile.Read(ref written);
        var deadline = Environment.TickCount64 + timeoutMs;
        while (Volatile.Read(ref flushed) < target && Environment.TickCount64 < deadline) Thread.Sleep(10);
    }

    static JsonObject? NumbersOnly(object? data)
    {
        if (data is null) return null;
        JsonNode? node;
        try { node = data as JsonNode ?? JsonSerializer.SerializeToNode(data); }
        catch (Exception) { return null; }
        if (node is not JsonObject source) return null;
        var kept = new JsonObject();
        foreach (var (key, value) in source)
        {
            if (value is not JsonValue scalar) continue;
            var kind = scalar.GetValueKind();
            if (kind is JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False) kept[Safe(key)] = scalar.DeepClone();
        }
        return kept.Count == 0 ? null : kept;
    }

    static string Safe(string? value) =>
        value is { Length: > 0 and <= 96 } && value.All(c => char.IsAsciiLetterOrDigit(c) || c is '.' or '_' or '-' or ':') ? value : "invalid";

    static void Run()
    {
        foreach (var line in Pending.GetConsumingEnumerable())
        {
            try
            {
                var target = directory;
                if (target is not null)
                {
                    Directory.CreateDirectory(target);
                    var file = Path.Combine(target, "trace.ndjson");
                    var info = new FileInfo(file);
                    if (info.Exists && info.Length > MaxBytes) File.Move(file, Path.Combine(target, "trace.1.ndjson"), overwrite: true);
                    File.AppendAllText(file, line + "\n");
                }
            }
            catch (Exception) { /* Best effort. */ }
            Interlocked.Increment(ref flushed);
        }
    }
}
