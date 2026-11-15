using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

public sealed class CoreFault(string code, string message) : Exception(message)
{
    public string Code { get; } = code;
}

public static class Wire
{
    public const int MaxLineBytes = 2 * 1024 * 1024;
    public static readonly JsonSerializerOptions Json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
    public static JsonNode? Node(object? value) => JsonSerializer.SerializeToNode(value, Json);
    public static string Text(JsonNode? value, string key, int max = 512, bool required = true)
    {
        var item = value?[key];
        if (item is null && !required) return "";
        if (item is not JsonValue scalar || !scalar.TryGetValue<string>(out var text) ||
            (required && string.IsNullOrWhiteSpace(text)) || text.Length > max)
            throw new CoreFault("invalid_params", "A required field is invalid.");
        return text;
    }
    public static bool Boolean(JsonNode? value, string key) => value?[key] is JsonValue scalar && scalar.TryGetValue<bool>(out var result) && result;
    public static bool? OptionalBoolean(JsonNode? value, string key) => value?[key] is JsonValue scalar && scalar.TryGetValue<bool>(out var result) ? result : null;
    public static int Number(JsonNode? value, string key, int fallback = 0) => value?[key] is JsonValue scalar && scalar.TryGetValue<int>(out var result) ? result : fallback;
    public static string NewToken() => Convert.ToHexString(RandomNumberGenerator.GetBytes(32)).ToLowerInvariant();
    public static bool TokenEquals(string supplied, string expected) => supplied.Length == expected.Length &&
        CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(supplied), Encoding.UTF8.GetBytes(expected));
    public static async Task<string?> ReadLineAsync(Stream stream, CancellationToken cancellation)
    {
        using var line = new MemoryStream();
        var buffer = new byte[4096];
        int read;
        while ((read = await stream.ReadAsync(buffer, cancellation)) != 0)
        {
            var newline = Array.IndexOf(buffer, (byte)'\n', 0, read);
            var count = newline < 0 ? read : newline;
            if (line.Length + count > MaxLineBytes) throw new CoreFault("message_too_large", "Message exceeds the protocol limit.");
            line.Write(buffer, 0, count);
            if (newline >= 0) return new UTF8Encoding(false, true).GetString(line.ToArray()).TrimEnd('\r');
        }
        return line.Length == 0 ? null : new UTF8Encoding(false, true).GetString(line.ToArray());
    }
    public static async Task WriteAsync(Stream stream, object value, CancellationToken cancellation)
    {
        var bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(value, Json) + "\n");
        if (bytes.Length > MaxLineBytes) throw new CoreFault("message_too_large", "Response exceeds the protocol limit.");
        await stream.WriteAsync(bytes, cancellation);
        await stream.FlushAsync(cancellation);
    }
}

public static class RuntimeIdentity
{
    public const string Version = "0.1.4";
    public static string UserKey => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
        WindowsIdentity.GetCurrent().User?.Value ?? throw new CoreFault("identity_unavailable", "Windows identity is unavailable."))))[..16].ToLowerInvariant();
    public static string PipeName => "office-in-editor-v1-" + UserKey;
    public static string MutexName => "Local\\office-in-editor-v1-" + UserKey;
    public static string DataRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeInEditor");
    public static object Endpoint => new { pipeName = PipeName, protocolVersion = 1, version = Version };
    public static ProcessStartInfo Child(string argument)
    {
        var executable = Environment.ProcessPath ?? throw new CoreFault("runtime_unavailable", "Worker executable is unavailable.");
        var start = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true };
        if (Path.GetFileNameWithoutExtension(executable).Equals("dotnet", StringComparison.OrdinalIgnoreCase))
            start.ArgumentList.Add(typeof(RuntimeIdentity).Assembly.Location);
        start.ArgumentList.Add(argument);
        return start;
    }
}

public static class PipeClient
{
    public static async Task<JsonNode?> CallAsync(string pipeName, string method, JsonNode? parameters, CancellationToken cancellation)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        timeout.CancelAfter(TimeSpan.FromSeconds(55));
        await using var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        await pipe.ConnectAsync(timeout.Token);
        var id = Guid.NewGuid().ToString("N");
        await Wire.WriteAsync(pipe, new { v = 1, id, method, @params = parameters }, timeout.Token);
        var line = await Wire.ReadLineAsync(pipe, timeout.Token) ?? throw new CoreFault("connection_closed", "The worker closed the connection.");
        var reply = JsonNode.Parse(line) ?? throw new CoreFault("invalid_response", "Worker response is invalid.");
        if (reply["id"]?.GetValue<string>() != id) throw new CoreFault("invalid_response", "Worker response identity is invalid.");
        if (!Wire.Boolean(reply, "ok"))
            throw new CoreFault(Wire.Text(reply["error"], "code"), Wire.Text(reply["error"], "message", 256));
        return reply["result"]?.DeepClone();
    }
    public static async Task EnsureAsync(CancellationToken cancellation)
    {
        if (await AvailableAsync(cancellation)) return;
        WindowsPipePeer.PreventStandardHandleInheritance();
        using (var child = Process.Start(RuntimeIdentity.Child("--broker")))
        {
            // The persistent worker must not inherit the short-lived launcher's stdout pipe.
            child?.StandardInput.Close(); child?.StandardOutput.Dispose(); child?.StandardError.Dispose();
        }
        var deadline = DateTime.UtcNow.AddSeconds(8);
        while (DateTime.UtcNow < deadline)
        {
            cancellation.ThrowIfCancellationRequested();
            if (await AvailableAsync(cancellation)) return;
            await Task.Delay(100, cancellation);
        }
        throw new CoreFault("broker_unavailable", "The private worker did not become ready.");
    }
    private static async Task<bool> AvailableAsync(CancellationToken cancellation)
    {
        try
        {
            using var shortTimeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
            shortTimeout.CancelAfter(200);
            await CallAsync(RuntimeIdentity.PipeName, "health", null, shortTimeout.Token);
            return true;
        }
        catch (Exception error) when (error is IOException or OperationCanceledException or CoreFault or TimeoutException) { return false; }
    }
}

internal static partial class WindowsPipePeer
{
    [LibraryImport("kernel32.dll")]
    private static partial IntPtr GetStdHandle(int kind);
    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetHandleInformation(IntPtr handle, uint mask, uint flags);
    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GetNamedPipeClientProcessId(IntPtr pipe, out uint processId);
    public static int ProcessId(NamedPipeServerStream pipe) => GetNamedPipeClientProcessId(pipe.SafePipeHandle.DangerousGetHandle(), out var id) ? checked((int)id) : 0;
    public static void PreventStandardHandleInheritance()
    {
        // CreateProcess can inherit unrelated inheritable handles in addition to its redirected
        // standard streams. Do not leave the launcher's parent stdout open in the broker.
        foreach (var kind in new[] { -10, -11, -12 })
        {
            var handle = GetStdHandle(kind);
            if (handle != IntPtr.Zero && handle != new IntPtr(-1)) SetHandleInformation(handle, 1, 0);
        }
    }
}

public sealed class ProcessLease : IDisposable
{
    private readonly Process process;
    private readonly Action onExit;
    private int disposed;
    public int Pid { get; }
    public long CreationTicks { get; }
    public ProcessLease(int pid, Action onExit)
    {
        Pid = pid;
        this.onExit = onExit;
        try
        {
            process = Process.GetProcessById(pid);
            CreationTicks = process.StartTime.ToUniversalTime().Ticks;
            process.Exited += Exited;
            process.EnableRaisingEvents = true;
            if (process.HasExited) throw new CoreFault("owner_unavailable", "The registration owner has exited.");
        }
        catch (CoreFault) { throw; }
        catch (Exception) { throw new CoreFault("owner_unavailable", "The registration owner is unavailable."); }
    }
    private void Exited(object? sender, EventArgs args) { if (Volatile.Read(ref disposed) == 0) onExit(); }
    public void Dispose()
    {
        if (Interlocked.Exchange(ref disposed, 1) != 0) return;
        process.Exited -= Exited;
        process.Dispose();
    }
}

internal static class DebugNdjson
{
    private const int Capacity = 64;
    private static readonly object Gate = new();
    private static readonly Queue<object> Events = new();
    private static readonly System.Net.Http.HttpClient Http = new() { Timeout = TimeSpan.FromMilliseconds(250) };
    private static readonly string[] Ingest =
    [
        "http://127.0.0.1:7327/ingest/455da3f8-8da7-4d14-9d39-307205a1b093",
        "http://127.0.0.1:7884/ingest/93344ccb-bd6f-4fa7-8e2d-f58c925f8e0b"
    ];

    internal static void Write(string hypothesisId, string location, string message, object? data)
    {
        var timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        try
        {
            var line = System.Text.Json.JsonSerializer.Serialize(new
            {
                sessionId = "4dd893",
                hypothesisId,
                location,
                message,
                data,
                timestamp
            });
            foreach (var url in Ingest)
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, url)
                {
                    Content = new System.Net.Http.StringContent(line, Encoding.UTF8, "application/json")
                };
                request.Headers.TryAddWithoutValidation("X-Debug-Session-Id", "4dd893");
                _ = Http.SendAsync(request);
            }
        }
        catch { /* Debug ingest must never block Office. */ }
        if (!AppContext.TryGetSwitch("OfficeInEditor.Diagnostics", out var enabled) || !enabled) return;
        lock (Gate)
        {
            if (Events.Count == Capacity) Events.Dequeue();
            Events.Enqueue(new { hypothesisId, location, message, timestamp });
        }
    }

    internal static object[] Snapshot() { lock (Gate) return Events.ToArray(); }
}
