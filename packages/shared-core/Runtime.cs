using Microsoft.Win32.SafeHandles;
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
    public static Task<JsonNode?> CallAsync(string pipeName, string method, JsonNode? parameters, CancellationToken cancellation) =>
        RequestAsync(pipeName, null, method, parameters, cancellation);

    public static Task<JsonNode?> CallOfficeAsync(string pipeName, int officePid, string method, JsonNode? parameters, CancellationToken cancellation) =>
        RequestAsync(pipeName, officePid, method, parameters, cancellation);

    private static async Task<JsonNode?> RequestAsync(string pipeName, int? officePid, string method, JsonNode? parameters, CancellationToken cancellation)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        timeout.CancelAfter(TimeSpan.FromSeconds(55));
        await using var pipe = officePid is null
            ? new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly)
            : new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous, TokenImpersonationLevel.Identification);
        await pipe.ConnectAsync(timeout.Token);
        // The Framework add-in ACL uses the account SID. CurrentUserOnly uses the token's
        // default owner SID on Windows, which can be the Administrators group instead.
        if (officePid is int expected && !WindowsPipePeer.IsRegisteredOfficeServer(pipe, expected))
            throw new CoreFault("office_peer_unverified", "The callback pipe does not belong to the registered Office process and account.");
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
    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GetNamedPipeServerProcessId(SafePipeHandle pipe, out uint serverProcessId);
    [LibraryImport("kernel32.dll", SetLastError = true)]
    private static partial SafeProcessHandle OpenProcess(uint access, [MarshalAs(UnmanagedType.Bool)] bool inherit, uint processId);
    [LibraryImport("advapi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool OpenProcessToken(SafeProcessHandle process, uint access, out SafeAccessTokenHandle token);

    internal static bool IsRegisteredOfficeServer(NamedPipeClientStream pipe, int expectedPid)
    {
        if (expectedPid <= 0 || !GetNamedPipeServerProcessId(pipe.SafePipeHandle, out var serverPid) || serverPid != expectedPid) return false;
        const uint queryLimitedInformation = 0x1000;
        const uint tokenQuery = 0x0008;
        using var process = OpenProcess(queryLimitedInformation, false, serverPid);
        if (process.IsInvalid || !OpenProcessToken(process, tokenQuery, out var token)) return false;
        using (token)
        using (var server = new WindowsIdentity(token.DangerousGetHandle()))
        using (var caller = WindowsIdentity.GetCurrent())
            return server.User is { } serverSid && caller.User is { } callerSid && serverSid.Equals(callerSid);
    }

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
        // Process.Dispose takes the Process object's lock, which .NET holds while that process's Exited callback runs.
        // Leases on one pid (a host and its chat) exit together: one cleanup holds the Core lock and disposes the other
        // lease while that lease's callback holds its Process lock and waits for the Core lock, so a synchronous
        // Dispose here deadlocked the whole broker. The disposed flag already stops onExit; release the handle later.
        ThreadPool.UnsafeQueueUserWorkItem(static state => ((Process)state!).Dispose(), process);
    }
}

internal static class DebugNdjson
{
    private const int Capacity = 64;
    private static readonly object Gate = new();
    private static readonly Queue<object> Events = new();

    internal static void Write(string hypothesisId, string location, string message, object? data)
    {
        if (!AppContext.TryGetSwitch("OfficeInEditor.Diagnostics", out var enabled) || !enabled) return;
        lock (Gate)
        {
            if (Events.Count == Capacity) Events.Dequeue();
            Events.Enqueue(new
            {
                hypothesisId = Stage(hypothesisId, 32),
                location = Stage(location, 96),
                message = Stage(message, 64),
                timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
            });
        }
    }

    internal static object[] Snapshot() { lock (Gate) return Events.ToArray(); }

    // Stage names are bounded identifiers, not arbitrary text. Payloads are never
    // inspected, formatted or retained; they may contain document content or paths.
    private static string Stage(string? value, int maximum)
    {
        if (string.IsNullOrEmpty(value) || value.Length > maximum) return "invalid-stage";
        foreach (var character in value)
            if (!char.IsAsciiLetterOrDigit(character) && character is not ('.' or '_' or '-' or ':'))
                return "invalid-stage";
        return value;
    }

    // Keep legacy probes on the same explicit opt-in, bounded in-memory path.
    // No event handler starts a file writer, background flush or network request.
    internal static void Probe(string hypothesisId, string location, string message, object? data)
    {
        try { Write(hypothesisId, location, message, data); }
        catch { /* Diagnostics never disrupt a native callback or Office operation. */ }
    }
}
