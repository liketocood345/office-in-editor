using System.IO.Pipes;
using System.Runtime.CompilerServices;
using System.Text.Json.Nodes;

[assembly: InternalsVisibleTo("OfficeInEditor.Core.Tests")]

namespace OfficeInEditor.Core;

public static class Program
{
    public static int Main(string[] args)
    {
        try
        {
            if (!OperatingSystem.IsWindows()) throw new CoreFault("unsupported_platform", "This private worker requires Windows.");
            NativeWindow.TryEnableDpi();
            switch (args.FirstOrDefault())
            {
                case "--endpoint":
                    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(RuntimeIdentity.Endpoint, Wire.Json));
                    return 0;
                case "--ensure-broker":
                    PipeClient.EnsureAsync(CancellationToken.None).GetAwaiter().GetResult();
                    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(RuntimeIdentity.Endpoint, Wire.Json));
                    return 0;
                case "--shutdown-if-idle":
                    return ShutdownIfIdleAsync().GetAwaiter().GetResult();
                case "--broker":
                    using (var mutex = new Mutex(false, RuntimeIdentity.MutexName))
                    {
                        bool owns;
                        try { owns = mutex.WaitOne(0); } catch (AbandonedMutexException) { owns = true; }
                        if (!owns) return 0;
                        try { RunBrokerAsync().GetAwaiter().GetResult(); }
                        finally { mutex.ReleaseMutex(); }
                    }
                    return 0;
                case "--mcp":
                    new McpServer().RunAsync().GetAwaiter().GetResult();
                    return 0;
                default:
                    Console.Error.WriteLine("Use --endpoint, --ensure-broker, --broker, --mcp, or --shutdown-if-idle.");
                    return 2;
            }
        }
        catch (Exception error)
        {
            // Diagnostics deliberately omit exception details, paths, parameters and document content.
            Console.Error.WriteLine(error is CoreFault fault ? fault.Code : "worker_failed");
            return 1;
        }
    }

    private static async Task<int> ShutdownIfIdleAsync()
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(2));
        try
        {
            var result = await PipeClient.CallAsync(RuntimeIdentity.PipeName, "runtime.shutdownIfIdle", null, timeout.Token);
            Console.WriteLine(result?.ToJsonString());
            return result?["status"]?.GetValue<string>() == "shutting_down" ? 0 : 3;
        }
        catch (Exception error) when (error is IOException or TimeoutException or OperationCanceledException)
        {
            // No existing connection can be reached. Never start a worker during uninstall.
            return 0;
        }
    }

    private static async Task RunBrokerAsync()
    {
        using var stop = new CancellationTokenSource();
        await using var state = new CoreState(RuntimeIdentity.DataRoot);
        await using var app = new AppServer(state);
        await app.StartAsync(stop.Token);
        state.AppOrigin = app.Origin;
        state.BecameIdle += () => stop.Cancel();
        state.ScheduleIdleShutdown();
        using var clients = new SemaphoreSlim(16);
        var requests = new List<Task>();
        try
        {
            while (!stop.IsCancellationRequested)
            {
                await clients.WaitAsync(stop.Token);
                var pipe = new NamedPipeServerStream(RuntimeIdentity.PipeName, PipeDirection.InOut,
                    16, PipeTransmissionMode.Byte, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly, 4096, 4096);
                try { await pipe.WaitForConnectionAsync(stop.Token); }
                catch { await pipe.DisposeAsync(); clients.Release(); throw; }
                requests.RemoveAll(task => task.IsCompleted);
                requests.Add(HandleAsync(pipe, state, stop.Token).ContinueWith(_ => clients.Release(), TaskScheduler.Default));
            }
        }
        catch (OperationCanceledException) when (stop.IsCancellationRequested) { }
        finally
        {
            stop.Cancel();
            await Task.WhenAny(Task.WhenAll(requests), Task.Delay(TimeSpan.FromSeconds(3)));
        }
    }

    private static async Task HandleAsync(NamedPipeServerStream pipe, CoreState state, CancellationToken stopped)
    {
        await using (pipe)
        {
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(stopped);
            timeout.CancelAfter(TimeSpan.FromSeconds(65));
            JsonNode? id = null;
            var method = "";
            try
            {
                var line = await Wire.ReadLineAsync(pipe, timeout.Token);
                if (line is null) return;
                var request = JsonNode.Parse(line) ?? throw new CoreFault("invalid_request", "Request is invalid.");
                id = request["id"]?.DeepClone();
                if (Wire.Number(request, "v") != 1) throw new CoreFault("protocol_mismatch", "Protocol version is unsupported.");
                method = Wire.Text(request, "method");
                var result = await state.InvokeAsync(method, request["params"], WindowsPipePeer.ProcessId(pipe), timeout.Token);
                await Wire.WriteAsync(pipe, new { id, ok = true, result }, timeout.Token);
            }
            catch (Exception error)
            {
                // Failed office.open aborts only its own site in OleDocumentHost. Do not
                // AbortOrphan sibling tabs or other Office kinds.
                try
                {
                    var fault = error as CoreFault;
                    await Wire.WriteAsync(pipe, new { id, ok = false, error = new { code = fault?.Code ?? "request_failed", message = fault?.Message ?? "The request could not be completed." } }, timeout.Token);
                }
                catch (Exception) { /* Disconnected callers cannot receive a response. Never replay their operation. */ }
            }
        }
    }
}
