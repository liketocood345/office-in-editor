using System.IO.Pipes;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

/// <summary>
/// Per-session stream pipe: one current-user connection per surface. A client proves itself with the session's
/// stream token, a surface token of this session and the same process id that opened that surface. Events are queued
/// in order; frames keep only the latest unsent one, so a slow tab sees fresh pixels instead of a backlog.
/// </summary>
internal sealed class FocusStream : IAsyncDisposable
{
    public const int MaxClients = 8;
    const int MaxQueuedEvents = 256;
    static readonly TimeSpan HelloTimeout = TimeSpan.FromSeconds(5);
    readonly string pipeName, token;
    readonly Func<string, int, bool> admit;
    readonly Func<JsonObject> snapshot;
    readonly CancellationTokenSource stop = new();
    readonly object gate = new();
    readonly List<Client> clients = [];
    readonly Task accept;

    public FocusStream(string pipeName, string token, Func<string, int, bool> admit, Func<JsonObject> snapshot)
    {
        this.pipeName = pipeName; this.token = token; this.admit = admit; this.snapshot = snapshot;
        accept = Task.Run(AcceptAsync);
    }

    public int ClientCount { get { lock (gate) return clients.Count; } }

    public void Broadcast(JsonObject value)
    {
        var bytes = FocusContract.EncodeEvent(value);
        lock (gate) foreach (var client in clients.ToArray()) if (!client.Enqueue(bytes)) Drop(client);
    }

    /// <summary>Publishes a frame to every client; returns false when the frame is too large and was dropped.</summary>
    public bool PublishFrame(FocusContract.FrameHeader header, ReadOnlySpan<byte> image)
    {
        var bytes = FocusContract.EncodeFrame(header, image);
        if (bytes is null) return false;
        lock (gate) foreach (var client in clients) client.SetFrame(bytes);
        return true;
    }

    public void Disconnect(string surfaceToken)
    {
        lock (gate) foreach (var client in clients.Where(item => item.SurfaceToken == surfaceToken).ToArray()) Drop(client);
    }

    void Drop(Client client) { clients.Remove(client); client.Close(); }

    async Task AcceptAsync()
    {
        while (!stop.IsCancellationRequested)
        {
            NamedPipeServerStream pipe;
            try
            {
                pipe = new NamedPipeServerStream(pipeName, PipeDirection.InOut, MaxClients, PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly, 4096, 1 << 20);
            }
            catch (IOException) { await Task.Delay(250, stop.Token).ConfigureAwait(false); continue; }
            try { await pipe.WaitForConnectionAsync(stop.Token).ConfigureAwait(false); }
            catch { await pipe.DisposeAsync(); if (stop.IsCancellationRequested) return; continue; }
            _ = Task.Run(() => AdmitAsync(pipe));
        }
    }

    async Task AdmitAsync(NamedPipeServerStream pipe)
    {
        Client? client = null;
        try
        {
            using var hello = CancellationTokenSource.CreateLinkedTokenSource(stop.Token);
            hello.CancelAfter(HelloTimeout);
            var line = await Wire.ReadLineAsync(pipe, hello.Token).ConfigureAwait(false);
            var request = FocusContract.ParseHello(line ?? throw new CoreFault("invalid_params", "Hello is missing."));
            var pid = WindowsPipePeer.ProcessId(pipe);
            if (!Wire.TokenEquals(request.Token, token) || !admit(request.SurfaceToken, pid)) throw new CoreFault("unauthorized", "Stream token is invalid.");
            client = new Client(pipe, request.SurfaceToken);
            lock (gate)
            {
                if (clients.Count >= MaxClients) throw new CoreFault("quota_exceeded", "Too many stream clients.");
                // One connection per surface: a reconnect replaces the old one.
                foreach (var old in clients.Where(item => item.SurfaceToken == request.SurfaceToken).ToArray()) Drop(old);
                clients.Add(client);
                client.Enqueue(FocusContract.EncodeEvent(snapshot()));
            }
            await client.PumpAsync(stop.Token).ConfigureAwait(false);
        }
        catch (Exception) { /* A bad hello or a closed tab ends only this connection. */ }
        finally
        {
            if (client is not null) lock (gate) clients.Remove(client);
            client?.Close();
            await pipe.DisposeAsync();
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (stop.IsCancellationRequested) return;
        stop.Cancel();
        lock (gate) { foreach (var client in clients) client.Close(); clients.Clear(); }
        try { await accept.ConfigureAwait(false); } catch { }
        stop.Dispose();
    }

    sealed class Client(NamedPipeServerStream pipe, string surfaceToken)
    {
        readonly Queue<byte[]> events = new();
        readonly SemaphoreSlim signal = new(0);
        byte[]? frame;
        bool closed;
        public string SurfaceToken { get; } = surfaceToken;

        public bool Enqueue(byte[] bytes)
        {
            lock (events)
            {
                if (closed) return true;
                if (events.Count >= MaxQueuedEvents) return false;
                events.Enqueue(bytes);
            }
            signal.Release();
            return true;
        }

        public void SetFrame(byte[] bytes)
        {
            bool wasEmpty;
            lock (events) { if (closed) return; wasEmpty = frame is null; frame = bytes; }
            if (wasEmpty) signal.Release();
        }

        public void Close()
        {
            lock (events) { if (closed) return; closed = true; }
            signal.Release();
            try { pipe.Dispose(); } catch { }
        }

        public async Task PumpAsync(CancellationToken stopped)
        {
            while (true)
            {
                await signal.WaitAsync(stopped).ConfigureAwait(false);
                while (true)
                {
                    byte[]? next;
                    lock (events)
                    {
                        if (closed) return;
                        // Events first so a state change is never overtaken by a frame of the new park.
                        if (events.Count > 0) next = events.Dequeue();
                        else { next = frame; frame = null; }
                    }
                    if (next is null) break;
                    await pipe.WriteAsync(next, stopped).ConfigureAwait(false);
                }
                await pipe.FlushAsync(stopped).ConfigureAwait(false);
            }
        }
    }
}
