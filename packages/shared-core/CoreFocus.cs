using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using Microsoft.Win32.SafeHandles;

namespace OfficeInEditor.Core;

/// <summary>An Office document window found by the add-in. <c>Hwnd</c> is a top-level frame of <c>Pid</c>.</summary>
internal readonly record struct FocusOfficeWindow(int Pid, long Hwnd, string DocumentId, string PipeName);

/// <summary>Calls into the Office add-in for the document behind a focus session.</summary>
internal interface IFocusOfficeLink
{
    Task<FocusOfficeWindow> ResolveAsync(string application, string path, CancellationToken cancellation);
    /// <summary>True or false when the add-in answered, null when it could not be reached.</summary>
    Task<bool?> DirtyAsync(FocusOfficeWindow window, CancellationToken cancellation);
    Task CloseDocumentAsync(FocusOfficeWindow window, bool save, CancellationToken cancellation);
}

internal sealed class FocusSurface(string token, string hostId, int pid, int editorPid, FocusContract.SurfaceSize size)
{
    public string Token { get; } = token;
    public string HostId { get; } = hostId;
    public int Pid { get; } = pid;
    public int EditorPid { get; } = editorPid;
    public FocusContract.SurfaceSize Size { get; set; } = size;
}

/// <summary>One Office document shown in one or more tabs. Collections are guarded by CoreState's lock; machine and ops by the dispatcher.</summary>
internal sealed class FocusSession(string id, string application, string fileKey, FocusOfficeWindow window, string streamToken) : IFocusTicker
{
    public string Id { get; } = id;
    public string Application { get; } = application;
    public string FileKey { get; } = fileKey;
    public FocusOfficeWindow Window { get; } = window;
    public string StreamToken { get; } = streamToken;
    public Dictionary<string, FocusSurface> Surfaces { get; } = new(StringComparer.Ordinal);
    public string ActiveSurface { get; set; } = "";
    public FocusStream? Stream { get; set; }
    public FocusStateMachine? Machine { get; set; }
    public Win32FocusWindowOps? Ops { get; set; }
    public IFocusCapture Capture { get; set; } = NoFocusCapture.Instance;
    public bool Ending { get; set; }
    volatile JsonObject snapshot = FocusContract.StateEvent("releasing", "open", 0, 0);
    volatile string pixelPolicy = "always";

    public string State => snapshot["state"]!.GetValue<string>();
    public int ActivationId => snapshot["activationId"]!.GetValue<int>();
    public int ParkId => snapshot["parkId"]!.GetValue<int>();
    public string PixelPolicy { get => pixelPolicy; set => pixelPolicy = value; }
    public JsonObject Snapshot() => (JsonObject)snapshot.DeepClone();

    public void Emit(JsonObject value)
    {
        if (value["type"]?.GetValue<string>() == "state") snapshot = (JsonObject)value.DeepClone();
        Stream?.Broadcast(value);
    }

    long frameSeq;

    /// <summary>Forwards a capture frame when it belongs to the current park; frames of an older park are dropped.</summary>
    public bool PublishFrame(int parkId, int width, int height, string format, byte[] image)
    {
        if (Stream is null || State != "parked" || parkId != ParkId || width <= 0 || height <= 0) return false;
        return Stream.PublishFrame(new FocusContract.FrameHeader(Interlocked.Increment(ref frameSeq), parkId, width, height, format), image);
    }

    public void PublishFrameAction(int parkId, int width, int height, string format, byte[] image) => PublishFrame(parkId, width, height, format, image);

    public void Tick(long now) => Machine?.Tick(now);
    public void ForegroundChanged(long now) => Machine?.ForegroundChanged(now);
}

public sealed partial class CoreState
{
    private const int MaxFocusSessions = 16;
    private readonly Dictionary<string, FocusSession> focusSessions = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> focusByFile = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Task<FocusSession>> focusOpening = new(StringComparer.Ordinal);
    private FocusDispatcher? focusDispatcher;
    private FocusRecovery? focusRecovery;

    /// <summary>
    /// Advertised as <c>health.capabilities.focusHybrid</c>: on when the capture helper ships next to Core. The lab
    /// can force it with OIE_FOCUS_HYBRID=1 to exercise window control without frames.
    /// </summary>
    internal static bool FocusHybridEnabled { get; set; } =
        Environment.GetEnvironmentVariable("OIE_FOCUS_HYBRID") == "1" || FocusCaptureHost.Locate() is not null;
    internal Func<FocusSession, IFocusCapture> FocusCaptureFactory { get; set; } = session => new FocusCaptureHost(
        session.Window.Hwnd, session.Window.Pid, session.PublishFrameAction, code => session.Emit(FocusContract.NoticeEvent(code)));
    internal IFocusOfficeLink? FocusLinkOverride { get; set; }

    private FocusRecovery Recovery => focusRecovery ??= new FocusRecovery(Path.GetDirectoryName(referenceDirectory)!);
    private IFocusOfficeLink FocusLink => FocusLinkOverride ?? new AddinFocusLink(this);
    private FocusDispatcher Dispatcher { get { lock (sync) return focusDispatcher ??= new FocusDispatcher(); } }

    private void RestoreFocusOrphans()
    {
        try { Recovery.RestoreOrphans(); } catch (Exception) { /* Startup must not fail on recovery. */ }
    }

    private async Task<JsonNode?> InvokeFocusAsync(string method, JsonNode? parameters, int pid, CancellationToken cancellation)
    {
        switch (method)
        {
            case FocusContract.Open: return await FocusOpenAsync(parameters, pid, cancellation);
            case FocusContract.Activate:
            {
                var request = FocusContract.ParseActivate(parameters);
                FocusSession session; FocusSurface surface;
                lock (sync) (session, surface) = AuthorizeFocusSurface(request.SessionId, request.SurfaceToken, parameters, pid);
                var state = await Dispatcher.InvokeAsync(() =>
                {
                    UseSurface(session, surface);
                    return session.Machine!.Activate(request.ActivationId, request.PointX, request.PointY, Environment.TickCount64);
                });
                return Wire.Node(new { state, activationId = session.ActivationId });
            }
            case FocusContract.Layout:
            {
                var request = FocusContract.ParseLayout(parameters);
                FocusSession session; FocusSurface surface;
                lock (sync)
                {
                    (session, surface) = AuthorizeFocusSurface(request.SessionId, request.SurfaceToken, parameters, pid);
                    surface.Size = request.Size;
                }
                await Dispatcher.InvokeAsync(() =>
                {
                    // Only the surface Office docks into drives the window; other tabs on the file just draw frames.
                    if (session.ActiveSurface == surface.Token) { session.Ops!.UpdateSize(surface.Size); session.Machine!.Layout(request.Visible, Environment.TickCount64); }
                    else if (request.Visible && session.State == "parked") { UseSurface(session, surface); session.Machine!.Layout(true, Environment.TickCount64); }
                });
                return Wire.Node(new { state = session.State });
            }
            case FocusContract.Policy:
            {
                var request = FocusContract.ParsePolicy(parameters);
                FocusSession session;
                lock (sync) session = AuthorizeFocusSession(request.SessionId, parameters, pid);
                await Dispatcher.InvokeAsync(() => session.Machine!.SetPolicy(request.PixelPolicy));
                session.PixelPolicy = request.PixelPolicy;
                return Wire.Node(new { pixelPolicy = request.PixelPolicy });
            }
            case FocusContract.Wheel:
            {
                var request = FocusContract.ParseWheel(parameters);
                FocusSession session; FocusSurface surface;
                lock (sync) (session, surface) = AuthorizeFocusSurface(request.SessionId, request.SurfaceToken, parameters, pid);
                // Only the tab Office is parked for maps 1:1 onto the parked window.
                var forwarded = await Dispatcher.InvokeAsync(() => session.ActiveSurface == surface.Token &&
                    session.Machine!.Wheel(request.PointX, request.PointY, request.DeltaY, request.DeltaX, request.Ctrl, request.Shift));
                return Wire.Node(new { forwarded });
            }
            case FocusContract.Close: return await FocusCloseAsync(parameters, pid, cancellation);
            case FocusContract.Status:
            {
                var id = FocusContract.ParseStatus(parameters);
                FocusSession session; int surfaces;
                lock (sync) { session = AuthorizeFocusSession(id, parameters, pid); surfaces = session.Surfaces.Count; }
                bool? dirty = session.State == "lost" ? null : await TryDirtyAsync(session, cancellation);
                return Wire.Node(new { state = session.State, application = session.Application, surfaces, dirty,
                    activationId = session.ActivationId, parkId = session.ParkId, pixelPolicy = session.PixelPolicy });
            }
            default: throw new CoreFault("method_not_found", "This operation is not supported.");
        }
    }

    private async Task<JsonNode?> FocusOpenAsync(JsonNode? parameters, int pid, CancellationToken cancellation)
    {
        var request = FocusContract.ParseOpen(parameters);
        string hostId;
        lock (sync) hostId = AuthorizeHost(parameters, pid).Id;
        var (fullPath, application) = FocusDocument(request.Path);
        var fileKey = FileIdentity(fullPath);
        Task<FocusSession> opening;
        bool creator;
        lock (sync)
        {
            if (focusSessions.Values.Any(item => item.Surfaces.ContainsKey(request.SurfaceToken)))
                throw new CoreFault("identity_conflict", "The surface already belongs to a focus session.");
            if (focusByFile.TryGetValue(fileKey, out var existingId) && focusSessions.TryGetValue(existingId, out var existing) && existing.State != "lost" && !existing.Ending)
                return JoinFocusSession(existing, request, hostId, pid, shared: true);
            creator = !focusOpening.TryGetValue(fileKey, out var pending);
            opening = pending!;
            if (creator)
            {
                if (focusSessions.Count + focusOpening.Count >= MaxFocusSessions) throw new CoreFault("quota_exceeded", "The focus session budget is full.");
                opening = StartFocusSessionAsync(fullPath, application, fileKey, request, cancellation);
                focusOpening[fileKey] = opening;
            }
        }
        FocusSession session;
        try { session = await opening; }
        finally { if (creator) lock (sync) focusOpening.Remove(fileKey); }
        lock (sync)
        {
            if (creator)
            {
                if (disposed) throw new CoreFault("focus_unavailable", "Core is stopping.");
                focusSessions[session.Id] = session;
                focusByFile[fileKey] = session.Id;
                session.ActiveSurface = request.SurfaceToken;
            }
            return JoinFocusSession(session, request, hostId, pid, shared: !creator);
        }
    }

    private JsonNode? JoinFocusSession(FocusSession session, FocusContract.OpenRequest request, string hostId, int pid, bool shared)
    {
        if (session.Surfaces.Count >= FocusStream.MaxClients) throw new CoreFault("quota_exceeded", "Too many tabs show this document.");
        session.Surfaces[request.SurfaceToken] = new FocusSurface(request.SurfaceToken, hostId, pid, request.EditorPid, request.Size);
        return Wire.Node(new
        {
            sessionId = session.Id, application = session.Application, shared, state = session.State,
            stream = new { pipeName = FocusContract.StreamPipeName(RuntimeIdentity.PipeName, session.Id), token = session.StreamToken },
        });
    }

    private async Task<FocusSession> StartFocusSessionAsync(string fullPath, string application, string fileKey, FocusContract.OpenRequest request, CancellationToken cancellation)
    {
        await Task.Yield();
        var window = await FocusLink.ResolveAsync(application, fullPath, cancellation);
        lock (sync)
            if (focusSessions.Values.Any(item => item.Window.Hwnd == window.Hwnd))
                throw new CoreFault("identity_conflict", "The Office window already belongs to another focus session.");
        var session = new FocusSession(FocusContract.NewSessionId(), application, fileKey, window, Wire.NewToken());
        session.PixelPolicy = request.PixelPolicy;
        session.Stream = new FocusStream(FocusContract.StreamPipeName(RuntimeIdentity.PipeName, session.Id), session.StreamToken,
            (surface, peer) => { lock (sync) return session.Surfaces.TryGetValue(surface, out var owner) && owner.Pid == peer; },
            session.Snapshot);
        try
        {
            await Dispatcher.InvokeAsync(() =>
            {
                var hwnd = new IntPtr(window.Hwnd);
                if (!Win32FocusWindowOps.IsOfficeFrame(hwnd, window.Pid, application))
                    throw new CoreFault("office_window_unavailable", "The Office document window could not be verified.");
                session.Ops = new Win32FocusWindowOps(hwnd, (uint)window.Pid, request.SurfaceToken, request.Size, Recovery);
                session.Capture = FocusCaptureFactory(session);
                session.Machine = new FocusStateMachine(session.Ops, session.Capture, session.Emit, request.PixelPolicy, Environment.TickCount64);
            });
        }
        catch
        {
            // The frame may have been moved out of view at cold start; never leave it hidden when the session fails.
            Win32FocusWindowOps.ReleasePrePark(new IntPtr(window.Hwnd), Recovery);
            await session.Stream.DisposeAsync();
            throw;
        }
        Dispatcher.Add(session);
        return session;
    }

    private async Task<JsonNode?> FocusCloseAsync(JsonNode? parameters, int pid, CancellationToken cancellation)
    {
        var request = FocusContract.ParseClose(parameters);
        FocusSession session;
        lock (sync)
        {
            (session, var surface) = AuthorizeFocusSurface(request.SessionId, request.SurfaceToken, parameters, pid);
            if (session.Surfaces.Count > 1)
            {
                session.Surfaces.Remove(surface.Token);
                session.Stream?.Disconnect(surface.Token);
                if (session.ActiveSurface == surface.Token)
                {
                    var next = session.Surfaces.Values.First();
                    _ = Dispatcher.InvokeAsync(() => { session.Machine!.Layout(false, Environment.TickCount64); UseSurface(session, next); });
                }
                return Wire.Node(new { closed = true, dirty = false, remainingSurfaces = session.Surfaces.Count });
            }
        }
        if (session.State == "lost")
        {
            await EndFocusSessionAsync(session, "closed", restore: false);
            return Wire.Node(new { closed = true, dirty = false, remainingSurfaces = 0 });
        }
        var dirty = await TryDirtyAsync(session, cancellation);
        if (request.Decision is null && dirty == true)
            return Wire.Node(new { closed = false, dirty = true, remainingSurfaces = 1 });
        // Put Office back on screen first, so anything it shows while closing is visible.
        await Dispatcher.InvokeAsync(() => session.Machine!.Restore());
        if (dirty is null && request.Decision is null)
        {
            // The add-in cannot be reached: leave Office open where it was rather than risk unsaved work.
            await EndFocusSessionAsync(session, "closed", restore: false);
            return Wire.Node(new { closed = true, dirty = false, remainingSurfaces = 0 });
        }
        try { await FocusLink.CloseDocumentAsync(session.Window, request.Decision == "save", cancellation); }
        catch (CoreFault) { throw new CoreFault("focus_close_failed", "Office did not close the document. It is still open."); }
        await EndFocusSessionAsync(session, "closed", restore: false);
        return Wire.Node(new { closed = true, dirty = false, remainingSurfaces = 0 });
    }

    private async Task<bool?> TryDirtyAsync(FocusSession session, CancellationToken cancellation)
    {
        try { return await FocusLink.DirtyAsync(session.Window, cancellation); }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { throw; }
        catch (Exception) { return null; }
    }

    /// <summary>Ends a session. Office is never closed here; <paramref name="restore"/> puts its window back on screen.</summary>
    private async Task EndFocusSessionAsync(FocusSession session, string reason, bool restore)
    {
        lock (sync)
        {
            if (session.Ending) return;
            session.Ending = true;
            focusSessions.Remove(session.Id);
            if (focusByFile.TryGetValue(session.FileKey, out var id) && id == session.Id) focusByFile.Remove(session.FileKey);
        }
        CoreTrace.Mark("focus-end", "begin");
        var dispatcher = Dispatcher;
        dispatcher.Remove(session);
        try
        {
            await dispatcher.InvokeAsync(() =>
            {
                if (restore) session.Machine?.Restore();
                else session.Capture.Stop();
                session.Ops?.ForgetRecovery();
            });
        }
        catch (Exception) { /* The dispatcher is stopping; recovery restores the window on the next start. */ }
        CoreTrace.Mark("focus-end", "restored");
        (session.Capture as IDisposable)?.Dispose();
        CoreTrace.Mark("focus-end", "capture-disposed");
        session.Stream?.Broadcast(FocusContract.ClosedEvent(reason));
        await Task.Delay(50);
        if (session.Stream is not null) await session.Stream.DisposeAsync();
        CoreTrace.Mark("focus-end", "stream-disposed");
        ScheduleIdleShutdown();
    }

    /// <summary>A host went away: drop its surfaces and end sessions nobody shows any more (Office stays open).</summary>
    private void DetachFocusHost(string hostId)
    {
        foreach (var session in focusSessions.Values.ToArray())
        {
            foreach (var surface in session.Surfaces.Values.Where(item => item.HostId == hostId).ToArray())
            {
                session.Surfaces.Remove(surface.Token);
                session.Stream?.Disconnect(surface.Token);
            }
            if (session.Surfaces.Count == 0) _ = EndFocusSessionAsync(session, "closed", restore: true);
            else if (!session.Surfaces.ContainsKey(session.ActiveSurface))
            {
                var next = session.Surfaces.Values.First();
                _ = Dispatcher.InvokeAsync(() => UseSurface(session, next));
            }
        }
    }

    private void DisposeFocus()
    {
        CoreTrace.Mark("focus-dispose", "begin");
        var sessions = focusSessions.Values.ToArray();
        focusSessions.Clear(); focusByFile.Clear();
        var dispatcher = focusDispatcher;
        if (dispatcher is null) return;
        try
        {
            dispatcher.InvokeAsync(() =>
            {
                foreach (var session in sessions) { session.Machine?.Restore(); session.Ops?.ForgetRecovery(); }
            }).Wait(TimeSpan.FromSeconds(2));
        }
        catch (Exception) { /* Recovery restores the windows on the next start. */ }
        foreach (var session in sessions)
        {
            (session.Capture as IDisposable)?.Dispose();
            _ = session.Stream?.DisposeAsync().AsTask();
        }
        CoreTrace.Mark("focus-dispose", "dispatcher-stopping");
        dispatcher.Dispose();
        CoreTrace.Mark("focus-dispose", "end");
    }

    private static void UseSurface(FocusSession session, FocusSurface surface)
    {
        session.ActiveSurface = surface.Token;
        session.Ops!.UseSurface(surface.Token, surface.Size);
    }

    private (FocusSession, FocusSurface) AuthorizeFocusSurface(string sessionId, string surfaceToken, JsonNode? parameters, int pid)
    {
        var host = AuthorizeHost(parameters, pid);
        if (!focusSessions.TryGetValue(sessionId, out var session) || session.Ending) throw new CoreFault("focus_session_not_found", "The focus session is unavailable.");
        if (!session.Surfaces.TryGetValue(surfaceToken, out var surface) || surface.HostId != host.Id || surface.Pid != pid)
            throw new CoreFault("unauthorized", "The surface does not belong to this host.");
        return (session, surface);
    }

    private FocusSession AuthorizeFocusSession(string sessionId, JsonNode? parameters, int pid)
    {
        var host = AuthorizeHost(parameters, pid);
        if (!focusSessions.TryGetValue(sessionId, out var session) || session.Ending) throw new CoreFault("focus_session_not_found", "The focus session is unavailable.");
        if (!session.Surfaces.Values.Any(item => item.HostId == host.Id && item.Pid == pid))
            throw new CoreFault("unauthorized", "The focus session does not belong to this host.");
        return session;
    }

    private static (string FullPath, string Application) FocusDocument(string path)
    {
        if (!Path.IsPathFullyQualified(path) || path.StartsWith("\\\\", StringComparison.Ordinal) ||
            Uri.TryCreate(path, UriKind.Absolute, out var uri) && uri.IsUnc)
            throw new CoreFault("invalid_path", "Choose an existing local Office document.");
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new CoreFault("file_not_found", "The selected document is unavailable.");
        if (Path.GetFileName(fullPath).StartsWith('~')) throw new CoreFault("office_cache_file", "This is an Office cache file and should not be opened.");
        if (!OfficeDocumentKind.TryClassify(fullPath, out var kind) || !ApplicationIsValid(kind.Application))
            throw new CoreFault("unsupported_format", "The focus route supports Word, Excel and PowerPoint documents.");
        return (fullPath, kind.Application);
    }

    /// <summary>Volume serial + file index, so links and different spellings of one file share a session. Falls back to the normalized path.</summary>
    internal static string FileIdentity(string fullPath)
    {
        try
        {
            using var handle = File.OpenHandle(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            if (GetFileInformationByHandle(handle, out var info))
                return $"id:{info.VolumeSerialNumber:x8}:{info.FileIndexHigh:x8}{info.FileIndexLow:x8}";
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { }
        return "path:" + Hash(fullPath.ToUpperInvariant());
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FileInformation
    {
        public uint Attributes;
        public long Created, Accessed, Written;
        public uint VolumeSerialNumber, SizeHigh, SizeLow, Links, FileIndexHigh, FileIndexLow;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle file, out FileInformation information);

    private Office[] OfficesFor(string application)
    {
        lock (sync) return offices.Values.Where(item => item.Application == application).ToArray();
    }

    /// <summary>Asks the registered Office add-ins for the document window, starting Office once when no add-in has the file.</summary>
    private sealed class AddinFocusLink(CoreState state) : IFocusOfficeLink
    {
        /// <summary>
        /// Overall budget for finding the document window, Office cold start included. It keeps focus.open well under
        /// 30 s; a warm Office answers in about a second, a cold one in 5–10 s.
        /// </summary>
        internal static readonly TimeSpan ResolveDeadline = TimeSpan.FromSeconds(25);
        static readonly TimeSpan CallTimeout = TimeSpan.FromSeconds(3);

        public async Task<FocusOfficeWindow> ResolveAsync(string application, string path, CancellationToken cancellation)
        {
            using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
            deadline.CancelAfter(ResolveDeadline);
            using var watching = CancellationTokenSource.CreateLinkedTokenSource(deadline.Token);
            Task<IntPtr>? watcher = null;
            var launched = false;
            try
            {
                while (true)
                {
                    foreach (var office in state.OfficesFor(application))
                    {
                        var found = await TryResolveAsync(office, path, deadline.Token);
                        if (found is null) continue;
                        // Keep a pre-parked frame only when it is the document window the session takes.
                        var parkedFrame = await StopWatcherAsync(watcher, watching);
                        if (parkedFrame != IntPtr.Zero && parkedFrame.ToInt64() != found.Value.Hwnd) Win32FocusWindowOps.ReleasePrePark(parkedFrame, state.Recovery);
                        return found.Value;
                    }
                    if (!launched)
                    {
                        launched = true;
                        var opened = await Task.Run(() => NativeOfficeLauncher.Open(path, "external"), deadline.Token);
                        var launchedPid = Wire.Number(opened, "pid");
                        if (launchedPid > 0) watcher = Task.Run(() => WatchNewFrameAsync(launchedPid, application, watching.Token));
                    }
                    await Task.Delay(250, deadline.Token);
                }
            }
            catch (Exception error)
            {
                // Never leave a document hidden off screen when the session could not take it.
                var parkedFrame = await StopWatcherAsync(watcher, watching);
                if (parkedFrame != IntPtr.Zero) Win32FocusWindowOps.ReleasePrePark(parkedFrame, state.Recovery);
                // Any cancellation from our own deadline (mid-call, mid-launch or mid-wait) is the same outcome.
                if (error is OperationCanceledException && !cancellation.IsCancellationRequested)
                    throw new CoreFault("office_window_unavailable", "Office did not report the document window. Check that the Office add-in is enabled.");
                throw;
            }
        }

        /// <summary>
        /// A cold-started Office shows its frame on screen before the add-in reports; this moves it out of view the
        /// moment it appears. It polls on its own, so slow add-in calls during Office startup do not delay it.
        /// </summary>
        async Task<IntPtr> WatchNewFrameAsync(int pid, string application, CancellationToken cancellation)
        {
            while (!cancellation.IsCancellationRequested)
            {
                var frame = Win32FocusWindowOps.PreParkNewFrame(pid, application, state.Recovery);
                if (frame != IntPtr.Zero) return frame;
                try { await Task.Delay(15, cancellation); } catch (OperationCanceledException) { break; }
            }
            return IntPtr.Zero;
        }

        static async Task<IntPtr> StopWatcherAsync(Task<IntPtr>? watcher, CancellationTokenSource watching)
        {
            if (watcher is null) return IntPtr.Zero;
            watching.Cancel();
            try { return await watcher; } catch (Exception) { return IntPtr.Zero; }
        }
        /// <summary>Null when this add-in does not have the document (yet). Throws when the add-in can never answer.</summary>
        static async Task<FocusOfficeWindow?> TryResolveAsync(Office office, string path, CancellationToken cancellation)
        {
            using var call = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
            call.CancelAfter(CallTimeout);
            JsonNode? result;
            try { result = await PipeClient.CallOfficeAsync(office.PipeName, office.Pid, "resolveDocumentWindow", Wire.Node(new { path }), call.Token); }
            catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { throw; }
            // An add-in that does not know the method is an older build: waiting will not help, so fail now.
            catch (CoreFault fault) when (fault.Code == "METHOD_NOT_FOUND")
            {
                throw new CoreFault("office_addin_outdated", "The Office add-in is an older version. Restart Office so it loads the installed version.");
            }
            catch (Exception) { return null; }
            if (result?["hwnd"] is not JsonValue handle || !handle.TryGetValue<long>(out var hwnd) || hwnd == 0) return null;
            var documentId = Wire.Text(result, "documentId", 160, false);
            return documentId.Length == 0 ? null : new FocusOfficeWindow(office.Pid, hwnd, documentId, office.PipeName);
        }

        public async Task<bool?> DirtyAsync(FocusOfficeWindow window, CancellationToken cancellation)
        {
            using var call = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
            call.CancelAfter(CallTimeout);
            var result = await PipeClient.CallOfficeAsync(window.PipeName, window.Pid, "documentState", Wire.Node(new { documentId = window.DocumentId }), call.Token);
            return Wire.OptionalBoolean(result, "dirty");
        }

        public async Task CloseDocumentAsync(FocusOfficeWindow window, bool save, CancellationToken cancellation)
        {
            using var call = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
            call.CancelAfter(TimeSpan.FromSeconds(30));
            JsonNode? result;
            try { result = await PipeClient.CallOfficeAsync(window.PipeName, window.Pid, "closeDocument", Wire.Node(new { documentId = window.DocumentId, save }), call.Token); }
            catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { throw; }
            catch (Exception) { throw new CoreFault("office_close_failed", "Office did not close the document."); }
            if (!Wire.Boolean(result, "closed")) throw new CoreFault("office_close_failed", "Office did not close the document.");
        }
    }
}
