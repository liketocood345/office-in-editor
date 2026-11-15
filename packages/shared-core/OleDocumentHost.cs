using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
using System.Text;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace OfficeInEditor.Core;

internal static class OleDocumentHost
{
    private const int OleIverbShow = -1;
    private const int OleIverbPrimary = 0;
    private const int OleIverbInPlaceActivate = -5;
    private const int OleCloseNosave = 1;
    private const int OleClosePromptSave = 2;
    private const uint RpcECantCalloutInInputSyncCall = 0x8001010D;
    private const int OleRenderNone = 0;
    private const int StgmCreate = 0x00001000;
    private const int StgmReadWrite = 0x00000002;
    private const int StgmShareExclusive = 0x00000010;
    private const uint WsPopup = 0x80000000;
    private const uint WsClipChildren = 0x02000000;
    private const uint WsVisible = 0x10000000;
    private const uint WsExToolwindow = 0x00000080;
    private const uint WsExNoActivate = 0x08000000;
    private const int SwShow = 5;
    private const int SwHide = 0;
    private const int GwlExStyle = -20;
    private const int SOk = 0;
    private const int ENotimpl = unchecked((int)0x80004001);
    private const int ENointerface = unchecked((int)0x80004002);
    private const int OleWhichMkObjFull = 3;
    private const uint OleCmdIdSave = 3;
    private const uint OleCmdFSupported = 1;
    private const uint OleCmdFEnabled = 2;
    private const uint WmKeyDown = 0x0100;
    private const int VkS = 0x53;
    private const int VkControl = 0x11;
    private static readonly IntPtr HwndMessage = new(-3);
    private static readonly object Gate = new();
    private static readonly ConcurrentDictionary<string, object> OleExclusiveByKind = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, KindHost> kinds = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, KindHost> sessionHosts = new(StringComparer.Ordinal);
    private static readonly ConcurrentDictionary<string, ConcurrentDictionary<string, byte>> sourceOwners = new(StringComparer.OrdinalIgnoreCase);
    private static string lastFailure = "";
    private sealed record HostTrace(string Stage, long ElapsedMs);
    private static readonly Queue<HostTrace> trace = new();
    private static long traceStart;
    internal static JsonNode Diagnostics { get { lock (Gate) return Wire.Node(new { lastFailure, trace = trace.ToArray() })!; } }
    private static void Trace(string stage)
    {
        lock (Gate)
        {
            if (trace.Count >= 32) trace.Dequeue();
            trace.Enqueue(new HostTrace(stage, (long)Stopwatch.GetElapsedTime(traceStart).TotalMilliseconds));
        }
    }
    private static readonly ConcurrentDictionary<string, CancellationTokenSource> opening = new();
    private static int openingCount;
    internal static string LastFailure { get { lock (Gate) return lastFailure; } }
    private const uint SwpNoActivate = 0x0010;
    private const uint SwpShowWindow = 0x0040;
    private const uint SwpHideWindow = 0x0080;
    private const uint WmSize = 0x0005;
    private const uint WmActivate = 0x0006;

    internal static bool HasSession
    {
        get
        {
            if (Volatile.Read(ref openingCount) > 0) return true;
            foreach (var host in kinds.Values) if (host.Busy) return true;
            return false;
        }
    }

    internal static bool IsServerPid(int pid)
    {
        if (pid <= 0) return false;
        foreach (var host in kinds.Values)
        {
            lock (host.Gate)
                foreach (var session in host.Sessions.Values)
                    if (session.ServerPid == pid) return true;
        }
        return false;
    }
    internal static JsonNode SourceForProcess(int pid)
    {
        if (SourceBindingsForProcess(pid)["bindings"] is JsonArray bindings && bindings.Count == 1)
            return new JsonObject { ["bound"] = true, ["sessionId"] = bindings[0]!["sessionId"]!.DeepClone(),
                ["sourcePath"] = bindings[0]!["sourcePath"]!.DeepClone(), ["hwnd"] = bindings[0]!["hwnd"]!.DeepClone() };
        return new JsonObject { ["bound"] = false };
    }

    internal static bool IsSessionForProcess(string sessionId, int pid)
    {
        if (pid <= 0 || !sessionHosts.TryGetValue(sessionId, out var host)) return false;
        lock (host.Gate) return host.Sessions.TryGetValue(sessionId, out var session) && session.Activated && session.ServerPid == pid;
    }

    internal static JsonNode SourceBindingsForProcess(int pid)
    {
        var bindings = new List<object>();
        foreach (var host in kinds.Values)
            lock (host.Gate)
                foreach (var session in host.Sessions.Values)
                    if (session.Activated && session.ServerPid == pid && session.Site.NativeObjectModelHandle != IntPtr.Zero)
                        bindings.Add(new { sessionId = session.Id, hwnd = session.Site.NativeObjectModelHandle.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture),
                            sourcePath = session.Path, sourceBound = session.Site.SourceBound, sourceVerified = session.Site.SourceIdentityVerified,
                            periodicSave = session.Site.PeriodicSave, saveIntervalSeconds = session.Site.SaveIntervalSeconds,
                            saveOnClose = session.Site.Shared?.SaveOnClose ?? session.Site.SaveOnClose });
        return Wire.Node(new { bindings })!;
    }

    internal static bool TryMap(string path, out string application, out string progId)
    {
        application = ""; progId = "";
        if (!OfficeDocumentKind.TryClassify(path, out var kind)) return false;
        application = kind.Application;
        progId = kind.ProgId;
        return true;
    }

    internal static JsonNode? TryOpen(string path, string application, int width = 0, int height = 0, string? sessionId = null, string? surfaceToken = null, CancellationToken cancellation = default, bool bindSource = false, bool visible = true, int jsX = 0, int jsY = 0, double pixelRatio = 1, string? backupId = null)
    {
        sessionId ??= "ole-" + Guid.NewGuid().ToString("N");
        if (!sessionId.StartsWith("ole-", StringComparison.Ordinal) || !Guid.TryParseExact(sessionId[4..], "N", out _))
            throw new CoreFault("invalid_session", "Invalid Office session identity.");
        if (!OfficeDocumentKind.TryClassify(path, out var kind))
            throw new CoreFault("unsupported_format", "This file is not an embeddable Office document.");
        var host = kinds.GetOrAdd(kind.Key, key => new KindHost(key));
        if (sessionHosts.Count >= 16) throw new CoreFault("quota_exceeded", "The native document session budget is full.");
        var sourceKey = Path.GetFullPath(path);
        var existing = CountSource(sourceKey);
        ClaimSource(sourceKey, sessionId);
        Interlocked.Increment(ref openingCount);
        Interlocked.Increment(ref host.Opening);
        using var stop = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        opening[sessionId] = stop;
        sessionHosts[sessionId] = host;
        lock (Gate) { trace.Clear(); traceStart = Stopwatch.GetTimestamp(); }
        Trace("sta_queued");
        // #region agent log
        DebugNdjson.Write("B", "OleDocumentHost.TryOpen", "kind_sta", new
        {
            kindKey = kind.Key, application = kind.Application, requested = application,
            sessionPrefix = sessionId.Length > 12 ? sessionId[4..12] : sessionId,
            openingForKind = host.Opening, sessionsForKind = Volatile.Read(ref host.SessionCount),
            kindCount = kinds.Count, existing, shareExisting = existing > 0
        });
        // #endregion
        try
        {
            return host.Sta.Invoke(() =>
            {
                var primary = FindLiveSession(sourceKey, sessionId);
                if (primary is not null)
                    return AttachViewOnSta(primary, kind, host, width, height, sessionId, surfaceToken ?? "", stop.Token, visible, jsX, jsY, pixelRatio);
                return OpenOnSta(path, kind, host, width, height, sessionId, surfaceToken ?? "", stop.Token, bindSource, visible, jsX, jsY, pixelRatio, backupId);
            });
        }
        catch (CoreFault) { throw; }
        catch (OperationCanceledException) { throw; }
        catch (Exception error)
        {
            if (string.IsNullOrEmpty(LastFailure)) RecordFailure("invoke", error);
            return null;
        }
        finally
        {
            opening.TryRemove(sessionId, out _);
            Interlocked.Decrement(ref openingCount);
            Interlocked.Decrement(ref host.Opening);
            lock (host.Gate)
                if (!host.Sessions.ContainsKey(sessionId))
                {
                    sessionHosts.TryRemove(sessionId, out _);
                    ReleaseSource(sourceKey, sessionId);
                }
        }
    }

    internal static bool TryHandleWindow(string method, JsonNode? value, out JsonNode? result)
    {
        result = null;
        var requestedId = Wire.Text(value, "officeId", 160, false);
        if (method == "office.detach_window" && opening.TryGetValue(requestedId, out var pending))
        {
            try { pending.Cancel(); } catch (ObjectDisposedException) { }
            result = Wire.Node(new { attached = false, cancellationRequested = true });
            if (!sessionHosts.TryGetValue(requestedId, out _)) return true;
        }
        if (string.IsNullOrEmpty(requestedId) || !sessionHosts.TryGetValue(requestedId, out var host))
        {
            if (requestedId.StartsWith("ole-", StringComparison.Ordinal)) throw new CoreFault("office_session_expired", "Office session is unavailable.");
            return false;
        }
        Session? session;
        lock (host.Gate) host.Sessions.TryGetValue(requestedId, out session);
        if (session is null)
        {
            if (method == "office.detach_window")
            {
                result = host.Sta.Invoke(() =>
                {
                    Session? created;
                    lock (host.Gate) host.Sessions.TryGetValue(requestedId, out created);
                    if (created is not null) CloseOnSta(created, host, discard: true);
                    return Wire.Node(new { attached = false, closed = true, cancellationRequested = true });
                });
                return true;
            }
            throw new CoreFault("office_session_expired", "Office session is unavailable.");
        }
        try
        {
            result = host.Sta.Invoke(() => HandleWindowOnSta(method, value, session, host));
            return result is not null;
        }
        catch (CoreFault) { throw; }
    }

    private static void RecordFailure(string step, Exception error)
    {
        var code = error is COMException com
            ? "com:" + unchecked((uint)com.HResult).ToString("X8")
            : "ex:" + error.GetType().Name;
        lock (Gate) lastFailure = step + ":" + code;
        Trace("failed_" + step + "_" + code);
    }

    private static JsonNode OpenOnSta(string path, OfficeDocumentKind kind, KindHost host, int width, int height, string sessionId, string surfaceToken, CancellationToken cancellation, bool bindSource, bool visible, int jsX, int jsY, double pixelRatio = 1, string? backupId = null)
    {
        cancellation.ThrowIfCancellationRequested();
        Trace("sta_started");
        if (width < 32 || height < 32 || !WebviewSurface.IsToken(surfaceToken))
            throw new CoreFault("invalid_host_window", "A valid editor webview surface is required.");
        OleSite? site = null;
        var step = "window";
        List<IntPtr>? existingFrames = null;
        HashSet<uint>? protectPids = null;
        var exclusive = false;
        try
        {
            site = new OleSite { DocumentPath = path, SourceBound = bindSource, Sta = host.Sta, PixelRatio = pixelRatio,
                Files = NativeSessionFiles.Create(path, sessionId, backupId), ForceDirty = !string.IsNullOrEmpty(backupId), OwnObject = true, SaveOnClose = true };
            // #region agent log
            DebugNdjson.Write("E", "OleDocumentHost.OpenOnSta", "shared_primary", new { kindKey = kind.Key, working = Path.GetFileName(site.Files.WorkingPath) });
            // #endregion
            site.CreateWindow();
            Trace("container_created");
            step = "place";
            site.Place(surfaceToken, false, width, height, visible ? 6000 : 400, jsX, jsY);
            NativeWindow.PumpWait(30);
            EnterOleExclusive(cancellation, kind.Key);
            exclusive = true;
            protectPids = ProtectedPids(site.Handle);
            existingFrames = SnapshotVisibleTopLevel();
            step = "storage";
            Guid iid = typeof(IOleObject).GUID;
            object unknown;
            var sourceLength = new FileInfo(path).Length;
            // #region agent log
            DebugNdjson.Write("K", "OleDocumentHost.OpenOnSta", "bind_choice", new
            {
                bindSource, via = bindSource ? "BindMoniker" : "OleCreateFromFile",
                existingFrames = existingFrames.Count, kindKey = kind.Key,
                clsid = kind.Clsid.ToString("B"), getClassHr = kind.GetClassFileHr
            });
            // #endregion
            if (bindSource)
            {
                step = "bind_source";
                Trace("bind_source_started");
                // #region agent log
                DebugNdjson.Write("K", "OleDocumentHost.OpenOnSta", "bind_enter", new { via = "BindMoniker" });
                // #endregion
                Marshal.ThrowExceptionForHR(CreateFileMoniker(path, out var moniker));
                try { Marshal.ThrowExceptionForHR(BindMoniker(moniker, 0, ref iid, out unknown)); }
                finally { Marshal.Release(moniker); }
                Trace("bind_source_completed");
                // #region agent log
                DebugNdjson.Write("K", "OleDocumentHost.OpenOnSta", "bind_done", new { via = "BindMoniker" });
                // #endregion
            }
            else
            {
                IStorage storage = CreateStorage();
                site.Storage = storage;
                Guid clsid = kind.Clsid;
                if (clsid == Guid.Empty && kind.ProgId.Length > 0 && CLSIDFromProgID(kind.ProgId, out var fromProg) == SOk)
                    clsid = fromProg;
                storage.SetClass(ref clsid);
                step = "create";
                Trace("ole_create_started");
                if (clsid == Guid.Empty)
                    throw new CoreFault("ole_embed_failed", "This file has no OLE document server.");
                var embedPath = site.Files.WorkingPath;
                NativeWindow.PumpWait(30);
                // #region agent log
                DebugNdjson.Write("D", "OleDocumentHost.OpenOnSta", "bind_enter", new { via = "OleCreateFromFile", clsid = clsid.ToString("B"), kindKey = kind.Key, sourceLength, embedLength = new FileInfo(embedPath).Length, staged = true, exclusiveKind = kind.Key });
                // #endregion
                Marshal.ThrowExceptionForHR(OleCreateFromFile(ref clsid, embedPath, ref iid, OleRenderNone, IntPtr.Zero, site, storage, out unknown));
                cancellation.ThrowIfCancellationRequested();
                var sourceAfterCreate = new FileInfo(path).Length;
                Trace("ole_create_completed");
                // #region agent log
                DebugNdjson.Write("S", "OleDocumentHost.OpenOnSta", "bind_done", new { via = "OleCreateFromFile", kindKey = kind.Key, sourceLength, sourceAfterCreate, sourceChanged = sourceAfterCreate != sourceLength });
                // #endregion
                SweepNewFrames(existingFrames ?? [], site.Handle, protectPids ?? []);
            }
            step = "cast";
            var oleObject = (IOleObject)unknown;
            site.Object = oleObject;
            if (site.AdviseCookie == 0)
            {
                oleObject.Advise(site, out var cookie);
                site.AdviseCookie = cookie;
            }
            if (bindSource)
            {
                var source = oleObject as IPersistFile ?? unknown as IPersistFile;
                string? original = null;
                var getCur = source is null ? "no_persist" : "ok";
                try { if (source is not null) source.GetCurFile(out original); }
                catch (Exception error)
                {
                    getCur = error.GetType().Name;
                    // #region agent log
                    DebugNdjson.Write("H", "OleDocumentHost.OpenOnSta", "getcurfile_threw", new { type = error.GetType().Name, hr = error.HResult });
                    // #endregion
                }
                var match = SameLocalFile(original, path);
                // #region agent log
                DebugNdjson.Write("H", "OleDocumentHost.OpenOnSta", "bind_identity", new { persist = source is not null, getCur, curEmpty = string.IsNullOrEmpty(original), match });
                // #endregion
                if (source is not null && match)
                {
                    site.SourceIdentityVerified = true;
                    try { site.RefreshSourceStamp(); } catch { site.SourceIdentityVerified = false; }
                    oleObject.Advise(site, out var adviseCookie);
                    site.AdviseCookie = adviseCookie;
                }
            }
            step = "site";
            Trace("set_client_site_started");
            oleObject.SetClientSite(site);
            try { site.RefreshSourceStamp(); } catch { }
            Trace("set_client_site_completed");
            if (CreateFileMoniker(path, out var fileMoniker) == SOk)
            {
                try { oleObject.SetMoniker(OleWhichMkObjFull, fileMoniker); }
                catch { /* Some servers reject a container-assigned file moniker. */ }
                finally { Marshal.Release(fileMoniker); }
            }
            if (unknown is IPersistFile persist)
            {
                try
                {
                    persist.Save(path, true);
                    persist.SaveCompleted(path);
                    site.RefreshSourceStamp();
                    // #region agent log
                    DebugNdjson.Write("B", "OleDocumentHost.OpenOnSta", "persist_bound", new { saved = true });
                    // #endregion
                }
                catch (Exception error)
                {
                    // #region agent log
                    DebugNdjson.Write("B", "OleDocumentHost.OpenOnSta", "persist_bind_fail", new { type = error.GetType().Name, hr = error.HResult });
                    // #endregion
                }
            }
            Trace("set_host_names_started");
            // Word requires host names even for a file-bound DocObject; without
            // them it treats the view as a link and opens an external frame.
            oleObject.SetHostNames("Office in Editor", bindSource ? path : Path.GetFileName(path));
            Trace("set_host_names_completed");
            Marshal.ThrowExceptionForHR(OleSetContainedObject(unknown, 1));
            Trace("contained_object_completed");
            if (!OleIsRunning(unknown)) Marshal.ThrowExceptionForHR(OleRun(unknown));
            Trace("ole_run_completed");
            site.Object = oleObject;
            cancellation.ThrowIfCancellationRequested();
            site.Shared = new SharedDocument { Object = oleObject, Storage = site.Storage, Files = site.Files, Views = 1, SaveOnClose = true };
            var session = new Session { Id = sessionId, Application = kind.Application, Path = path, Site = site, SurfaceToken = surfaceToken,
                Width = width, Height = height, PixelRatio = pixelRatio, Visible = visible, JsX = jsX, JsY = jsY };
            lock (host.Gate)
            {
                host.Sessions[sessionId] = session;
                host.SessionCount = host.Sessions.Count;
            }
            sessionHosts[sessionId] = host;
            session.Attached = visible;
            step = "activate";
            Trace("activation_started");
            try { site.ActivateInPlace(); }
            finally { SweepNewFrames(existingFrames ?? [], site.Handle, protectPids ?? []); }
            Trace("activation_completed");
            session.Activated = true;
            GetWindowThreadProcessId(site.ViewHandle == IntPtr.Zero ? site.Handle : site.ViewHandle, out var documentPid);
            session.ServerPid = (int)documentPid;
            site.ResolveNativeObjectModel(kind.Application);
            if (site.NativeObjectModelHandle != IntPtr.Zero)
            {
                GetWindowThreadProcessId(site.NativeObjectModelHandle, out documentPid);
                session.ServerPid = (int)documentPid;
            }
            lock (Gate) lastFailure = "";
            // #region agent log
            DebugNdjson.Write("A", "OleDocumentHost.OpenOnSta", "open_ok", new { kindKey = kind.Key, application = kind.Application, pid = session.ServerPid, inplace = true });
            // #endregion
            if (cancellation.IsCancellationRequested)
            {
                site.Place(surfaceToken, false, width, height);
                CloseOnSta(session, host);
                throw new OperationCanceledException(cancellation);
            }
            if (visible) site.Place(surfaceToken, true, width, height, 6000, jsX, jsY);
            site.ConfigurePeriodicSave(true, 60);
            // #region agent log
            DebugNdjson.Write("A", "OleDocumentHost.OpenOnSta", "open_reveal", new { visible, shown = site.Visible, kindKey = kind.Key, periodicSave = site.PeriodicSave, interval = site.SaveIntervalSeconds, saveOnClose = site.SaveOnClose });
            // #endregion
            return Wire.Node(new
            {
                opened = true,
                mode = "embedded",
                application = kind.Application,
                kind = kind.Key,
                pid = session.ServerPid,
                inplace = true,
                officeId = session.Id,
                dirty = site.IsDirty,
                hwnd = site.Handle.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture)
            }) ?? throw new InvalidOperationException();
        }
        catch (Exception error)
        {
            RecordFailure(step, error);
            // #region agent log
            DebugNdjson.Write("F", "OleDocumentHost.OpenOnSta", "open_failed", new { step, lastFailure = LastFailure, type = error.GetType().Name, hr = error.HResult, kindKey = kind.Key });
            // #endregion
            bool retained;
            lock (host.Gate) retained = host.Sessions.TryGetValue(sessionId, out var owned) && ReferenceEquals(owned.Site, site);
            if (retained)
            {
                try { AbortOnSta(sessionId, host); }
                catch
                {
                    lock (host.Gate)
                    {
                        if (host.Sessions.TryGetValue(sessionId, out var leftover) && ReferenceEquals(leftover.Site, site))
                        {
                            host.Sessions.Remove(sessionId);
                            host.SessionCount = host.Sessions.Count;
                        }
                    }
                    sessionHosts.TryRemove(sessionId, out _);
                    site?.Dispose();
                }
            }
            else site?.Abort();
            if (site is not null && existingFrames is not null)
                try { SweepNewFrames(existingFrames, site.Handle, protectPids ?? []); } catch { }
            throw;
        }
        finally
        {
            if (exclusive) ExitOleExclusive(kind.Key);
        }
    }

    private static JsonNode AttachViewOnSta(Session primary, OfficeDocumentKind kind, KindHost host, int width, int height, string sessionId, string surfaceToken, CancellationToken cancellation, bool visible, int jsX, int jsY, double pixelRatio)
    {
        cancellation.ThrowIfCancellationRequested();
        if (width < 32 || height < 32 || !WebviewSurface.IsToken(surfaceToken))
            throw new CoreFault("invalid_host_window", "A valid editor webview surface is required.");
        primary.Site.Shared ??= new SharedDocument
        {
            Object = primary.Site.Object, Storage = primary.Site.Storage, Files = primary.Site.Files,
            Views = Math.Max(1, CountSharedViews(host, primary.Site)), SaveOnClose = primary.Site.SaveOnClose
        };
        var shared = primary.Site.Shared;
        var target = new OleSite
        {
            DocumentPath = primary.Path, SourceBound = false, Sta = host.Sta, PixelRatio = pixelRatio,
            Files = shared.Files!, Shared = shared, OwnObject = false, Object = shared.Object, Storage = shared.Storage,
            SaveOnClose = shared.SaveOnClose
        };
        target.CreateWindow();
        target.Place(surfaceToken, false, width, height, 400, jsX, jsY);
        var mode = "alias";
        var cloned = false;
        try { cloned = primary.Site.TryCloneView(target); }
        catch { cloned = false; }
        if (!cloned)
            try { cloned = target.TryCreateDocumentView(); }
            catch { cloned = false; }
        Session session;
        if (cloned)
        {
            mode = "clone_view";
            shared.Views++;
            session = new Session
            {
                Id = sessionId, Application = kind.Application, Path = primary.Path, Site = target,
                SurfaceToken = surfaceToken, Width = width, Height = height, PixelRatio = pixelRatio,
                Visible = visible, JsX = jsX, JsY = jsY, ServerPid = primary.ServerPid, Activated = true, Attached = visible
            };
        }
        else
        {
            target.AbandonUnusedWindow();
            shared.Views++;
            session = new Session
            {
                Id = sessionId, Application = kind.Application, Path = primary.Path, Site = primary.Site,
                SurfaceToken = surfaceToken, Width = width, Height = height, PixelRatio = pixelRatio,
                Visible = visible, JsX = jsX, JsY = jsY, ServerPid = primary.ServerPid,
                SharedAlias = true, Activated = primary.Activated, Attached = visible
            };
        }
        lock (host.Gate)
        {
            host.Sessions[sessionId] = session;
            host.SessionCount = host.Sessions.Count;
        }
        sessionHosts[sessionId] = host;
        if (cloned)
        {
            if (visible) target.Place(surfaceToken, true, width, height, 1500, jsX, jsY);
        }
        else PlaceSharedSite(host, primary.Site);
        // #region agent log
        DebugNdjson.Write("A", "OleDocumentHost.AttachViewOnSta", "share_ok", new
        {
            kindKey = kind.Key, application = kind.Application, mode, views = shared.Views,
            sessionPrefix = sessionId.Length > 12 ? sessionId[4..12] : sessionId,
            primaryPrefix = primary.Id.Length > 12 ? primary.Id[4..12] : primary.Id, visible
        });
        // #endregion
        return Wire.Node(new
        {
            opened = true, mode = "embedded", application = kind.Application, kind = kind.Key,
            pid = session.ServerPid, inplace = true, officeId = session.Id, dirty = session.Site.IsDirty,
            hwnd = session.Site.Handle.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture),
            shared = true, viewMode = mode
        }) ?? throw new InvalidOperationException();
    }

    private static Session? FindLiveSession(string sourceKey, string exceptId)
    {
        if (!sourceOwners.TryGetValue(sourceKey, out var owners)) return null;
        foreach (var id in owners.Keys)
        {
            if (id == exceptId) continue;
            if (!sessionHosts.TryGetValue(id, out var host)) continue;
            lock (host.Gate)
                if (host.Sessions.TryGetValue(id, out var session))
                    return session;
        }
        return null;
    }

    private static int CountSharedViews(KindHost host, OleSite site)
    {
        lock (host.Gate)
            return host.Sessions.Values.Count(item => SameDocumentSite(item, site));
    }

    private static bool SameDocumentSite(Session session, OleSite site) =>
        ReferenceEquals(session.Site, site)
        || (session.Site.Shared is not null && site.Shared is not null && ReferenceEquals(session.Site.Shared, site.Shared));

    private static int CountSiteViews(KindHost host, OleSite site)
    {
        lock (host.Gate)
            return host.Sessions.Values.Count(item => ReferenceEquals(item.Site, site));
    }

    private static void PlaceSharedSite(KindHost host, OleSite site)
    {
        Session? placeFrom = null;
        var show = false;
        lock (host.Gate)
        {
            foreach (var item in host.Sessions.Values)
            {
                if (!ReferenceEquals(item.Site, site)) continue;
                if (item.Visible) { show = true; placeFrom = item; }
                else placeFrom ??= item;
            }
        }
        if (placeFrom is null) return;
        site.Place(placeFrom.SurfaceToken, show, placeFrom.Width, placeFrom.Height, show ? 1500 : 400, placeFrom.JsX, placeFrom.JsY);
    }

    private static object OleLock(string kindKey) => OleExclusiveByKind.GetOrAdd(kindKey, static _ => new object());
    private static void EnterOleExclusive(CancellationToken cancellation, string kindKey)
    {
        var gate = OleLock(kindKey);
        var started = Stopwatch.GetTimestamp();
        while (!Monitor.TryEnter(gate, 40))
        {
            cancellation.ThrowIfCancellationRequested();
            NativeWindow.PumpWait(40);
        }
        // #region agent log
        DebugNdjson.Write("L", "OleDocumentHost.OpenOnSta", "ole_exclusive_enter", new { kindKey, lockScope = "kind", waitMs = (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds });
        // #endregion
    }

    private static void ExitOleExclusive(string kindKey)
    {
        if (OleExclusiveByKind.TryGetValue(kindKey, out var gate) && Monitor.IsEntered(gate)) Monitor.Exit(gate);
        // #region agent log
        DebugNdjson.Write("L", "OleDocumentHost.OpenOnSta", "ole_exclusive_exit", new { kindKey, lockScope = "kind" });
        // #endregion
    }

    private static int CountSource(string sourceKey) =>
        sourceOwners.TryGetValue(sourceKey, out var owners) ? owners.Count : 0;

    private static void ClaimSource(string sourceKey, string sessionId) =>
        sourceOwners.GetOrAdd(sourceKey, static _ => new ConcurrentDictionary<string, byte>(StringComparer.Ordinal)).TryAdd(sessionId, 0);

    private static void ReleaseSource(string sourceKey, string sessionId)
    {
        if (!sourceOwners.TryGetValue(sourceKey, out var owners)) return;
        owners.TryRemove(sessionId, out _);
        if (owners.IsEmpty) sourceOwners.TryRemove(sourceKey, out _);
    }

    private static JsonNode? HandleWindowOnSta(string method, JsonNode? value, Session session, KindHost host)
    {
        lock (host.Gate)
            if (!host.Sessions.TryGetValue(session.Id, out var current) || !ReferenceEquals(current, session))
                throw new CoreFault("office_session_expired", "The Office session is no longer active.");
        if (method == "office.describe_window")
            return Wire.Node(new { hwnd = session.Site.Handle.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture), processId = Environment.ProcessId, attached = session.Attached, inplace = true });
        if (method == "office.status_window") return session.Site.Status();
        if (method == "office.configure_save_window")
        {
            if (value?["enabled"] is not null)
                session.Site.ConfigurePeriodicSave(Wire.Boolean(value, "enabled"), Wire.Number(value, "seconds", session.Site.SaveIntervalSeconds));
            if (value?["saveOnClose"] is not null)
            {
                session.Site.SaveOnClose = Wire.Boolean(value, "saveOnClose");
                if (session.Site.Shared is not null) session.Site.Shared.SaveOnClose = session.Site.SaveOnClose;
            }
            return session.Site.Status();
        }
        if (method == "office.backup_window") return Wire.Node(new { backupId = session.Site.Backup() });
        if (method == "office.release_backup_window")
        {
            NativeSessionFiles.ReleaseBackup(Wire.Text(value, "backupId", 128));
            return Wire.Node(new { released = true });
        }
        if (method == "office.revert_window")
        {
            var visible = session.Site.Visible;
            CloseOnSta(session, host, discard: true);
            if (!OfficeDocumentKind.TryClassify(session.Path, out var kind)) throw new CoreFault("unsupported_format", "The document server is unavailable.");
            ClaimSource(Path.GetFullPath(session.Path), session.Id);
            try
            {
                return OpenOnSta(session.Path, kind, host, session.Width, session.Height, session.Id, session.SurfaceToken,
                    CancellationToken.None, false, visible, session.JsX, session.JsY, session.PixelRatio);
            }
            catch
            {
                ReleaseSource(Path.GetFullPath(session.Path), session.Id);
                throw;
            }
        }
        if (method is "office.attach_window" or "office.layout_window")
        {
            bool show = Wire.OptionalBoolean(value, "visible") ?? true;
            var token = Wire.Text(value, "surfaceToken", 80, false);
            if (WebviewSurface.IsToken(token)) session.SurfaceToken = token;
            if (!WebviewSurface.IsToken(session.SurfaceToken))
                throw new CoreFault("invalid_host_window", "The editor tab surface is unavailable.");
            session.Width = Wire.Number(value, "width", session.Width);
            session.Height = Wire.Number(value, "height", session.Height);
            session.JsX = Wire.Number(value, "x", session.JsX);
            session.JsY = Wire.Number(value, "y", session.JsY);
            session.Visible = show;
            if (value?["pixelRatio"] is JsonValue scale && scale.TryGetValue<double>(out var ratio) && ratio is >= 0.5 and <= 8)
                session.PixelRatio = session.Site.PixelRatio = ratio;
            if (session.SharedAlias || CountSiteViews(host, session.Site) > 1)
                PlaceSharedSite(host, session.Site);
            else
                session.Site.Place(session.SurfaceToken, show, Wire.Number(value, "width"), Wire.Number(value, "height"), show ? 1500 : 400, session.JsX, session.JsY);
            session.Attached = show;
            if (show && !session.Activated)
            {
                var existing = SnapshotVisibleTopLevel();
                try { session.Site.ActivateInPlace(); }
                finally { HideNewFrames(existing, session.Site.Handle, ProtectedPids(session.Site.Handle)); }
                session.Activated = true;
            }
            return Wire.Node(new { attached = show, hwnd = session.Site.Handle.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture), inplace = true, pid = session.ServerPid });
        }
        if (method == "office.detach_window")
        {
            CloseOnSta(session, host, Wire.OptionalBoolean(value, "discard") ?? false);
            return Wire.Node(new { attached = false, inplace = true, closed = true });
        }
        if (method == "office.save_window")
        {
            Marshal.ThrowExceptionForHR(session.Site.SaveObject());
            return session.Site.Status(saved: true);
        }
        return null;
    }

    internal static void AbortOrphan()
    {
        foreach (var host in kinds.Values)
        {
            string[] ids;
            lock (host.Gate) ids = host.Sessions.Keys.ToArray();
            foreach (var id in ids)
                try { host.Sta.Invoke(() => { AbortOnSta(id, host); return true; }); }
                catch { /* A disconnected caller must not leave a host window. */ }
        }
    }

    private static void AbortOnSta(string sessionId, KindHost host)
    {
        Session? session;
        lock (host.Gate)
        {
            if (!host.Sessions.TryGetValue(sessionId, out session)) return;
            host.Sessions.Remove(sessionId);
            host.SessionCount = host.Sessions.Count;
        }
        sessionHosts.TryRemove(sessionId, out _);
        ReleaseSource(Path.GetFullPath(session.Path), session.Id);
        session.Site.Abort();
    }

    private static void CloseOnSta(Session session, KindHost host, bool discard = false)
    {
        lock (host.Gate)
            if (!host.Sessions.TryGetValue(session.Id, out var current) || !ReferenceEquals(current, session))
                return;
        session.Visible = false;
        session.Attached = false;
        var site = session.Site;
        var remaining = CountSharedViews(host, site);
        var lastView = remaining <= 1;
        try { site.Place(session.SurfaceToken, false, session.Width, session.Height, 400, session.JsX, session.JsY); }
        catch { /* Hide before save/close so the popup cannot outlive the tab. */ }
        // #region agent log
        DebugNdjson.Write("C", "OleDocumentHost.CloseOnSta", "close_begin", new
        {
            lastView, remaining, alias = session.SharedAlias, ownObject = site.OwnObject,
            saveOnClose = site.Shared?.SaveOnClose ?? site.SaveOnClose, discard, dirty = site.IsDirty,
            sessionPrefix = session.Id.Length > 12 ? session.Id[4..12] : session.Id
        });
        // #endregion
        if (!lastView)
        {
            lock (host.Gate)
            {
                host.Sessions.Remove(session.Id);
                host.SessionCount = host.Sessions.Count;
            }
            sessionHosts.TryRemove(session.Id, out _);
            ReleaseSource(Path.GetFullPath(session.Path), session.Id);
            if (site.Shared is not null) site.Shared.Views = Math.Max(0, site.Shared.Views - 1);
            if (CountSiteViews(host, site) == 0)
            {
                site.OwnObject = false;
                site.Dispose();
            }
            else
                try { PlaceSharedSite(host, site); } catch { }
            return;
        }
        var saveOnClose = site.Shared?.SaveOnClose ?? site.SaveOnClose;
        if (saveOnClose && !discard)
        {
            try { site.SaveObject(); }
            catch { /* Closing the last tab still destroys the surface. */ }
        }
        try
        {
            var sourceBeforeClose = 0L;
            try { sourceBeforeClose = new FileInfo(session.Path).Length; } catch { }
            var closed = site.Object?.Close(OleCloseNosave) ?? SOk;
            var sourceAfterClose = 0L;
            try { sourceAfterClose = new FileInfo(session.Path).Length; } catch { }
            // #region agent log
            DebugNdjson.Write("S", "OleDocumentHost.CloseOnSta", "close_done", new { closed, sourceBeforeClose, sourceAfterClose, sourceChanged = sourceAfterClose != sourceBeforeClose, lastView });
            // #endregion
        }
        catch
        {
            try { site.Files.MarkRecoveryNeeded(); } catch { }
        }
        site.ReleaseView();
        try { site.Object?.SetClientSite(null!); } catch { }
        try { if (site.Object is not null) OleSetContainedObject(site.Object, 0); } catch { }
        lock (host.Gate)
        {
            host.Sessions.Remove(session.Id);
            host.SessionCount = host.Sessions.Count;
        }
        sessionHosts.TryRemove(session.Id, out _);
        ReleaseSource(Path.GetFullPath(session.Path), session.Id);
        try { site.Files.DisposeClean(); } catch { }
        site.OwnObject = true;
        site.Dispose();
    }

    private static string StagingCopy(string path)
    {
        var root = Path.Combine(Path.GetTempPath(), "OfficeInEditor", "embed");
        Directory.CreateDirectory(root);
        var stamp = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(path.ToUpperInvariant())))[..16];
        var dir = Path.Combine(root, stamp);
        Directory.CreateDirectory(dir);
        var dest = Path.Combine(dir, Path.GetFileName(path));
        File.Copy(path, dest, overwrite: true);
        return dest;
    }

    private static IStorage CreateStorage()
    {
        Marshal.ThrowExceptionForHR(CreateILockBytesOnHGlobal(IntPtr.Zero, 1, out var bytes));
        try
        {
            Marshal.ThrowExceptionForHR(StgCreateDocfileOnILockBytes(bytes, StgmCreate | StgmReadWrite | StgmShareExclusive, 0, out var storage));
            return storage;
        }
        finally { Marshal.Release(bytes); }
    }

    private static bool SameLocalFile(string? left, string right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;
        try
        {
            if (!Path.IsPathFullyQualified(left))
            {
                var dir = Path.GetDirectoryName(right);
                if (!string.IsNullOrEmpty(dir)) left = Path.Combine(dir, left);
            }
            return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private sealed class SharedDocument
    {
        internal IOleObject? Object;
        internal IStorage? Storage;
        internal NativeSessionFiles? Files;
        internal int Views;
        internal bool SaveOnClose = true;
    }

    private sealed class KindHost
    {
        internal readonly string Key;
        internal readonly object Gate = new();
        internal readonly Dictionary<string, Session> Sessions = new(StringComparer.Ordinal);
        internal readonly StaWorker Sta;
        internal int Opening;
        internal int SessionCount;
        internal bool Busy => Volatile.Read(ref Opening) > 0 || Volatile.Read(ref SessionCount) > 0;
        internal KindHost(string key)
        {
            Key = key;
            Sta = new StaWorker(key, () => Busy);
        }
    }

    private sealed class Session
    {
        internal required string Id, Application, Path;
        internal required OleSite Site;
        internal string SurfaceToken = "";
        internal int ServerPid;
        internal int Width, Height;
        internal double PixelRatio = 1;
        internal bool Attached;
        internal bool Activated;
        internal bool Visible;
        internal bool SharedAlias;
        internal int JsX, JsY;
    }

    private sealed class StaWorker
    {
        private readonly BlockingCollection<Work> queue = new(32);
        private readonly AutoResetEvent workReady = new(false);
        private readonly Func<bool> busy;
        internal StaWorker(string name, Func<bool> busy)
        {
            this.busy = busy;
            var thread = new Thread(Run) { IsBackground = true, Name = "OieOleSta-" + name };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }
        internal T Invoke<T>(Func<T> work)
        {
            var item = new Work(() => work());
            if (!queue.TryAdd(item)) throw new CoreFault("office_busy", "The Office work queue is full.");
            workReady.Set();
            return (T)item.Completion.Task.GetAwaiter().GetResult()!;
        }
        internal void Post(Action work)
        {
            var item = new Work(() => { work(); return null; });
            if (!queue.TryAdd(item)) throw new CoreFault("office_busy", "The Office work queue is full.");
            // Observe deferred failures; a layout error must not crash the STA.
            _ = item.Completion.Task.ContinueWith(task => { _ = task.Exception; }, TaskContinuationOptions.OnlyOnFaulted);
            workReady.Set();
        }
        private readonly PumpFilter filter = new();
        private void Run()
        {
            OleInitialize(IntPtr.Zero);
            CoRegisterMessageFilter(filter, out _);
            var handles = new[] { workReady.SafeWaitHandle.DangerousGetHandle() };
            while (true)
            {
                Pump();
                while (queue.TryTake(out var item))
                {
                    try { item.Completion.SetResult(item.Callback()); }
                    catch (Exception error) { item.Completion.SetException(error); }
                    Pump();
                }
                MsgWaitForMultipleObjects(1, handles, false, busy() ? 30u : 250u, QsAllInput);
            }
        }
        private static void Pump()
        {
            while (PeekMessage(out var msg, IntPtr.Zero, 0, 0, 1))
            {
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }
        }
        private sealed class Work
        {
            internal readonly Func<object?> Callback;
            internal readonly TaskCompletionSource<object?> Completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
            internal Work(Func<object?> callback) => Callback = callback;
        }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    private sealed class OleSite : StandardOleMarshalObject, ICustomQueryInterface, IOleClientSite, IAdviseSink, IOleWindow, IOleInPlaceSite, IOleInPlaceUIWindow, IOleInPlaceFrame, IOleDocumentSite, IOleParseDisplayName, IOleContainer, IOleCommandTarget, IDisposable
    {
        private readonly NativeWindow.WndProc wndProc;
        private readonly string className = "OieOleHost_" + Guid.NewGuid().ToString("N");
        internal IntPtr Handle { get; private set; }
        private IOleDocumentView? documentView;
        private bool layoutQueued;
        private bool disposed;
        private bool viewActivated, activatingView;
        private readonly int ownerThread = Environment.CurrentManagedThreadId;
        internal IOleObject? Object;
        internal IStorage? Storage;
        internal string DocumentPath = "";
        internal bool OwnObject = true;
        internal SharedDocument? Shared;
        internal bool SaveOnClose = true;
        internal bool SourceBound;
        internal bool SourceIdentityVerified;
        internal NativeSessionFiles Files = null!;
        internal bool ForceDirty;
        internal double PixelRatio = 1;
        internal int AdviseCookie;
        internal StaWorker Sta = null!;
        internal IntPtr ViewHandle { get; private set; }
        internal IntPtr NativeObjectModelHandle { get; private set; }
        internal bool PeriodicSave { get; private set; }
        internal int SaveIntervalSeconds { get; private set; } = 60;
        private string lastSaveError = "";
        private int saveQueued;
        private bool saving;
        private int queryStatusCalls;
        private long sourceLength, sourceLastWrite;
        internal void RefreshSourceStamp()
        {
            var file = new FileInfo(DocumentPath);
            sourceLength = file.Length; sourceLastWrite = file.LastWriteTimeUtc.Ticks;
        }
        internal bool Visible;
        private Timer? silentSave;
        private IntPtr nestedOwner;
        private int nestedX, nestedY, nestedW, nestedH, ownerLeft, ownerTop;
        private string nestedHow = "";
        private bool placing;
        private int borderL, borderT, borderR, borderB;
        private bool applyingLayout;
        private bool keepInPlaceWhileHidden;
        private IOleInPlaceActiveObject? activeObject;
        internal OleSite() => wndProc = WndProc;
        internal void CreateWindow()
        {
            var windowClass = new NativeWindow.WndClassEx
            {
                cbSize = (uint)Marshal.SizeOf<NativeWindow.WndClassEx>(),
                lpfnWndProc = wndProc,
                hInstance = NativeWindow.GetModuleHandle(null),
                lpszClassName = className
            };
            if (NativeWindow.RegisterClassEx(ref windowClass) == 0) throw new InvalidOperationException();
            Handle = NativeWindow.CreateWindowEx(WsExToolwindow, className, "",
                WsPopup | WsClipChildren, 0, 0, 32, 32, IntPtr.Zero, IntPtr.Zero, windowClass.hInstance, IntPtr.Zero);
            if (Handle == IntPtr.Zero) throw new InvalidOperationException();
        }
        internal bool IsDirty
        {
            get
            {
                if (ForceDirty) return true;
                try { return Object is not IPersistFile file || file.IsDirty() != 1; }
                catch { return true; }
            }
        }
        internal JsonNode Status(bool saved = false) => Wire.Node(new { dirty = IsDirty, saved,
            periodicSave = PeriodicSave, saveIntervalSeconds = SaveIntervalSeconds, saveError = lastSaveError,
            saveOnClose = Shared?.SaveOnClose ?? SaveOnClose })!;
        internal void ConfigurePeriodicSave(bool enabled, int seconds)
        {
            if (seconds is < 5 or > 3600) throw new CoreFault("invalid_save_interval", "Choose an interval between 5 and 3600 seconds.");
            silentSave?.Dispose(); silentSave = null;
            PeriodicSave = enabled; SaveIntervalSeconds = seconds; lastSaveError = "";
            if (!enabled || disposed) return;
            silentSave = new Timer(_ =>
            {
                if (disposed || Interlocked.CompareExchange(ref saveQueued, 1, 0) != 0) return;
                try { Sta.Post(() => {
                    try
                    {
                        var dirty = !disposed && PeriodicSave && IsDirty;
                        // #region agent log
                        DebugNdjson.Write("A", "OleSite.PeriodicSave", "tick", new { dirty, enabled = PeriodicSave });
                        // #endregion
                        if (dirty && SaveObject() < 0) PausePeriodicSave();
                    }
                    finally { Interlocked.Exchange(ref saveQueued, 0); }
                }); }
                catch { Interlocked.Exchange(ref saveQueued, 0); }
            }, null, seconds * 1000, seconds * 1000);
        }
        private void PausePeriodicSave()
        {
            PeriodicSave = false; silentSave?.Dispose(); silentSave = null;
        }
        internal string Backup()
        {
            if (Object is not IPersistFile file) throw new CoreFault("backup_unavailable", "The native document cannot be backed up.");
            return Files.Backup(path => { file.Save(path, false); file.SaveCompleted(null); });
        }
        internal void ResolveNativeObjectModel(string application)
        {
            NativeObjectModelHandle = IntPtr.Zero;
            if (application != "word") return;
            NativeWindow.EnumChildWindows(Handle, (window, _) => {
                var name = new StringBuilder(64); NativeWindow.GetClassName(window, name, name.Capacity);
                if (name.ToString() == "_WwG") NativeObjectModelHandle = window;
                return true;
            }, IntPtr.Zero);
        }
        internal NativeRect ClientRect()
        {
            NativeWindow.GetClientRect(Handle, out var rect);
            return rect;
        }
        internal NativeRect ViewportRect()
        {
            var frame = ClientRect();
            var view = new NativeRect
            {
                Left = frame.Left + borderL,
                Top = frame.Top + borderT,
                Right = Math.Max(frame.Left + borderL, frame.Right - borderR),
                Bottom = Math.Max(frame.Top + borderT, frame.Bottom - borderB)
            };
            return view;
        }
        private void ApplyViewLayout()
        {
            if (disposed || documentView is null || applyingLayout) return;
            applyingLayout = true;
            try
            {
                var view = ViewportRect();
                var frame = ClientRect();
                try { Marshal.ThrowExceptionForHR(documentView.SetRect(ref view)); } catch { }
                try { activeObject?.ResizeBorder(ref frame, this, 1); } catch { }
            }
            finally { applyingLayout = false; }
        }
        internal void Place(string token, bool show, int width = 0, int height = 0, int resolveMs = 6000, int jsX = 0, int jsY = 0)
        {
            using var physicalCoordinates = WebviewSurface.EnterPhysicalCoordinates();
            var started = Stopwatch.GetTimestamp();
            // #region agent log
            DebugNdjson.Write("H", "OleSite.Place", "place_start", new { tokenLen = token.Length, show, width, height, cached = nestedOwner != IntPtr.Zero, resolveMs });
            // #endregion
            if (!show)
            {
                keepInPlaceWhileHidden = true;
                Visible = false;
                var hideW = width >= 32 ? width : 32;
                var hideH = height >= 32 ? height : 32;
                NativeWindow.SetWindowPos(Handle, IntPtr.Zero, 0, 0, hideW, hideH, SwpNoActivate | SwpHideWindow);
                // #region agent log
                DebugNdjson.Write("I", "OleSite.Place", "hidden_keep_ole", new { elapsedMs = (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds, keepUi = true, hideW, hideH, borderT, borderB });
                // #endregion
                return;
            }
            placing = true;
            var becomingVisible = !Visible;
            keepInPlaceWhileHidden = false;
            try
            {
                WebviewSurface.Placement surface;
                var sizeDrift = width >= 32 && height >= 32 && (Math.Abs(nestedW - width) > 48 || Math.Abs(nestedH - height) > 48);
                var usedCache = !becomingVisible && !sizeDrift
                    && nestedOwner != IntPtr.Zero && NativeWindow.IsWindow(nestedOwner) && WebviewSurface.IsTrustedHow(nestedHow)
                    && WebviewSurface.FitsEditor(new WebviewSurface.Placement(nestedOwner, nestedX, nestedY, nestedW, nestedH, nestedHow), width, height);
                if (usedCache)
                {
                    NativeWindow.GetWindowRect(nestedOwner, out var ownerNow);
                    nestedX += ownerNow.Left - ownerLeft;
                    nestedY += ownerNow.Top - ownerTop;
                    ownerLeft = ownerNow.Left;
                    ownerTop = ownerNow.Top;
                    surface = new WebviewSurface.Placement(nestedOwner, nestedX, nestedY, nestedW, nestedH, nestedHow + "-cache");
                }
                else
                {
                    var resolved = WebviewSurface.Resolve(token, resolveMs, width, height);
                    if (resolved is null || !WebviewSurface.IsTrustedHow(resolved.Value.How)
                        || !WebviewSurface.FitsEditor(resolved.Value, width, height))
                    {
                        // #region agent log
                        DebugNdjson.Write("G", "OleSite.Place", "surface_miss", new { tokenLen = token.Length, show, resolveMs, jsX, jsY, jsIgnored = true, how = resolved?.How, resolvedW = resolved?.Width, resolvedH = resolved?.Height, expectW = width, expectH = height });
                        // #endregion
                        Visible = false;
                        ShowWindow(Handle, SwHide);
                        if (Object is null)
                            throw new CoreFault("invalid_host_window", "The editor tab surface is unavailable.");
                        return;
                    }
                    surface = resolved.Value;
                    nestedOwner = surface.Parent;
                    nestedX = surface.X;
                    nestedY = surface.Y;
                    nestedW = surface.Width;
                    nestedH = surface.Height;
                    nestedHow = surface.How;
                    NativeWindow.GetWindowRect(nestedOwner, out var placedOwner);
                    ownerLeft = placedOwner.Left;
                    ownerTop = placedOwner.Top;
                }
                NativeWindow.OwnPopup(Handle, surface.Parent);
                NativeWindow.GetWindowRect(nestedOwner, out var ownerBox);
                NativeWindow.GetWindowRect(Handle, out var before);
                var same = Visible
                    && before.Left == surface.X && before.Top == surface.Y
                    && (before.Right - before.Left) == surface.Width
                    && (before.Bottom - before.Top) == surface.Height;
                Visible = true;
                if (!same && !NativeWindow.SetWindowPos(Handle, IntPtr.Zero, surface.X, surface.Y, surface.Width, surface.Height,
                    SwpNoActivate | SwpShowWindow))
                    throw new InvalidOperationException("Office layout failed.");
                NativeWindow.GetWindowRect(Handle, out var placed);
                var host = NativeWindow.DescribeHost(Handle, surface.Parent);
                // #region agent log
                DebugNdjson.Write("G", "OleSite.Place", "nested", new
                {
                    how = surface.How, x = surface.X, y = surface.Y, width = surface.Width, height = surface.Height, show,
                    usedCache, skippedPos = same, jsX, jsY, hostMode = "owned-popup", ownerClass = WebviewSurface.ClassName(surface.Parent),
                    ownerLeft = ownerBox.Left, ownerTop = ownerBox.Top, ownerRight = ownerBox.Right, ownerBottom = ownerBox.Bottom,
                    vx = NativeWindow.GetSystemMetrics(76), vy = NativeWindow.GetSystemMetrics(77),
                    vw = NativeWindow.GetSystemMetrics(78), vh = NativeWindow.GetSystemMetrics(79),
                    placedLeft = placed.Left, placedTop = placed.Top, placedRight = placed.Right, placedBottom = placed.Bottom,
                    popup = host.Popup, child = host.Child, isChildOfOwner = host.IsChildOfOwner,
                    parent = host.Parent, hwndParent = host.HwndParent, style = host.Style,
                    insetLeft = placed.Left - ownerBox.Left, insetTop = placed.Top - ownerBox.Top,
                    trusted = WebviewSurface.IsTrustedHow(surface.How),
                    elapsedMs = (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds
                });
                // #endregion
                if (!same) QueueViewLayout();
                if (becomingVisible) RestoreUiActivation();
            }
            finally { placing = false; }
        }
        private void ReleaseUiActivation()
        {
            try { documentView?.UIActivate(0); } catch { }
            try { if (Object is IOleInPlaceObject inplace) inplace.UIDeactivate(); } catch { }
        }
        private void RestoreUiActivation()
        {
            if (documentView is null) return;
            var started = Stopwatch.GetTimestamp();
            try
            {
                documentView.UIActivate(1);
                ApplyViewLayout();
                documentView.Show(1);
            }
            catch { }
            // #region agent log
            DebugNdjson.Write("H", "OleSite.RestoreUiActivation", "restore_ui", new { elapsedMs = (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds, borderT, borderB, visible = Visible });
            // #endregion
        }
        internal void ActivateInPlace()
        {
            if (Object is null) throw new InvalidOperationException();
            var rect = ClientRect();
            var isDoc = Object is IOleDocument;
            var isInplace = Object is IOleInPlaceObject;
            // #region agent log
            DebugNdjson.Write("A", "OleDocumentHost.ActivateInPlace", "do_verb_start", new { verb = "show", isDoc, isInplace });
            // #endregion
            if (!isDoc && !isInplace)
                throw new CoreFault("ole_embed_failed", "This Office file cannot be hosted in the editor tab.");
            int shown;
            try { shown = Object.DoVerb(OleIverbShow, IntPtr.Zero, this, 0, Handle, ref rect); }
            catch (Exception error)
            {
                // #region agent log
                DebugNdjson.Write("A", "OleDocumentHost.ActivateInPlace", "do_verb_threw", new { type = error.GetType().Name, hr = error.HResult });
                // #endregion
                throw;
            }
            // #region agent log
            DebugNdjson.Write("A", "OleDocumentHost.ActivateInPlace", "do_verb_done", new { shown });
            // #endregion
            if (shown != SOk)
            {
                int primary = Object.DoVerb(OleIverbPrimary, IntPtr.Zero, this, 0, Handle, ref rect);
                int inplace = primary == SOk ? SOk : Object.DoVerb(OleIverbInPlaceActivate, IntPtr.Zero, this, 0, Handle, ref rect);
                // #region agent log
                DebugNdjson.Write("A", "OleDocumentHost.ActivateInPlace", "do_verb_inplace", new { primary, inplace, isDoc });
                // #endregion
                if (primary != SOk && inplace != SOk && !isDoc)
                    Marshal.ThrowExceptionForHR(shown);
            }
            if (!viewActivated) CompleteActivation();
            var hosted = TryBindViewHandle();
            var host = NativeWindow.DescribeHost(Handle, nestedOwner);
            var frame = ClientRect();
            var view = ViewportRect();
            // #region agent log
            DebugNdjson.Write("E", "OleDocumentHost.ActivateInPlace", "view_bind", new
            {
                hosted, view = ViewHandle.ToInt64(), hostHwnd = Handle.ToInt64(),
                popup = host.Popup, child = host.Child,
                borderL, borderT, borderR, borderB,
                frameW = frame.Right - frame.Left, frameH = frame.Bottom - frame.Top,
                viewL = view.Left, viewT = view.Top, viewR = view.Right, viewB = view.Bottom
            });
            // #endregion
            if (!hosted)
                throw new InvalidOperationException("The Office view is not hosted in the native container.");
            Trace("native_child_verified");
        }
        private bool TryBindViewHandle()
        {
            if (documentView is IOleWindow viewWindow && viewWindow.GetWindow(out var viewHwnd) == SOk && IsHostedChild(viewHwnd))
            {
                ViewHandle = viewHwnd;
                return true;
            }
            if (Object is IOleInPlaceObject inplace && inplace.GetWindow(out var inplaceHwnd) == SOk && IsHostedChild(inplaceHwnd))
            {
                ViewHandle = inplaceHwnd;
                return true;
            }
            IntPtr child = IntPtr.Zero;
            NativeWindow.EnumChildWindows(Handle, (window, _) =>
            {
                if (window == Handle) return true;
                child = window;
                return false;
            }, IntPtr.Zero);
            if (!IsHostedChild(child)) return false;
            ViewHandle = child;
            return true;
        }
        private bool IsHostedChild(IntPtr hwnd) =>
            hwnd != IntPtr.Zero && (hwnd == Handle || NativeWindow.IsChild(Handle, hwnd));
        internal void Abort()
        {
            try
            {
                ReleaseView();
                if (OwnObject)
                {
                    try { Object?.Close(OleCloseNosave); } catch { }
                    try { Object?.SetClientSite(null!); } catch { }
                    try { if (Object is not null) OleSetContainedObject(Object, 0); } catch { }
                }
            }
            finally { Dispose(); }
        }
        private void QueueViewLayout()
        {
            if (disposed || documentView is null || layoutQueued) return;
            layoutQueued = true;
            Sta.Post(() => {
                layoutQueued = false;
                if (disposed || documentView is null) return;
                ApplyViewLayout();
            });
        }
        public void Dispose()
        {
            if (Handle == IntPtr.Zero && disposed) return;
            disposed = true;
            if (OwnObject && AdviseCookie != 0) { try { Object?.Unadvise(AdviseCookie); } catch { } AdviseCookie = 0; }
            else AdviseCookie = 0;
            ReleaseView();
            if (OwnObject)
            {
                try { if (Object is not null) Marshal.ReleaseComObject(Object); } catch { }
                try { if (Storage is not null) Marshal.ReleaseComObject(Storage); } catch { }
            }
            Object = null;
            Storage = null;
            try { silentSave?.Dispose(); } catch { }
            silentSave = null;
            if (Handle != IntPtr.Zero)
            {
                NativeWindow.DestroyWindow(Handle);
                Handle = IntPtr.Zero;
                NativeWindow.UnregisterClass(className, NativeWindow.GetModuleHandle(null));
            }
        }
        internal void ReleaseView()
        {
            if (documentView is not null)
            {
                try { documentView.UIActivate(0); documentView.Show(0); documentView.CloseView(0); documentView.SetInPlaceSite(null!); } catch { }
                if (!ReferenceEquals(documentView, Object)) try { Marshal.ReleaseComObject(documentView); } catch { }
                documentView = null;
            }
            ViewHandle = IntPtr.Zero;
            viewActivated = false;
        }
        internal bool TryCloneView(OleSite target)
        {
            if (documentView is null || Object is null) return false;
            var hr = documentView.Clone(target, out var cloned);
            // #region agent log
            DebugNdjson.Write("A", "OleSite.TryCloneView", "clone", new { hr });
            // #endregion
            if (hr != SOk || cloned is null) return false;
            target.documentView = cloned;
            target.Object = Object;
            target.Storage = Storage;
            target.Files = Files;
            target.Shared = Shared;
            target.OwnObject = false;
            try { cloned.SetInPlaceSite(target); } catch { }
            try { cloned.UIActivate(1); } catch { }
            target.ApplyViewLayout();
            try { cloned.Show(1); } catch { }
            target.viewActivated = true;
            return true;
        }
        internal bool TryCreateDocumentView()
        {
            if (Object is not IOleDocument document) return false;
            var hr = document.CreateView(this, IntPtr.Zero, 0, out var view);
            // #region agent log
            DebugNdjson.Write("A", "OleSite.TryCreateDocumentView", "create_view", new { hr });
            // #endregion
            if (hr != SOk || view is null) return false;
            documentView = view;
            try { view.SetInPlaceSite(this); } catch { }
            try { view.UIActivate(1); } catch { }
            ApplyViewLayout();
            try { view.Show(1); } catch { }
            viewActivated = true;
            return true;
        }
        internal void AbandonUnusedWindow()
        {
            ReleaseView();
            if (Handle != IntPtr.Zero)
            {
                NativeWindow.DestroyWindow(Handle);
                Handle = IntPtr.Zero;
                NativeWindow.UnregisterClass(className, NativeWindow.GetModuleHandle(null));
            }
            disposed = true;
            Object = null;
            Storage = null;
            Files = null!;
        }
        private IntPtr WndProc(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam)
        {
            if (msg == WmSize)
                QueueViewLayout();
            if (msg == WmActivate)
            {
                var active = (wParam.ToInt64() & 0xFFFF) != 0;
                if (placing || keepInPlaceWhileHidden || !Visible)
                    return NativeWindow.DefWindowProc(hwnd, msg, wParam, lParam);
                // #region agent log
                DebugNdjson.Write("I", "OleSite.WndProc", "wm_activate", new { active, visible = Visible });
                // #endregion
                if (active) RestoreUiActivation();
                else ReleaseUiActivation();
            }
            return NativeWindow.DefWindowProc(hwnd, msg, wParam, lParam);
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            if (iid == typeof(IOleCommandTarget).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IOleCommandTarget), CustomQueryInterfaceMode.Ignore);
                // #region agent log
                DebugNdjson.Write("B", "OleSite.GetInterface", "qi_cmdtarget", new { ok = ppv != IntPtr.Zero });
                // #endregion
                return CustomQueryInterfaceResult.Handled;
            }
            ppv = IntPtr.Zero;
            return CustomQueryInterfaceResult.NotHandled;
        }
        public int SaveObject()
        {
            if (Object is not IPersistFile file) return ENotimpl;
            if (saving) return SOk;
            saving = true;
            try
            {
                Trace("source_save_started");
                var current = new FileInfo(DocumentPath);
                if (!current.Exists || current.IsReadOnly)
                    throw new CoreFault("document_conflict", "The source file is missing or read-only.");
                if (sourceLength > 0 && current.Length != sourceLength && current.LastWriteTimeUtc.Ticks != sourceLastWrite)
                    throw new CoreFault("document_conflict", "The source file changed outside this session.");
                file.Save(DocumentPath, true);
                try { file.SaveCompleted(DocumentPath); } catch { /* Some servers omit SaveCompleted. */ }
                ForceDirty = false;
                lastSaveError = "";
                RefreshSourceStamp();
                Trace("source_save_completed");
                // #region agent log
                DebugNdjson.Write("B", "OleSite.SaveObject", "save_ok", new { length = new FileInfo(DocumentPath).Length, bound = SourceBound, saveOnClose = SaveOnClose });
                // #endregion
                return SOk;
            }
            catch (Exception error)
            {
                lastSaveError = error is CoreFault fault ? fault.Code : "save_failed";
                RecordFailure("save_source", error);
                // #region agent log
                DebugNdjson.Write("B", "OleSite.SaveObject", "save_fail", new { type = error.GetType().Name, hr = error.HResult });
                // #endregion
                return Marshal.GetHRForException(error);
            }
            finally { saving = false; }
        }
        public int GetMoniker(int dwAssign, int dwWhichMoniker, out IntPtr ppmk)
        {
            ppmk = IntPtr.Zero;
            if (string.IsNullOrEmpty(DocumentPath)) return ENotimpl;
            var hr = CreateFileMoniker(DocumentPath, out ppmk);
            // #region agent log
            DebugNdjson.Write("B", "OleSite.GetMoniker", "moniker", new { dwAssign, dwWhichMoniker, hr });
            // #endregion
            return hr;
        }
        public int GetContainer(out IntPtr ppContainer)
        {
            ppContainer = Marshal.GetComInterfaceForObject(this, typeof(IOleContainer));
            return SOk;
        }
        public int ParseDisplayName(IntPtr pbc, string pszDisplayName, out uint pchEaten, out IntPtr ppmkOut)
        {
            pchEaten = 0; ppmkOut = IntPtr.Zero; return ENotimpl;
        }
        public int EnumObjects(uint grfFlags, out IntPtr ppenum) { ppenum = IntPtr.Zero; return ENotimpl; }
        public int LockContainer(int fLock) => SOk;
        public int ShowObject() { if (Visible) ShowWindow(Handle, SwShow); return SOk; }
        public int OnShowWindow(int fShow) => SOk;
        public int RequestNewObjectLayout() => SOk;
        public void OnDataChange(ref FORMATETC format, ref STGMEDIUM stgmed) { }
        public void OnViewChange(int dwAspect, int lindex) { }
        public void OnRename(object pmk) { }
        public void OnSave()
        {
            if (SourceBound && SourceIdentityVerified) try { RefreshSourceStamp(); } catch { SourceIdentityVerified = false; }
        }
        public void OnClose() { }
        public int GetWindow(out IntPtr phwnd) { phwnd = Handle; return SOk; }
        public int ContextSensitiveHelp(int fEnterMode) => SOk;
        public int CanInPlaceActivate() => SOk;
        public int OnInPlaceActivate() => SOk;
        public int OnUIActivate() => SOk;
        public int GetWindowContext(out IOleInPlaceFrame ppFrame, out IOleInPlaceUIWindow? ppDoc, out NativeRect lprcPosRect, out NativeRect lprcClipRect, ref OleInPlaceFrameInfo lpFrameInfo)
        {
            Trace("site_window_context");
            ppFrame = this;
            ppDoc = null;
            lprcPosRect = ViewportRect();
            lprcClipRect = lprcPosRect;
            lpFrameInfo.cb = Marshal.SizeOf<OleInPlaceFrameInfo>();
            lpFrameInfo.fMDIApp = 0;
            lpFrameInfo.hwndFrame = Handle;
            lpFrameInfo.haccel = IntPtr.Zero;
            lpFrameInfo.cAccelEntries = 0;
            return SOk;
        }
        public int Scroll(NativeSize scrollExtant) => ENotimpl;
        public int OnUIDeactivate(int fUndoable) => SOk;
        public int OnInPlaceDeactivate() => SOk;
        public int DiscardUndoState() => ENotimpl;
        public int DeactivateAndUndo() => ENotimpl;
        public int OnPosRectChange(ref NativeRect lprcPosRect)
        {
            QueueViewLayout();
            return SOk;
        }
        public int GetBorder(out NativeRect lprectBorder) { lprectBorder = ClientRect(); return SOk; }
        public int RequestBorderSpace(IntPtr pborderwidths)
        {
            if (pborderwidths == IntPtr.Zero) return SOk;
            var requested = Marshal.PtrToStructure<NativeRect>(pborderwidths);
            var frame = ClientRect();
            int viewW = (frame.Right - frame.Left) - Math.Max(0, requested.Left) - Math.Max(0, requested.Right);
            int viewH = (frame.Bottom - frame.Top) - Math.Max(0, requested.Top) - Math.Max(0, requested.Bottom);
            return viewW >= 32 && viewH >= 32 ? SOk : unchecked((int)0x80004005);
        }
        public int SetBorderSpace(IntPtr pborderwidths)
        {
            int nextL = 0, nextT = 0, nextR = 0, nextB = 0;
            if (pborderwidths != IntPtr.Zero)
            {
                var requested = Marshal.PtrToStructure<NativeRect>(pborderwidths);
                nextL = Math.Max(0, requested.Left);
                nextT = Math.Max(0, requested.Top);
                nextR = Math.Max(0, requested.Right);
                nextB = Math.Max(0, requested.Bottom);
            }
            var changed = nextL != borderL || nextT != borderT || nextR != borderR || nextB != borderB;
            borderL = nextL; borderT = nextT; borderR = nextR; borderB = nextB;
            if (changed)
            {
                // #region agent log
                DebugNdjson.Write("M", "OleSite.SetBorderSpace", "border", new { borderL, borderT, borderR, borderB, nullPtr = pborderwidths == IntPtr.Zero });
                // #endregion
                ApplyViewLayout();
            }
            return SOk;
        }
        public int SetActiveObject(IntPtr pActiveObject, string? pszObjName)
        {
            activeObject = null;
            if (pActiveObject != IntPtr.Zero)
            {
                try { activeObject = (IOleInPlaceActiveObject)Marshal.GetObjectForIUnknown(pActiveObject); }
                catch { activeObject = null; }
            }
            return SOk;
        }
        public int InsertMenus(IntPtr hmenuShared, IntPtr lpMenuWidths) => ENotimpl;
        public int SetMenu(IntPtr hmenuShared, IntPtr holemenu, IntPtr hwndActiveObject) => SOk;
        public int RemoveMenus(IntPtr hmenuShared) => SOk;
        public int SetStatusText(string pszStatusText) => SOk;
        public int EnableModeless(int fEnable) => SOk;
        public int TranslateAccelerator(IntPtr lpmsg, ushort wID)
        {
            if (lpmsg == IntPtr.Zero) return ENotimpl;
            var msg = Marshal.PtrToStructure<NativeMessage>(lpmsg);
            if (msg.message == WmKeyDown && msg.wParam.ToInt64() == VkS && (GetKeyState(VkControl) & 0x8000) != 0)
            {
                // #region agent log
                DebugNdjson.Write("B", "OleSite.TranslateAccelerator", "ctrl_s", new { wID });
                // #endregion
                SaveObject();
                return SOk;
            }
            return ENotimpl;
        }
        public int QueryStatus(IntPtr pguidCmdGroup, uint cCmds, IntPtr prgCmds, IntPtr pCmdText)
        {
            if (prgCmds == IntPtr.Zero || cCmds == 0) return ENotimpl;
            var group = pguidCmdGroup == IntPtr.Zero ? Guid.Empty : Marshal.PtrToStructure<Guid>(pguidCmdGroup);
            var size = Marshal.SizeOf<OleCmd>();
            var enabledSave = false;
            var ids = new uint[Math.Min(cCmds, 12)];
            for (uint i = 0; i < cCmds; i++)
            {
                var slot = IntPtr.Add(prgCmds, (int)(i * size));
                var cmd = Marshal.PtrToStructure<OleCmd>(slot);
                if (i < ids.Length) ids[i] = cmd.CmdId;
                if (cmd.CmdId == OleCmdIdSave)
                {
                    cmd.Cmdf = OleCmdFSupported | OleCmdFEnabled;
                    Marshal.StructureToPtr(cmd, slot, false);
                    enabledSave = true;
                }
            }
            if (Interlocked.Increment(ref queryStatusCalls) <= 12 || enabledSave)
            {
                // #region agent log
                DebugNdjson.Write("B", "OleSite.QueryStatus", enabledSave ? "save_enabled" : "cmds", new { group = group.ToString("B"), cCmds, ids, enabledSave });
                // #endregion
            }
            return SOk;
        }
        public int Exec(IntPtr pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            var group = pguidCmdGroup == IntPtr.Zero ? Guid.Empty : Marshal.PtrToStructure<Guid>(pguidCmdGroup);
            if (nCmdID == OleCmdIdSave)
            {
                // #region agent log
                DebugNdjson.Write("B", "OleSite.Exec", "save", new { group = group.ToString("B"), nCmdexecopt });
                // #endregion
                return SaveObject();
            }
            return ENotimpl;
        }
        public int ActivateMe(IOleDocumentView? pViewToActivate)
        {
            Trace(Environment.CurrentManagedThreadId == ownerThread ? "site_activate_on_sta" : "site_activate_wrong_thread");
            try
            {
                documentView ??= pViewToActivate;
                CompleteActivation();
                return SOk;
            }
            catch (COMException error) when (unchecked((uint)error.HResult) == RpcECantCalloutInInputSyncCall)
            {
                Trace("activate_deferred_input_sync");
                Sta.Post(() =>
                {
                    try { CompleteActivation(); }
                    catch (Exception later) { RecordFailure("activate_deferred", later); }
                });
                return SOk;
            }
            catch (Exception error) { RecordFailure("activate_callback", error); return Marshal.GetHRForException(error); }
        }
        private void CompleteActivation()
        {
            if (viewActivated || activatingView) return;
            activatingView = true;
            try
            {
                Trace("view_activate_after_callback");
                var view = documentView;
                if (view is null && Object is IOleDocument document)
                {
                    int created = document.CreateView(this, IntPtr.Zero, 0, out view);
                    Marshal.ThrowExceptionForHR(created);
                }
                if (view is null) throw new InvalidOperationException("The document view is unavailable.");
                documentView = view;
                Trace("view_set_site");
                Marshal.ThrowExceptionForHR(view.SetInPlaceSite(this));
                Trace("view_ui_activate");
                Marshal.ThrowExceptionForHR(view.UIActivate(1));
                ApplyViewLayout();
                Trace("view_show");
                Marshal.ThrowExceptionForHR(view.Show(1));
                ApplyViewLayout();
                viewActivated = true;
            }
            catch (Exception error) { RecordFailure("activate_view", error); throw; }
            finally { activatingView = false; }
        }
    }

    [ComImport, Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleObject
    {
        void SetClientSite(IOleClientSite pClientSite);
        void GetClientSite(out IOleClientSite ppClientSite);
        void SetHostNames([MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);
        [PreserveSig] int Close(int dwSaveOption);
        void SetMoniker(int dwWhichMoniker, IntPtr pmk);
        void GetMoniker(int dwAssign, int dwWhichMoniker, out IntPtr ppmk);
        void InitFromData(IntPtr pDataObject, int fCreation, int dwReserved);
        void GetClipboardData(int dwReserved, out IntPtr ppDataObject);
        [PreserveSig] int DoVerb(int iVerb, IntPtr lpmsg, IOleClientSite pActiveSite, int lindex, IntPtr hwndParent, ref NativeRect lprcPosRect);
        void EnumVerbs(out IntPtr ppEnumOleVerb);
        void Update();
        void IsUpToDate();
        void GetUserClassID(out Guid pClsid);
        void GetUserType(int dwFormOfType, [MarshalAs(UnmanagedType.LPWStr)] out string pszUserType);
        void SetExtent(int dwDrawAspect, ref NativeSize pSizel);
        void GetExtent(int dwDrawAspect, out NativeSize pSizel);
        void Advise(IAdviseSink pAdvSink, out int pdwConnection);
        void Unadvise(int dwConnection);
        void EnumAdvise(out IntPtr ppenumAdvise);
        void GetMiscStatus(int dwAspect, out int pdwStatus);
        void SetColorScheme(IntPtr pLogpal);
    }

    [ComImport, Guid("00000118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleClientSite
    {
        [PreserveSig] int SaveObject();
        [PreserveSig] int GetMoniker(int dwAssign, int dwWhichMoniker, out IntPtr ppmk);
        [PreserveSig] int GetContainer(out IntPtr ppContainer);
        [PreserveSig] int ShowObject();
        [PreserveSig] int OnShowWindow(int fShow);
        [PreserveSig] int RequestNewObjectLayout();
    }

    [ComImport, Guid("0000010F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IAdviseSink
    {
        [PreserveSig] void OnDataChange(ref FORMATETC format, ref STGMEDIUM stgmed);
        [PreserveSig] void OnViewChange(int dwAspect, int lindex);
        [PreserveSig] void OnRename([MarshalAs(UnmanagedType.IUnknown)] object pmk);
        [PreserveSig] void OnSave();
        [PreserveSig] void OnClose();
    }

    [ComImport, Guid("00000119-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleInPlaceSite
    {
        [PreserveSig] int GetWindow(out IntPtr phwnd);
        [PreserveSig] int ContextSensitiveHelp(int fEnterMode);
        [PreserveSig] int CanInPlaceActivate();
        [PreserveSig] int OnInPlaceActivate();
        [PreserveSig] int OnUIActivate();
        [PreserveSig] int GetWindowContext(out IOleInPlaceFrame ppFrame, out IOleInPlaceUIWindow? ppDoc, out NativeRect lprcPosRect, out NativeRect lprcClipRect, ref OleInPlaceFrameInfo lpFrameInfo);
        [PreserveSig] int Scroll(NativeSize scrollExtant);
        [PreserveSig] int OnUIDeactivate(int fUndoable);
        [PreserveSig] int OnInPlaceDeactivate();
        [PreserveSig] int DiscardUndoState();
        [PreserveSig] int DeactivateAndUndo();
        [PreserveSig] int OnPosRectChange(ref NativeRect lprcPosRect);
    }

    [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleWindow
    {
        [PreserveSig] int GetWindow(out IntPtr window);
        [PreserveSig] int ContextSensitiveHelp(int enterMode);
    }

    [ComImport, Guid("0000011A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleParseDisplayName
    {
        [PreserveSig] int ParseDisplayName(IntPtr context, [MarshalAs(UnmanagedType.LPWStr)] string name, out uint eaten, out IntPtr moniker);
    }

    [ComImport, Guid("00000115-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleInPlaceUIWindow
    {
        [PreserveSig] int GetWindow(out IntPtr phwnd);
        [PreserveSig] int ContextSensitiveHelp(int fEnterMode);
        [PreserveSig] int GetBorder(out NativeRect lprectBorder);
        [PreserveSig] int RequestBorderSpace(IntPtr pborderwidths);
        [PreserveSig] int SetBorderSpace(IntPtr pborderwidths);
        [PreserveSig] int SetActiveObject(IntPtr pActiveObject, [MarshalAs(UnmanagedType.LPWStr)] string? pszObjName);
    }

    [ComImport, Guid("B722BCCB-4E68-101B-A2BC-00AA00404770"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleCommandTarget
    {
        [PreserveSig] int QueryStatus(IntPtr pguidCmdGroup, uint cCmds, IntPtr prgCmds, IntPtr pCmdText);
        [PreserveSig] int Exec(IntPtr pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct OleCmd { public uint CmdId; public uint Cmdf; }

    [ComImport, Guid("00000116-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleInPlaceFrame
    {
        [PreserveSig] int GetWindow(out IntPtr phwnd);
        [PreserveSig] int ContextSensitiveHelp(int fEnterMode);
        [PreserveSig] int GetBorder(out NativeRect lprectBorder);
        [PreserveSig] int RequestBorderSpace(IntPtr pborderwidths);
        [PreserveSig] int SetBorderSpace(IntPtr pborderwidths);
        [PreserveSig] int SetActiveObject(IntPtr pActiveObject, [MarshalAs(UnmanagedType.LPWStr)] string? pszObjName);
        [PreserveSig] int InsertMenus(IntPtr hmenuShared, IntPtr lpMenuWidths);
        [PreserveSig] int SetMenu(IntPtr hmenuShared, IntPtr holemenu, IntPtr hwndActiveObject);
        [PreserveSig] int RemoveMenus(IntPtr hmenuShared);
        [PreserveSig] int SetStatusText([MarshalAs(UnmanagedType.LPWStr)] string pszStatusText);
        [PreserveSig] int EnableModeless(int fEnable);
        [PreserveSig] int TranslateAccelerator(IntPtr lpmsg, ushort wID);
    }

    [ComImport, Guid("00000117-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleInPlaceActiveObject
    {
        [PreserveSig] int GetWindow(out IntPtr phwnd);
        [PreserveSig] int ContextSensitiveHelp(int fEnterMode);
        [PreserveSig] int TranslateAccelerator(IntPtr lpmsg);
        [PreserveSig] int OnFrameWindowActivate(int fActivate);
        [PreserveSig] int OnDocWindowActivate(int fActivate);
        [PreserveSig] int ResizeBorder(ref NativeRect prcBorder, IOleInPlaceUIWindow pUIWindow, int fFrameWindow);
        [PreserveSig] int EnableModeless(int fEnable);
    }

    [ComImport, Guid("00000113-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleInPlaceObject
    {
        [PreserveSig] int GetWindow(out IntPtr phwnd);
        [PreserveSig] int ContextSensitiveHelp(int fEnterMode);
        [PreserveSig] int InPlaceDeactivate();
        [PreserveSig] int UIDeactivate();
        [PreserveSig] int SetObjectRects(ref NativeRect lprcPosRect, ref NativeRect lprcClipRect);
        [PreserveSig] int ReactivateAndUndo();
    }

    [ComImport, Guid("0000011B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleContainer
    {
        [PreserveSig] int ParseDisplayName(IntPtr pbc, [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, out uint pchEaten, out IntPtr ppmkOut);
        [PreserveSig] int EnumObjects(uint grfFlags, out IntPtr ppenum);
        [PreserveSig] int LockContainer(int fLock);
    }

    [ComImport, Guid("B722BCC7-4E68-101B-A2BC-00AA00404770"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleDocumentSite
    {
        [PreserveSig] int ActivateMe(IOleDocumentView? pViewToActivate);
    }

    [ComImport, Guid("B722BCC5-4E68-101B-A2BC-00AA00404770"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleDocument
    {
        [PreserveSig] int CreateView(IOleInPlaceSite pIPSite, IntPtr pstm, uint dwReserved, out IOleDocumentView ppView);
        [PreserveSig] int GetDocMiscStatus(out uint pdwStatus);
        [PreserveSig] int EnumViews(out IntPtr ppEnum, out IOleDocumentView ppView);
    }

    [ComImport, Guid("B722BCC6-4E68-101B-A2BC-00AA00404770"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleDocumentView
    {
        [PreserveSig] int SetInPlaceSite(IOleInPlaceSite pIPSite);
        [PreserveSig] int GetInPlaceSite(out IOleInPlaceSite ppIPSite);
        [PreserveSig] int GetDocument([MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
        [PreserveSig] int SetRect(ref NativeRect prcView);
        [PreserveSig] int GetRect(out NativeRect prcView);
        [PreserveSig] int SetRectComplex(ref NativeRect prcView, ref NativeRect prcHScroll, ref NativeRect prcVScroll, ref NativeRect prcSizeBox);
        [PreserveSig] int Show(int fShow);
        [PreserveSig] int UIActivate(int fUIActivate);
        [PreserveSig] int Open();
        [PreserveSig] int CloseView(uint dwReserved);
        [PreserveSig] int SaveViewState(IntPtr pstm);
        [PreserveSig] int ApplyViewState(IntPtr pstm);
        [PreserveSig] int Clone(IOleInPlaceSite pIPSiteNew, out IOleDocumentView ppViewNew);
    }

    [ComImport, Guid("0000000B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IStorage
    {
        void CreateStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, int grfMode, int reserved1, int reserved2, out IStream ppstm);
        void OpenStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr reserved1, int grfMode, int reserved2, out IStream ppstm);
        void CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, int grfMode, int reserved1, int reserved2, out IStorage ppstg);
        void OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IStorage? pstgPriority, int grfMode, IntPtr snbExclude, int reserved, out IStorage ppstg);
        void CopyTo(int ciidExclude, IntPtr rgiidExclude, IntPtr snbExclude, IStorage pstgDest);
        void MoveElementTo([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IStorage pstgDest, [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName, int grfFlags);
        void Commit(int grfCommitFlags);
        void Revert();
        void EnumElements(int reserved1, IntPtr reserved2, int reserved3, out IntPtr ppenum);
        void DestroyElement([MarshalAs(UnmanagedType.LPWStr)] string pwcsName);
        void RenameElement([MarshalAs(UnmanagedType.LPWStr)] string pwcsOldName, [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName);
        void SetElementTimes([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, ref FILETIME pctime, ref FILETIME patime, ref FILETIME pmtime);
        void SetClass(ref Guid clsid);
        void SetStateBits(int grfStateBits, int grfMask);
        void Stat(out STATSTG pstatstg, int grfStatFlag);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeRect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeSize { public int Cx, Cy; }
    [StructLayout(LayoutKind.Sequential)]
    internal struct OleInPlaceFrameInfo
    {
        public int cb, fMDIApp;
        public IntPtr hwndFrame, haccel;
        public int cAccelEntries;
    }

    private static int ServerPid(IOleObject oleObject)
    {
        try
        {
            if (oleObject is IOleInPlaceObject inplace && inplace.GetWindow(out var hwnd) == SOk && hwnd != IntPtr.Zero)
            {
                GetWindowThreadProcessId(hwnd, out var pid);
                if (pid != 0) return (int)pid;
            }
        }
        catch { }
        return 0;
    }

    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000016-0000-0000-C000-000000000046")]
    private interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int callType, IntPtr task, int tickCount, IntPtr taskInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr task, int tickCount, int rejectType);
        [PreserveSig] int MessagePending(IntPtr task, int tickCount, int pendingType);
    }

    [ComVisible(true)]
    private sealed class PumpFilter : IOleMessageFilter
    {
        public int HandleInComingCall(int callType, IntPtr task, int tickCount, IntPtr taskInfo) => 0;
        public int RetryRejectedCall(IntPtr task, int tickCount, int rejectType) => 150;
        public int MessagePending(IntPtr task, int tickCount, int pendingType)
        {
            NativeWindow.PumpWait(10);
            return 2;
        }
    }

    [DllImport("ole32.dll")] private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
    [DllImport("ole32.dll")] private static extern int OleInitialize(IntPtr reserved);
    [DllImport("ole32.dll")] private static extern int OleRun([MarshalAs(UnmanagedType.IUnknown)] object pUnknown);
    [DllImport("ole32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool OleIsRunning([MarshalAs(UnmanagedType.IUnknown)] object unknown);
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)] private static extern int CreateFileMoniker(string path, out IntPtr moniker);
    [DllImport("ole32.dll")] private static extern int BindMoniker(IntPtr moniker, uint options, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object result);
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)] private static extern int CLSIDFromProgID(string progId, out Guid clsid);
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int OleCreateFromFile(ref Guid rclsid, string lpszFileName, ref Guid riid, uint renderopt,
        IntPtr pFormatEtc, IOleClientSite pClientSite, IStorage pstg, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObj);
    [DllImport("ole32.dll")] private static extern int OleSetContainedObject([MarshalAs(UnmanagedType.IUnknown)] object pUnknown, int fContained);
    [DllImport("ole32.dll")] private static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, int fDeleteOnRelease, out IntPtr ppLkbyt);
    [DllImport("ole32.dll")] private static extern int StgCreateDocfileOnILockBytes(IntPtr plkbyt, int grfMode, int reserved, out IStorage ppstgOpen);
    private static HashSet<uint> ProtectedPids(IntPtr host)
    {
        var pids = new HashSet<uint> { (uint)Environment.ProcessId };
        var parent = NativeWindow.GetParent(host);
        if (parent != IntPtr.Zero)
        {
            NativeWindow.GetWindowThreadProcessId(parent, out var parentPid);
            if (parentPid != 0) pids.Add(parentPid);
        }
        NativeWindow.GetWindowThreadProcessId(host, out var hostPid);
        if (hostPid != 0) pids.Add(hostPid);
        return pids;
    }

    private static bool IsOwnContainer(IntPtr window)
    {
        var name = new StringBuilder(256);
        NativeWindow.GetClassName(window, name, name.Capacity);
        return name.ToString().StartsWith("OieOleHost_", StringComparison.Ordinal);
    }

    private static List<IntPtr> SnapshotVisibleTopLevel()
    {
        var frames = new List<IntPtr>();
        NativeWindow.EnumProc callback = (window, _) =>
        {
            if (NativeWindow.IsWindowVisible(window) && !IsOwnContainer(window))
                frames.Add(window);
            return true;
        };
        NativeWindow.EnumWindows(callback, IntPtr.Zero);
        GC.KeepAlive(callback);
        return frames;
    }

    private static void HideNewFrames(List<IntPtr> existing, IntPtr host, HashSet<uint> protectPids)
    {
        var hidden = 0;
        var skipped = 0;
        var hiddenClasses = new List<string>();
        NativeWindow.EnumProc callback = (window, _) =>
        {
            if (window == host || NativeWindow.IsChild(host, window) || existing.Contains(window) || IsOwnContainer(window))
                return true;
            var owned = NativeWindow.GetWindow(window, 4);
            if (owned == host || (owned != IntPtr.Zero && NativeWindow.IsChild(host, owned)))
            {
                skipped++;
                return true;
            }
            NativeWindow.GetWindowThreadProcessId(window, out var pid);
            if (pid == 0 || protectPids.Contains(pid))
            {
                skipped++;
                return true;
            }
            var className = WebviewSurface.ClassName(window);
            if (className is "#32770" or "NUIDialog"
                || className.StartsWith("bosa_sdm", StringComparison.OrdinalIgnoreCase)
                || className.Contains("Dialog", StringComparison.OrdinalIgnoreCase))
            {
                skipped++;
                return true;
            }
            if (!NativeWindow.IsWindowVisible(window)) return true;
            if (hiddenClasses.Count < 8) hiddenClasses.Add(WebviewSurface.ClassName(window));
            ShowWindow(window, SwHide);
            hidden++;
            return true;
        };
        NativeWindow.EnumWindows(callback, IntPtr.Zero);
        GC.KeepAlive(callback);
        // #region agent log
        DebugNdjson.Write("C", "OleDocumentHost.HideNewFrames", "hide_new_frames", new { existing = existing.Count, hidden, skipped, protect = protectPids.Count, hiddenClasses });
        // #endregion
    }

    private static void SweepNewFrames(List<IntPtr> existing, IntPtr host, HashSet<uint> protectPids)
    {
        HideNewFrames(existing, host, protectPids);
        NativeWindow.PumpWait(120);
        HideNewFrames(existing, host, protectPids);
    }

    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("user32.dll")] private static extern short GetKeyState(int virtualKey);
    private const uint QsAllInput = 0x04FF;
    [DllImport("user32.dll")] private static extern bool PeekMessage(out NativeMessage msg, IntPtr hwnd, uint min, uint max, uint remove);
    [DllImport("user32.dll")] private static extern bool TranslateMessage(ref NativeMessage msg);
    [DllImport("user32.dll")] private static extern IntPtr DispatchMessage(ref NativeMessage msg);
    [DllImport("user32.dll")] private static extern uint MsgWaitForMultipleObjects(uint count, IntPtr[] handles, bool waitAll, uint milliseconds, uint wakeMask);
    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage { public IntPtr hwnd; public uint message; public IntPtr wParam, lParam; public uint time; public int x, y; }
}

internal static class NativeWindow
{
    internal delegate IntPtr WndProc(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct WndClassEx
    {
        public uint cbSize, style;
        public WndProc lpfnWndProc;
        public int cbClsExtra, cbWndExtra;
        public IntPtr hInstance, hIcon, hCursor, hbrBackground;
        public string lpszMenuName, lpszClassName;
        public IntPtr hIconSm;
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern ushort RegisterClassEx(ref WndClassEx windowClass);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern IntPtr CreateWindowEx(uint exStyle, string className, string windowName, uint style,
        int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);
    [DllImport("user32.dll")] internal static extern bool DestroyWindow(IntPtr hwnd);
    [DllImport("user32.dll")] internal static extern bool IsChild(IntPtr parent, IntPtr child);
    internal delegate bool EnumChildProc(IntPtr window, IntPtr data);
    internal delegate bool EnumProc(IntPtr window, IntPtr data);
    [DllImport("user32.dll")] internal static extern bool EnumChildWindows(IntPtr parent, EnumChildProc callback, IntPtr data);
    [DllImport("user32.dll")] internal static extern bool EnumWindows(EnumProc callback, IntPtr data);
    [DllImport("user32.dll")] internal static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetClassName(IntPtr window, StringBuilder name, int size);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetWindowText(IntPtr window, StringBuilder text, int maxCount);
    [DllImport("user32.dll")] internal static extern IntPtr SendMessage(IntPtr window, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] internal static extern bool MoveWindow(IntPtr hwnd, int x, int y, int width, int height, bool repaint);
    [DllImport("user32.dll")] internal static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int width, int height, uint flags);
    [DllImport("user32.dll")] internal static extern bool GetClientRect(IntPtr hwnd, out OleDocumentHost.NativeRect rect);
    [DllImport("user32.dll")] internal static extern bool GetWindowRect(IntPtr hwnd, out OleDocumentHost.NativeRect rect);
    [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] internal static extern IntPtr GetParent(IntPtr hwnd);
    [DllImport("user32.dll")] internal static extern IntPtr GetWindow(IntPtr hwnd, uint command);
    [DllImport("user32.dll")] internal static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] internal static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll", SetLastError = true)] internal static extern IntPtr SetParent(IntPtr child, IntPtr parent);
    [DllImport("user32.dll")] internal static extern int GetWindowLong(IntPtr hwnd, int index);
    [DllImport("user32.dll")] internal static extern int SetWindowLong(IntPtr hwnd, int index, int value);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] internal static extern IntPtr GetWindowLongPtr(IntPtr hwnd, int index);
    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")] internal static extern IntPtr SetWindowLongPtr(IntPtr hwnd, int index, IntPtr value);
    [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")] internal static extern bool ClientToScreen(IntPtr hwnd, ref Point point);
    [DllImport("user32.dll")] private static extern bool SetProcessDpiAwarenessContext(IntPtr value);
    [StructLayout(LayoutKind.Sequential)]
    internal struct Point { public int X, Y; }
    internal static void TryEnableDpi()
    {
        try { SetProcessDpiAwarenessContext(new IntPtr(-4)); } catch { }
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern IntPtr DefWindowProc(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern bool UnregisterClass(string className, IntPtr instance);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] internal static extern IntPtr GetModuleHandle(string? name);

    private const int GwlStyle = -16;
    private const int GwlExStyle = -20;
    private const int GwlpHwndParent = -8;
    private const int WsChild = 0x40000000;
    private const int WsPopupStyle = unchecked((int)0x80000000);
    private const int WsClipSiblings = 0x04000000;
    private const int WsClipChildrenStyle = 0x02000000;
    private const int WsExToolwindow = 0x00000080;
    private const int WsExNoActivate = 0x08000000;
    private const uint SwpNoMove = 0x0002;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoZOrder = 0x0004;
    private const uint SwpFrameChanged = 0x0020;
    private const uint SwpNoActivate = 0x0010;

    internal static void PumpWait(int milliseconds)
    {
        var until = Environment.TickCount64 + Math.Max(0, milliseconds);
        while (Environment.TickCount64 < until)
        {
            NativeMessage msg;
            while (PeekMessage(out msg, IntPtr.Zero, 0, 0, 1))
            {
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }
            MsgWaitForMultipleObjects(0, Array.Empty<IntPtr>(), false, 50, 0x04FF);
        }
    }

    internal readonly record struct HostDesc(int Style, bool Popup, bool Child, long Parent, long HwndParent, bool IsChildOfOwner);

    internal static HostDesc DescribeHost(IntPtr hwnd, IntPtr owner)
    {
        if (hwnd == IntPtr.Zero || !IsWindow(hwnd))
            return new HostDesc(0, false, false, 0, 0, false);
        var style = GetWindowLong(hwnd, GwlStyle);
        return new HostDesc(
            style,
            (style & WsPopupStyle) != 0,
            (style & WsChild) != 0,
            GetParent(hwnd).ToInt64(),
            GetWindowLongPtr(hwnd, GwlpHwndParent).ToInt64(),
            owner != IntPtr.Zero && IsChild(owner, hwnd));
    }

    internal static void OwnPopup(IntPtr popup, IntPtr owner)
    {
        if (popup == IntPtr.Zero || owner == IntPtr.Zero) throw new InvalidOperationException("Office layout failed.");
        if (GetParent(popup) != IntPtr.Zero) SetParent(popup, IntPtr.Zero);
        var style = GetWindowLong(popup, GwlStyle);
        SetWindowLong(popup, GwlStyle, (style & ~WsChild) | WsPopupStyle | WsClipSiblings | WsClipChildrenStyle);
        var ex = GetWindowLong(popup, GwlExStyle);
        SetWindowLong(popup, GwlExStyle, (ex | WsExToolwindow) & ~WsExNoActivate);
        SetWindowPos(popup, IntPtr.Zero, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoZOrder | SwpNoActivate | SwpFrameChanged);
        SetWindowLongPtr(popup, GwlpHwndParent, owner);
    }

    internal static void AdoptChild(IntPtr child, IntPtr parent)
    {
        if (child == IntPtr.Zero || parent == IntPtr.Zero) throw new InvalidOperationException("Office layout failed.");
        var style = GetWindowLong(child, GwlStyle);
        SetWindowLong(child, GwlStyle, (style & ~WsPopupStyle) | WsChild | WsClipSiblings | WsClipChildrenStyle);
        var ex = GetWindowLong(child, GwlExStyle);
        SetWindowLong(child, GwlExStyle, ex & ~WsExToolwindow & ~WsExNoActivate);
        SetWindowPos(child, IntPtr.Zero, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoZOrder | SwpNoActivate | SwpFrameChanged);
        SetParent(child, parent);
        if (GetParent(child) != parent) throw new InvalidOperationException("Office layout failed.");
    }

    [DllImport("user32.dll")] private static extern bool PeekMessage(out NativeMessage msg, IntPtr hwnd, uint min, uint max, uint remove);
    [DllImport("user32.dll")] private static extern bool TranslateMessage(ref NativeMessage msg);
    [DllImport("user32.dll")] private static extern IntPtr DispatchMessage(ref NativeMessage msg);
    [DllImport("user32.dll")] private static extern uint MsgWaitForMultipleObjects(uint count, IntPtr[] handles, bool waitAll, uint milliseconds, uint wakeMask);
    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage { public IntPtr hwnd; public uint message; public IntPtr wParam, lParam; public uint time; public int x, y; }
}
