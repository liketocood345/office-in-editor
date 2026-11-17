using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

[assembly: ComVisible(false)]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("OfficeInEditor.Addin.Smoke")]

namespace OfficeInEditor.OfficeAddin
{
    [ComVisible(true), Guid("6B4072D4-3EB1-4CAB-995D-4F922B19D4B1")]
    [ProgId("OfficeInEditor.Addin"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class OfficeAddin : IDTExtensibility2, IRibbonExtensibility
    {
        private sealed class ChatTarget
        {
            internal string Id, Label, Epoch;
        }
        private readonly CoreClient core = new CoreClient();
        private readonly string instanceId = Guid.NewGuid().ToString("N");
        private object connectedApplication;
        private OfficeDispatcher dispatcher;
        private OfficeCallbackServer callback;
        private SelectionCapture selection;
        private LocalAutoSave autoSave;
        private readonly WindowHost windows = new WindowHost();
        private string applicationKind = "unsupported";
        private string pipeName;
        private object ribbon;
        private readonly ChatRouteCache routes = new ChatRouteCache();
        private List<ChatTarget> targets = new List<ChatTarget>();
        private string status = "";
        private int refreshingBindings;
        private int busy;
        private volatile bool disconnected;
        private const int ConnectModeStartup = 1;

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            try
            {
                TraceLoad("OnConnection mode=" + connectMode);
                connectedApplication = application;
                // Official shared add-in: COMAddIns.Item(ProgId).Object is this IDispatch,
                // so ribbon onAction and host automation reach RefreshChats / AddToChat.
                try
                {
                    ((dynamic)addInInst).Object = this;
                    TraceLoad("COMAddIn.Object assigned");
                }
                catch { TraceLoad("COMAddIn.Object failed"); }
                // Official Shared Add-in pattern: OnStartupComplete is not raised when the
                // user connects the add-in after Word has already started.
                if (connectMode != ConnectModeStartup) OnStartupComplete(ref custom);
            }
            catch { status = "Office integration is unavailable. No document changes were made."; }
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                TraceLoad("OnStartupComplete");
                StartBridge();
            }
            catch { status = "Office integration is unavailable. No document changes were made."; }
        }

        private void StartBridge()
        {
            if (disconnected || connectedApplication == null || dispatcher != null) return;
            if (applicationKind == "unsupported")
            {
                try { applicationKind = DetectApplication(Convert.ToString(((dynamic)connectedApplication).Name)); }
                catch { return; }
            }
            if (applicationKind == "unsupported") return;
            var application = connectedApplication;
            dispatcher = new OfficeDispatcher();
            selection = new SelectionCapture(application, applicationKind, instanceId);
            autoSave = new LocalAutoSave(application, applicationKind, dispatcher, text => { status = text; Invalidate(); }, documents: selection.Documents);
            autoSave.Configure(true, 60);
            var operations = new DocumentOperations(application, applicationKind, selection.Documents);
            pipeName = "office-in-editor-office-" + Process.GetCurrentProcess().Id + "-" + instanceId;
            callback = new OfficeCallbackServer(pipeName, dispatcher, (method, parameters) =>
            {
                if (method == "captureSelection") return selection.Capture(Json.Text(parameters, "documentId"));
                if (method == "bindChatTarget") return BindChatTarget(parameters);
                if (method == "refreshDocumentBindings")
                {
                    if (Json.Get(parameters, "bindings") != null)
                    {
                        object described = UpdateDocumentBindings(parameters);
                        Task.Run(async () => { try { await RegisterDocumentBindingsAsync(described).ConfigureAwait(false); } catch { } });
                        return described;
                    }
                    BeginRefreshDocumentBindings(); return Json.Object("refreshing", true);
                }
                if (method == "describeWindow") return windows.DescribeWindow(application);
                if (method == "attachWindow") return windows.Attach(parameters, application);
                if (method == "layoutWindow") return windows.Layout(parameters, application);
                if (method == "detachWindow") return windows.Detach(application);
                if (method == "inspectDocument")
                {
                    var info = Json.Map(operations.Dispatch(method, parameters));
                    var embedded = selection.Documents.Embedded.Find(selection.Documents.Resolve(Json.Text(info, "documentId")));
                    info["localAutoSave"] = Json.Object("periodic", embedded?.Periodic ?? autoSave.Periodic,
                        "intervalSeconds", embedded?.Seconds ?? autoSave.Seconds,
                        "scope", embedded == null ? "active-local-document" : "bound-embedded-document", "clickSave", embedded == null);
                    return info;
                }
                return operations.Dispatch(method, parameters);
            });
            callback.Start();
            Task.Run(async () =>
            {
                try
                {
                    await RegisterAsync(false).ConfigureAwait(false);
                    await RefreshDocumentBindingsAsync().ConfigureAwait(false);
                    await LoadTargetsAsync(false).ConfigureAwait(false);
                }
                catch { }
            });
        }

        private EmbeddedDocumentBindings.Binding ActiveBinding()
        { try { return selection?.Documents.Embedded.Find(selection.Documents.Active()); } catch { return null; } }
        public bool GetSavePeriodic(object control) => ActiveBinding()?.Periodic ?? autoSave?.Periodic ?? false;
        public bool GetSaveOnClose(object control) => ActiveBinding()?.SaveOnClose ?? true;
        public bool GetAckAlerts(object control) => ActiveBinding()?.AckAlerts ?? true;
        public string GetSaveInterval(object control) => (ActiveBinding()?.Seconds ?? autoSave?.Seconds ?? 60).ToString();
        public void SetSaveOnClose(object control, bool enabled)
        {
            var binding = ActiveBinding();
            if (binding != null) ConfigureEmbeddedSaveOnClose(binding, enabled);
            Invalidate();
        }
        public void SetAckAlerts(object control, bool enabled)
        {
            var binding = ActiveBinding();
            if (binding != null) ConfigureEmbeddedAckAlerts(binding, enabled);
            Invalidate();
        }
        public void SetSavePeriodic(object control, bool enabled)
        {
            var binding = ActiveBinding();
            if (binding == null) autoSave?.Configure(enabled, autoSave.Seconds);
            else ConfigureEmbeddedSave(binding, enabled, binding.Seconds);
            Invalidate();
        }
        public void SetSaveInterval(object control, string text)
        {
            if (int.TryParse(text, out int seconds) && seconds >= 5 && seconds <= 3600)
            {
                var binding = ActiveBinding();
                if (binding == null) autoSave?.Configure(autoSave.Periodic, seconds);
                else ConfigureEmbeddedSave(binding, binding.Periodic, seconds);
            }
            Invalidate();
        }

        internal static string DetectApplication(string name)
        {
            if (name == null) return "unsupported";
            if (name.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0) return "word";
            if (name.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0) return "excel";
            if (name.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0) return "powerpoint";
            return "unsupported";
        }

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                TraceLoad("GetCustomUI ribbonId=" + ribbonId);
                string kind = RibbonMarkup.FromRibbonId(ribbonId);
                if (kind != "unsupported") applicationKind = kind;
                return RibbonMarkup.Build(applicationKind);
            }
            catch { return RibbonMarkup.Build(applicationKind); }
        }
        public void OnRibbonLoad(object ribbonUi)
        {
            try { TraceLoad("OnRibbonLoad"); ribbon = ribbonUi; Invalidate(); } catch { }
        }
        public bool GetSendEnabled(object control)
        {
            try
            {
                return !disconnected && selection != null && Volatile.Read(ref busy) == 0;
            }
            catch { return false; }
        }
        public bool GetRefreshEnabled(object control)
        {
            try { return !disconnected && dispatcher != null && Volatile.Read(ref busy) == 0; }
            catch { return false; }
        }
        public bool GetFileSaveEnabled(object control)
        {
            try { return !disconnected; }
            catch { return true; }
        }
        public void RepurposeFileSave(object control, ref bool cancelDefault)
        {
            cancelDefault = true;
            try { BeginRibbonSave(); } catch { }
        }
        public string GetStatusLabel(object control) { try { return status; } catch { return ""; } }
        public string GetTargetLabel(object control)
        {
            try { var route = CurrentRoute(); return route.Id == null ? "Choose chat" : "Chat: " + route.Label; }
            catch { return "Choose chat"; }
        }

        public string GetChatMenu(object control)
        {
            try { if (dispatcher != null && !disconnected && Volatile.Read(ref busy) == 0) BeginRefreshChats(); return BuildChatMenu(); }
            catch { return "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"></menu>"; }
        }
        private string BuildChatMenu()
        {
            var xml = new StringBuilder("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
            xml.Append("<button id=\"OieRefreshFromMenu\" label=\"Refresh chats\" onAction=\"RefreshChats\" getEnabled=\"GetRefreshEnabled\"/>");
            for (int i = 0; i < targets.Count; i++)
            {
                var target = targets[i];
                xml.Append("<button id=\"OieTarget").Append(i).Append("\" label=\"")
                    .Append(RibbonMarkup.Escape(target.Label)).Append("\" tag=\"")
                    .Append(RibbonMarkup.Escape(Json.Encode(Json.Object("id", target.Id, "epoch", target.Epoch)))).Append("\" onAction=\"ChooseChat\"/>");
            }
            xml.Append("<button id=\"OieClearTarget\" label=\"Clear selected chat\" onAction=\"ClearChat\"/>");
            return xml.Append("</menu>").ToString();
        }

        public void ChooseChat(object control)
        {
            if (disconnected || Volatile.Read(ref busy) != 0) return;
            try
            {
                object choice = Json.Serializer().DeserializeObject(ReadControlTag(control) ?? "{}");
                string id = Json.Text(choice, "id"), epoch = Json.Text(choice, "epoch");
                var selected = targets.Find(t => t.Id == id && t.Epoch == epoch);
                var route = CurrentRoute();
                route.Id = selected?.Id; route.Label = selected?.Label; route.Epoch = selected?.Epoch; route.Manual = true;
                status = "";
                TraceLoad("ChooseChat hit=" + (selected != null));
                Invalidate();
            }
            catch { status = ""; Invalidate(); }
        }
        private static string ReadControlTag(object control)
        {
            if (control == null) return null;
            if (control is string text) return text;
            try { return Convert.ToString(((dynamic)control).Tag); }
            catch { return Convert.ToString(control); }
        }
        public void ClearChat(object control)
        {
            try
            {
                if (Volatile.Read(ref busy) != 0) return;
                var route = CurrentRoute(); route.Id = null; route.Label = null; route.Epoch = null; route.Manual = true; status = ""; Invalidate();
            }
            catch { }
        }

        public void RefreshChats(object control)
        {
            try { TraceLoad("RefreshChats"); BeginRefreshChats(); } catch { Interlocked.Exchange(ref busy, 0); }
        }
        private void BeginRefreshChats()
        {
            if (disconnected || dispatcher == null || Interlocked.CompareExchange(ref busy, 1, 0) != 0) return;
            Task.Run(async () =>
            {
                try { await LoadTargetsAsync(true).ConfigureAwait(false); }
                catch (Exception error) { await ShowErrorAsync(error).ConfigureAwait(false); }
                finally { Interlocked.Exchange(ref busy, 0); await OnUiAsync(Invalidate).ConfigureAwait(false); }
            });
        }
        private async Task LoadTargetsAsync(bool ensure)
        {
            await RegisterAsync(ensure).ConfigureAwait(false);
            object listed = await core.CallAsync("chat.list", Json.Object()).ConfigureAwait(false);
            var list = ParseTargets(listed);
            object fallback = await core.CallAsync("chat.defaultTarget", Json.Object()).ConfigureAwait(false);
            await OnUiAsync(() =>
            {
                targets = list;
                var route = CurrentRoute();
                if (route.Id != null && !targets.Exists(t => t.Id == route.Id && t.Epoch == route.Epoch))
                { route.Id = null; route.Label = null; route.Epoch = null; }
                if (!route.Manual && ActiveBinding() == null)
                {
                    var next = TargetFrom(fallback);
                    route.Id = next?.Id; route.Label = next?.Label; route.Epoch = next?.Epoch;
                }
                status = "";
            }).ConfigureAwait(false);
        }
        private static ChatTarget TargetFrom(object value)
        {
            string id = Json.Text(value, "id"), label = Json.Text(value, "label");
            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(label)) return null;
            return new ChatTarget { Id = id, Label = label, Epoch = Convert.ToString(Json.Get(value, "epoch")) };
        }
        private object BindChatTarget(object parameters)
        {
            string documentId = Json.Text(parameters, "documentId"), sessionId = Json.Text(parameters, "sessionId");
            object document = selection.Documents.Resolve(documentId);
            var embedded = selection.Documents.Embedded.Find(document);
            if (embedded == null || embedded.SessionId != sessionId) return Json.Object("bound", false);
            var target = TargetFrom(parameters);
            if (target == null || !long.TryParse(Convert.ToString(Json.Get(parameters, "routeVersion")), out long version))
                return Json.Object("bound", false);
            if (!routes.Bind(documentId, sessionId, target.Id, target.Label, target.Epoch, version, Json.Text(parameters, "brokerEpoch")))
                return Json.Object("bound", false);
            targets.RemoveAll(item => item.Id == target.Id);
            targets.Add(target);
            status = "";
            Invalidate();
            return Json.Object("bound", true);
        }

        public void AddToChat(object control)
        {
            try
            {
                TraceLoad("AddToChat");
                BeginAddToChat();
            }
            catch { Interlocked.Exchange(ref busy, 0); }
        }
        private void BeginRibbonSave()
        {
            if (disconnected || selection == null || Interlocked.CompareExchange(ref busy, 1, 0) != 0) return;
            string documentId = null, sessionId = null;
            object document;
            try
            {
                document = selection.Documents.Active();
                documentId = selection.Documents.Bind(document);
                sessionId = selection.Documents.Embedded.Find(document)?.SessionId;
            }
            catch
            {
                Interlocked.Exchange(ref busy, 0);
                return;
            }
            if (string.IsNullOrEmpty(sessionId))
            {
                try
                {
                    string path = Convert.ToString(((dynamic)document).FullName);
                    LocalAutoSave.ClickSave(connectedApplication, document, path, applicationKind);
                }
                catch { }
                Interlocked.Exchange(ref busy, 0);
                Invalidate();
                return;
            }
            Task.Run(async () =>
            {
                try
                {
                    await core.CallAsync("office.save_window", Json.Object("officeId", sessionId, "documentId", documentId)).ConfigureAwait(false);
                }
                catch { }
                finally { Interlocked.Exchange(ref busy, 0); await OnUiAsync(Invalidate).ConfigureAwait(false); }
            });
        }
        private void BeginAddToChat()
        {
            if (disconnected || selection == null || Interlocked.CompareExchange(ref busy, 1, 0) != 0) return;
            string documentId, sessionId, targetId, targetEpoch;
            bool manual;
            long routeVersion;
            try
            {
                object document = selection.Documents.Active();
                documentId = selection.Documents.Bind(document);
                sessionId = selection.Documents.Embedded.Find(document)?.SessionId;
                var route = routes.For(documentId);
                manual = route.Manual; targetId = route.Id; targetEpoch = route.Epoch; routeVersion = route.Version;
                if (manual && targetId == null) { Interlocked.Exchange(ref busy, 0); return; }
            }
            catch
            {
                Interlocked.Exchange(ref busy, 0);
                return;
            }
            string requestId = Guid.NewGuid().ToString("N");
            Task.Run(async () =>
            {
                try
                {
                    await RegisterAsync(false).ConfigureAwait(false);
                    if (string.IsNullOrEmpty(sessionId))
                    {
                        try { await RefreshDocumentBindingsAsync().ConfigureAwait(false); }
                        catch { }
                        try
                        {
                            await OnUiAsync(() =>
                            {
                                object document = selection.Documents.Resolve(documentId);
                                sessionId = selection.Documents.Embedded.Find(document)?.SessionId;
                            }).ConfigureAwait(false);
                        }
                        catch { }
                    }
                    await core.CallAsync("office.document.bind", Json.Object("officeId", instanceId,
                        "documentId", documentId, "sessionId", sessionId)).ConfigureAwait(false);
                    // Resolve the destination first in Core; captureSelection is then
                    // called back with this exact documentId, and Core rechecks the route.
                    object delivered = await core.CallAsync("reference.captureAndDeliver", Json.Object("officeId", instanceId,
                        "documentId", documentId, "sessionId", sessionId, "routeMode", manual ? "manual" : "auto",
                        "targetId", manual ? targetId : null, "targetEpoch", manual ? targetEpoch : null,
                        "routeVersion", routeVersion, "requestId", requestId)).ConfigureAwait(false);
                    bool queued = Json.IsTrue(Json.Get(delivered, "queued"));
                    string reason = Json.Text(delivered, "reason");
                    await OnUiAsync(() =>
                    {
                        status = queued ? "已投递到聊天输入框"
                            : reason == "target_unavailable" ? "没有可用聊天"
                            : reason == "route_changed" ? "聊天目标已变化"
                            : "未投递";
                    }).ConfigureAwait(false);
                }
                catch (Exception error)
                {
                    await OnUiAsync(() => status = "未投递").ConfigureAwait(false);
                    if (manual) await ShowErrorAsync(error).ConfigureAwait(false);
                }
                finally { Interlocked.Exchange(ref busy, 0); await OnUiAsync(Invalidate).ConfigureAwait(false); }
            });
        }

        private ChatRouteCache.Route CurrentRoute()
        {
            try { return routes.For(selection.Documents.Bind(selection.Documents.Active())); }
            catch { return routes.For(null); }
        }
        private object UpdateDocumentBindings(object value)
        {
            var mapped = applicationKind == "word" ? selection.Documents.Embedded.Update(value) : new List<EmbeddedDocumentBindings.Binding>();
            var result = new List<object>();
            foreach (var item in mapped)
                result.Add(Json.Object("documentId", selection.Documents.Bind(item.Document), "sessionId", item.SessionId));
            return Json.Object("bindings", result);
        }
        private void BeginRefreshDocumentBindings()
        {
            if (disconnected || Interlocked.CompareExchange(ref refreshingBindings, 1, 0) != 0) return;
            Task.Run(async () =>
            {
                try { await RefreshDocumentBindingsAsync().ConfigureAwait(false); }
                catch { }
                finally { Interlocked.Exchange(ref refreshingBindings, 0); }
            });
        }
        private async Task RefreshDocumentBindingsAsync()
        {
            object value = await core.CallAsync("office.source.resolve", Json.Object("officeId", instanceId)).ConfigureAwait(false);
            object described = null;
            await OnUiAsync(() => described = UpdateDocumentBindings(value)).ConfigureAwait(false);
            await RegisterDocumentBindingsAsync(described).ConfigureAwait(false);
        }
        private async Task RegisterDocumentBindingsAsync(object described)
        {
            foreach (object item in Json.Array(Json.Get(described, "bindings")))
                await core.CallAsync("office.document.bind", Json.Object("officeId", instanceId,
                    "documentId", Json.Text(item, "documentId"), "sessionId", Json.Text(item, "sessionId"))).ConfigureAwait(false);
        }
        private void ConfigureEmbeddedSave(EmbeddedDocumentBindings.Binding binding, bool enabled, int seconds)
        {
            // UI callbacks must return before Core invokes its own OLE object.
            string documentId = selection.Documents.Bind(binding.Document), sessionId = binding.SessionId;
            Task.Run(async () =>
            {
                try
                {
                    await core.CallAsync("office.configure_save_window", Json.Object("officeId", sessionId,
                        "documentId", documentId, "enabled", enabled, "seconds", seconds, "saveOnClose", binding.SaveOnClose)).ConfigureAwait(false);
                    await OnUiAsync(() =>
                    {
                        var current = selection.Documents.Embedded.Find(selection.Documents.Resolve(documentId));
                        if (current != null && current.SessionId == sessionId) { current.Periodic = enabled; current.Seconds = seconds; }
                        status = "";
                    }).ConfigureAwait(false);
                }
                catch (Exception error) { await ShowErrorAsync(error).ConfigureAwait(false); }
            });
        }
        private void ConfigureEmbeddedSaveOnClose(EmbeddedDocumentBindings.Binding binding, bool enabled)
        {
            string documentId = selection.Documents.Bind(binding.Document), sessionId = binding.SessionId;
            Task.Run(async () =>
            {
                try
                {
                    await core.CallAsync("office.configure_save_window", Json.Object("officeId", sessionId,
                        "documentId", documentId, "saveOnClose", enabled)).ConfigureAwait(false);
                    await OnUiAsync(() =>
                    {
                        var current = selection.Documents.Embedded.Find(selection.Documents.Resolve(documentId));
                        if (current != null && current.SessionId == sessionId) current.SaveOnClose = enabled;
                        status = "";
                    }).ConfigureAwait(false);
                }
                catch (Exception error) { await ShowErrorAsync(error).ConfigureAwait(false); }
            });
        }
        private void ConfigureEmbeddedAckAlerts(EmbeddedDocumentBindings.Binding binding, bool enabled)
        {
            string documentId = selection.Documents.Bind(binding.Document), sessionId = binding.SessionId;
            Task.Run(async () =>
            {
                try
                {
                    await core.CallAsync("office.configure_save_window", Json.Object("officeId", sessionId,
                        "documentId", documentId, "ackAlerts", enabled)).ConfigureAwait(false);
                    await OnUiAsync(() =>
                    {
                        var current = selection.Documents.Embedded.Find(selection.Documents.Resolve(documentId));
                        if (current != null && current.SessionId == sessionId) current.AckAlerts = enabled;
                        status = "";
                    }).ConfigureAwait(false);
                }
                catch (Exception error) { await ShowErrorAsync(error).ConfigureAwait(false); }
            });
        }

        private async Task RegisterAsync(bool ensure)
        {
            if (disconnected || pipeName == null) throw new BridgeException("OFFICE_DISCONNECTED");
            await core.CallAsync("office.register", Json.Object("officeId", instanceId,
                "application", applicationKind, "pipeName", pipeName, "pid", Process.GetCurrentProcess().Id), ensure)
                .ConfigureAwait(false);
            if (disconnected)
            {
                // A connection callback may finish after shutdown's initial unregister request.
                // Remove that late registration without blocking Office or replaying any selection.
                try { await core.CallAsync("office.unregister", Json.Object("officeId", instanceId)).ConfigureAwait(false); }
                catch { }
                throw new BridgeException("OFFICE_DISCONNECTED");
            }
        }
        private static List<ChatTarget> ParseTargets(object result)
        {
            object values = Json.Get(result, "targets") ?? Json.Get(result, "chats") ?? result;
            var parsed = new List<ChatTarget>();
            foreach (object item in Json.Array(values))
            {
                string id = Json.Text(item, "id"), label = Json.Text(item, "label");
                if (string.IsNullOrEmpty(id) || id.Length > 128 || string.IsNullOrEmpty(label)) continue;
                if (label.Length > 120) label = label.Substring(0, 117) + "..."; // display-only title, not selection data
                var clean = new StringBuilder();
                foreach (char c in label) if (!char.IsControl(c)) clean.Append(c);
                // Titles are display-only. Stable target identity remains in Id/Tag for routing,
                // and a short identity suffix distinguishes chats with identical titles.
                parsed.Add(new ChatTarget { Id = id, Label = clean.ToString(), Epoch = Convert.ToString(Json.Get(item, "epoch")) });
                if (parsed.Count > 256) throw new BridgeException("TOO_MANY_CHAT_TARGETS");
            }
            return parsed;
        }
        private async Task OnUiAsync(Action action)
        {
            if (disconnected || dispatcher == null) return;
            using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(3)))
                try { await dispatcher.InvokeAsync(() => { if (!disconnected) { action(); Invalidate(); } return true; },
                    timeout.Token).ConfigureAwait(false); } catch { }
        }
        private Task ShowErrorAsync(Exception error) => OnUiAsync(() => status = ErrorStatus(error));
        private static string ErrorStatus(Exception error)
        {
            if (error is BridgeException bridge)
                return "Request stopped: " + CoreClient.SafeError(bridge.Code) + ". No automatic retry.";
            if (error is TimeoutException || error is OperationCanceledException)
                return "Request outcome unknown or Office busy. No automatic retry.";
            return "Office bridge unavailable. Refresh chats when ready. No automatic retry.";
        }
        private void Invalidate()
        {
            if (ribbon == null || disconnected) return;
            try { ((dynamic)ribbon).Invalidate(); } catch { }
        }

        public void OnDisconnection(int removeMode, ref Array custom) => Disconnect();
        public void OnBeginShutdown(ref Array custom) => Disconnect();
        public void OnAddInsUpdate(ref Array custom) { }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern void OutputDebugString(string message);
        private static void TraceLoad(string step)
        {
            try
            {
                OutputDebugString("OfficeInEditor.Addin " + step);
                if (Environment.GetEnvironmentVariable("OFFICE_IN_EDITOR_TRACE") != "1") return;
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs", "OfficeInEditor", "current", "office", "load-trace.txt");
                string line = DateTime.UtcNow.ToString("o") + " " + step + "\r\n";
                if (File.Exists(path) && new FileInfo(path).Length > 65536)
                    File.WriteAllText(path, line);
                else File.AppendAllText(path, line);
            }
            catch { }
        }
        private void Disconnect()
        {
            if (disconnected) return;
            disconnected = true;
            autoSave?.Dispose();
            try { windows.RestoreAll(); } catch { }
            callback?.Dispose();
            dispatcher?.Dispose();
            routes.Clear(); targets.Clear(); selection?.Documents.Clear(); selection = null; ribbon = null;
            // No Quit/Close/Save, process termination, COM final-release, or global UI reset.
            Task.Run(async () => { try { await core.CallAsync("office.unregister", Json.Object("officeId", instanceId))
                        .ConfigureAwait(false); } catch { } });
        }
    }
}
