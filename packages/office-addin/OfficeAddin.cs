using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
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
        private OfficeDispatcher dispatcher;
        private OfficeCallbackServer callback;
        private SelectionCapture selection;
        private string applicationKind = "unsupported";
        private string pipeName;
        private object ribbon;
        private ChatTarget selected;
        private List<ChatTarget> targets = new List<ChatTarget>();
        private string status = "Choose a registered chat. Native embedding and local AutoSave redirect: unsupported.";
        private int busy;
        private volatile bool disconnected;

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            try
            {
                string name = Convert.ToString(((dynamic)application).Name);
                applicationKind = DetectApplication(name);
                if (applicationKind == "unsupported") return;
                dispatcher = new OfficeDispatcher();
                selection = new SelectionCapture(application, applicationKind, instanceId);
                var operations = new DocumentOperations(application, applicationKind, selection.Documents);
                pipeName = "office-in-editor-office-" + Process.GetCurrentProcess().Id + "-" + instanceId;
                callback = new OfficeCallbackServer(pipeName, dispatcher, (method, parameters) =>
                    method == "captureSelection" ? selection.Capture() : operations.Dispatch(method, parameters));
                callback.Start();
                // Passive registration only. Normal Office launch never spawns Core or writes documents/state.
                Task.Run(async () => { try { await RegisterAsync(false).ConfigureAwait(false); } catch { } });
            }
            catch { status = "Office integration is unavailable. No document changes were made."; }
        }

        internal static string DetectApplication(string name)
        {
            if (name == null) return "unsupported";
            if (name.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0) return "word";
            if (name.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0) return "excel";
            if (name.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0) return "powerpoint";
            return "unsupported";
        }

        public string GetCustomUI(string ribbonId) => RibbonMarkup.Build(applicationKind);
        public void OnRibbonLoad(object ribbonUi) { ribbon = ribbonUi; Invalidate(); }
        public bool GetSendEnabled(object control) => !disconnected && selected != null &&
            selection != null && Volatile.Read(ref busy) == 0;
        public bool GetRefreshEnabled(object control) => !disconnected && dispatcher != null &&
            Volatile.Read(ref busy) == 0;
        public string GetStatusLabel(object control) => status;
        public string GetTargetLabel(object control) => selected == null ? "Choose chat" : "Chat: " + selected.Label;

        public string GetChatMenu(object control)
        {
            var xml = new StringBuilder("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
            xml.Append("<button id=\"OieRefreshFromMenu\" label=\"Refresh registered chats\" onAction=\"RefreshChats\" getEnabled=\"GetRefreshEnabled\"/>");
            if (targets.Count == 0)
                xml.Append("<button id=\"OieNoTargets\" label=\"Register this chat using its Office MCP App\" enabled=\"false\"/>");
            for (int i = 0; i < targets.Count; i++)
            {
                var target = targets[i];
                xml.Append("<button id=\"OieTarget").Append(i).Append("\" label=\"")
                    .Append(RibbonMarkup.Escape(target.Label)).Append("\" tag=\"")
                    .Append(RibbonMarkup.Escape(target.Id)).Append("\" onAction=\"ChooseChat\"/>");
            }
            xml.Append("<button id=\"OieClearTarget\" label=\"Clear selected chat\" onAction=\"ClearChat\"/>");
            return xml.Append("</menu>").ToString();
        }

        public void ChooseChat(object control)
        {
            if (disconnected || Volatile.Read(ref busy) != 0) return;
            try
            {
                string id = Convert.ToString(((dynamic)control).Tag);
                selected = targets.Find(t => t.Id == id);
                status = selected == null ? "No registered chat selected." : "Selected chat. Add to Chat queues context; it never submits a message.";
                Invalidate();
            }
            catch { status = "Chat target unavailable. Refresh chats."; Invalidate(); }
        }
        public void ClearChat(object control)
        {
            if (Volatile.Read(ref busy) != 0) return;
            selected = null; status = "No registered chat selected."; Invalidate();
        }

        public void RefreshChats(object control)
        {
            if (disconnected || dispatcher == null || Interlocked.CompareExchange(ref busy, 1, 0) != 0) return;
            status = "Refreshing registered chats..."; Invalidate();
            Task.Run(async () =>
            {
                try
                {
                    await RegisterAsync(true).ConfigureAwait(false);
                    object result = await core.CallAsync("chat.list", Json.Object()).ConfigureAwait(false);
                    var list = ParseTargets(result);
                    await OnUiAsync(() =>
                    {
                        targets = list;
                        if (selected != null) selected = targets.Find(t => t.Id == selected.Id && t.Epoch == selected.Epoch);
                        status = targets.Count == 0 ? "No registered chats. Register this specific chat using its MCP App."
                            : "Choose a registered chat from the Chat menu.";
                    }).ConfigureAwait(false);
                }
                catch (Exception error) { await ShowErrorAsync(error).ConfigureAwait(false); }
                finally { Interlocked.Exchange(ref busy, 0); await OnUiAsync(Invalidate).ConfigureAwait(false); }
            });
        }

        public void AddToChat(object control)
        {
            if (disconnected || selected == null || selection == null || Interlocked.CompareExchange(ref busy, 1, 0) != 0) return;
            object captured;
            string targetId = selected.Id;
            try { captured = selection.Capture(); }
            catch (Exception error)
            { Interlocked.Exchange(ref busy, 0); status = ErrorStatus(error); Invalidate(); return; }
            // Capture on the Office STA before changing focus or dispatching asynchronous work.
            string requestId = Guid.NewGuid().ToString("N");
            status = "Queuing selected context..."; Invalidate();
            Task.Run(async () =>
            {
                try
                {
                    await RegisterAsync(false).ConfigureAwait(false);
                    await core.CallAsync("reference.deliverSelection", Json.Object("selection", captured,
                        "targetId", targetId, "requestId", requestId)).ConfigureAwait(false);
                    await OnUiAsync(() => status = "Reference queued. Chat acceptance is pending; no message was submitted.")
                        .ConfigureAwait(false);
                }
                catch (Exception error) { await ShowErrorAsync(error).ConfigureAwait(false); }
                finally { Interlocked.Exchange(ref busy, 0); await OnUiAsync(Invalidate).ConfigureAwait(false); }
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
                string display = clean + " [" + ShortIdentity(id);
                string hostId = Json.Text(item, "hostId");
                if (!string.IsNullOrEmpty(hostId)) display += "; host " + ShortIdentity(hostId);
                display += "]";
                parsed.Add(new ChatTarget { Id = id, Label = display, Epoch = Convert.ToString(Json.Get(item, "epoch")) });
                if (parsed.Count > 256) throw new BridgeException("TOO_MANY_CHAT_TARGETS");
            }
            return parsed;
        }
        private static string ShortIdentity(string identity)
        {
            var clean = new StringBuilder();
            foreach (char c in identity)
            {
                if (!char.IsControl(c)) clean.Append(c);
                if (clean.Length == 8) break;
            }
            return clean.ToString();
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
        public void OnStartupComplete(ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        private void Disconnect()
        {
            if (disconnected) return;
            disconnected = true;
            callback?.Dispose();
            dispatcher?.Dispose();
            selected = null; targets.Clear(); selection = null; ribbon = null;
            // No Quit/Close/Save, process termination, COM final-release, or global UI reset.
            Task.Run(async () => { try { await core.CallAsync("office.unregister", Json.Object("officeId", instanceId))
                        .ConfigureAwait(false); } catch { } });
        }
    }
}
