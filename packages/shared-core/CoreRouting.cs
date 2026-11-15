using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

public sealed partial class CoreState
{
    private sealed record SessionOwner(string HostId, string Lease, string Epoch, int HostPid, int ServerPid);
    private sealed record DocumentBinding(string OfficeId, string DocumentId, string SessionId, string HostId);
    private sealed record DocumentRouteState(string TargetId, string TargetEpoch, long Version);
    private sealed record CaptureRequest(string Fingerprint, TaskCompletionSource<JsonNode?> Completion);
    private readonly Dictionary<string, SessionOwner> sessionOwners = new(StringComparer.Ordinal);
    private readonly Dictionary<string, DocumentBinding> documentBindings = new(StringComparer.Ordinal);
    private readonly Dictionary<string, DocumentRouteState> documentRoutes = new(StringComparer.Ordinal);
    private readonly Dictionary<string, CaptureRequest> captureRequests = new(StringComparer.Ordinal);
    private long routeSequence;
    private static string DocumentKey(string officeId, string documentId) => officeId + "\0" + documentId;
    private static double ReadPixelRatio(JsonNode? value)
    {
        if (value?["pixelRatio"] is null) return 1;
        if (value["pixelRatio"] is not JsonValue scalar || !scalar.TryGetValue<double>(out var ratio) || !double.IsFinite(ratio) || ratio < 0.5 || ratio > 8)
            throw new CoreFault("invalid_params", "The editor pixel ratio is invalid.");
        return ratio;
    }

    private Host AuthorizeHost(JsonNode? value, int pid)
    {
        var id = Wire.Text(value, "hostId", 128);
        if (!hosts.TryGetValue(id, out var host) || host.Owner.Pid != pid ||
            !Wire.TokenEquals(Wire.Text(value, "hostLease", 128), host.Lease) ||
            Wire.Text(value, "hostEpoch", 128) != host.Epoch)
            throw new CoreFault("unauthorized", "The editor host lease is invalid.");
        return host;
    }

    private void ValidateOpeningHost(JsonNode? value, int pid)
    {
        if (Wire.Text(value, "mode", 16, false) != "embedded") return;
        var host = AuthorizeHost(value, pid);
        var id = Wire.Text(value, "sessionId", 64);
        if (!id.StartsWith("ole-", StringComparison.Ordinal) || !Guid.TryParseExact(id[4..], "N", out _))
            throw new CoreFault("invalid_session", "The editor session identity is invalid.");
        if (sessionOwners.ContainsKey(id)) throw new CoreFault("identity_conflict", "The session already has an owner.");
        if (sessionOwners.Count >= 64) throw new CoreFault("quota_exceeded", "The session ownership budget is full.");
        sessionOwners.Add(id, new SessionOwner(host.Id, host.Lease, host.Epoch, pid, 0));
    }

    private void RememberOpenedSession(JsonNode? value, JsonNode? opened, int pid)
    {
        if (Wire.Text(value, "mode", 16, false) != "embedded") return;
        var host = AuthorizeHost(value, pid);
        var id = Wire.Text(value, "sessionId", 64);
        if (!sessionOwners.TryGetValue(id, out var owner) || owner.HostId != host.Id || owner.Epoch != host.Epoch)
            throw new CoreFault("office_session_expired", "The editor session owner expired during open.");
        if (Wire.Text(opened, "officeId", 160) != id || !Wire.Boolean(opened, "inplace"))
            throw new CoreFault("invalid_response", "The native session identity does not match.");
        sessionOwners[id] = owner with { ServerPid = Wire.Number(opened, "pid") };
    }

    private void ForgetSession(string id)
    {
        sessionOwners.Remove(id);
        foreach (var item in documentBindings.Where(item => item.Value.SessionId == id).ToArray())
        {
            documentBindings.Remove(item.Key);
            documentRoutes.Remove(item.Key);
        }
    }

    private void ValidateSessionRequest(string method, JsonNode? value, int pid)
    {
        var id = Wire.Text(value, "officeId", 160);
        if (!id.StartsWith("ole-", StringComparison.Ordinal)) return;
        if (!sessionOwners.TryGetValue(id, out var owner)) throw new CoreFault("office_session_expired", "The session ownership has expired.");
        if (pid == owner.HostPid)
        {
            var host = AuthorizeHost(value, pid);
            if (host.Id != owner.HostId || host.Epoch != owner.Epoch || host.Lease != owner.Lease)
                throw new CoreFault("unauthorized", "This session belongs to another host lease.");
            return;
        }
        var documentId = Wire.Text(value, "documentId", 160);
        if (!documentBindings.Values.Any(item => item.SessionId == id && item.DocumentId == documentId &&
            offices.TryGetValue(item.OfficeId, out var office) && office.Pid == pid))
            throw new CoreFault("unauthorized", "This document does not own the native session.");
        if (method is not ("office.save_window" or "office.configure_save_window" or "office.describe_window" or "office.status_window"))
            throw new CoreFault("unauthorized", "Only the owning editor may change the native window lifecycle.");
    }

    private JsonNode BindDocument(JsonNode? value, int pid)
    {
        var office = GetOffice(Wire.Text(value, "officeId", 128));
        if (office.Pid != pid) throw new CoreFault("unauthorized", "Document bindings must come from the owning Office process.");
        var documentId = Wire.Text(value, "documentId", 160);
        var sessionId = Wire.Text(value, "sessionId", 64, false);
        var hostId = "";
        if (sessionId.Length > 0)
        {
            if (!OleDocumentHost.IsSessionForProcess(sessionId, pid) || !sessionOwners.TryGetValue(sessionId, out var session) || session.ServerPid != pid)
                throw new CoreFault("unauthorized", "The native session does not belong to this Office document process.");
            hostId = session.HostId;
        }
        var key = DocumentKey(office.Id, documentId);
        if (!documentBindings.ContainsKey(key) && documentBindings.Count >= 128)
            throw new CoreFault("quota_exceeded", "The document binding budget is full.");
        var binding = new DocumentBinding(office.Id, documentId, sessionId, hostId);
        if (documentBindings.TryGetValue(key, out var old) && old != binding) documentRoutes.Remove(key);
        documentBindings[key] = binding;
        return RouteSnapshot(binding);
    }

    private JsonNode DocumentRoute(JsonNode? value, int pid)
    {
        var office = GetOffice(Wire.Text(value, "officeId", 128));
        if (office.Pid != pid) throw new CoreFault("unauthorized", "Document routing belongs to the Office process.");
        var key = DocumentKey(office.Id, Wire.Text(value, "documentId", 160));
        if (!documentBindings.TryGetValue(key, out var binding)) return Wire.Node(new { brokerEpoch, bound = false })!;
        return RouteSnapshot(binding);
    }

    private JsonNode RouteSnapshot(DocumentBinding binding)
    {
        var key = DocumentKey(binding.OfficeId, binding.DocumentId);
        documentRoutes.TryGetValue(key, out var route);
        Chat? chat = null;
        if (route is not null && chats.TryGetValue(route.TargetId, out var target) && target.Epoch == route.TargetEpoch) chat = target;
        return Wire.Node(new { brokerEpoch, bound = chat is not null, embedded = binding.SessionId.Length > 0,
            documentId = binding.DocumentId, sessionId = binding.SessionId, routeVersion = route?.Version ?? 0,
            target = chat is null ? null : PublicChat(chat) })!;
    }

    private Chat? ResolveTouchedChat(JsonNode? value, int pid, Host? delegatedHost = null)
    {
        var ownerId = Wire.Text(value, "ownerId", 160);
        var targetId = Wire.Text(value, "targetId", 128, false);
        Chat? chat;
        if (targetId.Length > 0)
        {
            if (!chats.TryGetValue(targetId, out chat)) return null;
            if (chat.OwnerId != ownerId || (delegatedHost is null ? chat.Owner.Pid != pid : chat.HostId != delegatedHost.Id))
                throw new CoreFault("unauthorized", "The initiating chat belongs to another owner.");
            var epoch = Wire.Text(value, "targetEpoch", 128, false);
            if (epoch.Length > 0 && chat.Epoch != epoch) throw new CoreFault("unauthorized", "The initiating chat epoch expired.");
            return chat;
        }
        // A stdio process can serve multiple chats. Never guess its most recent registration.
        var owned = chats.Values.Where(item => item.OwnerId == ownerId && item.Owner.Pid == pid).Take(2).ToArray();
        return owned.Length == 1 ? owned[0] : null;
    }

    private async Task<JsonNode?> TouchRouteAsync(JsonNode? value, int pid, CancellationToken cancellation, Host? delegatedHost = null)
    {
        Chat? chat;
        DocumentBinding? binding = null;
        DocumentRouteState? route = null;
        Office? office = null;
        lock (sync)
        {
            chat = ResolveTouchedChat(value, pid, delegatedHost);
            if (chat is null) return Wire.Node(new { bound = false, attributed = false, reason = "chat_unattributed" });
            MarkActivity(chat);
            var officeId = Wire.Text(value, "officeId", 128, false);
            var documentId = Wire.Text(value, "documentId", 160, false);
            if (officeId.Length > 0 && documentId.Length == 0)
            {
                var candidates = documentBindings.Values.Where(item => item.OfficeId == officeId && item.HostId == chat.HostId).Take(2).ToArray();
                if (candidates.Length == 1) documentId = candidates[0].DocumentId;
            }
            if (officeId.Length > 0 && documentId.Length > 0 && documentBindings.TryGetValue(DocumentKey(officeId, documentId), out binding) && binding.SessionId.Length > 0)
            {
                if (chat.HostId.Length == 0 || chat.HostId != binding.HostId)
                    throw new CoreFault("unauthorized", "The chat and embedded document must belong to the same editor host.");
                route = new DocumentRouteState(chat.Id, chat.Epoch, ++routeSequence);
                documentRoutes[DocumentKey(officeId, documentId)] = route;
                office = GetOffice(officeId);
            }
        }
        if (binding is null || route is null || office is null)
            return Wire.Node(new { bound = false, attributed = true, targetId = chat.Id, reason = "document_unbound" });
        var pushed = false;
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        deadline.CancelAfter(TimeSpan.FromMilliseconds(500));
        try
        {
            var answer = await PipeClient.CallAsync(office.PipeName, "bindChatTarget", Wire.Node(new { id = chat.Id, label = chat.Label, epoch = chat.Epoch,
                documentId = binding.DocumentId, sessionId = binding.SessionId, routeVersion = route.Version, brokerEpoch }), deadline.Token);
            pushed = Wire.Boolean(answer, "bound");
        }
        catch { /* The authoritative route remains committed; cache acknowledgement is separate. */ }
        return Wire.Node(new { bound = true, attributed = true, cacheAcknowledged = pushed, targetId = chat.Id, routeVersion = route.Version, brokerEpoch });
    }

    private (Chat? Chat, long Version) ResolveDeliveryTarget(JsonNode? value, DocumentBinding binding)
    {
        var mode = Wire.Text(value, "routeMode", 16);
        if (mode == "manual")
        {
            var id = Wire.Text(value, "targetId", 128);
            var epoch = Wire.Text(value, "targetEpoch", 128);
            return chats.TryGetValue(id, out var selected) && selected.Epoch == epoch ? (selected, 0) : (null, 0);
        }
        if (mode != "auto") throw new CoreFault("invalid_params", "The routing mode is invalid.");
        if (binding.SessionId.Length > 0 && documentRoutes.TryGetValue(DocumentKey(binding.OfficeId, binding.DocumentId), out var route)
            && chats.TryGetValue(route.TargetId, out var routed) && routed.Epoch == route.TargetEpoch
            && (routed.HostId.Length == 0 || binding.HostId.Length == 0 || routed.HostId == binding.HostId))
            return (routed, route.Version);
        return (chats.Values.OrderByDescending(item => item.ActivitySequence).FirstOrDefault(), 0);
    }

    private async Task<JsonNode?> CaptureAndDeliverAsync(JsonNode? value, int pid, CancellationToken cancellation)
    {
        var officeId = Wire.Text(value, "officeId", 128);
        var documentId = Wire.Text(value, "documentId", 160);
        var requestId = Wire.Text(value, "requestId", 128);
        var fingerprint = Hash(value!.ToJsonString());
        CaptureRequest request;
        Office office;
        DocumentBinding binding;
        bool leader;
        lock (sync)
        {
            office = GetOffice(officeId);
            if (office.Pid != pid) throw new CoreFault("unauthorized", "Selection delivery must originate in the owning Office process.");
            if (!documentBindings.TryGetValue(DocumentKey(officeId, documentId), out binding!))
                throw new CoreFault("document_unbound", "Register the exact document before capturing its selection.");
            leader = !captureRequests.TryGetValue(requestId, out request!);
            if (!leader && request.Fingerprint != fingerprint) throw new CoreFault("idempotency_conflict", "The capture request identity has different parameters.");
            if (leader)
            {
                if (captureRequests.Count >= 1024) throw new CoreFault("quota_exceeded", "The bounded capture request budget is full.");
                request = new CaptureRequest(fingerprint, new(TaskCreationOptions.RunContinuationsAsynchronously));
                captureRequests.Add(requestId, request);
            }
        }
        if (!leader) return (await request.Completion.Task.WaitAsync(cancellation))?.DeepClone();
        try
        {
            Chat? target;
            long version;
            lock (sync) (target, version) = ResolveDeliveryTarget(value, binding);
            if (target is null)
            {
                // #region agent log
                DebugNdjson.Write("C", "CoreRouting.CaptureAndDeliver", "no_target", new { chats = chats.Count, documentId });
                // #endregion
                var skipped = Wire.Node(new { queued = false, reason = "target_unavailable", brokerEpoch });
                request.Completion.TrySetResult(skipped);
                return skipped?.DeepClone();
            }
            var selection = await CallbackAsync(office, "captureSelection", Wire.Node(new { documentId }), cancellation);
            if (Wire.Text(selection, "documentId", 160) != documentId)
                throw new CoreFault("document_mismatch", "The selection belongs to another document.");
            JsonNode? result;
            lock (sync)
            {
                var currentOffice = GetOffice(officeId);
                if (currentOffice.Pid != pid || !documentBindings.TryGetValue(DocumentKey(officeId, documentId), out var currentBinding) || currentBinding != binding)
                    throw new CoreFault("document_mismatch", "The document binding changed during capture.");
                var current = ResolveDeliveryTarget(value, binding);
                if (current.Chat is null || current.Chat.Id != target.Id || current.Chat.Epoch != target.Epoch || current.Version != version)
                    result = Wire.Node(new { queued = false, reason = "route_changed", brokerEpoch });
                else
                {
                    var delivered = DeliverSelection(Wire.Node(new { selection, targetId = target.Id, requestId }))!.AsObject();
                    delivered["queued"] = true;
                    delivered["routeVersion"] = version;
                    delivered["brokerEpoch"] = brokerEpoch;
                    // #region agent log
                    DebugNdjson.Write("C", "CoreRouting.CaptureAndDeliver", "queued", new { hasTarget = true });
                    // #endregion
                    result = delivered;
                }
            }
            request.Completion.TrySetResult(result?.DeepClone());
            return result;
        }
        catch (Exception error)
        {
            request.Completion.TrySetException(error);
            // Observe the cached failure even when no duplicate request waits on it.
            _ = request.Completion.Task.Exception;
            throw;
        }
    }

    private JsonNode? SaveEmbeddedDocument(Office office, string documentId, string sessionId)
    {
        lock (sync)
        {
            if (!documentBindings.TryGetValue(DocumentKey(office.Id, documentId), out var binding) || binding.SessionId != sessionId ||
                !OleDocumentHost.IsSessionForProcess(sessionId, office.Pid))
                throw new CoreFault("document_mismatch", "The embedded save does not match the bound document.");
        }
        if (!OleDocumentHost.TryHandleWindow("office.save_window", Wire.Node(new { officeId = sessionId, documentId }), out var saved))
            throw new CoreFault("office_session_expired", "The embedded document session expired.");
        return saved;
    }
}
