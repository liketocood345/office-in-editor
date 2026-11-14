using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Channels;

namespace OfficeInEditor.Core;

public sealed partial class CoreState : IAsyncDisposable
{
    public const int MaxReferences = 512;
    public const long MaxReferenceBytes = 64L * 1024 * 1024;
    public const int MaxSingleReferenceBytes = 1024 * 1024;
    private readonly object sync = new();
    private readonly string referenceDirectory;
    private readonly Dictionary<string, Reference> references = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Host> hosts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Office> offices = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Chat> chats = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Delivery> deliveries = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Deduplication> deduplication = new(StringComparer.Ordinal);
    private readonly Dictionary<string, RuntimeOwner> runtimeOwners = new(StringComparer.Ordinal);
    private readonly Timer idleTimer;
    private readonly OperationJournal operations;
    private long referenceBytes;
    private bool storageBlocked;
    private int activeRequests;
    private bool disposed;
    private bool shutdownRequested;
    public string AppOrigin { get; set; } = "";
    public event Action? BecameIdle;

    private sealed record Reference(string Id, JsonObject Selection, string CreatedAt, string Hash, long Bytes);
    private sealed record Host(string Id, string Kind, string Label, string Lease, string Epoch, ProcessLease Owner);
    private sealed record Office(string Id, string Application, string PipeName, int Pid, string RegisteredAt, string Registration, ProcessLease Owner);
    private sealed record RuntimeOwner(string Id, string Lease, ProcessLease Process);
    private sealed class Chat
    {
        public required string Id, OwnerId, Label, HostId, Epoch, Token;
        public required ProcessLease Owner;
        public Channel<bool> Signals { get; } = Channel.CreateBounded<bool>(new BoundedChannelOptions(1) { FullMode = BoundedChannelFullMode.DropWrite });
        public bool Connected;
    }
    private sealed class Delivery
    {
        public required string RequestId, RefId, TargetId, CreatedAt;
        public string Status = "queued";
        public string? LastError;
    }
    private sealed record Deduplication(string Fingerprint, string RefId, string TargetId);

    public CoreState(string dataRoot)
    {
        referenceDirectory = Path.Combine(Path.GetFullPath(dataRoot), "references");
        operations = new OperationJournal(Path.Combine(Path.GetFullPath(dataRoot), "operations"));
        idleTimer = new Timer(_ =>
        {
            lock (sync)
                if (!disposed && activeRequests == 0 && hosts.Count == 0 && offices.Count == 0 && chats.Count == 0 && runtimeOwners.Count == 0)
                    BecameIdle?.Invoke();
        });
        LoadReferences();
    }
    public void ScheduleIdleShutdown()
    {
        lock (sync)
            if (!disposed) idleTimer.Change(hosts.Count + offices.Count + chats.Count + runtimeOwners.Count + activeRequests == 0 ?
                (shutdownRequested ? TimeSpan.FromMilliseconds(250) : TimeSpan.FromSeconds(20)) : Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
    }
    private static string Timestamp() => DateTimeOffset.UtcNow.ToString("O");
    private static string Hash(string input) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(input))).ToLowerInvariant();
    private static bool ApplicationIsValid(string application) => application is "word" or "excel" or "powerpoint";

    public async Task<JsonNode?> InvokeAsync(string method, JsonNode? parameters, int ownerPid, CancellationToken cancellation)
    {
        lock (sync) { ObjectDisposedException.ThrowIf(disposed, this); activeRequests++; idleTimer.Change(Timeout.Infinite, Timeout.Infinite); }
        try
        {
            if (method == "office.capture")
            {
                Office office;
                lock (sync) office = GetOffice(Wire.Text(parameters, "officeId"));
                JsonNode? selection;
                try { selection = await PipeClient.CallAsync(office.PipeName, "captureSelection", null, cancellation); }
                catch (Exception) { throw new CoreFault("office_capture_unavailable", "Office could not capture a selection. No reference was created."); }
                lock (sync) return Metadata(Capture(selection));
            }
            if (method == "office.open") return NativeOfficeLauncher.Open(Wire.Text(parameters, "path", 32767));
            if (method is "office.inspect_document" or "office.apply_edit" or "office.revert_edit" or "office.save_document")
                return await InvokeOfficeOperationAsync(method, parameters, cancellation);
            lock (sync)
            {
                return method switch
                {
                    "health" => Health(),
                    "runtime.acquireOwner" => RegisterRuntimeOwner(parameters, ownerPid),
                    "runtime.releaseOwner" => ReleaseRuntimeOwner(parameters),
                    "runtime.shutdownIfIdle" => ShutdownIfIdle(),
                    "office.get_operation" => operations.GetPublic(parameters),
                    "operation.release" => operations.Release(parameters),
                    "host.register" => RegisterHost(parameters, ownerPid),
                    "host.unregister" => UnregisterHost(parameters),
                    "host.list" => Wire.Node(hosts.Values.Select(item => new { hostId = item.Id, kind = item.Kind, label = item.Label }).ToArray()),
                    "chat.register" => RegisterChat(parameters, ownerPid),
                    "chat.list" => Wire.Node(chats.Values.Select(PublicChat).ToArray()),
                    "chat.unregister" => UnregisterChat(parameters),
                    "office.register" => RegisterOffice(parameters, ownerPid),
                    "office.unregister" => UnregisterOffice(parameters),
                    "office.list" => Wire.Node(offices.Values.Select(item => new { officeId = item.Id, application = item.Application, pid = item.Pid, registeredAt = item.RegisteredAt }).ToArray()),
                    "office.route.get" => Wire.Node(new { mode = "standalone", supported = false, reason = "native_embedding_unavailable" }),
                    "reference.capture" => Metadata(Capture(parameters?["selection"] ?? parameters)),
                    "reference.list" => new JsonArray(references.Values.OrderByDescending(item => item.CreatedAt).Select(Metadata).ToArray()),
                    "reference.get" => FullReference(GetReference(Wire.Text(parameters, "refId"))),
                    "reference.release" => ReleaseReference(parameters),
                    "reference.deliver" => Deliver(parameters),
                    "reference.deliverSelection" => DeliverSelection(parameters),
                    "delivery.list" => ListDeliveries(Wire.Text(parameters, "targetId"), Wire.Text(parameters, "token", 128)),
                    "delivery.ack" => Ack(parameters, Wire.Text(parameters, "token", 128)),
                    _ => throw new CoreFault("method_not_found", "This operation is not supported.")
                };
            }
        }
        finally { lock (sync) activeRequests--; ScheduleIdleShutdown(); }
    }
    private JsonNode? Health() => Wire.Node(new
    {
        protocolVersion = 1,
        version = RuntimeIdentity.Version,
        capabilities = new { nativeOfficeEmbedding = false, nativeAutoSaveLocalRedirect = false, documentWrite = true, documentWriteRequiresAdapterCapability = true, selectionCapture = true, mcpApps = true,
            supportedWriteKinds = AllowedWriteKinds, manualSave = true, guardedRevert = true },
        appOrigin = AppOrigin,
        appPort = Uri.TryCreate(AppOrigin, UriKind.Absolute, out var appUri) ? appUri.Port : 0,
        owners = new { hosts = hosts.Count, office = offices.Count, chats = chats.Count, apps = chats.Values.Count(item => item.Connected), mcp = runtimeOwners.Count },
        storage = new { referenceCount = references.Count, referenceBytes, maxReferences = MaxReferences, maxReferenceBytes = MaxReferenceBytes, storageBlocked,
            operationRecords = operations.Count, operationStorageBlocked = operations.Blocked }
    });
    private JsonNode? RegisterRuntimeOwner(JsonNode? value, int pid)
    {
        var id = Wire.Text(value, "ownerId", 160);
        if (!runtimeOwners.ContainsKey(id) && runtimeOwners.Count >= 32) throw new CoreFault("quota_exceeded", "Runtime owner capacity is reached.");
        var lease = Wire.NewToken();
        var process = NewOwner(pid, () => RemoveRuntimeOwner(id, lease));
        if (runtimeOwners.Remove(id, out var old)) old.Process.Dispose();
        runtimeOwners[id] = new RuntimeOwner(id, lease, process);
        return Wire.Node(new { ownerId = id, lease });
    }
    private void RemoveRuntimeOwner(string id, string lease)
    {
        if (runtimeOwners.TryGetValue(id, out var owner) && owner.Lease == lease) { runtimeOwners.Remove(id); owner.Process.Dispose(); }
    }
    private JsonNode? ReleaseRuntimeOwner(JsonNode? value)
    {
        RemoveRuntimeOwner(Wire.Text(value, "ownerId", 160), Wire.Text(value, "lease", 128));
        return Wire.Node(new { released = true });
    }
    private JsonNode? ShutdownIfIdle()
    {
        if (hosts.Count + offices.Count + chats.Count + runtimeOwners.Count != 0 || activeRequests > 1)
            return Wire.Node(new { status = "busy" });
        shutdownRequested = true;
        return Wire.Node(new { status = "shutting_down" });
    }
    private ProcessLease NewOwner(int ownerPid, Action cleanup)
    {
        if (ownerPid <= 0) throw new CoreFault("owner_unavailable", "A live process owner is required.");
        return new ProcessLease(ownerPid, () => { lock (sync) { if (!disposed) cleanup(); } ScheduleIdleShutdown(); });
    }
    private JsonNode? RegisterHost(JsonNode? value, int pid)
    {
        var id = Wire.Text(value, "hostId", 128);
        var kind = Wire.Text(value, "kind", 16);
        if (kind is not ("vscode" or "cursor")) throw new CoreFault("invalid_params", "Host kind is unsupported.");
        if (hosts.TryGetValue(id, out var existingHost) && existingHost.Owner.Pid != pid)
            throw new CoreFault("identity_conflict", "This host identity belongs to another live process.");
        if (!hosts.ContainsKey(id) && hosts.Count >= 32) throw new CoreFault("quota_exceeded", "Host registration capacity is reached.");
        var lease = Wire.NewToken(); var epoch = Guid.NewGuid().ToString("N");
        var owner = NewOwner(pid, () => RemoveHost(id, lease));
        if (hosts.Remove(id, out var old)) old.Owner.Dispose();
        hosts[id] = new Host(id, kind, Wire.Text(value, "label", 160), lease, epoch, owner);
        return Wire.Node(new { hostId = id, lease, epoch });
    }
    private void RemoveHost(string id, string? lease = null)
    {
        if (!hosts.TryGetValue(id, out var host) || lease is not null && host.Lease != lease) return;
        hosts.Remove(id); host.Owner.Dispose();
        foreach (var chat in chats.Values.Where(item => item.HostId == id).ToArray()) RemoveChat(chat.Id);
    }
    private JsonNode? UnregisterHost(JsonNode? value)
    {
        var id = Wire.Text(value, "hostId", 128);
        if (hosts.TryGetValue(id, out var host) && !Wire.TokenEquals(Wire.Text(value, "lease", 128), host.Lease))
            throw new CoreFault("unauthorized", "Host ownership is invalid.");
        RemoveHost(id); return Wire.Node(new { removed = true });
    }
    private JsonNode? RegisterOffice(JsonNode? value, int callerPid)
    {
        var id = Wire.Text(value, "officeId", 128); var application = Wire.Text(value, "application", 16);
        var pipeName = Wire.Text(value, "pipeName", 180); var pid = Wire.Number(value, "pid");
        if (!ApplicationIsValid(application) || !pipeName.StartsWith("office-in-editor-office-" + pid + "-", StringComparison.Ordinal) ||
            pipeName.Any(character => !(char.IsAsciiLetterOrDigit(character) || character is '-' or '_')) || pid <= 0)
            throw new CoreFault("invalid_params", "Office registration is invalid.");
        if (callerPid != pid) throw new CoreFault("unauthorized", "Office registration must originate from its process.");
        if (offices.TryGetValue(id, out var existingOffice) && existingOffice.Pid != pid)
            throw new CoreFault("identity_conflict", "This Office identity belongs to another live process.");
        if (!offices.ContainsKey(id) && offices.Count >= 32) throw new CoreFault("quota_exceeded", "Office registration capacity is reached.");
        var registration = Guid.NewGuid().ToString("N");
        var owner = NewOwner(pid, () => RemoveOffice(id, registration));
        if (offices.Remove(id, out var old)) old.Owner.Dispose();
        offices[id] = new Office(id, application, pipeName, pid, Timestamp(), registration, owner);
        return Wire.Node(new { officeId = id, registered = true });
    }
    private void RemoveOffice(string id, string? registration = null)
    {
        if (offices.TryGetValue(id, out var office) && (registration is null || office.Registration == registration)) { offices.Remove(id); office.Owner.Dispose(); }
    }
    private JsonNode? UnregisterOffice(JsonNode? value) { RemoveOffice(Wire.Text(value, "officeId", 128)); return Wire.Node(new { removed = true }); }
    private Office GetOffice(string id) => offices.TryGetValue(id, out var office) ? office : throw new CoreFault("office_unavailable", "The Office instance is unavailable.");
    private static object PublicChat(Chat chat) => new { id = chat.Id, label = chat.Label, hostId = chat.HostId, epoch = chat.Epoch, ownerId = chat.OwnerId, online = true };
    private JsonNode? RegisterChat(JsonNode? value, int pid)
    {
        if (chats.Count >= 32) throw new CoreFault("quota_exceeded", "Chat registration capacity is reached.");
        var hostId = Wire.Text(value, "hostId", 128, false);
        if (hostId != "" && !hosts.ContainsKey(hostId)) throw new CoreFault("host_unavailable", "The requested host is unavailable.");
        var id = Guid.NewGuid().ToString("N");
        var chat = new Chat { Id = id, OwnerId = Wire.Text(value, "ownerId", 160), Label = Wire.Text(value, "label", 160), HostId = hostId,
            Epoch = Guid.NewGuid().ToString("N"), Token = Wire.NewToken(), Owner = NewOwner(pid, () => RemoveChat(id)) };
        chats.Add(id, chat);
        var result = (JsonObject)Wire.Node(PublicChat(chat))!;
        result["token"] = chat.Token; result["appOrigin"] = AppOrigin;
        return result;
    }
    private Chat Authorize(string targetId, string token)
    {
        if (!chats.TryGetValue(targetId, out var chat) || !Wire.TokenEquals(token, chat.Token))
            throw new CoreFault("unauthorized", "The App target or authorization is invalid.");
        return chat;
    }
    private JsonNode? UnregisterChat(JsonNode? value)
    {
        var chat = Authorize(Wire.Text(value, "targetId", 128, false) is { Length: > 0 } target ? target : Wire.Text(value, "id", 128), Wire.Text(value, "token", 128));
        if (Wire.Text(value, "ownerId", 160) != chat.OwnerId) throw new CoreFault("unauthorized", "Chat ownership is invalid.");
        RemoveChat(chat.Id); return Wire.Node(new { removed = true });
    }
    private void RemoveChat(string id)
    {
        if (!chats.Remove(id, out var chat)) return;
        chat.Signals.Writer.TryComplete(); chat.Owner.Dispose();
        foreach (var delivery in deliveries.Values.Where(item => item.TargetId == id && item.Status == "queued")) { delivery.Status = "failed"; delivery.LastError = "target_offline"; }
    }
    private Reference GetReference(string id) => references.TryGetValue(id, out var reference) ? reference : throw new CoreFault("reference_not_found", "The reference is unavailable.");
    private static JsonNode Metadata(Reference reference)
    {
        var value = reference.Selection;
        return Wire.Node(new { refId = reference.Id, id = reference.Id, application = Wire.Text(value, "application"), documentId = Wire.Text(value, "documentId"),
            documentName = Wire.Text(value, "documentName"), kind = Wire.Text(value, "kind"), createdAt = reference.CreatedAt,
            textLength = Wire.Text(value, "text", MaxSingleReferenceBytes, false).Length, sha256 = reference.Hash,
            dirty = Wire.Boolean(value, "dirty"), autoSaveAvailable = Wire.OptionalBoolean(value, "autoSaveAvailable"), autoSaveEnabled = Wire.OptionalBoolean(value, "autoSaveEnabled"),
            autoSaveState = Wire.Text(value, "autoSaveState", 32, false) })!;
    }
    private static JsonNode FullReference(Reference reference)
    {
        var result = (JsonObject)Metadata(reference);
        result["text"] = reference.Selection["text"]?.DeepClone();
        result["locator"] = reference.Selection["locator"]?.DeepClone();
        if (reference.Selection["sourcePath"] is { } source) result["sourcePath"] = source.DeepClone();
        return result;
    }
    private Reference Capture(JsonNode? selection)
    {
        if (selection is not JsonObject input) throw new CoreFault("invalid_params", "Selection is invalid.");
        var application = Wire.Text(input, "application", 16);
        if (!ApplicationIsValid(application)) throw new CoreFault("invalid_params", "Application is unsupported.");
        var clean = new JsonObject { ["application"] = application, ["documentId"] = Wire.Text(input, "documentId", 160),
            ["documentName"] = Wire.Text(input, "documentName", 512), ["kind"] = Wire.Text(input, "kind", 64),
            ["text"] = Wire.Text(input, "text", MaxSingleReferenceBytes, false), ["locator"] = input["locator"]?.DeepClone(),
            ["dirty"] = Wire.Boolean(input, "dirty"), ["autoSaveAvailable"] = Wire.Node(Wire.OptionalBoolean(input, "autoSaveAvailable")), ["autoSaveEnabled"] = Wire.Node(Wire.OptionalBoolean(input, "autoSaveEnabled")),
            ["autoSaveState"] = Wire.Text(input, "autoSaveState", 32, false) };
        if (input["sourcePath"] is not null) clean["sourcePath"] = Wire.Text(input, "sourcePath", 32767, false);
        var id = Guid.NewGuid().ToString("N"); var createdAt = Timestamp(); var hash = Hash(clean.ToJsonString());
        var envelope = new JsonObject { ["schemaVersion"] = 1, ["refId"] = id, ["createdAt"] = createdAt, ["sha256"] = hash, ["selection"] = clean.DeepClone() };
        var bytes = Encoding.UTF8.GetBytes(envelope.ToJsonString());
        if (bytes.Length > MaxSingleReferenceBytes) throw new CoreFault("reference_too_large", "Select a smaller range. Reference size is limited.");
        if (storageBlocked || references.Count >= MaxReferences || referenceBytes + bytes.Length > MaxReferenceBytes)
            throw new CoreFault("quota_exceeded", "Reference storage is full or requires explicit repair. Release references before capturing more.");
        EnsureSafeDirectory();
        var path = ReferencePath(id);
        try
        {
            using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough);
            file.Write(bytes); file.Flush(true);
        }
        catch { storageBlocked = true; throw new CoreFault("storage_unavailable", "The immutable reference could not be stored. No delivery was queued."); }
        var reference = new Reference(id, clean, createdAt, hash, bytes.Length);
        references.Add(id, reference); referenceBytes += bytes.Length;
        return reference;
    }
    private string ReferencePath(string id)
    {
        if (!Guid.TryParseExact(id, "N", out _)) throw new CoreFault("invalid_params", "Reference identity is invalid.");
        return Path.Combine(referenceDirectory, id + ".json");
    }
    private void EnsureSafeDirectory()
    {
        for (DirectoryInfo? current = new(referenceDirectory); current is not null; current = current.Parent)
            if (current.Exists && (current.Attributes & FileAttributes.ReparsePoint) != 0) throw new CoreFault("unsafe_storage", "Reference storage cannot traverse a linked directory.");
        Directory.CreateDirectory(referenceDirectory);
    }
    private void LoadReferences()
    {
        if (!Directory.Exists(referenceDirectory)) return;
        try
        {
            EnsureSafeDirectory();
            var files = Directory.EnumerateFiles(referenceDirectory).Take(MaxReferences + 1).ToArray();
            if (files.Length > MaxReferences) { storageBlocked = true; return; }
            foreach (var file in files)
            {
                var info = new FileInfo(file);
                if ((info.Attributes & FileAttributes.ReparsePoint) != 0 || info.Length > MaxSingleReferenceBytes ||
                    !Guid.TryParseExact(Path.GetFileNameWithoutExtension(file), "N", out _) || info.Extension != ".json") { storageBlocked = true; continue; }
                referenceBytes += info.Length;
                var value = JsonNode.Parse(File.ReadAllText(file, Encoding.UTF8));
                var id = Wire.Text(value, "refId", 32);
                if (Path.GetFileNameWithoutExtension(file) != id || value?["selection"] is not JsonObject selection || Hash(selection.ToJsonString()) != Wire.Text(value, "sha256", 64))
                { storageBlocked = true; continue; }
                references.Add(id, new Reference(id, (JsonObject)selection.DeepClone(), Wire.Text(value, "createdAt", 64), Wire.Text(value, "sha256", 64), info.Length));
            }
            if (referenceBytes > MaxReferenceBytes) storageBlocked = true;
        }
        catch { storageBlocked = true; }
    }
    private JsonNode? ReleaseReference(JsonNode? value)
    {
        if (!Wire.Boolean(value, "confirm")) throw new CoreFault("confirmation_required", "Explicit reference release confirmation is required.");
        var reference = GetReference(Wire.Text(value, "refId", 32));
        EnsureSafeDirectory(); var path = ReferencePath(reference.Id);
        if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) throw new CoreFault("unsafe_storage", "Linked references cannot be released automatically.");
        try { File.Delete(path); }
        catch { throw new CoreFault("storage_unavailable", "The reference could not be released."); }
        references.Remove(reference.Id); referenceBytes -= reference.Bytes;
        foreach (var entry in deliveries.Where(item => item.Value.RefId == reference.Id).Select(item => item.Key).ToArray()) deliveries.Remove(entry);
        // Keep bounded request tombstones until worker exit: a released operation must never capture again on replay.
        return Wire.Node(new { released = true, refId = reference.Id });
    }
    private static JsonNode DeliveryNode(Delivery delivery) => Wire.Node(new { requestId = delivery.RequestId, refId = delivery.RefId, targetId = delivery.TargetId,
        status = delivery.Status, createdAt = delivery.CreatedAt, lastError = delivery.LastError })!;
    private Chat ValidateTarget(string targetId) => chats.TryGetValue(targetId, out var target) ? target : throw new CoreFault("target_offline", "Select an online registered chat.");
    private Delivery Queue(string refId, string targetId, string requestId, string fingerprint)
    {
        var target = ValidateTarget(targetId);
        GetReference(refId);
        if (deduplication.TryGetValue(requestId, out var prior))
        {
            if (prior.Fingerprint != fingerprint) throw new CoreFault("idempotency_conflict", "The request identity already has different parameters.");
            if (deliveries.TryGetValue(requestId, out var old)) return old;
            throw new CoreFault("reference_released", "This request was already completed and its reference was released.");
        }
        if (deliveries.Count >= 512 || deduplication.Count >= 1024 || deliveries.Values.Count(item => item.TargetId == targetId && item.Status == "queued") >= 32)
            throw new CoreFault("quota_exceeded", "The bounded delivery queue is full.");
        var delivery = new Delivery { RequestId = requestId, RefId = refId, TargetId = targetId, CreatedAt = Timestamp() };
        deliveries.Add(requestId, delivery); deduplication.Add(requestId, new Deduplication(fingerprint, refId, targetId));
        target.Signals.Writer.TryWrite(true);
        return delivery;
    }
    private JsonNode? Deliver(JsonNode? value)
    {
        var refId = Wire.Text(value, "refId", 32); var targetId = Wire.Text(value, "targetId", 128); var requestId = Wire.Text(value, "requestId", 128);
        return DeliveryNode(Queue(refId, targetId, requestId, Hash("reference:" + refId + ":" + targetId)));
    }
    private JsonNode? DeliverSelection(JsonNode? value)
    {
        var targetId = Wire.Text(value, "targetId", 128); var requestId = Wire.Text(value, "requestId", 128);
        var selection = value?["selection"] ?? throw new CoreFault("invalid_params", "Selection is required.");
        var fingerprint = Hash("selection:" + targetId + ":" + selection.ToJsonString());
        if (deduplication.TryGetValue(requestId, out var prior))
        {
            if (prior.Fingerprint != fingerprint) throw new CoreFault("idempotency_conflict", "The request identity already has different parameters.");
            return new JsonObject { ["reference"] = Metadata(GetReference(prior.RefId)), ["delivery"] = DeliveryNode(deliveries[requestId]) };
        }
        ValidateTarget(targetId);
        if (deliveries.Count >= 512 || deduplication.Count >= 1024 || deliveries.Values.Count(item => item.TargetId == targetId && item.Status == "queued") >= 32)
            throw new CoreFault("quota_exceeded", "The bounded delivery queue is full.");
        var reference = Capture(selection);
        return new JsonObject { ["reference"] = Metadata(reference), ["delivery"] = DeliveryNode(Queue(reference.Id, targetId, requestId, fingerprint)) };
    }
    public JsonNode? ListDeliveries(string targetId, string token)
    {
        lock (sync) { Authorize(targetId, token); return new JsonArray(deliveries.Values.Where(item => item.TargetId == targetId).Select(DeliveryNode).ToArray()); }
    }
    public JsonNode? Ack(JsonNode? value, string token)
    {
        lock (sync)
        {
            var targetId = Wire.Text(value, "targetId", 128); Authorize(targetId, token);
            var requestId = Wire.Text(value, "requestId", 128);
            if (!deliveries.TryGetValue(requestId, out var delivery) || delivery.TargetId != targetId || delivery.RefId != Wire.Text(value, "refId", 32))
                throw new CoreFault("delivery_not_found", "The delivery is unavailable.");
            var status = Wire.Text(value, "status", 32);
            if (status is not ("attached" or "unsupported" or "failed")) throw new CoreFault("invalid_params", "Acknowledgement status is invalid.");
            if (delivery.Status == "attached") return DeliveryNode(delivery);
            delivery.Status = status == "unsupported" ? "queued" : status;
            delivery.LastError = status == "unsupported" ? "host_context_unsupported" : null;
            return DeliveryNode(delivery);
        }
    }
    public ChannelReader<bool> ConnectApp(string targetId, string token)
    {
        lock (sync)
        {
            var chat = Authorize(targetId, token);
            if (chat.Connected) throw new CoreFault("app_already_connected", "This target already has an App connection.");
            chat.Connected = true; return chat.Signals.Reader;
        }
    }
    public void DisconnectApp(string targetId)
    {
        // An App disconnect invalidates its route. A still-running MCP proxy does not imply that
        // this particular chat App can accept context; it must be registered explicitly again.
        lock (sync) RemoveChat(targetId);
        ScheduleIdleShutdown();
    }
    public JsonNode[] PendingEvents(string targetId, string token)
    {
        lock (sync)
        {
            Authorize(targetId, token);
            return deliveries.Values.Where(item => item.TargetId == targetId && item.Status == "queued").Select(item =>
            {
                var reference = GetReference(item.RefId);
                return Wire.Node(new { requestId = item.RequestId, refId = item.RefId, text = Wire.Text(reference.Selection, "text", MaxSingleReferenceBytes, false),
                    documentName = Wire.Text(reference.Selection, "documentName"), kind = Wire.Text(reference.Selection, "kind"), locator = reference.Selection["locator"]?.DeepClone() })!;
            }).ToArray();
        }
    }
    public ValueTask DisposeAsync()
    {
        lock (sync)
        {
            if (disposed) return ValueTask.CompletedTask;
            disposed = true; idleTimer.Dispose();
            foreach (var value in hosts.Values) value.Owner.Dispose();
            foreach (var value in offices.Values) value.Owner.Dispose();
            foreach (var value in chats.Values) { value.Owner.Dispose(); value.Signals.Writer.TryComplete(); }
            foreach (var value in runtimeOwners.Values) value.Process.Dispose();
        }
        return ValueTask.CompletedTask;
    }
}
