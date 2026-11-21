using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

internal sealed class OperationRecord(string key, JsonObject value, long bytes)
{
    public string Key { get; } = key;
    public JsonObject Value { get; set; } = value;
    public long MetadataBytes { get; set; } = bytes;
}

internal sealed class OperationJournal
{
    private const long MetadataBudget = 4L * 1024 * 1024;
    private const long RecoveryBudget = 256L * 1024 * 1024;
    private const int MaxRecords = 1024;
    private const int MaxMetadataRecord = 256 * 1024;
    private readonly string directory;
    private readonly Dictionary<string, OperationRecord> entries = new(StringComparer.Ordinal);
    private long metadataBytes, recoveryBytes;
    public bool Blocked { get; private set; }
    public int Count => entries.Count;
    public OperationJournal(string directory) { this.directory = directory; Load(); }
    private static string Key(string documentId, string operationId) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(documentId + "\0" + operationId))).ToLowerInvariant();
    public OperationRecord? Find(string documentId, string operationId) => entries.GetValueOrDefault(Key(documentId, operationId));
    public bool DocumentHasUnknown(string documentId) => entries.Values.Any(item => Wire.Text(item.Value, "documentId") == documentId &&
        Wire.Text(item.Value, "status") is "executing" or "unknown" or "reverting" or "saving");
    public OperationRecord Begin(string officeId, string documentId, string operationId, string fingerprint, string intent, JsonNode? operation,
        JsonNode? preparation = null)
    {
        if (Blocked || entries.Count >= MaxRecords) throw new CoreFault("operation_quota_exceeded", "Operation metadata storage is full or requires explicit repair.");
        var key = Key(documentId, operationId);
        if (entries.ContainsKey(key)) throw new CoreFault("idempotency_conflict", "The operation identity already exists.");
        var recovery = preparation?["recovery"];
        var recoveryData = recovery is null ? null : Encoding.UTF8.GetBytes(recovery.ToJsonString());
        if (recoveryData?.Length > Wire.MaxLineBytes) throw new CoreFault("recovery_too_large", "The recovery payload exceeds the callback protocol budget.");
        if (recoveryData is not null && (entries.Values.Count(item => item.Value["recoveryFile"] is not null) >= 4 || recoveryBytes + recoveryData.Length > RecoveryBudget))
            throw new CoreFault("recovery_quota_exceeded", "Release an explicitly reviewed recovery record before another edit.");
        var value = new JsonObject { ["operationId"] = operationId, ["documentId"] = documentId, ["officeId"] = officeId, ["fingerprint"] = fingerprint,
            ["intent"] = intent, ["operation"] = operation?.DeepClone(), ["status"] = "prepared", ["createdAt"] = DateTimeOffset.UtcNow.ToString("O"),
            ["updatedAt"] = DateTimeOffset.UtcNow.ToString("O"), ["beforeHash"] = preparation?["beforeHash"]?.DeepClone(),
            ["prepared"] = preparation?["prepared"]?.DeepClone(), ["recoveryFile"] = recoveryData is null ? null : key + ".recovery",
            ["recoverySha256"] = recoveryData is null ? null : Convert.ToHexString(SHA256.HashData(recoveryData)).ToLowerInvariant(),
            ["recoveryBytes"] = recoveryData?.Length };
        var record = new OperationRecord(key, value, 0);
        Persist(record, value, false);
        entries.Add(key, record);
        if (recoveryData is not null)
        {
            try
            {
                var path = Path.Combine(directory, key + ".recovery");
                using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
                { file.Write(recoveryData); file.Flush(true); }
                var verified = File.ReadAllBytes(path);
                if (!SHA256.HashData(verified).AsSpan().SequenceEqual(SHA256.HashData(recoveryData))) throw new IOException("Recovery verification failed.");
                recoveryBytes += recoveryData.Length;
            }
            catch { Blocked = true; record.Value["status"] = "unknown"; throw new CoreFault("recovery_storage_failed", "Recovery could not be made durable. No edit was dispatched."); }
        }
        return record;
    }
    public void Update(OperationRecord record, Action<JsonObject> update)
    {
        var next = (JsonObject)record.Value.DeepClone(); update(next); next["updatedAt"] = DateTimeOffset.UtcNow.ToString("O");
        Persist(record, next, true);
    }
    public void Unknown(OperationRecord record, string errorCode)
    {
        try { Update(record, next => { next["status"] = "unknown"; next["errorCode"] = errorCode; }); }
        catch { record.Value["status"] = "unknown"; record.Value["errorCode"] = "journal_update_failed"; Blocked = true; }
    }
    public JsonNode ReadRecovery(OperationRecord record)
    {
        SafeDirectory();
        var expected = record.Key + ".recovery";
        if (Wire.Text(record.Value, "recoveryFile", 100) != expected) throw new CoreFault("recovery_unavailable", "The recovery record is invalid.");
        var info = new FileInfo(Path.Combine(directory, expected));
        if (!info.Exists || info.Length > Wire.MaxLineBytes || (info.Attributes & FileAttributes.ReparsePoint) != 0)
            throw new CoreFault("recovery_unavailable", "The recovery record is unavailable.");
        var bytes = File.ReadAllBytes(info.FullName);
        if (Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant() != Wire.Text(record.Value, "recoverySha256", 64))
            throw new CoreFault("recovery_unavailable", "Recovery integrity verification failed.");
        return JsonNode.Parse(bytes) ?? throw new CoreFault("recovery_unavailable", "Recovery is invalid.");
    }
    public JsonNode Public(OperationRecord record)
    {
        var value = (JsonObject)record.Value.DeepClone();
        value["hasRecovery"] = value["recoveryFile"] is not null;
        foreach (var key in new[] { "fingerprint", "prepared", "recoveryFile" }) value.Remove(key);
        return value;
    }
    public JsonNode GetPublic(JsonNode? parameters)
    {
        var operationId = Wire.Text(parameters, "operationId", 128);
        var documentId = Wire.Text(parameters, "documentId", 160, false);
        var matches = entries.Values.Where(item => Wire.Text(item.Value, "operationId", 128) == operationId &&
            (documentId == "" || Wire.Text(item.Value, "documentId", 160) == documentId)).Take(2).ToArray();
        if (matches.Length != 1) throw new CoreFault("operation_not_found", "Specify a unique operation and document identity.");
        return Public(matches[0]);
    }
    public JsonNode Release(JsonNode? parameters)
    {
        if (!Wire.Boolean(parameters, "confirm")) throw new CoreFault("confirmation_required", "Explicit operation recovery release confirmation is required.");
        var documentId = Wire.Text(parameters, "documentId", 160); var operationId = Wire.Text(parameters, "operationId", 128);
        var record = Find(documentId, operationId) ?? throw new CoreFault("operation_not_found", "Operation is unavailable.");
        if (Wire.Text(record.Value, "status") is "executing" or "saving" or "reverting") throw new CoreFault("operation_busy", "This operation is still executing.");
        var recoveryFile = record.Value["recoveryFile"]?.GetValue<string>();
        // Preserve a small durable tombstone: a released operation identity can never execute again.
        Update(record, value =>
        {
            value["status"] = "released";
            foreach (var field in new[] { "operation", "prepared", "result", "beforeHash", "afterHash" }) value.Remove(field);
        });
        if (recoveryFile is not null)
        {
            SafeDirectory();
            if (recoveryFile != record.Key + ".recovery") throw new CoreFault("unsafe_storage", "Recovery ownership is invalid.");
            var path = Path.Combine(directory, recoveryFile);
            if (File.Exists(path))
            {
                var info = new FileInfo(path);
                if ((info.Attributes & FileAttributes.ReparsePoint) != 0) throw new CoreFault("unsafe_storage", "Linked recovery cannot be removed.");
                var size = info.Length;
                File.Delete(path); recoveryBytes = Math.Max(0, recoveryBytes - size);
            }
            Update(record, value => value.Remove("recoveryFile"));
        }
        return Public(record);
    }
    private void Persist(OperationRecord record, JsonObject next, bool useControlReserve)
    {
        SafeDirectory();
        var bytes = Encoding.UTF8.GetBytes(next.ToJsonString());
        // Reserve the replacement temporary file as well as the old metadata until rename completes.
        if (Blocked || bytes.Length > MaxMetadataRecord || metadataBytes + bytes.Length > (useControlReserve ? MetadataBudget : MetadataBudget - 1024 * 1024))
            throw new CoreFault("operation_quota_exceeded", "Operation metadata budget is exhausted.");
        var temporary = Path.Combine(directory, record.Key + ".tmp");
        if (File.Exists(temporary)) { Blocked = true; throw new CoreFault("operation_storage_blocked", "A previous journal update requires explicit recovery."); }
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
            { stream.Write(bytes); stream.Flush(true); }
            File.Move(temporary, Path.Combine(directory, record.Key + ".json"), true);
            metadataBytes += bytes.Length - record.MetadataBytes; record.MetadataBytes = bytes.Length; record.Value = next;
        }
        catch { Blocked = true; throw new CoreFault("operation_storage_failed", "The operation journal could not be made durable."); }
    }
    private void SafeDirectory()
    {
        for (DirectoryInfo? item = new(directory); item is not null; item = item.Parent)
            if (item.Exists && (item.Attributes & FileAttributes.ReparsePoint) != 0) throw new CoreFault("unsafe_storage", "Operation storage cannot traverse linked directories.");
        Directory.CreateDirectory(directory);
    }
    private void Load()
    {
        if (!Directory.Exists(directory)) return;
        try
        {
            SafeDirectory();
            var files = Directory.EnumerateFiles(directory).Take(MaxRecords + 6).ToArray();
            if (files.Length > MaxRecords + 4) { Blocked = true; return; }
            foreach (var path in files)
            {
                var info = new FileInfo(path);
                if ((info.Attributes & FileAttributes.ReparsePoint) != 0) { Blocked = true; continue; }
                if (info.Extension == ".recovery") { recoveryBytes += info.Length; if (info.Length > Wire.MaxLineBytes) Blocked = true; continue; }
                if (info.Extension != ".json" || info.Length > MaxMetadataRecord || Path.GetFileNameWithoutExtension(path).Length != 64)
                { Blocked = true; continue; }
                metadataBytes += info.Length;
                var value = JsonNode.Parse(File.ReadAllText(path, Encoding.UTF8)) as JsonObject ?? throw new CoreFault("invalid_journal", "Operation metadata is invalid.");
                var key = Key(Wire.Text(value, "documentId", 160), Wire.Text(value, "operationId", 128));
                if (key != Path.GetFileNameWithoutExtension(path)) { Blocked = true; continue; }
                if (Wire.Text(value, "status") is "prepared" or "executing" or "saving" or "reverting")
                { value["status"] = "unknown"; value["errorCode"] = "interrupted_operation_not_replayed"; }
                entries.Add(key, new OperationRecord(key, value, info.Length));
            }
            var expectedRecovery = entries.Values.Select(item => item.Value["recoveryFile"]?.GetValue<string>()).Where(item => item is not null).ToHashSet(StringComparer.Ordinal);
            if (files.Where(path => Path.GetExtension(path) == ".recovery").Any(path => !expectedRecovery.Contains(Path.GetFileName(path))) ||
                expectedRecovery.Count > 4 || recoveryBytes > RecoveryBudget || metadataBytes > MetadataBudget || entries.Count > MaxRecords) Blocked = true;
        }
        catch { Blocked = true; }
    }
}
