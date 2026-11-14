using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

public sealed partial class CoreState
{
    private static readonly string[] AllowedWriteKinds = ["word.insert_text", "excel.set_range_values", "excel.set_range_formulas", "powerpoint.add_slide", "powerpoint.add_textbox"];
    private sealed class DocumentGate { public readonly SemaphoreSlim Semaphore = new(1); public int Users; }
    private readonly Dictionary<string, DocumentGate> documentGates = new(StringComparer.Ordinal);

    private async Task<JsonNode?> InvokeOfficeOperationAsync(string method, JsonNode? value, CancellationToken cancellation)
    {
        Office office;
        var officeId = Wire.Text(value, "officeId", 128);
        lock (sync) office = GetOffice(officeId);
        var documentId = Wire.Text(value, "documentId", 160, method != "office.inspect_document");
        if (method == "office.inspect_document")
            return await CallbackAsync(office, "inspectDocument", Wire.Node(new { documentId = documentId == "" ? null : documentId }), cancellation);
        var operationId = Wire.Text(value, "operationId", 128);
        DocumentGate gate;
        lock (sync)
        {
            if (!documentGates.TryGetValue(documentId, out gate!))
            {
                if (documentGates.Count >= 128) throw new CoreFault("quota_exceeded", "Too many documents have pending operations.");
                documentGates.Add(documentId, gate = new DocumentGate());
            }
            gate.Users++;
        }
        var acquired = false;
        try
        {
            await gate.Semaphore.WaitAsync(cancellation); acquired = true;
            return method switch
            {
                "office.apply_edit" => await ApplyEditAsync(office, documentId, operationId, value?["operation"], cancellation),
                "office.revert_edit" => await RevertEditAsync(office, documentId, operationId, cancellation),
                "office.save_document" => await SaveDocumentAsync(office, documentId, operationId, cancellation),
                _ => throw new CoreFault("method_not_found", "Operation is unsupported.")
            };
        }
        finally
        {
            lock (sync)
            {
                if (acquired) gate.Semaphore.Release();
                if (--gate.Users == 0) { documentGates.Remove(documentId); gate.Semaphore.Dispose(); }
            }
        }
    }
    private static async Task<JsonNode?> CallbackAsync(Office office, string method, JsonNode? parameters, CancellationToken cancellation)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        timeout.CancelAfter(TimeSpan.FromSeconds(17));
        try { return await PipeClient.CallAsync(office.PipeName, method, parameters, timeout.Token); }
        catch (CoreFault fault) { throw new CoreFault(fault.Code, "The Office adapter rejected the operation. Inspect its state before continuing."); }
        catch (Exception) { throw new CoreFault("office_outcome_unknown", "The Office response did not complete. Do not replay a possible write."); }
    }
    private async Task<JsonNode> InspectForWriteAsync(Office office, string documentId, string? kind, bool reverting, CancellationToken cancellation)
    {
        var inspected = await CallbackAsync(office, "inspectDocument", Wire.Node(new { documentId }), cancellation)
            ?? throw new CoreFault("office_capability_unavailable", "The Office adapter did not report document capabilities.");
        if (Wire.Text(inspected, "documentId", 160) != documentId || Wire.Text(inspected, "application", 16) != office.Application)
            throw new CoreFault("document_mismatch", "The Office document identity does not match.");
        var capabilities = inspected["capabilities"];
        if (kind is not null)
        {
            var reported = capabilities?["writeKinds"] ?? inspected["writeKinds"] ?? inspected["operationKinds"];
            if (reported is not JsonArray array || !array.Any(item => item is JsonValue scalar && scalar.TryGetValue<string>(out var name) && name == kind))
                throw new CoreFault("operation_unsupported", "This Office document does not permit the requested edit.");
        }
        else if (!Wire.Boolean(capabilities, reverting ? "revert" : "manualSave"))
            throw new CoreFault("operation_unsupported", "The Office adapter does not permit this operation.");
        return inspected;
    }
    private static string Fingerprint(string intent, string documentId, JsonNode? value) => Convert.ToHexString(SHA256.HashData(
        Encoding.UTF8.GetBytes(intent + "\0" + documentId + "\0" + (value?.ToJsonString() ?? "")))).ToLowerInvariant();
    private JsonNode? ExistingOperation(string documentId, string operationId, string fingerprint)
    {
        var old = operations.Find(documentId, operationId);
        if (old is null) return null;
        if (Wire.Text(old.Value, "fingerprint", 64) != fingerprint) throw new CoreFault("idempotency_conflict", "The operation identity has different parameters.");
        return operations.Public(old);
    }
    private void CheckUnknown(string documentId)
    {
        if (operations.DocumentHasUnknown(documentId)) throw new CoreFault("document_operation_unknown", "Resolve the document's unknown operation before another write.");
    }
    private async Task<JsonNode?> ApplyEditAsync(Office office, string documentId, string operationId, JsonNode? operation, CancellationToken cancellation)
    {
        var kind = Wire.Text(operation, "kind", 64);
        if (!AllowedWriteKinds.Contains(kind, StringComparer.Ordinal) || operation?["args"] is not JsonObject)
            throw new CoreFault("operation_unsupported", "Only declared, recoverable Office edit kinds are permitted.");
        if (!kind.StartsWith(office.Application + ".", StringComparison.Ordinal)) throw new CoreFault("operation_unsupported", "The edit kind does not match the Office application.");
        if (Encoding.UTF8.GetByteCount(operation.ToJsonString()) > 128 * 1024) throw new CoreFault("operation_too_large", "The bounded edit payload is too large.");
        var fingerprint = Fingerprint("edit", documentId, operation);
        lock (sync)
        {
            var previous = ExistingOperation(documentId, operationId, fingerprint);
            if (previous is not null) return previous;
            CheckUnknown(documentId);
        }
        await InspectForWriteAsync(office, documentId, kind, false, cancellation);
        // prepareEdit is read-only. It binds an exact target and returns the compensation payload.
        var preparation = await CallbackAsync(office, "prepareEdit", Wire.Node(new { documentId, operationId, operation }), cancellation);
        if (preparation is null || Wire.Text(preparation, "documentId", 160) != documentId ||
            Wire.Text(preparation, "application", 16) != office.Application || preparation["recovery"] is null || preparation["prepared"] is null)
            throw new CoreFault("recovery_unavailable", "The Office adapter did not prepare a complete recoverable edit.");
        var beforeHash = Wire.Text(preparation, "beforeHash", 128);
        OperationRecord record;
        lock (sync)
        {
            record = operations.Begin(office.Id, documentId, operationId, fingerprint, "edit", operation, preparation);
            operations.Update(record, next => next["status"] = "executing");
        }
        try
        {
            var applied = await CallbackAsync(office, "applyEdit", Wire.Node(new { documentId, operationId, operation, prepared = preparation["prepared"]!.DeepClone(), beforeHash }), cancellation);
            var afterHash = Wire.Text(applied, "afterHash", 128);
            lock (sync)
            {
                operations.Update(record, next => { next["status"] = "applied"; next["afterHash"] = afterHash; next["result"] = applied?["result"]?.DeepClone(); });
                return operations.Public(record);
            }
        }
        catch (Exception error)
        {
            lock (sync) { operations.Unknown(record, error is CoreFault fault ? fault.Code : "write_outcome_unknown"); return operations.Public(record); }
        }
    }
    private async Task<JsonNode?> SaveDocumentAsync(Office office, string documentId, string operationId, CancellationToken cancellation)
    {
        var fingerprint = Fingerprint("manual_save", documentId, null);
        lock (sync)
        {
            var previous = ExistingOperation(documentId, operationId, fingerprint);
            if (previous is not null) return previous;
            CheckUnknown(documentId);
        }
        await InspectForWriteAsync(office, documentId, null, false, cancellation);
        OperationRecord record;
        lock (sync)
        {
            record = operations.Begin(office.Id, documentId, operationId, fingerprint, "manual_save", null);
            operations.Update(record, next => next["status"] = "saving");
        }
        try
        {
            var result = await CallbackAsync(office, "saveDocument", Wire.Node(new { documentId, operationId, intent = "manual_save" }), cancellation);
            lock (sync)
            {
                operations.Update(record, next => { next["status"] = "saved"; next["result"] = result?.DeepClone(); });
                return operations.Public(record);
            }
        }
        catch (Exception error)
        {
            lock (sync) { operations.Unknown(record, error is CoreFault fault ? fault.Code : "save_outcome_unknown"); return operations.Public(record); }
        }
    }
    private async Task<JsonNode?> RevertEditAsync(Office office, string documentId, string operationId, CancellationToken cancellation)
    {
        OperationRecord record;
        lock (sync)
        {
            record = operations.Find(documentId, operationId) ?? throw new CoreFault("operation_not_found", "The operation is unavailable.");
            var status = Wire.Text(record.Value, "status");
            if (status == "reverted") return operations.Public(record);
            if (status != "applied" || Wire.Text(record.Value, "intent") != "edit") throw new CoreFault("revert_unavailable", "Only a confirmed applied edit can be reverted.");
            CheckUnknown(documentId);
        }
        await InspectForWriteAsync(office, documentId, null, true, cancellation);
        JsonNode recovery;
        string afterHash;
        JsonNode? result;
        lock (sync)
        {
            recovery = operations.ReadRecovery(record); afterHash = Wire.Text(record.Value, "afterHash", 128); result = record.Value["result"]?.DeepClone();
            operations.Update(record, next => next["status"] = "reverting");
        }
        try
        {
            var reverted = await CallbackAsync(office, "revertEdit", Wire.Node(new { documentId, operationId, recovery, afterHash, result }), cancellation);
            lock (sync)
            {
                operations.Update(record, next => { next["status"] = "reverted"; next["revertResult"] = reverted?.DeepClone(); });
                return operations.Public(record);
            }
        }
        catch (CoreFault fault) when (fault.Code.Contains("CONFLICT", StringComparison.OrdinalIgnoreCase) ||
            fault.Code.Contains("HASH", StringComparison.OrdinalIgnoreCase) || fault.Code == "document_changed")
        {
            lock (sync)
            {
                operations.Update(record, next => { next["status"] = "applied"; next["errorCode"] = "revert_conflict"; });
                throw new CoreFault("revert_conflict", "The document changed after this edit. No automatic rollback is permitted.");
            }
        }
        catch (Exception error)
        {
            lock (sync) { operations.Unknown(record, error is CoreFault fault ? fault.Code : "revert_outcome_unknown"); return operations.Public(record); }
        }
    }
}
