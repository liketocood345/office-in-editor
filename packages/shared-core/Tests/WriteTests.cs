using System.IO.Pipes;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using OfficeInEditor.Core;

internal static class WriteTests
{
    public static async Task RunAsync(string root, Action<bool, string> check, Func<Func<Task>, string, Task> fault)
    {
        await using var fake = new FakeOffice(root);
        await fake.StartAsync();
        await using var core = new CoreState(root);
        Task<JsonNode?> Call(string method, JsonNode? value = null) => core.InvokeAsync(method, value, Environment.ProcessId, CancellationToken.None);
        async Task Register(CoreState state) => await state.InvokeAsync("office.register", Wire.Node(new { officeId = "write-test-office", application = "word", pipeName = fake.PipeName, pid = Environment.ProcessId }), Environment.ProcessId, CancellationToken.None);
        JsonNode Args(string id, string text = " inserted", string document = "write-doc") => Wire.Node(new { officeId = "write-test-office", documentId = document, operationId = id,
            operation = new { kind = "word.insert_text", args = new { offset = 0, text } } })!;
        JsonNode Key(string id, string document = "write-doc") => Wire.Node(new { officeId = "write-test-office", documentId = document, operationId = id })!;
        await Register(core);
        var inspected = await Call("office.inspect_document", Wire.Node(new { officeId = "write-test-office", documentId = "write-doc" }));
        check(inspected?["capabilities"]?["writeKinds"]?.AsArray().Count == 1, "write capability is provided by the Office adapter");
        var edit = await Call("office.apply_edit", Args("edit-one"));
        check(edit?["status"]?.GetValue<string>() == "applied" && fake.ApplyCount == 1 && fake.SaveCount == 0, "Agent edit applies without automatic saving");
        check(fake.DurableRecoveryObserved, "recovery and executing journal are durable before Office apply");
        await Call("office.apply_edit", Args("edit-one"));
        check(fake.ApplyCount == 1, "same operation identity never applies twice");
        await fault(async () => { await Call("office.apply_edit", Args("edit-one", "different")); }, "idempotency_conflict");
        var count = fake.ApplyCount;
        await using (var replacement = new CoreState(root))
        {
            await Register(replacement);
            var replay = await replacement.InvokeAsync("office.apply_edit", Args("edit-one"), Environment.ProcessId, CancellationToken.None);
            check(replay?["status"]?.GetValue<string>() == "applied" && fake.ApplyCount == count, "persisted operation prevents replay after worker replacement");
        }
        var reverted = await Call("office.revert_edit", Key("edit-one"));
        check(reverted?["status"]?.GetValue<string>() == "reverted" && fake.Text("write-doc") == "original", "explicit revert uses stored compensation");
        await Call("office.revert_edit", Key("edit-one"));
        check(fake.RevertCount == 1, "revert is not replayed");
        await Call("office.apply_edit", Args("edit-two"));
        var postEdit = fake.Text("write-doc"); fake.SetText("write-doc", postEdit + " user change");
        await fault(async () => { await Call("office.revert_edit", Key("edit-two")); }, "revert_conflict");
        check(fake.Text("write-doc").EndsWith(" user change", StringComparison.Ordinal), "revert refuses subsequent user changes");
        foreach (var id in new[] { "edit-one", "edit-two" }) await Call("operation.release", Wire.Node(new { documentId = "write-doc", operationId = id, confirm = true }));
        var released = await Call("office.apply_edit", Args("edit-one"));
        check(released?["status"]?.GetValue<string>() == "released" && fake.ApplyCount == 2, "released recovery retains an operation tombstone");
        for (var index = 0; index < 4; index++) await Call("office.apply_edit", Args("quota-edit-" + index));
        count = fake.ApplyCount;
        await fault(async () => { await Call("office.apply_edit", Args("quota-overflow")); }, "recovery_quota_exceeded");
        check(fake.ApplyCount == count && Directory.GetFiles(Path.Combine(root, "operations"), "*.recovery").Length == 4, "four-record recovery quota stops another mutation");
        for (var index = 0; index < 4; index++) await Call("operation.release", Wire.Node(new { documentId = "write-doc", operationId = "quota-edit-" + index, confirm = true }));
        fake.AllowWrites = false;
        await fault(async () => { await Call("office.apply_edit", Args("capability-denied")); }, "operation_unsupported");
        check(fake.ApplyCount == count, "missing document capability blocks a write"); fake.AllowWrites = true;
        fake.DisconnectNextApply = true;
        var unknown = await Call("office.apply_edit", Args("unknown-edit", " uncertain", "unknown-doc"));
        check(unknown?["status"]?.GetValue<string>() == "unknown", "lost response produces a durable unknown outcome");
        count = fake.ApplyCount;
        await Call("office.apply_edit", Args("unknown-edit", " uncertain", "unknown-doc"));
        check(fake.ApplyCount == count, "unknown edit is never replayed");
        await fault(async () => { await Call("office.apply_edit", Args("unknown-next", "next", "unknown-doc")); }, "document_operation_unknown");
        await using (var replacement = new CoreState(root))
        {
            await Register(replacement);
            var outcome = await replacement.InvokeAsync("office.get_operation", Key("unknown-edit", "unknown-doc"), Environment.ProcessId, CancellationToken.None);
            check(outcome?["status"]?.GetValue<string>() == "unknown", "unknown outcome survives replacement");
            await replacement.InvokeAsync("office.apply_edit", Args("unknown-edit", " uncertain", "unknown-doc"), Environment.ProcessId, CancellationToken.None);
            check(fake.ApplyCount == count, "replacement never replays an unknown operation");
        }
        var saved = await Call("office.save_document", Key("manual-save", "save-doc"));
        await Call("office.save_document", Key("manual-save", "save-doc"));
        check(saved?["status"]?.GetValue<string>() == "saved" && saved?["intent"]?.GetValue<string>() == "manual_save" && fake.SaveCount == 1, "manual save is explicit and durably deduplicated");
        check((await Call("health"))?["capabilities"]?["nativeAutoSaveLocalRedirect"]?.GetValue<bool>() == false, "write tools do not claim native AutoSave redirection");
    }

    private sealed class FakeOffice(string root) : IAsyncDisposable
    {
        private readonly CancellationTokenSource stop = new();
        private readonly Dictionary<string, string> documents = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> preparations = new(StringComparer.Ordinal);
        private Task? loop;
        public string PipeName { get; } = "office-in-editor-office-" + Environment.ProcessId + "-" + Guid.NewGuid().ToString("N");
        public int ApplyCount, SaveCount, RevertCount;
        public bool DisconnectNextApply, DurableRecoveryObserved, AllowWrites = true;
        public string Text(string documentId) => documents.GetValueOrDefault(documentId, "original");
        public void SetText(string id, string text) => documents[id] = text;
        private string Hash(string documentId) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Text(documentId)))).ToLowerInvariant();
        public Task StartAsync()
        {
            loop = Task.Run(async () =>
            {
                while (!stop.IsCancellationRequested)
                {
                    await using var pipe = new NamedPipeServerStream(PipeName, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
                    try
                    {
                        await pipe.WaitForConnectionAsync(stop.Token);
                        var request = JsonNode.Parse((await Wire.ReadLineAsync(pipe, stop.Token))!);
                        var method = Wire.Text(request, "method"); var parameters = request?["params"];
                        var documentId = Wire.Text(parameters, "documentId", 160);
                        JsonNode? result;
                        var dropped = false;
                        try
                        {
                            switch (method)
                            {
                                case "inspectDocument": result = Wire.Node(new { documentId, application = "word", documentName = "test.docx", capabilities = new { writeKinds = AllowWrites ? new[] { "word.insert_text" } : [], manualSave = true, revert = true } }); break;
                                case "prepareEdit":
                                    var nonce = Guid.NewGuid().ToString("N"); preparations[Wire.Text(parameters, "operationId")] = nonce;
                                    result = Wire.Node(new { documentId, application = "word", beforeHash = Hash(documentId), recovery = new { text = Text(documentId) }, prepared = new { nonce } }); break;
                                case "applyEdit":
                                    var operationId = Wire.Text(parameters, "operationId");
                                    if (preparations.GetValueOrDefault(operationId) != parameters?["prepared"]?["nonce"]?.GetValue<string>() || Hash(documentId) != Wire.Text(parameters, "beforeHash")) throw new CoreFault("DOCUMENT_CONFLICT", "No change.");
                                    DurableRecoveryObserved = Directory.GetFiles(Path.Combine(root, "operations"), "*.json").Any(file =>
                                        JsonNode.Parse(File.ReadAllText(file))?["status"]?.GetValue<string>() == "executing") && Directory.GetFiles(Path.Combine(root, "operations"), "*.recovery").Length > 0;
                                    documents[documentId] = Wire.Text(parameters?["operation"]?["args"], "text") + Text(documentId); ApplyCount++;
                                    if (DisconnectNextApply) { DisconnectNextApply = false; dropped = true; result = null; break; }
                                    result = Wire.Node(new { documentId, afterHash = Hash(documentId), result = new { inserted = true } }); break;
                                case "revertEdit":
                                    if (Hash(documentId) != Wire.Text(parameters, "afterHash")) throw new CoreFault("DOCUMENT_CONFLICT", "No change.");
                                    documents[documentId] = Wire.Text(parameters?["recovery"], "text"); RevertCount++;
                                    result = Wire.Node(new { documentId, reverted = true }); break;
                                case "saveDocument": SaveCount++; result = Wire.Node(new { documentId, saved = true, intent = "manual_save" }); break;
                                default: throw new CoreFault("UNSUPPORTED", "Unsupported callback.");
                            }
                            if (dropped) continue;
                            await Wire.WriteAsync(pipe, new { id = request!["id"]!.GetValue<string>(), ok = true, result }, stop.Token);
                        }
                        catch (CoreFault fault)
                        {
                            await Wire.WriteAsync(pipe, new { id = request!["id"]!.GetValue<string>(), ok = false, error = new { code = fault.Code, message = "Office rejected the operation." } }, stop.Token);
                        }
                    }
                    catch (OperationCanceledException) when (stop.IsCancellationRequested) { break; }
                }
            });
            return Task.CompletedTask;
        }
        public async ValueTask DisposeAsync()
        {
            stop.Cancel();
            if (loop is not null) await loop;
            stop.Dispose();
        }
    }
}
