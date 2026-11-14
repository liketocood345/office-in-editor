using System.Diagnostics;
using System.IO.Pipes;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
using OfficeInEditor.Core;

var testId = Guid.NewGuid().ToString("N");
var testRoot = Path.Combine(Path.GetTempPath(), "OfficeInEditor-tests-" + testId);
var passed = 0;
void Check(bool condition, string name)
{
    if (!condition) throw new InvalidOperationException("FAILED: " + name);
    passed++; Console.WriteLine("PASS " + name);
}
async Task Fault(Func<Task> operation, string code)
{
    try { await operation(); throw new InvalidOperationException("Expected fault: " + code); }
    catch (CoreFault error) { Check(error.Code == code, code); }
}
JsonNode Selection(string text) => Wire.Node(new { application = "word", documentId = "test-document", documentName = "test.docx", sourcePath = Path.Combine(testRoot, "original.docx"),
    kind = "text", text, locator = new { start = 2, end = 7 }, dirty = true, autoSaveAvailable = (bool?)null, autoSaveEnabled = false, autoSaveState = "off_or_unavailable" })!;

try
{
    await using var state = new CoreState(Path.Combine(testRoot, "data"));
    Check(!Directory.Exists(testRoot), "construction does not create storage");
    Task<JsonNode?> Call(string method, JsonNode? value = null) => state.InvokeAsync(method, value, Environment.ProcessId, CancellationToken.None);
    var health = await Call("health");
    Check(health?["capabilities"]?["nativeOfficeEmbedding"]?.GetValue<bool>() == false && health?["capabilities"]?["nativeAutoSaveLocalRedirect"]?.GetValue<bool>() == false, "unsupported capabilities are explicit");
    Check((await Call("office.route.get"))?["mode"]?.GetValue<string>() == "standalone", "no guessed embedded route");

    var host = await Call("host.register", Wire.Node(new { hostId = "test-host", kind = "vscode", label = "Test host" }));
    Check((await Call("host.list"))?.AsArray().Count == 1, "host registration");
    await Fault(async () => { await Call("host.unregister", Wire.Node(new { hostId = "test-host", lease = "wrong" })); }, "unauthorized");
    var target = await Call("chat.register", Wire.Node(new { ownerId = "test-owner", hostId = "test-host", label = "Test chat" }));
    var targetId = target!["id"]!.GetValue<string>(); var token = target["token"]!.GetValue<string>();
    Check(!(await Call("chat.list"))!.ToJsonString().Contains(token, StringComparison.Ordinal), "chat directory never returns tokens");
    Check((await Call("runtime.shutdownIfIdle"))?["status"]?.GetValue<string>() == "busy", "busy shutdown does not affect owners");

    await using var app = new AppServer(state);
    await app.StartAsync(CancellationToken.None); state.AppOrigin = app.Origin;
    using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(8) };
    var preflight = new HttpRequestMessage(HttpMethod.Options, app.Origin + "/app/ack");
    preflight.Headers.Add("Origin", "null"); preflight.Headers.Add("Access-Control-Request-Headers", "authorization,content-type");
    var preflightReply = await http.SendAsync(preflight);
    Check(preflightReply.StatusCode == HttpStatusCode.NoContent && preflightReply.Headers.GetValues("Access-Control-Allow-Private-Network").Single() == "true" && !preflightReply.Headers.Contains("Access-Control-Allow-Credentials"), "private-network preflight without blanket credentials");
    var badHost = new HttpRequestMessage(HttpMethod.Get, app.Origin + "/app/events?targetId=" + targetId); badHost.Headers.Host = "other.invalid";
    Check((await http.SendAsync(badHost)).StatusCode == HttpStatusCode.Forbidden, "Host validation rejects rebinding");
    var streamRequest = new HttpRequestMessage(HttpMethod.Get, app.Origin + "/app/events?targetId=" + targetId);
    streamRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token); streamRequest.Headers.Add("Origin", "null");
    using var streamReply = await http.SendAsync(streamRequest, HttpCompletionOption.ResponseHeadersRead);
    Check(streamReply.IsSuccessStatusCode && streamReply.Content.Headers.ContentType?.MediaType == "text/event-stream", "authenticated push stream");
    using var streamReader = new StreamReader(await streamReply.Content.ReadAsStreamAsync());
    var text = "A selected paragraph\n中文选区 Δ";
    var selection = Selection(text);
    var deliveryArgs = Wire.Node(new { selection, targetId, requestId = "delivery-test-1" });
    var delivered = await Call("reference.deliverSelection", deliveryArgs);
    var refId = delivered!["reference"]!["refId"]!.GetValue<string>();
    Check(delivered["delivery"]?["status"]?.GetValue<string>() == "queued", "queued is not attached");
    var replay = await Call("reference.deliverSelection", deliveryArgs);
    Check(replay?["reference"]?["refId"]?.GetValue<string>() == refId && (await Call("reference.list"))?.AsArray().Count == 1, "repeated right-click request captures once");
    await Fault(async () => { await Call("reference.deliverSelection", Wire.Node(new { selection = Selection("different"), targetId, requestId = "delivery-test-1" })); }, "idempotency_conflict");
    JsonNode? pushed = null;
    using (var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5)))
        while (await streamReader.ReadLineAsync(deadline.Token) is { } line)
            if (line.StartsWith("data: ", StringComparison.Ordinal)) { pushed = JsonNode.Parse(line[6..]); break; }
    Check(pushed?["refId"]?.GetValue<string>() == refId && pushed?["text"]?.GetValue<string>() == text && pushed?["sourcePath"] is null, "stream preserves text and excludes absolute source path");
    var ackArgs = Wire.Node(new { targetId, requestId = "delivery-test-1", refId, status = "attached" });
    var badAck = new HttpRequestMessage(HttpMethod.Post, app.Origin + "/app/ack") { Content = new StringContent(ackArgs!.ToJsonString(), Encoding.UTF8, "application/json") };
    badAck.Headers.Authorization = new AuthenticationHeaderValue("Bearer", new string('0', 64));
    Check((await http.SendAsync(badAck)).StatusCode == HttpStatusCode.Unauthorized, "wrong App token cannot acknowledge");
    var ack = new HttpRequestMessage(HttpMethod.Post, app.Origin + "/app/ack") { Content = new StringContent(ackArgs.ToJsonString(), Encoding.UTF8, "application/json") };
    ack.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    var accepted = JsonNode.Parse(await (await http.SendAsync(ack)).Content.ReadAsStringAsync());
    Check(accepted?["status"]?.GetValue<string>() == "attached", "only authenticated host acknowledgement marks attached");
    var stored = await Call("reference.get", Wire.Node(new { refId }));
    Check(stored?["text"]?.GetValue<string>() == text && stored?["autoSaveAvailable"] is null, "immutable content and unknown AutoSave state");
    await Fault(async () => { await Call("reference.capture", Selection(new string('x', CoreState.MaxSingleReferenceBytes))); }, "reference_too_large");
    Check((await Call("reference.list"))?.AsArray().Count == 1, "oversized capture makes no stored reference");

    var officePipe = "office-in-editor-office-" + Environment.ProcessId + "-" + testId;
    var officeServer = Task.Run(async () =>
    {
        await using var pipe = new NamedPipeServerStream(officePipe, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        await pipe.WaitForConnectionAsync(deadline.Token);
        var request = JsonNode.Parse((await Wire.ReadLineAsync(pipe, deadline.Token))!);
        await Wire.WriteAsync(pipe, new { id = request!["id"]!.GetValue<string>(), ok = true, result = Selection("Mock Office selection") }, deadline.Token);
    });
    await Call("office.register", Wire.Node(new { officeId = "test-office", application = "word", pipeName = officePipe, pid = Environment.ProcessId }));
    var captured = await Call("office.capture", Wire.Node(new { officeId = "test-office" })); await officeServer;
    Check(captured?["documentId"]?.GetValue<string>() == "test-document", "Office callback pipe captures without launching Office");
    await using (var reloaded = new CoreState(Path.Combine(testRoot, "data")))
        Check((await reloaded.InvokeAsync("reference.list", null, Environment.ProcessId, CancellationToken.None))?.AsArray().Count == 2, "references survive broker replacement");

    for (var index = 0; index < 32; index++) await Call("reference.deliver", Wire.Node(new { refId, targetId, requestId = "queue-" + index }));
    await Fault(async () => { await Call("reference.deliverSelection", Wire.Node(new { selection = Selection("Do not capture"), targetId, requestId = "queue-overflow" })); }, "quota_exceeded");
    Check((await Call("reference.list"))?.AsArray().Count == 2, "full queue does not leave an extra capture");
    await Fault(async () => { await Call("reference.release", Wire.Node(new { refId, confirm = false })); }, "confirmation_required");
    Directory.CreateDirectory(testRoot); var original = Path.Combine(testRoot, "original.docx"); await File.WriteAllTextAsync(original, "original sentinel");
    await Call("reference.release", Wire.Node(new { refId, confirm = true }));
    Check(File.Exists(original) && await File.ReadAllTextAsync(original) == "original sentinel", "release never deletes the source document");
    await Fault(async () => { await Call("reference.deliverSelection", deliveryArgs); }, "reference_not_found");
    Check((await Call("reference.list"))?.AsArray().Count == 1, "released request cannot capture again");
    await Call("office.unregister", Wire.Node(new { officeId = "test-office" }));
    await Call("chat.unregister", Wire.Node(new { targetId, ownerId = "test-owner", token }));
    await Call("host.unregister", Wire.Node(new { hostId = "test-host", lease = host!["lease"]!.GetValue<string>() }));
    var stopped = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously); state.BecameIdle += () => stopped.TrySetResult();
    Check((await Call("runtime.shutdownIfIdle"))?["status"]?.GetValue<string>() == "shutting_down", "idle shutdown is explicit and bounded");
    await stopped.Task.WaitAsync(TimeSpan.FromSeconds(2));
    await WriteTests.RunAsync(Path.Combine(testRoot, "write-transactions"), Check, Fault);
    Console.WriteLine("ALL " + passed + " CORE CHECKS PASSED");
}
finally
{
    // Delete only this test's verified, generated temporary root, never any source document outside it.
    var expected = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "OfficeInEditor-tests-" + testId));
    if (Directory.Exists(testRoot) && Path.GetFullPath(testRoot) == expected && Guid.TryParseExact(testId, "N", out _) &&
        (File.GetAttributes(testRoot) & FileAttributes.ReparsePoint) == 0) Directory.Delete(testRoot, true);
}
