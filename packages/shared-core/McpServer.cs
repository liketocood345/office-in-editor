using System.Text;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

public sealed class McpServer
{
    private const string ResourceUri = "ui://office-in-editor/chat";
    private readonly string ownerId = "mcp-" + Guid.NewGuid().ToString("N");
    private readonly string hostId = Environment.GetEnvironmentVariable("OFFICE_IN_EDITOR_HOST_ID") ?? "";
    private string? ownerLease;
    private static readonly HashSet<string> Versions = ["2024-11-05", "2025-03-26", "2025-06-18", "2025-11-25", "2026-07-28"];

    public async Task RunAsync()
    {
        using var input = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false, true));
        using var output = Console.OpenStandardOutput();
        var reader = new McpLineReader(input);
        try
        {
            while (await reader.ReadAsync() is { } line)
            {
                JsonNode? id = null;
                try
                {
                    var request = JsonNode.Parse(line) ?? throw new CoreFault("invalid_request", "MCP request is invalid.");
                    if (request["id"] is null) continue; // initialized/cancelled notifications have no response.
                    id = request["id"]!.DeepClone();
                    var method = Wire.Text(request, "method");
                    var result = await HandleAsync(method, request["params"]);
                    await Wire.WriteAsync(output, new { jsonrpc = "2.0", id, result }, CancellationToken.None);
                }
                catch (Exception error)
                {
                    var code = error is CoreFault { Code: "method_not_found" } ? -32601 : error is System.Text.Json.JsonException ? -32700 : -32603;
                    await Wire.WriteAsync(output, new { jsonrpc = "2.0", id, error = new { code, message = error is CoreFault fault ? fault.Message : "The MCP request could not be completed." } }, CancellationToken.None);
                }
            }
        }
        finally
        {
            if (ownerLease is not null)
                try { await PipeClient.CallAsync(RuntimeIdentity.PipeName, "runtime.releaseOwner", Wire.Node(new { ownerId, lease = ownerLease }), CancellationToken.None); }
                catch (Exception) { /* Process ownership also releases through the broker's process-exit event. */ }
        }
    }
    private async Task<JsonNode?> HandleAsync(string method, JsonNode? parameters)
    {
        switch (method)
        {
            case "initialize":
                await PipeClient.EnsureAsync(CancellationToken.None);
                ownerLease = (await PipeClient.CallAsync(RuntimeIdentity.PipeName, "runtime.acquireOwner", Wire.Node(new { ownerId }), CancellationToken.None))?["lease"]?.GetValue<string>();
                var requested = Wire.Text(parameters, "protocolVersion", 64, false);
                return Wire.Node(new { protocolVersion = Versions.Contains(requested) ? requested : "2025-11-25",
                    capabilities = new { tools = new { listChanged = false }, resources = new { subscribe = false, listChanged = false } },
                    serverInfo = new { name = "office-in-editor", version = RuntimeIdentity.Version },
                    instructions = "Use explicitly registered Office instances. Inspect documents before capability-gated edits. Edit/save operation IDs are durable and never replay unknown outcomes; inspect the operation instead. Revert refuses subsequent user changes. Native AutoSave local redirection is unavailable. The editor tab hosts the Office window. Word save uses the same SaveObject path as the Save button; periodic save silently clicks that path every 60s. Selection and agent edits stay on the Office add-in and MCP. Register this chat with office_register_chat, then Word Add to Chat fills the composer without sending." });
            case "ping": return new JsonObject();
            case "tools/list": return new JsonObject { ["tools"] = ToolDefinitions() };
            case "resources/list":
                return Wire.Node(new { resources = new[] { new { uri = ResourceUri, name = "Office reference inbox", mimeType = "text/html;profile=mcp-app", description = "An explicitly registered chat App receives selection references without sending messages." } } });
            case "resources/templates/list": return Wire.Node(new { resourceTemplates = Array.Empty<object>() });
            case "resources/read": return await ReadResourceAsync(parameters);
            case "tools/call":
                try { return await CallToolAsync(parameters); }
                catch (Exception error)
                {
                    var message = error is CoreFault fault ? fault.Code + ": " + fault.Message : "operation_failed: The operation did not complete. Do not automatically replay it.";
                    return Wire.Node(new { isError = true, content = new[] { new { type = "text", text = message } } });
                }
            default: throw new CoreFault("method_not_found", "MCP method is not supported.");
        }
    }
    private async Task<JsonNode?> BrokerAsync(string method, JsonNode? parameters = null)
    {
        await PipeClient.EnsureAsync(CancellationToken.None);
        return await PipeClient.CallAsync(RuntimeIdentity.PipeName, method, parameters, CancellationToken.None);
    }
    private async Task NoteRouteAsync(string? targetId, JsonNode? arguments = null, string? targetEpoch = null)
    {
        try { await BrokerAsync("routing.touch", Wire.Node(new { ownerId, targetId = targetId ?? "", targetEpoch = targetEpoch ?? "",
            officeId = Wire.Text(arguments, "officeId", 128, false), documentId = Wire.Text(arguments, "documentId", 160, false) })); }
        catch { }
    }
    private static string MetaTargetId(JsonNode? parameters)
    {
        return parameters?["_meta"] is JsonObject meta && meta["officeInEditor"] is JsonObject office
            ? Wire.Text(office, "targetId", 128, false)
            : "";
    }
    private async Task<JsonNode?> CallToolAsync(JsonNode? parameters)
    {
        var name = Wire.Text(parameters, "name");
        var args = parameters?["arguments"] as JsonObject ?? new JsonObject();
        if (name == "office_register_chat")
        {
            var label = Wire.Text(args, "label", 160, false);
            var target = await BrokerAsync("chat.register", Wire.Node(new { ownerId, hostId, label = label == "" ? "Office reference chat" : label }));
            await NoteRouteAsync(target?["id"]?.GetValue<string>());
            var publicTarget = target!.DeepClone(); publicTarget.AsObject().Remove("token"); publicTarget.AsObject().Remove("appOrigin");
            return new JsonObject
            {
                ["content"] = Wire.Node(new[] { new { type = "text", text = "Registered this chat." } }),
                ["structuredContent"] = new JsonObject { ["target"] = publicTarget, ["registeredChatsOnly"] = true },
                ["_meta"] = new JsonObject { ["officeInEditor"] = new JsonObject
                {
                    ["targetId"] = target["id"]!.DeepClone(), ["token"] = target["token"]!.DeepClone(),
                    ["appOrigin"] = target["appOrigin"]!.DeepClone(), ["epoch"] = target["epoch"]!.DeepClone(), ["label"] = target["label"]!.DeepClone()
                } }
            };
        }
        await NoteRouteAsync(MetaTargetId(parameters), args,
            parameters?["_meta"]?["officeInEditor"]?["epoch"]?.GetValue<string>());
        var method = name switch
        {
            "office_list_documents" => "office.list", "office_capture_selection" => "office.capture", "office_list_references" => "reference.list",
            "office_get_reference" => "reference.get", "office_capabilities" => "health", "office_app_ack" => "delivery.ack",
            "office_inspect_document" => "office.inspect_document", "office_apply_edit" => "office.apply_edit", "office_get_operation" => "office.get_operation",
            "office_revert_edit" => "office.revert_edit", "office_save_document" => "office.save_document", "office_release_operation" => "operation.release",
            _ => throw new CoreFault("method_not_found", "This tool is not supported.")
        };
        var result = await BrokerAsync(method, args);
        if (name == "office_get_reference" && result is JsonObject reference) reference.Remove("sourcePath");
        return new JsonObject { ["isError"] = name is "office_apply_edit" or "office_revert_edit" or "office_save_document" && result?["status"]?.GetValue<string>() == "unknown",
            ["content"] = Wire.Node(new[] { new { type = "text", text = result?.ToJsonString() ?? "null" } }),
            ["structuredContent"] = new JsonObject { ["result"] = result?.DeepClone() } };
    }
    private async Task<JsonNode?> ReadResourceAsync(JsonNode? value)
    {
        if (Wire.Text(value, "uri", 128) != ResourceUri) throw new CoreFault("resource_not_found", "Resource is unavailable.");
        var health = await BrokerAsync("health");
        var origin = Wire.Text(health, "appOrigin", 128);
        var path = Path.Combine(AppContext.BaseDirectory, "mcp-app.html");
        string html;
        if (File.Exists(path) && new FileInfo(path).Length <= 1024 * 1024) html = await File.ReadAllTextAsync(path, Encoding.UTF8);
        else html = "<!doctype html><html><meta charset=\"utf-8\"><body>Office reference App is unavailable in this build. No reference has been attached.</body></html>";
        return new JsonObject { ["contents"] = new JsonArray(new JsonObject
        {
            ["uri"] = ResourceUri, ["mimeType"] = "text/html;profile=mcp-app", ["text"] = html,
            ["_meta"] = new JsonObject { ["ui"] = new JsonObject { ["csp"] = new JsonObject { ["connectDomains"] = Wire.Node(new[] { origin }), ["resourceDomains"] = new JsonArray() } } }
        }) };
    }
    private static JsonArray ToolDefinitions()
    {
        JsonObject Tool(string name, string description, bool readOnly, JsonObject properties, string[] required, bool app = false)
        {
            var tool = new JsonObject { ["name"] = name, ["description"] = description,
                ["inputSchema"] = new JsonObject { ["type"] = "object", ["properties"] = properties, ["required"] = Wire.Node(required), ["additionalProperties"] = false },
                ["annotations"] = Wire.Node(new { readOnlyHint = readOnly, destructiveHint = name == "office_release_operation",
                    idempotentHint = readOnly || name is "office_apply_edit" or "office_save_document" or "office_revert_edit", openWorldHint = false }) };
            if (name == "office_register_chat" || app)
                tool["_meta"] = new JsonObject { ["ui"] = new JsonObject { ["resourceUri"] = ResourceUri, ["visibility"] = Wire.Node(app ? new[] { "app" } : new[] { "model", "app" }) }, ["ui/resourceUri"] = ResourceUri };
            return tool;
        }
        static JsonObject Field(int max = 160) => new() { ["type"] = "string", ["maxLength"] = max };
        return new JsonArray(
            Tool("office_list_documents", "List live, explicitly registered Office Add-in instances. Does not launch Office.", true, new(), []),
            Tool("office_capture_selection", "Capture the current native Office selection as one bounded immutable reference. Does not edit or save the document.", false, new() { ["officeId"] = Field() }, ["officeId"]),
            Tool("office_list_references", "List immutable selection reference metadata.", true, new(), []),
            Tool("office_get_reference", "Read the captured content and semantic locator of a reference. Absolute source paths are excluded.", true, new() { ["refId"] = Field(32) }, ["refId"]),
            Tool("office_inspect_document", "Inspect an Office instance's bound documents and actual per-document edit/manual-save/revert capabilities. Does not edit or save.", true,
                new() { ["officeId"] = Field(128), ["documentId"] = Field() }, ["officeId"]),
            Tool("office_apply_edit", "Apply one capability-approved edit with a caller-generated unique operationId. Core durably stores recovery before dispatch. If status is unknown, inspect the operation; never retry under a fresh ID.", false,
                new() { ["officeId"] = Field(128), ["documentId"] = Field(), ["operationId"] = Field(128), ["operation"] = new JsonObject
                { ["type"] = "object", ["properties"] = new JsonObject { ["kind"] = new JsonObject { ["type"] = "string", ["enum"] = Wire.Node(new[] { "word.insert_text", "excel.set_range_values", "excel.set_range_formulas", "powerpoint.add_slide", "powerpoint.add_textbox" }) }, ["args"] = new JsonObject { ["type"] = "object" } },
                    ["required"] = Wire.Node(new[] { "kind", "args" }), ["additionalProperties"] = false } }, ["officeId", "documentId", "operationId", "operation"]),
            Tool("office_get_operation", "Read a durable operation outcome without replaying it. Unknown outcomes require inspection, not another write.", true,
                new() { ["operationId"] = Field(128), ["documentId"] = Field() }, ["operationId"]),
            Tool("office_revert_edit", "Explicitly revert a confirmed applied operation only when Office verifies its post-edit hash still matches. Refuses to overwrite subsequent user changes.", false,
                new() { ["officeId"] = Field(128), ["documentId"] = Field(), ["operationId"] = Field(128) }, ["officeId", "documentId", "operationId"]),
            Tool("office_save_document", "Explicit manual Save through the bound Office document (same path as the Save button / IOleClientSite.SaveObject). Requires a unique operationId. Native AutoSave redirection is not implemented. Periodic silent Save is a separate host timer. Unknown save outcomes are not replayed.", false,
                new() { ["officeId"] = Field(128), ["documentId"] = Field(), ["operationId"] = Field(128) }, ["officeId", "documentId", "operationId"]),
            Tool("office_release_operation", "Release recovery ONLY after the user explicitly confirms it is no longer needed. The irreversible release retains a small deduplication tombstone and never deletes the original document.", false,
                new() { ["documentId"] = Field(), ["operationId"] = Field(128), ["confirm"] = new JsonObject { ["type"] = "boolean", ["const"] = true } }, ["documentId", "operationId", "confirm"]),
            Tool("office_capabilities", "Report supported operations. The editor tab hosts the Office window; selection and agent edits stay on the Office add-in and MCP. Native AutoSave local redirection remains unavailable.", true, new(), []),
            Tool("office_register_chat", "Register an Office reference App for this specific tool invocation. Only explicitly registered chats can be targeted. Open the returned App; it only updates context and never submits a message.", false, new() { ["label"] = Field() }, []),
            Tool("office_app_ack", "App-only acknowledgement after the App host accepts updateModelContext; never call on behalf of a model.", false,
                new() { ["targetId"] = Field(32), ["token"] = Field(64), ["requestId"] = Field(128), ["refId"] = Field(32), ["status"] = new JsonObject { ["type"] = "string", ["enum"] = Wire.Node(new[] { "attached", "unsupported", "failed" }) } },
                ["targetId", "token", "requestId", "refId", "status"], true));
    }

    private sealed class McpLineReader(StreamReader stream)
    {
        private readonly char[] buffer = new char[4096];
        private int position, length;
        public async Task<string?> ReadAsync()
        {
            var result = new StringBuilder();
            while (true)
            {
                if (position == length) { length = await stream.ReadAsync(buffer); position = 0; if (length == 0) return result.Length == 0 ? null : result.ToString(); }
                var newline = Array.IndexOf(buffer, '\n', position, length - position);
                var end = newline < 0 ? length : newline;
                if (result.Length + end - position > Wire.MaxLineBytes) throw new CoreFault("message_too_large", "MCP message exceeds the protocol limit.");
                result.Append(buffer, position, end - position); position = newline < 0 ? length : newline + 1;
                if (newline >= 0)
                {
                    var line = result.ToString().TrimEnd('\r');
                    if (Encoding.UTF8.GetByteCount(line) > Wire.MaxLineBytes) throw new CoreFault("message_too_large", "MCP message exceeds the protocol limit.");
                    return line;
                }
            }
        }
    }
}
