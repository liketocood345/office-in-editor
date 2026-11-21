using System.Text;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

// A short-lived observer. It never returns prompt text, extra context, approval changes,
// follow-up messages, or commands from an incoming event, and never logs the raw payload.
internal static class CursorHookBridge
{
    internal static JsonObject? Normalize(JsonNode? input)
    {
        var eventName = Wire.Text(input, "hook_event_name", 64);
        if (eventName is not ("beforeSubmitPrompt" or "beforeMCPExecution" or "afterMCPExecution" or "postToolUseFailure" or "sessionEnd")) return null;
        var conversationId = Wire.Text(input, "conversation_id", 160);
        var generationId = Wire.Text(input, "generation_id", 160, false);
        var result = new JsonObject { ["event"] = eventName, ["conversationId"] = conversationId, ["generationId"] = generationId };
        if (eventName is "beforeSubmitPrompt" or "sessionEnd") return result;
        var toolName = Wire.Text(input, "tool_name", 128, false);
        var server = Wire.Text(input, "mcp_server_name", 128, false);
        if (toolName.StartsWith("MCP:", StringComparison.Ordinal)) toolName = toolName[4..];
        if (server != "office-in-editor" || !toolName.StartsWith("office_", StringComparison.Ordinal)) return null;
        result["toolName"] = toolName;
        var args = ParseObject(input?["tool_input"]);
        foreach (var name in new[] { "officeId", "documentId", "operationId" })
        {
            var id = Wire.Text(args, name, 160, false);
            if (id.Length > 0) result[name] = id;
        }
        if (eventName == "afterMCPExecution" && toolName == "office_register_chat")
        {
            var reply = ParseObject(input?["result_json"]);
            var target = reply?["structuredContent"]?["target"] ?? reply?["result"]?["structuredContent"]?["target"];
            var id = Wire.Text(target, "id", 128, false);
            if (id.Length > 0) result["targetId"] = id;
        }
        return result;
    }
    private static JsonNode? ParseObject(JsonNode? value)
    {
        if (value is JsonObject) return value;
        if (value is JsonValue scalar && scalar.TryGetValue<string>(out var text) && Encoding.UTF8.GetByteCount(text) <= Wire.MaxLineBytes)
            try { return JsonNode.Parse(text) as JsonObject; } catch { }
        return null;
    }
    internal static async Task RunAsync()
    {
        try
        {
            var args = Environment.GetCommandLineArgs();
            var index = Array.IndexOf(args, "--hook-context");
            if (index < 0 || index + 1 >= args.Length) throw new CoreFault("invalid_params", "Hook context is required.");
            var path = Path.GetFullPath(args[index + 1]);
            if (path.StartsWith("\\\\", StringComparison.Ordinal) || !File.Exists(path) || new FileInfo(path).Length > 16384 ||
                (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) throw new CoreFault("invalid_params", "Hook context is invalid.");
            var context = JsonNode.Parse(await File.ReadAllTextAsync(path, Encoding.UTF8));
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            await using var input = Console.OpenStandardInput();
            using var content = new MemoryStream();
            var buffer = new byte[4096];
            int count;
            while ((count = await input.ReadAsync(buffer, deadline.Token)) > 0)
            {
                if (content.Length + count > Wire.MaxLineBytes) throw new CoreFault("message_too_large", "Hook event exceeds the budget.");
                content.Write(buffer, 0, count);
            }
            var incoming = JsonNode.Parse(content.ToArray());
            var normalized = Normalize(incoming);
            if (normalized is not null)
            {
                foreach (var field in new[] { "hostId", "hostLease", "hostEpoch" }) normalized[field] = Wire.Text(context, field, 128);
                try { await PipeClient.CallAsync(RuntimeIdentity.PipeName, "routing.hook", normalized, deadline.Token); }
                catch { /* Observing routing never starts Core or blocks an otherwise authorized tool. */ }
            }
            var eventName = Wire.Text(incoming, "hook_event_name", 64);
            // Required neutral output schemas. Other policy hooks remain responsible for their
            // own decisions; this observer never denies, asks, alters arguments or submits.
            Console.WriteLine(eventName == "beforeMCPExecution" ? "{\"permission\":\"allow\"}" : eventName == "beforeSubmitPrompt" ? "{\"continue\":true}" : "{}");
        }
        catch { Environment.ExitCode = 1; /* Nonzero is fail-open for this observer; no private diagnostics. */ }
    }
}

public sealed partial class CoreState
{
    private sealed record HookConversation(string TargetId, string GenerationId);
    private readonly Dictionary<string, HookConversation> hookConversations = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> hookLatestGeneration = new(StringComparer.Ordinal);

    private Host AuthorizeHookHost(JsonNode? value)
    {
        var id = Wire.Text(value, "hostId", 128);
        if (!hosts.TryGetValue(id, out var host) || host.Kind != "cursor" ||
            !Wire.TokenEquals(Wire.Text(value, "hostLease", 128), host.Lease) || Wire.Text(value, "hostEpoch", 128) != host.Epoch)
            throw new CoreFault("unauthorized", "The hook host capability is invalid.");
        return host;
    }
    private async Task<JsonNode?> HookActivityAsync(JsonNode? value, CancellationToken cancellation)
    {
        Host host;
        Chat? target = null;
        var eventName = Wire.Text(value, "event", 64);
        var conversationId = Wire.Text(value, "conversationId", 160);
        var generationId = Wire.Text(value, "generationId", 160, false);
        var toolName = Wire.Text(value, "toolName", 128, false);
        lock (sync)
        {
            host = AuthorizeHookHost(value);
            var key = host.Id + "\0" + conversationId;
            if (eventName == "sessionEnd")
            {
                hookConversations.Remove(key); hookLatestGeneration.Remove(key);
                return Wire.Node(new { observed = true });
            }
            if (eventName is "beforeSubmitPrompt" or "beforeMCPExecution")
            {
                if (!hookLatestGeneration.ContainsKey(key) && hookLatestGeneration.Count >= 256)
                    throw new CoreFault("quota_exceeded", "Hook conversation capacity is reached.");
                hookLatestGeneration[key] = generationId;
            }
            if (eventName == "afterMCPExecution" && toolName == "office_register_chat")
            {
                var id = Wire.Text(value, "targetId", 128, false);
                if (id.Length == 0 || !chats.TryGetValue(id, out target) || target.HostId != host.Id)
                    return Wire.Node(new { observed = false, reason = "registration_unattributed" });
                if (!hookConversations.ContainsKey(key) && hookConversations.Count >= 128)
                    throw new CoreFault("quota_exceeded", "Hook registration capacity is reached.");
                hookConversations[key] = new HookConversation(id, generationId);
            }
            else if (hookConversations.TryGetValue(key, out var route)) chats.TryGetValue(route.TargetId, out target);
            if (target is null || target.HostId != host.Id) return Wire.Node(new { observed = false, reason = "chat_unregistered" });
            if (hookLatestGeneration.TryGetValue(key, out var latest) && latest != generationId)
                return Wire.Node(new { observed = true, ignored = "stale_generation" });
            if (eventName == "beforeSubmitPrompt")
            {
                MarkActivity(target);
                return Wire.Node(new { observed = true, targetId = target.Id });
            }
            // A before hook already committed the source for failed calls. Late completion
            // events must not turn an old tool into the latest route again.
            if (eventName == "postToolUseFailure") return Wire.Node(new { observed = true });
            if (eventName == "afterMCPExecution" && toolName != "office_register_chat") return Wire.Node(new { observed = true });
        }
        return await TouchRouteAsync(Wire.Node(new { ownerId = target.OwnerId, targetId = target.Id, targetEpoch = target.Epoch,
            officeId = Wire.Text(value, "officeId", 128, false), documentId = Wire.Text(value, "documentId", 160, false) }), 0, cancellation, host);
    }
}
