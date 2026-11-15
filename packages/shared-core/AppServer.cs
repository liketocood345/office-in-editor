using System.Net;
using System.Text.Json.Nodes;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace OfficeInEditor.Core;

public sealed class AppServer(CoreState state) : IAsyncDisposable
{
    private WebApplication? application;
    public string Origin { get; private set; } = "";
    public async Task StartAsync(CancellationToken cancellation)
    {
        var builder = WebApplication.CreateSlimBuilder(new WebApplicationOptions { Args = [], ApplicationName = typeof(AppServer).Assembly.FullName,
            ContentRootPath = AppContext.BaseDirectory, EnvironmentName = "Production" });
        // Never consume a workspace's appsettings or inherited endpoint configuration.
        builder.Configuration.Sources.Clear();
        builder.Configuration.AddInMemoryCollection();
        builder.WebHost.UseSetting(WebHostDefaults.PreventHostingStartupKey, "true");
        builder.Logging.ClearProviders();
        builder.WebHost.ConfigureKestrel(options =>
        {
            options.Listen(IPAddress.Loopback, 0);
            options.Limits.MaxRequestBodySize = 64 * 1024;
            options.Limits.RequestHeadersTimeout = TimeSpan.FromSeconds(10);
            options.Limits.MaxConcurrentConnections = 64;
            options.Limits.MaxConcurrentUpgradedConnections = 32;
            options.AddServerHeader = false;
        });
        application = builder.Build();
        application.Run(HandleAsync);
        await application.StartAsync(cancellation);
        var server = application.Services.GetRequiredService<IServer>();
        Origin = server.Features.Get<IServerAddressesFeature>()!.Addresses.Single().TrimEnd('/');
    }
    private async Task HandleAsync(HttpContext context)
    {
        try
        {
            if (!IPAddress.IsLoopback(context.Connection.RemoteIpAddress ?? IPAddress.None) ||
                context.Request.Host.Value != new Uri(Origin).Authority || context.Request.Query.ContainsKey("token"))
            { context.Response.StatusCode = 403; return; }
            var origin = context.Request.Headers.Origin.ToString();
            if (origin != "" && !AllowedOrigin(origin)) { context.Response.StatusCode = 403; return; }
            context.Response.Headers.CacheControl = "no-store";
            context.Response.Headers["X-Content-Type-Options"] = "nosniff";
            if (origin != "")
            {
                context.Response.Headers.AccessControlAllowOrigin = origin;
                context.Response.Headers.Vary = "Origin";
            }
            if (context.Request.Method == "OPTIONS")
            {
                context.Response.Headers.AccessControlAllowMethods = "GET, POST, OPTIONS";
                context.Response.Headers.AccessControlAllowHeaders = "Authorization, Content-Type";
                context.Response.Headers["Access-Control-Allow-Private-Network"] = "true";
                context.Response.StatusCode = 204; return;
            }
            var authorization = context.Request.Headers.Authorization.ToString();
            if (!authorization.StartsWith("Bearer ", StringComparison.Ordinal) || authorization.Length != 71)
                throw new CoreFault("unauthorized", "App authorization is required.");
            var token = authorization[7..];
            if (context.Request.Path == "/app/events" && context.Request.Method == "GET")
            {
                var targetId = context.Request.Query["targetId"].ToString();
                var connection = state.ConnectApp(targetId, token);
                var signals = connection.Reader;
                try
                {
                    context.Response.ContentType = "text/event-stream";
                    context.Response.Headers["X-Accel-Buffering"] = "no";
                    await context.Response.StartAsync(context.RequestAborted);
                    await context.Response.WriteAsync(": connected\n\n", context.RequestAborted);
                    await context.Response.Body.FlushAsync(context.RequestAborted);
                    var emitted = new HashSet<string>(StringComparer.Ordinal);
                    while (!context.RequestAborted.IsCancellationRequested)
                    {
                        foreach (var entry in state.PendingEvents(targetId, token))
                        {
                            var requestId = Wire.Text(entry, "requestId", 128);
                            if (!emitted.Add(requestId)) continue;
                            await context.Response.WriteAsync("data: " + entry.ToJsonString() + "\n\n", context.RequestAborted);
                            await context.Response.Body.FlushAsync(context.RequestAborted);
                        }
                        if (!await signals.WaitToReadAsync(context.RequestAborted)) break;
                        while (signals.TryRead(out _)) { }
                    }
                }
                finally { state.DisconnectApp(targetId, connection.Id); }
                return;
            }
            if (context.Request.Path == "/app/ack" && context.Request.Method == "POST")
            {
                var value = await JsonNode.ParseAsync(context.Request.Body, cancellationToken: context.RequestAborted);
                await context.Response.WriteAsJsonAsync(state.Ack(value, token), cancellationToken: context.RequestAborted);
                return;
            }
            context.Response.StatusCode = 404;
        }
        catch (OperationCanceledException) when (context.RequestAborted.IsCancellationRequested) { }
        catch (Exception error)
        {
            if (context.Response.HasStarted) { context.Abort(); return; }
            context.Response.StatusCode = error is CoreFault { Code: "unauthorized" } ? 401 : 400;
            await context.Response.WriteAsJsonAsync(new { error = error is CoreFault fault ? fault.Code : "request_failed" });
        }
    }
    private static bool AllowedOrigin(string origin) => origin == "null" || Uri.TryCreate(origin, UriKind.Absolute, out var value) &&
        value.AbsolutePath == "/" && value.Query == "" && value.Fragment == "" && value.UserInfo == "" &&
        (value.Scheme == "https" || value.Scheme == "http" && IPAddress.TryParse(value.Host, out var ip) && IPAddress.IsLoopback(ip));
    public async ValueTask DisposeAsync()
    {
        if (application is null) return;
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(3));
        try { await application.StopAsync(timeout.Token); } catch (OperationCanceledException) { }
        await application.DisposeAsync();
    }
}
