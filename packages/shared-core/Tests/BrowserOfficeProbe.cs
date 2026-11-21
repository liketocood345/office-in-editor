using System.Diagnostics;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;

// A real standard-browser capability experiment, never a product editor.
// Only queries Office registration. Never creates a COM object, opens an Office
// document, captures pixels or sends input to an existing application.
internal static class BrowserOfficeProbe
{
    private sealed record Candidate(string Application, string ProgId, string Mime, string? Clsid);

    internal static async Task<int> RunAsync()
    {
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException();
        var executable = FindBrowser();
        var candidates = new[]
        {
            ReadCandidate("word", "Word.Document.12", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            ReadCandidate("excel", "Excel.Sheet.12", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            ReadCandidate("powerpoint", "PowerPoint.Show.12", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        };
        var probeId = Guid.NewGuid().ToString("N");
        var profile = Path.Combine(Path.GetTempPath(), "OfficeInEditor-browser-probe-" + probeId);
        var report = new TaskCompletionSource<JsonObject>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        var builder = WebApplication.CreateSlimBuilder(new WebApplicationOptions
        {
            Args = [], EnvironmentName = "Production", ApplicationName = typeof(BrowserOfficeProbe).Assembly.FullName,
            ContentRootPath = AppContext.BaseDirectory
        });
        builder.Configuration.Sources.Clear();
        builder.Configuration.AddInMemoryCollection();
        builder.Logging.ClearProviders();
        builder.WebHost.UseSetting(WebHostDefaults.PreventHostingStartupKey, "true");
        builder.WebHost.ConfigureKestrel(options =>
        {
            options.Listen(IPAddress.Loopback, 0);
            options.Limits.MaxRequestBodySize = 8192;
            options.Limits.MaxConcurrentConnections = 4;
            options.AddServerHeader = false;
        });
        await using var app = builder.Build();
        string origin = "";
        app.Run(async context =>
        {
            if (!IPAddress.IsLoopback(context.Connection.RemoteIpAddress ?? IPAddress.None)
                || context.Request.Host.Value != new Uri(origin).Authority)
            { context.Response.StatusCode = 403; return; }
            context.Response.Headers.CacheControl = "no-store";
            context.Response.Headers["X-Content-Type-Options"] = "nosniff";
            if (context.Request.Method == "GET" && context.Request.Path == "/")
            {
                var nonce = Convert.ToBase64String(RandomNumberGenerator.GetBytes(18));
                // object-src deliberately permits the local origin: this experiment
                // must not manufacture a negative result by blocking OBJECT in CSP.
                context.Response.Headers.ContentSecurityPolicy = "default-src 'self'; script-src 'nonce-" + nonce
                    + "'; style-src 'nonce-" + nonce + "'; object-src 'self'; base-uri 'none'; form-action 'none'";
                context.Response.ContentType = "text/html; charset=utf-8";
                await context.Response.WriteAsync(Html(candidates, nonce), context.RequestAborted);
                return;
            }
            if (context.Request.Method == "POST" && context.Request.Path == "/report"
                && context.Request.Headers.Origin == origin)
            {
                var result = await JsonNode.ParseAsync(context.Request.Body, cancellationToken: context.RequestAborted) as JsonObject;
                if (result is null || result["candidates"] is not JsonArray { Count: 3 })
                { context.Response.StatusCode = 400; return; }
                report.TrySetResult(result);
                context.Response.StatusCode = 204;
                return;
            }
            context.Response.StatusCode = 404;
        });
        Process? browser = null;
        var cleaned = false;
        try
        {
            await app.StartAsync(timeout.Token);
            origin = app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>()!.Addresses.Single().TrimEnd('/');
            Directory.CreateDirectory(profile);
            var start = new ProcessStartInfo(executable)
            {
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true
            };
            foreach (var argument in new[]
            {
                "--headless=new", "--no-first-run", "--no-default-browser-check", "--disable-extensions",
                "--disable-background-networking", "--disable-component-update", "--disable-sync",
                "--no-proxy-server", "--user-data-dir=" + profile, "--dump-dom", "--virtual-time-budget=1500", origin + "/"
            }) start.ArgumentList.Add(argument);
            browser = Process.Start(start) ?? throw new InvalidOperationException("Browser probe could not start.");
            // Drain both streams without retaining or publishing the browser's DOM,
            // profile paths or diagnostics. The structured result is posted once.
            var stdout = DrainAsync(browser.StandardOutput, timeout.Token);
            var stderr = DrainAsync(browser.StandardError, timeout.Token);
            var observed = await report.Task.WaitAsync(timeout.Token);
            await browser.WaitForExitAsync(timeout.Token);
            await Task.WhenAll(stdout, stderr);
            var checks = observed["candidates"]!.AsArray();
            var nativeOfficeEmbedded = checks.Any(item => item?["fallbackVisible"]?.GetValue<bool>() == false
                && observed["activeXAvailable"]?.GetValue<bool>() == true);
            // No native Office is launched by this probe. A positive capability
            // signal would still require a separate real edit/save acceptance test.
            var result = new
            {
                experiment = "local-office-ole-in-standard-browser",
                serviceStarted = true,
                serviceOrigin = origin,
                browser = Path.GetFileName(executable),
                browserVersion = FileVersionInfo.GetVersionInfo(executable).ProductVersion,
                registeredApplications = candidates.Select(item => new { item.Application, registered = item.Clsid is not null }),
                observed,
                nativeOfficeEmbedded,
                outcome = nativeOfficeEmbedded ? "requires_edit_acceptance" : "blocked_browser_ole_host_unavailable",
                officeStartedByProbe = false,
                userDocumentsOpened = 0,
                pixelTransfer = false
            };
            Console.WriteLine(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
            return 0; // Successful experiment is distinct from successful embedding.
        }
        finally
        {
            if (browser is not null)
            {
                try
                {
                    // Only this process, launched with an isolated profile, is owned
                    // by the experiment. Never terminate a named browser or Office.
                    if (!browser.HasExited) { browser.Kill(entireProcessTree: true); browser.WaitForExit(5000); }
                }
                finally { browser.Dispose(); }
            }
            using var stop = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            await app.StopAsync(stop.Token);
            cleaned = CleanProfile(profile, probeId);
            Console.WriteLine(JsonSerializer.Serialize(new { temporaryBrowserProfileRemoved = cleaned, serviceStopped = true }));
        }
    }

    private static Candidate ReadCandidate(string application, string progId, string mime)
    {
        using var key = Registry.ClassesRoot.OpenSubKey(progId + @"\CLSID", writable: false);
        return new(application, progId, mime, Guid.TryParse(key?.GetValue(null) as string, out var id) ? id.ToString("B") : null);
    }

    private static string FindBrowser()
    {
        foreach (var path in new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft", "Edge", "Application", "msedge.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft", "Edge", "Application", "msedge.exe")
        }) if (File.Exists(path)) return path;
        throw new InvalidOperationException("A standard Microsoft Edge installation is required for this experiment.");
    }

    private static async Task DrainAsync(StreamReader reader, CancellationToken cancellation)
    {
        var buffer = new char[4096];
        while (await reader.ReadAsync(buffer.AsMemory(), cancellation) != 0) { }
    }

    private static string Html(Candidate[] candidates, string nonce)
    {
        var objects = string.Join("\n", candidates.Select(item =>
            "<section><h2>" + item.Application + "</h2><object id=\"" + item.Application + "\" type=\"" + item.Mime
            + "\" classid=\"clsid:" + (item.Clsid ?? "").Trim('{', '}')
            + "\"><p id=\"fallback-" + item.Application + "\">Office native object was not activated by the browser.</p></object></section>"));
        return "<!doctype html><html><head><meta charset=\"utf-8\"><title>Office browser hosting experiment</title>"
            + "<style nonce=\"" + nonce + "\">body{font:16px sans-serif}object{display:block;width:640px;height:180px;border:1px solid gray}</style></head><body>"
            + "<h1>Local Office browser-host experiment</h1><p>This page is a capability probe, not an Office editor.</p>" + objects
            + "<script nonce=\"" + nonce + "\">const declared=" + JsonSerializer.Serialize(candidates, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase }) + ";"
            + "setTimeout(async()=>{const result={activeXAvailable:typeof window.ActiveXObject==='function',candidates:declared.map(c=>{"
            + "const f=document.getElementById('fallback-'+c.application);const r=f.getBoundingClientRect();return {application:c.application,"
            + "officeMimeRegistered:Array.from(navigator.mimeTypes).some(m=>m.type===c.mime),fallbackVisible:r.width>0&&r.height>0};})};"
            + "document.body.dataset.completed='true';await fetch('/report',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(result)});},300);"
            + "</script></body></html>";
    }

    private static bool CleanProfile(string profile, string probeId)
    {
        var root = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar);
        var expected = Path.GetFullPath(Path.Combine(root, "OfficeInEditor-browser-probe-" + probeId));
        if (!string.Equals(Path.GetFullPath(profile), expected, StringComparison.OrdinalIgnoreCase)
            || !expected.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) return false;
        if (!Directory.Exists(expected)) return true;
        try
        {
            // Inspect each directory before descending. Preserve anything crossing a
            // reparse point; recursive deletion is restricted to this exact fixture.
            var pending = new Stack<string>(); pending.Push(expected);
            var count = 0;
            while (pending.TryPop(out var current))
            {
                if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) return false;
                foreach (var entry in Directory.EnumerateFileSystemEntries(current))
                {
                    if (++count > 10000) return false;
                    var attributes = File.GetAttributes(entry);
                    if ((attributes & FileAttributes.ReparsePoint) != 0) return false;
                    if ((attributes & FileAttributes.Directory) != 0) pending.Push(entry);
                }
            }
            Directory.Delete(expected, recursive: true);
            return true;
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }
}
