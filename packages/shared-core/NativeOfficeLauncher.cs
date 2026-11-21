using System.Diagnostics;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

internal static class NativeOfficeLauncher
{
    const long HiddenStartGraceMs = 15_000;
    static readonly Dictionary<string, long> hiddenStarts = new(StringComparer.Ordinal);

    /// <summary>
    /// Starts Word with no document and no window, so a focus tab can have the add-in open its document out of view.
    /// True when Word is starting this way (now, or for another tab moments ago). False when the caller should open the
    /// file the usual way: not Word, or Word already runs (it may be one without the add-in).
    /// </summary>
    internal static bool StartHidden(string application)
    {
        if (application != "word") return false;
        lock (hiddenStarts)
        {
            var now = Environment.TickCount64;
            if (hiddenStarts.TryGetValue(application, out var at) && now - at < HiddenStartGraceMs) return true;
            var session = Process.GetCurrentProcess().SessionId;
            foreach (var running in Process.GetProcessesByName("WINWORD"))
                using (running) if (running.SessionId == session) return false;
            var executable = OfficeDocumentKind.FindAppPath("WINWORD.EXE");
            if (executable is null) return false;
            try
            {
                // /q: no splash; /n: no blank document. Hidden: Word shows no frame of its own.
                using var process = Process.Start(new ProcessStartInfo(executable, "/q /n") { UseShellExecute = true, WindowStyle = ProcessWindowStyle.Hidden });
                if (process is null) return false;
            }
            catch (Exception) { return false; }
            hiddenStarts[application] = now;
            DebugNdjson.Write("F", "NativeOfficeLauncher.StartHidden", "office_started_hidden", null);
            return true;
        }
    }

    public static JsonNode? Open(string path, string mode, bool fallback = false, int x = 0, int y = 0, int width = 0, int height = 0, string? sessionId = null, string? surfaceToken = null, CancellationToken cancellation = default, bool visible = true, double pixelRatio = 1, string? backupId = null, bool quiet = false)
    {
        if (mode is not ("external" or "embedded"))
            throw new CoreFault("invalid_open_mode", "Office open mode must be external or embedded.");
        if (!Path.IsPathFullyQualified(path) || path.StartsWith("\\\\", StringComparison.Ordinal) ||
            Uri.TryCreate(path, UriKind.Absolute, out var uri) && uri.IsUnc)
            throw new CoreFault("invalid_path", "Choose an existing local Office document.");
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new CoreFault("file_not_found", "The selected document is unavailable.");
        if (Path.GetFileName(fullPath).StartsWith('~'))
        {
            throw new CoreFault("office_cache_file", "This is an Office cache file and should not be opened.");
        }
        if (!OfficeDocumentKind.TryClassify(fullPath, out var kind))
            throw new CoreFault("unsupported_format", "This file is not an embeddable Office document.");
        _ = (x, y);
        if (mode == "embedded")
        {
            if (!WebviewSurface.IsToken(surfaceToken) || width < 32 || height < 32)
                throw new CoreFault("invalid_host_window", "A valid editor webview surface is required.");
            try
            {
                var hosted = OleDocumentHost.TryOpen(fullPath, kind.Application, width, height, sessionId, surfaceToken, cancellation, bindSource: false, visible: visible, jsX: x, jsY: y, pixelRatio: pixelRatio, backupId: backupId);
                if (hosted is not null) return hosted;
            }
            catch (OperationCanceledException) { throw; }
            catch (CoreFault fault)
            {
                // #region agent log
                DebugNdjson.Probe("I", "NativeOfficeLauncher.Open", fault.Code, new { code = fault.Code, kind = kind.Key });
                // #endregion
                throw;
            }
            catch
            {
                throw new CoreFault("ole_embed_failed", "The document could not be hosted in the editor tab.");
            }
            throw new CoreFault("ole_embed_failed", "The document could not be hosted in the editor tab.");
        }
        var executable = OfficeDocumentKind.FindAppPath(kind.ServerFile);
        if (executable is null && kind.LocalServerPath is not null && File.Exists(kind.LocalServerPath))
            executable = kind.LocalServerPath;
        if (executable is null) throw new CoreFault("office_not_found", "The installed Microsoft Office application could not be located.");
        try
        {
            var start = new ProcessStartInfo(executable) { UseShellExecute = false };
            // A focus tab hides Office until it is docked; Word's start-up splash would flash on screen meanwhile.
            if (quiet && kind.Application == "word") start.ArgumentList.Add("/q");
            start.ArgumentList.Add(fullPath);
            using var process = Process.Start(start);
            return Wire.Node(new { opened = true, mode, application = kind.Application, pid = process?.Id, inplace = false });
        }
        catch { throw new CoreFault("office_open_failed", "Microsoft Office could not open the selected document."); }
    }
}
