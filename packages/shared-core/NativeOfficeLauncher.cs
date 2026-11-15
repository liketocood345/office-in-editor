using System.Diagnostics;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

internal static class NativeOfficeLauncher
{
    public static JsonNode? Open(string path, string mode, bool fallback = false, int x = 0, int y = 0, int width = 0, int height = 0, string? sessionId = null, string? surfaceToken = null, CancellationToken cancellation = default, bool visible = true, double pixelRatio = 1, string? backupId = null)
    {
        if (mode is not ("external" or "embedded"))
            throw new CoreFault("invalid_open_mode", "Office open mode must be external or embedded.");
        if (!Path.IsPathFullyQualified(path) || path.StartsWith("\\\\", StringComparison.Ordinal) ||
            Uri.TryCreate(path, UriKind.Absolute, out var uri) && uri.IsUnc)
            throw new CoreFault("invalid_path", "Choose an existing local Office document.");
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new CoreFault("file_not_found", "The selected document is unavailable.");
        if (!OfficeDocumentKind.TryClassify(fullPath, out var kind))
            throw new CoreFault("unsupported_format", "This file is not an embeddable Office document.");
        _ = (x, y);
        if (mode == "embedded")
        {
            // #region agent log
            DebugNdjson.Write("A", "NativeOfficeLauncher.cs:embedded", "embedded_open", new
            {
                fallback, width, height, startProcess = false,
                hasToken = WebviewSurface.IsToken(surfaceToken),
                kindKey = kind.Key, application = kind.Application,
                clsid = kind.Clsid.ToString("B"), getClassHr = kind.GetClassFileHr,
                server = kind.ServerFile
            });
            // #endregion
            if (!WebviewSurface.IsToken(surfaceToken) || width < 32 || height < 32)
                throw new CoreFault("invalid_host_window", "A valid editor webview surface is required.");
            try
            {
                var hosted = OleDocumentHost.TryOpen(fullPath, kind.Application, width, height, sessionId, surfaceToken, cancellation, bindSource: false, visible: visible, jsX: x, jsY: y, pixelRatio: pixelRatio, backupId: backupId);
                if (hosted is not null) return hosted;
            }
            catch (OperationCanceledException) { throw; }
            catch (CoreFault) { throw; }
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
            start.ArgumentList.Add(fullPath);
            // #region agent log
            DebugNdjson.Write("B", "NativeOfficeLauncher.cs:external", "external_process_start", new { application = kind.Application, startProcess = true });
            // #endregion
            using var process = Process.Start(start);
            return Wire.Node(new { opened = true, mode, application = kind.Application, pid = process?.Id, inplace = false });
        }
        catch { throw new CoreFault("office_open_failed", "Microsoft Office could not open the selected document."); }
    }
}
