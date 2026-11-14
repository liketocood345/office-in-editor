using System.Diagnostics;
using System.Text.Json.Nodes;
using Microsoft.Win32;

namespace OfficeInEditor.Core;

internal static class NativeOfficeLauncher
{
    public static JsonNode? Open(string path)
    {
        if (!Path.IsPathFullyQualified(path) || path.StartsWith("\\\\", StringComparison.Ordinal) ||
            Uri.TryCreate(path, UriKind.Absolute, out var uri) && uri.IsUnc)
            throw new CoreFault("invalid_path", "Choose an existing local Office document.");
        var extension = Path.GetExtension(path).ToLowerInvariant();
        var (application, file) = extension switch
        {
            ".docx" => ("word", "WINWORD.EXE"), ".xlsx" => ("excel", "EXCEL.EXE"), ".pptx" => ("powerpoint", "POWERPNT.EXE"),
            _ => throw new CoreFault("unsupported_format", "Only local DOCX, XLSX and PPTX documents are supported.")
        };
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new CoreFault("file_not_found", "The selected document is unavailable.");
        string? executable = null;
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var root = RegistryKey.OpenBaseKey(hive, view);
            using var entry = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + file);
            if (entry?.GetValue(null) is string candidate && File.Exists(candidate.Trim('"')))
            { executable = candidate.Trim('"'); break; }
        }
        if (executable is null) throw new CoreFault("office_not_found", "The installed Microsoft Office application could not be located.");
        try
        {
            var start = new ProcessStartInfo(executable) { UseShellExecute = false };
            start.ArgumentList.Add(fullPath);
            using var process = Process.Start(start);
            return Wire.Node(new { opened = true, mode = "external", application });
        }
        catch { throw new CoreFault("office_open_failed", "Microsoft Office could not open the selected document."); }
    }
}
