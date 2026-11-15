using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OfficeInEditor.Core;

/// <summary>
/// File → OLE document-server identity. Kind key is the local server exe (winword.exe),
/// not an HWND class name. Outlook / OneDrive / mail-store types are out of scope.
/// </summary>
internal sealed class OfficeDocumentKind
{
    internal required string Key;
    internal required string Application;
    internal required string ProgId;
    internal Guid Clsid;
    internal string ServerFile = "";
    internal string? LocalServerPath;
    internal int GetClassFileHr;

    private static readonly HashSet<string> DeniedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".msg", ".eml", ".pst", ".ost", ".oft", ".olk", ".oab", ".ics", ".vcf", ".hol", ".nk2"
    };

    private static readonly HashSet<string> DeniedServers = new(StringComparer.OrdinalIgnoreCase)
    {
        "outlook.exe", "olk.exe", "onedrive.exe", "onedrivestandaloneupdater.exe",
        "onedriveupdater.exe", "ms-teams.exe", "teams.exe", "msteams.exe",
        "lync.exe", "communicator.exe", "lync99.exe"
    };

    internal static bool IsDeniedExtension(string extension) => DeniedExtensions.Contains(extension);

    internal static bool IsDeniedServer(string fileName)
    {
        var name = Path.GetFileName(fileName.Trim().Trim('"'));
        return name.Length > 0 && DeniedServers.Contains(name);
    }

    internal static bool TryClassify(string path, out OfficeDocumentKind kind)
    {
        kind = null!;
        var extension = Path.GetExtension(path);
        if (IsDeniedExtension(extension)) return false;
        Guid clsid = Guid.Empty;
        var getClassHr = 0;
        string progId = "";
        string? localServer = null;
        if (File.Exists(path))
        {
            getClassHr = GetClassFile(path, out clsid);
            if (getClassHr == 0 && clsid != Guid.Empty)
            {
                progId = ProgIdFrom(clsid);
                localServer = LocalServerFrom(clsid);
                if (IsDeniedProgId(progId) || (localServer is not null && IsDeniedServer(localServer)))
                    return false;
            }
        }
        if (clsid == Guid.Empty || string.IsNullOrEmpty(localServer))
        {
            if (!TryGuessFromExtension(extension, out var guessedApp, out var guessedProg, out var guessedExe))
                return false;
            if (clsid == Guid.Empty)
            {
                foreach (var candidate in new[] { guessedProg, guessedProg + ".12", guessedProg + ".8" })
                {
                    if (CLSIDFromProgID(candidate, out var fromProg) != 0 || fromProg == Guid.Empty) continue;
                    clsid = fromProg;
                    if (string.IsNullOrEmpty(progId)) progId = candidate;
                    localServer ??= LocalServerFrom(clsid);
                    break;
                }
            }
            kind = new OfficeDocumentKind
            {
                Key = guessedExe.ToLowerInvariant(),
                Application = guessedApp,
                ProgId = string.IsNullOrEmpty(progId) ? guessedProg : progId,
                Clsid = clsid,
                ServerFile = guessedExe,
                LocalServerPath = localServer,
                GetClassFileHr = getClassHr
            };
            return true;
        }
        var serverFile = Path.GetFileName(localServer.Trim().Trim('"'));
        if (serverFile.Length == 0) return false;
        kind = new OfficeDocumentKind
        {
            Key = serverFile.ToLowerInvariant(),
            Application = ApplicationFromServer(serverFile),
            ProgId = progId,
            Clsid = clsid,
            ServerFile = serverFile,
            LocalServerPath = localServer,
            GetClassFileHr = getClassHr
        };
        return !IsDeniedServer(serverFile) && !IsDeniedProgId(progId);
    }

    internal static bool TryGuessFromExtension(string extension, out string application, out string progId, out string serverFile)
    {
        application = ""; progId = ""; serverFile = "";
        (application, progId, serverFile) = extension.ToLowerInvariant() switch
        {
            ".docx" or ".doc" or ".docm" or ".dotx" or ".dotm" or ".rtf" => ("word", "Word.Document", "WINWORD.EXE"),
            ".xlsx" or ".xls" or ".xlsm" or ".xlsb" or ".xltx" or ".xltm" or ".csv" => ("excel", "Excel.Sheet", "EXCEL.EXE"),
            ".pptx" or ".ppt" or ".pptm" or ".potx" or ".potm" or ".ppsx" => ("powerpoint", "PowerPoint.Show", "POWERPNT.EXE"),
            ".vsdx" or ".vsd" or ".vstx" or ".vssx" or ".vstm" => ("visio", "Visio.Drawing", "VISIO.EXE"),
            ".one" or ".onepkg" => ("onenote", "OneNote.Section", "ONENOTE.EXE"),
            ".pub" => ("publisher", "Publisher.Document", "MSPUB.EXE"),
            ".mpp" => ("project", "MSProject.Project", "WINPROJ.EXE"),
            _ => ("", "", "")
        };
        return application.Length > 0;
    }

    private static bool IsDeniedProgId(string progId) =>
        progId.StartsWith("Outlook.", StringComparison.OrdinalIgnoreCase)
        || progId.StartsWith("OneDrive.", StringComparison.OrdinalIgnoreCase)
        || progId.StartsWith("Teams.", StringComparison.OrdinalIgnoreCase);

    private static string ApplicationFromServer(string serverFile) =>
        Path.GetFileName(serverFile).ToLowerInvariant() switch
        {
            "winword.exe" => "word",
            "excel.exe" => "excel",
            "powerpnt.exe" => "powerpoint",
            "visio.exe" or "visio32.exe" => "visio",
            "onenote.exe" or "onenotem.exe" => "onenote",
            "mspub.exe" => "publisher",
            "winproj.exe" => "project",
            var name => Path.GetFileNameWithoutExtension(name).ToLowerInvariant()
        };

    private static string ProgIdFrom(Guid clsid)
    {
        if (ProgIDFromCLSID(ref clsid, out var pointer) != 0 || pointer == IntPtr.Zero) return "";
        try { return Marshal.PtrToStringUni(pointer) ?? ""; }
        finally { Marshal.FreeCoTaskMem(pointer); }
    }

    private static string? LocalServerFrom(Guid clsid)
    {
        var id = clsid.ToString("B");
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var root = RegistryKey.OpenBaseKey(hive, view);
            foreach (var name in new[] { @"SOFTWARE\Classes\CLSID\" + id + @"\LocalServer32", @"SOFTWARE\Classes\CLSID\" + id + @"\LocalServer" })
            {
                using var key = root.OpenSubKey(name);
                if (key?.GetValue(null) is string raw && TrimServer(raw) is { Length: > 0 } trimmed)
                    return trimmed;
            }
        }
        return null;
    }

    internal static string? TrimServer(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        value = value.Trim();
        if (value.StartsWith('"'))
        {
            var end = value.IndexOf('"', 1);
            if (end > 0) return value[1..end];
        }
        var slash = value.IndexOf(" /", StringComparison.Ordinal);
        if (slash > 0) value = value[..slash];
        return value.Trim().Trim('"');
    }

    internal static string? FindAppPath(string serverFile)
    {
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var root = RegistryKey.OpenBaseKey(hive, view);
            using var entry = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + serverFile);
            if (entry?.GetValue(null) is string candidate && File.Exists(candidate.Trim('"')))
                return candidate.Trim('"');
        }
        return null;
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassFile(string szFileName, out Guid pclsid);

    [DllImport("ole32.dll")]
    private static extern int ProgIDFromCLSID(ref Guid clsid, out IntPtr lplpszProgID);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);
}
