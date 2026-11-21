param(
    [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run,
    [ValidateRange(1, 4)][int]$Attempt = 1
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$root = Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\editor-slot-' + $Run)
if ($PSScriptRoot -ine $root -or (Get-Process -Id $PID).SessionId -eq 0) { throw 'Invalid experiment context' }
$launch = Get-Content -LiteralPath (Join-Path $root 'launch.json') -Raw | ConvertFrom-Json
$opened = Get-Content -LiteralPath (Join-Path $root 'result.json') -Raw | ConvertFrom-Json
if ($launch.run -ne $Run -or -not $opened.passed) { throw 'No verified native session' }
$source = Join-Path $root 'synthetic.docx'
$resultPath = Join-Path $root $(if ($Attempt -eq 1) { 'save-result.json' } else { 'save-result-' + $Attempt + '.json' })
if (Test-Path -LiteralPath $resultPath) { throw 'The save test already ran' }
$utf8 = [Text.UTF8Encoding]::new($false)
$report = [ordered]@{ stage = 'scope'; manual = $false; periodic = $false }

function Read-SyntheticText {
    $zip = [IO.Compression.ZipFile]::OpenRead($source)
    try {
        $reader = [IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
        try { [xml]$part = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $namespace = [Xml.XmlNamespaceManager]::new($part.NameTable)
        $namespace.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        return (@($part.SelectNodes('//w:t', $namespace) | ForEach-Object { $_.InnerText }) -join '')
    } finally { $zip.Dispose() }
}

try {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $initial = 'Office in Editor synthetic slot test'
    if ((Read-SyntheticText) -cne $initial) { throw 'Synthetic fixture was edited externally' }
    $editor = Get-Process -Id $launch.pid
    if ($editor.Path -ine (Join-Path $root 'editor\Code.exe') -or $editor.MainWindowHandle -eq [IntPtr]::Zero) { throw 'Editor identity changed' }
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class OieLiveSave {
    private delegate bool Visit(IntPtr h, IntPtr p);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, Visit callback, IntPtr data);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder cls, int size);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object value);
    private static string Class(IntPtr hwnd) {
        var name = new StringBuilder(128);
        GetClassName(hwnd, name, name.Capacity);
        return name.ToString();
    }
    public static object FindWordWindow(IntPtr root) {
        var views = new List<IntPtr>();
        EnumChildWindows(root, (slot, _) => {
            if (Class(slot) != "OfficeInEditor.ExperimentalNativeSlot.1") return true;
            EnumChildWindows(slot, (h, __) => {
                if (Class(h) == "_WwG" && IsWindowVisible(h)) views.Add(h);
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        if (views.Count != 1) throw new InvalidOperationException("Expected one visible Word view");
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        object value;
        Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(views[0], 0xFFFFFFF0, ref iid, out value));
        return value;
    }
}
'@
    $report.stage = 'word-binding'
    $window = [OieLiveSave]::FindWordWindow($editor.MainWindowHandle)
    $doc = $window.Document
    $report.stage = 'word-path'
    $docPath = [IO.Path]::GetFullPath([string]$doc.FullName)
    $sessionBase = Join-Path $env:LOCALAPPDATA 'OfficeInEditor\NativeSessions'
    $relative = $docPath.Substring([Math]::Min($docPath.Length, $sessionBase.Length))
    if ($docPath -ine $source -and
        (-not $docPath.StartsWith($sessionBase + '\', [StringComparison]::OrdinalIgnoreCase) -or
         $relative -notmatch '^\\s-([a-f0-9]{32})\\working-\1\.docx$')) { throw 'Unexpected Word document path' }
    $report.stage = 'word-content'
    if (([string]$doc.Content.Text).Trim() -cne $initial) { throw 'Unexpected Word document content' }
    $report.stage = 'active-word-target'
    $word = $window.Application
    if ($word.Documents.Count -ne 1 -or [string]$word.ActiveDocument.FullName -ine $docPath) { throw 'FileSave would target an ambiguous document' }

    $report.stage = 'manual-save'
    $manualMarker = ' [manual commit]'
    $doc.Content.InsertAfter($manualMarker)
    $word.CommandBars.ExecuteMso('FileSave')
    if ((Read-SyntheticText) -cne ($initial + $manualMarker)) { throw 'FileSave did not commit to the source file' }
    $report.manual = $true

    $report.stage = 'periodic-save'
    $periodicMarker = ' [periodic commit]'
    $doc.Content.InsertAfter($periodicMarker)
    Start-Sleep -Seconds 65
    if ((Read-SyntheticText) -cne ($initial + $manualMarker + $periodicMarker)) { throw 'Periodic save did not commit to the source file' }
    $report.periodic = $true
    $report.stage = 'complete'
} catch {
    $report.error = $_.Exception.GetType().Name
    $report.hresult = '0x{0:X8}' -f $_.Exception.HResult
    if ($_.Exception.InnerException) {
        $report.inner = $_.Exception.InnerException.GetType().Name
        $report.innerHresult = '0x{0:X8}' -f $_.Exception.InnerException.HResult
    }
} finally {
    [IO.File]::WriteAllText($resultPath, ($report | ConvertTo-Json -Compress), $utf8)
}
