param([Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run)
$ErrorActionPreference='Stop'
$root=Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\focus-hybrid-'+$Run)
if($root -ine $PSScriptRoot -or (Get-Process -Id $PID).SessionId -eq 0){throw 'Expected the isolated interactive runner'}
$input=Get-Content -Raw -LiteralPath (Join-Path $root 'inputs.json')|ConvertFrom-Json
if($input.run -ne $Run -or $input.stem -ne ('synthetic-'+$Run) -or $input.smoke){throw 'Expected a dedicated interactive fixture'}
$sourceExtension=switch($input.app){'word'{'docx'} 'excel'{'xlsx'} 'powerpoint'{'pptx'} default{throw 'Unknown fixture application'}}
$source=Join-Path $root ($input.stem+'.'+$sourceExtension)
$office=Get-Process -Id $input.officePid
$editor=Get-Process -Id $input.codePid
$command=Get-CimInstance Win32_Process -Filter ('ProcessId='+[int]$office.Id)
if($editor.ProcessName -ne 'Code' -or $office.SessionId -ne $editor.SessionId -or
    -not $office.MainWindowTitle.StartsWith([string]$input.stem,[StringComparison]::OrdinalIgnoreCase) -or
    -not $command.CommandLine -or $command.CommandLine.IndexOf($source,[StringComparison]::OrdinalIgnoreCase) -lt 0){
    throw 'Could not establish the exact fixture window'
}
$officeWindow=$office.MainWindowHandle
$editorWindow=$editor.MainWindowHandle
if($officeWindow -eq [IntPtr]::Zero -or $editorWindow -eq [IntPtr]::Zero){throw 'Fixture window is unavailable'}

Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ParkedFocusProbe {
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr window, uint ancestor);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int n);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr window, out Rect rectangle);
    [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr dpi);
    [StructLayout(LayoutKind.Sequential)] struct Rect { internal int Left, Top, Right, Bottom; }
    public static void UsePhysicalCoordinates() { SetProcessDpiAwarenessContext(new IntPtr(-4)); }
    public static bool OwnsForeground(IntPtr window) { return GetAncestor(GetForegroundWindow(), 3) == window; }
    public static bool IsParked(IntPtr window) {
        Rect rectangle;
        return GetWindowRect(window, out rectangle) && rectangle.Right <= GetSystemMetrics(76);
    }
}
"@
[ParkedFocusProbe]::UsePhysicalCoordinates()
function Set-WindowFocus([IntPtr]$Window){
    $element=[Windows.Automation.AutomationElement]::FromHandle($Window)
    if(-not $element){throw 'Window did not expose an accessibility surface'}
    $element.SetFocus()
}
$out=[ordered]@{run=$Run;app=$input.app;phase='activate_editor';parkedBefore=$false;raised=$false;recovered=$false;
    offscreenForegroundAtDeadline=$false;editorRestored=$false}
try {
    if(-not [ParkedFocusProbe]::OwnsForeground($editorWindow)){Set-WindowFocus $editorWindow}
    if(-not [ParkedFocusProbe]::OwnsForeground($editorWindow)){throw 'Editor did not accept foreground'}
    $out.phase='verify_parked'
    $out.parkedBefore=[ParkedFocusProbe]::IsParked($officeWindow)
    if(-not $out.parkedBefore){throw 'The exact Office window was not parked'}
    $out.phase='raise_parked_office'
    $clock=[Diagnostics.Stopwatch]::StartNew()
    Set-WindowFocus $officeWindow
    do {
        $owns=[ParkedFocusProbe]::OwnsForeground($officeWindow)
        $parked=[ParkedFocusProbe]::IsParked($officeWindow)
        if($owns){$out.raised=$true}
        if($owns -and -not $parked){
            $out.recovered=$true
            $out.recoveredMs=$clock.ElapsedMilliseconds
            break
        }
        Start-Sleep -Milliseconds 25
    } while($clock.ElapsedMilliseconds -lt 850)
    $out.offscreenForegroundAtDeadline=[ParkedFocusProbe]::OwnsForeground($officeWindow) -and [ParkedFocusProbe]::IsParked($officeWindow)
    $out.phase='observed'
}catch{
    $out.errorType=$_.Exception.GetBaseException().GetType().Name
    $out.errorLine=$_.InvocationInfo.ScriptLineNumber
}finally{
    if([ParkedFocusProbe]::OwnsForeground($officeWindow)){
        try{Set-WindowFocus $editorWindow}catch{}
    }
    $out.editorRestored=[ParkedFocusProbe]::OwnsForeground($editorWindow)
    $file=Join-Path $root 'parked-focus-evidence.json'
    [IO.File]::WriteAllText($file,($out|ConvertTo-Json -Compress),[Text.UTF8Encoding]::new($false))
}
