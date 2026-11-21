param([Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run,
      [Parameter(Mandatory)][ValidateRange(1,2147483647)][int]$OfficePid,
      [Parameter(Mandatory)][ValidateSet('word','excel','powerpoint')][string]$App,[switch]$ExercisePointer)
$ErrorActionPreference='Stop'
$stem='synthetic-'+$Run
$root=Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\focus-hybrid-'+$Run)
$format=switch($App){'word'{'docx'} 'excel'{'xlsx'} 'powerpoint'{'pptx'}}
$paneClass=switch($App){'word'{'_WwG'} 'excel'{'EXCEL7'} 'powerpoint'{'paneClassDC'}}
$source=Join-Path $root ($stem+'.'+$format)
if($PSScriptRoot -ine $root -or -not (Test-Path -LiteralPath $source -PathType Leaf)){exit 2}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class FocusSaveProbe {
    delegate bool Visit(IntPtr hwnd, IntPtr context);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [StructLayout(LayoutKind.Sequential)] struct Point { internal int X, Y; }
    [DllImport("user32.dll")] static extern bool GetCursorPos(out Point point);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int capacity);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder text, int capacity);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr hwnd, Visit visit, IntPtr context);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid,
        [MarshalAs(UnmanagedType.IDispatch)] out object result);

    public static bool ForegroundMatches(uint expectedPid, string expectedTitle) {
        var hwnd = GetForegroundWindow();
        uint owner;
        GetWindowThreadProcessId(hwnd, out owner);
        if (owner != expectedPid) return false;
        var title = new StringBuilder(260);
        GetWindowText(hwnd, title, title.Capacity);
        return title.ToString().StartsWith(expectedTitle, StringComparison.OrdinalIgnoreCase);
    }

    public static bool ExercisePointer(uint pid, string title) {
        Point original;
        if (!ForegroundMatches(pid, title) || !GetCursorPos(out original) || !SetCursorPos(original.X + 8, original.Y)) return false;
        System.Threading.Thread.Sleep(230);
        Point current;
        if (GetCursorPos(out current) && current.X == original.X + 8 && current.Y == original.Y) SetCursorPos(original.X, original.Y);
        return true;
    }

    public static object FindDocumentWindow(uint expectedPid, string expectedTitle, string paneClass) {
        if (!ForegroundMatches(expectedPid, expectedTitle)) return null;
        var hwnd = GetForegroundWindow();
        object result = null;
        EnumChildWindows(hwnd, (child, ignored) => {
            var name = new StringBuilder(64);
            GetClassName(child, name, name.Capacity);
            if (!String.Equals(name.ToString(), paneClass, StringComparison.OrdinalIgnoreCase)) return true;
            var dispatch = new Guid("00020400-0000-0000-C000-000000000046");
            return AccessibleObjectFromWindow(child, 0xFFFFFFF0, ref dispatch, out result) != 0;
        }, IntPtr.Zero);
        return result;
    }
}
"@

try {
    if(-not [FocusSaveProbe]::ForegroundMatches($OfficePid,$stem)){exit 3}
    $window=[FocusSaveProbe]::FindDocumentWindow($OfficePid,$stem,$paneClass)
    $document=$null
    $application=$null
    if($window){
        $application=$window.Application
        $document=switch($App){
            'word' {$window.Document}
            'excel' {$application.ActiveWorkbook}
            'powerpoint' {$window.Presentation}
        }
    }
    if(-not $document -and $App -eq 'powerpoint'){
        $application=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        $document=$application.Presentations | Where-Object {[string]$_.FullName -ieq $source} | Select-Object -First 1
    }
    if(-not $document){exit 3}
    if([string]$document.FullName -ine $source){exit 4}
    if($document.ReadOnly){exit 5}
    $evidence=[ordered]@{app=$App;connected=$false;capturePolicyCallback=$false;pointerExercised=$false}
    try {
        $addin=$application.COMAddIns.Item('OfficeInEditor.Addin')
        $evidence.connected=[bool]$addin.Connect
        if($evidence.connected){$evidence.capturePolicyCallback=([string]$addin.Object.GetCapturePolicy($null)).StartsWith('OieCapture')}
    }catch{}
    if($ExercisePointer){$evidence.pointerExercised=[FocusSaveProbe]::ExercisePointer($OfficePid,$stem)}
    $evidence|ConvertTo-Json -Compress|Set-Content -Encoding UTF8 -LiteralPath (Join-Path $root 'addin-evidence.json')
    if(-not [FocusSaveProbe]::ForegroundMatches($OfficePid,$stem)){exit 3}
    switch($App){
        'word' {[void]$document.Content.InsertAfter(' OIE_FOCUS_SAVED')}
        'excel' {$document.Worksheets.Item(1).Range('E2').Value2='OIE_FOCUS_SAVED'}
        'powerpoint' {[void]$document.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertAfter(' OIE_FOCUS_SAVED')}
    }
    $document.Save()
    if(-not $document.Saved){exit 6}
}catch{exit 7}
