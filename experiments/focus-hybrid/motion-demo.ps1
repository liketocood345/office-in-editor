param([Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run,
      [ValidateRange(30,900)][int]$DurationSeconds=600)
$ErrorActionPreference='Stop'
$root=Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\focus-hybrid-'+$Run)
if($PSScriptRoot -ine $root -or (Get-Process -Id $PID).SessionId -eq 0){throw 'Unexpected motion lab context'}
$stem='synthetic-'+$Run
$file=Join-Path $root ($stem+'.pptx')
$gif=Join-Path $root 'focus-motion.gif'
$evidence=Join-Path $root 'motion-evidence.json'
$task='OfficeInEditor-FocusMotion-'+$Run
$state=[ordered]@{run=$Run;embedded=$false;durationSeconds=$DurationSeconds;offscreenMoves=0;backgroundMoves=0;state='opening'}
$picture=$null
$startX=$null
function Record {
    $partial=$evidence+'.partial'
    [IO.File]::WriteAllText($partial,($state|ConvertTo-Json -Compress),[Text.UTF8Encoding]::new($false))
    if([IO.File]::Exists($evidence)){[IO.File]::Replace($partial,$evidence,[NullString]::Value)}else{[IO.File]::Move($partial,$evidence)}
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class FocusMotionProbe {
    [StructLayout(LayoutKind.Sequential)] struct Rect { internal int Left, Top, Right, Bottom; }
    delegate bool Visit(IntPtr window, IntPtr context);
    [DllImport("user32.dll")] static extern bool EnumWindows(Visit visit, IntPtr context);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr window, uint command);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr window, uint command);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr w, out Rect r);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr w, out uint pid);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    public static IntPtr Find(uint expected, string stem) {
        IntPtr result = IntPtr.Zero;
        EnumWindows((window, ignored) => {
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            if (owner != expected || !IsWindowVisible(window) || GetWindow(window, 4) != IntPtr.Zero) return true;
            var title = new StringBuilder(260);
            GetWindowText(window, title, title.Capacity);
            if (!title.ToString().StartsWith(stem, StringComparison.OrdinalIgnoreCase)) return true;
            result = window;
            return false;
        }, IntPtr.Zero);
        return result;
    }
    public static int State(IntPtr window, uint expected) {
        uint owner;
        GetWindowThreadProcessId(window, out owner);
        if (owner != expected) return -1;
        if (GetAncestor(GetForegroundWindow(), 3) == window) return 0;
        Rect bounds;
        return GetWindowRect(window, out bounds) && bounds.Right <= GetSystemMetrics(76) ? 2 : 1;
    }
}
"@

try {
    $application=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $deck=$application.Presentations | Where-Object {[string]$_.FullName -ieq $file} | Select-Object -First 1
    if(-not $deck -or $deck.ReadOnly){throw 'The exact writable synthetic deck is unavailable'}
    $editor=Get-Process -Id (Get-Content -LiteralPath (Join-Path $root 'inputs.json') -Raw|ConvertFrom-Json).codePid
    $office=Get-Process -Name POWERPNT | Where-Object {$_.SessionId -eq $editor.SessionId -and
        $_.MainWindowTitle.StartsWith($stem,[StringComparison]::OrdinalIgnoreCase)} | Select-Object -First 1
    if(-not $office){throw 'The exact synthetic slide window is unavailable'}
    $window=[FocusMotionProbe]::Find($office.Id,$stem)
    if($window -eq [IntPtr]::Zero){throw 'No matching presentation frame'}
    $slide=$deck.Slides.Item(1)
    $width=[single]$deck.PageSetup.SlideWidth
    $height=[single]$deck.PageSetup.SlideHeight
    $startX=[single]($width*0.10)
    $picture=$slide.Shapes | Where-Object Name -eq 'OfficeInEditorFocusMotion' | Select-Object -First 1
    if(-not $picture){
        $picture=$slide.Shapes.AddPicture($gif,0,-1,$startX,[single]($height*0.48),[single]($width*0.39),[single]($width*0.39/4))
        $picture.Name='OfficeInEditorFocusMotion'
        $picture.AlternativeText='Focus preview diagnostic animated GIF'
    }
    $deck.Save()
    Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
    $sha=(Get-FileHash -LiteralPath $gif -Algorithm SHA256).Hash
    $fileStream=[IO.File]::Open($file,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    try {
        $pkg=[IO.Compression.ZipArchive]::new($fileStream,[IO.Compression.ZipArchiveMode]::Read,$true)
        try {
            foreach($entry in $pkg.Entries){
                if($entry.FullName -notmatch '^ppt/media/[^/]+\.gif$' -or $entry.Length -gt 524288){continue}
                $stream=$entry.Open()
                try {if((Get-FileHash -InputStream $stream -Algorithm SHA256).Hash -eq $sha){$state.embedded=$true;break}}
                finally{$stream.Dispose()}
            }
        }finally{$pkg.Dispose()}
    }finally{$fileStream.Dispose()}
    if(-not $state.embedded){throw 'The saved deck does not contain the original animated GIF'}
    $state.state='running'
    Record
    $clock=[Diagnostics.Stopwatch]::StartNew()
    $ticks=0
    while($clock.Elapsed.TotalSeconds -lt $DurationSeconds -and -not $editor.HasExited -and -not $office.HasExited){
        $probe=[FocusMotionProbe]::State($window,$office.Id)
        if($probe -lt 0){break}
        $state.sourceState=@('foreground','background','parked')[$probe]
        if($probe -ge 1){
            $progress=($state.backgroundMoves%12)/11.0
            $picture.Left=[single]($startX+$width*0.42*$progress)
            $state.backgroundMoves++
            if($probe -eq 2){$state.offscreenMoves++}
        }
        $ticks++
        if($ticks -le 4 -or $ticks%8 -eq 0){Record}
        Start-Sleep -Milliseconds 720
    }
    $state.state='finished'
}catch{
    $state.state='error'
    $state.errorType=$_.Exception.GetBaseException().GetType().Name
    $state.errorLine=$_.InvocationInfo.ScriptLineNumber
}finally{
    if($picture -and $null -ne $startX){try{$picture.Left=$startX}catch{}}
    Record
    Unregister-ScheduledTask -TaskName $task -Confirm:$false -ErrorAction SilentlyContinue
}
