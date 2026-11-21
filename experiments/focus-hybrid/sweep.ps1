param([Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9][a-zA-Z0-9._-]*$')][string]$Remote,
      [ValidateSet('word','excel','powerpoint')][string[]]$Apps=@('word','excel','powerpoint'),
      [ValidateSet('always','switch','mouse')][string[]]$Policies=@('always','switch','mouse'),
      [ValidateRange(1,40)][int]$Cycles=10,[ValidateRange(1,4)][int]$Repeat=2,[switch]$FailFast,
      [ValidateSet('click','raise','raise-baseline')][string[]]$Scenarios=@('click'))

$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$runRoot=Join-Path $repository '.local/focus-hybrid-runs'
$sweep=[Guid]::NewGuid().ToString('N')
$out=Join-Path $runRoot ('sweeps/'+$sweep)
New-Item -ItemType Directory -Path $out -Force|Out-Null
$utf8=[Text.UTF8Encoding]::new($false)
$combinations=@(foreach($app in $Apps){foreach($policy in $Policies){foreach($scenario in $Scenarios){@{app=$app;policy=$policy;scenario=$scenario}}}})
$cases=@(for($index=1;$index -le $Repeat;$index++){
    $ordered=@($combinations)
    if($index%2 -eq 0){[array]::Reverse($ordered)}
    foreach($case in $ordered){[ordered]@{app=$case.app;policy=$case.policy;scenario=$case.scenario;repeat=$index;status='pending'}}
})
$manifest=[ordered]@{sweep=$sweep;cycles=$Cycles;repeats=$Repeat;cases=$cases}
function Save-Manifest {
    $target=Join-Path $out 'manifest.json'
    $partial=$target+'.partial'
    [IO.File]::WriteAllText($partial,($manifest|ConvertTo-Json -Depth 7),$utf8)
    if([IO.File]::Exists($target)){[IO.File]::Replace($partial,$target,[NullString]::Value)}
    else{[IO.File]::Move($partial,$target)}
}

function Wait-Run([string]$Id) {
    if($Id -notmatch '^[a-f0-9]{32}$'){throw 'Unexpected run ID'}
    $body=@'
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$dir=Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\focus-hybrid-'+$runId)
$finished=Join-Path $dir 'finished.json'
$watch=[IO.FileSystemWatcher]::new($dir,'finished.json')
try{
 $watch.EnableRaisingEvents=$true
 $clock=[Diagnostics.Stopwatch]::StartNew()
 while($clock.ElapsedMilliseconds -lt $limit -and -not (Test-Path -LiteralPath $finished)){
  [void]$watch.WaitForChanged([IO.WatcherChangeTypes]::Created -bor [IO.WatcherChangeTypes]::Renamed,8000)
 }
 if(Test-Path -LiteralPath $finished){'finished'}else{'timeout'}
}finally{$watch.Dispose()}
'@
    $prefix='$runId = '''+$Id+'''; $limit = '+(75000+4500*$Cycles)+[Environment]::NewLine
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($prefix+$body))
    $state=& ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
    if($LASTEXITCODE -or $state -notin 'finished','timeout'){throw 'Remote finish wait failed'}
    $state
}

Save-Manifest
@{sweep=$sweep;cases=$cases.Length;cycles=$Cycles}|ConvertTo-Json -Compress
foreach($case in $cases){
    try {
        $case.status='starting'
        Save-Manifest
        $started=(& (Join-Path $PSScriptRoot 'remote.ps1') -Remote $Remote -Smoke -App $case.app -PixelPolicy $case.policy -Cycles $Cycles -Scenario $case.scenario)|ConvertFrom-Json
        $case.run=$started.run
        $case.status='running'
        Save-Manifest
        $state=Wait-Run $case.run
        $evidence=(& (Join-Path $PSScriptRoot 'remote.ps1') -Remote $Remote -Run $case.run -Collect)|ConvertFrom-Json
        if($state -eq 'finished' -and $evidence.finished){
            $runPath=Join-Path $runRoot $case.run
            & node (Join-Path $PSScriptRoot 'analyze.mjs') $runPath
            $case.status=if($LASTEXITCODE -eq 0){'passed'}else{'failed'}
        }else{$case.status='timeout'}
    }catch{
        $case.status='error'
        $case.errorType=$_.Exception.GetBaseException().GetType().Name
        Write-Warning ('Sweep control failed at '+$case.app+'/'+$case.policy+'/'+$case.scenario)
    }
    Save-Manifest
    $case|ConvertTo-Json -Compress
    if($case.status -in 'timeout','error' -or ($FailFast -and $case.status -ne 'passed')){break}
}
& node (Join-Path $PSScriptRoot 'analyze.mjs') --suite $out
exit $LASTEXITCODE
