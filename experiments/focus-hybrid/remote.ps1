param([Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9][a-zA-Z0-9._-]*$')][string]$Remote,
      [ValidatePattern('^[a-f0-9]{32}$')][string]$Run,[switch]$Smoke,
      [ValidateSet('word','excel','powerpoint')][string]$App='word',
      [ValidateSet('always','switch','mouse')][string]$PixelPolicy='always', [ValidateRange(0,40)][int]$Cycles=0,[switch]$Collect,
      [ValidateSet('click','raise','raise-baseline')][string]$Scenario='click')
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$readOnly=[bool]$Run
if($Cycles -and (-not $Smoke -or $readOnly)){throw 'Cycles require a new smoke run'}
if($Scenario -ne 'click' -and -not $Cycles){throw 'Raise scenarios require a cycled smoke run'}
if($Collect -and -not $readOnly){throw 'Collect requires an existing run ID'}
if(-not $Run){$Run=[Guid]::NewGuid().ToString('N')}
$repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$local=Join-Path $repository ('.local/focus-hybrid-runs/'+$Run)
function Invoke-Remote([string]$Body){
    $prefix='$ErrorActionPreference = ''Stop''; $ProgressPreference = ''SilentlyContinue''; $Run = '''+$Run+"'`r`n"+
       '$base = Join-Path $env:LOCALAPPDATA ''OfficeInEditorLab''' + "`r`n"+
       '$root = Join-Path $base (''focus-hybrid-'' + $Run)' + "`r`n"+
       '$app = '''+$App+"'`r`n"+'$pixelPolicy = '''+$PixelPolicy+"'`r`n"
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($prefix+$Body))
    $res=& ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
    if($LASTEXITCODE){throw 'Focus experiment remote command failed'}
    $res
}
if($readOnly){
    $collectLiteral=if($Collect){'$true'}else{'$false'}
    $readStatus=@'
$status=@{run=$Run}
if($CollectEvidence -and (Test-Path -LiteralPath $root -PathType Container)) {
 if(((Get-Item -LiteralPath $root).Attributes -band [IO.FileAttributes]::ReparsePoint) -or
    ((Get-Item -LiteralPath $base).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Redirected evidence directory'}
 $status.collectFrom=$root.Replace('\','/')
 $files=@('environment.json','result.json','finished.json','events.jsonl','events-word.jsonl','events-excel.jsonl','events-powerpoint.jsonl')
 $status.collectFiles=@($files|Where-Object {
  $item=Get-Item -LiteralPath (Join-Path $root $_) -ErrorAction SilentlyContinue
  $item -and -not $item.PSIsContainer -and -not ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -and $item.Length -le 8MB
 })
}
$inputs=Join-Path $root 'inputs.json'
if(Test-Path -LiteralPath $inputs -PathType Leaf){
 $processes=Get-Content -LiteralPath $inputs -Raw -Encoding UTF8 | ConvertFrom-Json
 $status.smoke=$processes.smoke
 $status.codeAlive=[bool](Get-Process -Id $processes.codePid -ErrorAction SilentlyContinue | Where-Object ProcessName -eq 'Code')
 $status.app=if($processes.app){$processes.app}else{'word'}
 $officePid=if($processes.officePid){$processes.officePid}else{$processes.wordPid}
 $status.officeAlive=[bool](Get-Process -Id $officePid -ErrorAction SilentlyContinue | Where-Object ProcessName -in 'WINWORD','EXCEL','POWERPNT')
}
foreach($name in 'result','finished'){
 $file=Join-Path $root ($name+'.json')
 if(Test-Path -LiteralPath $file -PathType Leaf){$status[$name]=Get-Content -LiteralPath $file -Raw -Encoding UTF8|ConvertFrom-Json}
}
$addin=Join-Path $root 'addin-evidence.json'
if(Test-Path -LiteralPath $addin -PathType Leaf){$status.addin=Get-Content -LiteralPath $addin -Raw -Encoding UTF8|ConvertFrom-Json}
if(-not $status.ContainsKey('result')){
 $journal=Join-Path $root 'events.jsonl'
 if(Test-Path -LiteralPath $journal -PathType Leaf){
  $status.events=@(Get-Content -LiteralPath $journal -Encoding UTF8 -Tail 20 | ForEach-Object {
   try {$_|ConvertFrom-Json -ErrorAction Stop} catch {}
  })
 }
}
$tabs=@{}
foreach($other in 'word','excel','powerpoint'){
 $journal=Join-Path $root ('events-'+$other+'.jsonl')
 if(Test-Path -LiteralPath $journal -PathType Leaf){
  $tabs[$other]=@(Get-Content -LiteralPath $journal -Encoding UTF8 -Tail 20 | ForEach-Object {
   try {$_|ConvertFrom-Json -ErrorAction Stop} catch {}
  })
 }
}
if($tabs.Count){$status.otherTabs=$tabs}
$status|ConvertTo-Json -Depth 15 -Compress
'@
    $json=Invoke-Remote ('$CollectEvidence = '+$collectLiteral+[Environment]::NewLine+$readStatus)
    $status=$json|ConvertFrom-Json
    if($Collect -and $status.collectFrom){
        if($status.collectFrom -notmatch ('^[A-Za-z]:/[A-Za-z0-9_./-]+/OfficeInEditorLab/focus-hybrid-'+$Run+'$') -or
            ($status.collectFrom -split '/') -contains '..'){throw 'Unexpected evidence source'}
        New-Item -ItemType Directory -Path $local -Force|Out-Null
        $sources=@($status.collectFiles | ForEach-Object {
            if($_ -notmatch '^(environment|finished|result)\.json$|^events(?:-(word|excel|powerpoint))?\.jsonl$'){throw 'Unexpected evidence filename'}
            $Remote+':'+$status.collectFrom+'/'+$_
        })
        if($sources.Count){
            & scp -q -o BatchMode=yes -o ConnectTimeout=8 @sources $local
            if($LASTEXITCODE){throw 'Focus experiment evidence collection failed'}
        }
        $status.PSObject.Properties.Remove('collectFrom')
        $status.PSObject.Properties.Remove('collectFiles')
        $json=$status|ConvertTo-Json -Depth 15 -Compress
    }
    if($status.finished -or $status.events){
        New-Item -ItemType Directory -Path $local -Force|Out-Null
        $name=if($status.finished){'ssh-evidence.json'}else{'ssh-live.json'}
        [IO.File]::WriteAllText((Join-Path $local $name),$json,[Text.UTF8Encoding]::new($false))
    }
    $json
    return
}
if(-not $Smoke){
    $running=Invoke-Remote @'
Get-ScheduledTask -TaskName 'OfficeInEditor-FocusHybrid-*' -ErrorAction SilentlyContinue | Where-Object State -eq 'Running' | ForEach-Object {
 $id=$_.TaskName.Substring('OfficeInEditor-FocusHybrid-'.Length)
 if($id -match '^[a-f0-9]{32}$'){
  $session=Join-Path $base ('focus-hybrid-'+$id)
  $action=@($_.Actions)[0]
  if($action.WorkingDirectory -ieq $session -and $action.Arguments -like ('* -App '+$app+' *') -and
      $action.Arguments -notmatch ' -Smoke(?:\s|$)' -and
      -not (Test-Path -LiteralPath (Join-Path $session 'result.json') -PathType Leaf)){$id}
 }
} | Select-Object -First 1
'@
    if($running){@{run=$running;mode='interactive';reused=$true}|ConvertTo-Json -Compress;return}
}
$extension=Join-Path $local 'extension'
New-Item -ItemType Directory -Path $extension -Force|Out-Null
foreach($file in 'package.json','extension.cjs'){Copy-Item -LiteralPath (Join-Path $PSScriptRoot $file) -Destination $extension}
Push-Location -LiteralPath $repository
try {
    # Single-quoted JS literals: Windows PowerShell 5.1 drops embedded double quotes from native arguments.
    $bundle='import {build} from ''esbuild''; await build({entryPoints:[''packages/adapters/vsix/src/mcp.ts''], bundle:true, platform:''node'', external:[''vscode''], format:''cjs'', outfile:process.argv[1]});'
    & node --input-type=module -e $bundle (Join-Path $extension 'mcp.cjs')
    if($LASTEXITCODE){throw 'Shared MCP provider bundle failed'}
}finally{Pop-Location}
$fixturePath=Join-Path $local 'fixtures'
New-Item -ItemType Directory -Path $fixturePath|Out-Null
& python (Join-Path $repository 'experiments/editor-native-host/suite-fixtures.py') $fixturePath $Run
if($LASTEXITCODE){throw 'Excel and PowerPoint fixture generation failed'}
Move-Item -LiteralPath (Join-Path $fixturePath 'synthetic.pptx') -Destination (Join-Path $fixturePath ('synthetic-'+$Run+'.pptx'))
$files=@($extension,(Join-Path $PSScriptRoot 'launch.ps1'),(Join-Path $PSScriptRoot 'open-fixture.ps1'),(Join-Path $PSScriptRoot 'save-probe.ps1'),
    (Join-Path $fixturePath ('synthetic-'+$Run+'.pptx')),(Join-Path $fixturePath ('synthetic-'+$Run+'.xlsx')),
    (Join-Path $repository '.local/build/focus-hybrid/native.exe'))
$zip=Join-Path $local 'payload.zip'
Compress-Archive -LiteralPath $files -DestinationPath $zip
$hash=(Get-FileHash -LiteralPath $zip).Hash
$remotePath=Invoke-Remote @'
New-Item -ItemType Directory -Path $base -Force|Out-Null
if((Get-Item -LiteralPath $base -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Redirected lab folder'}
if(@(Get-ChildItem -LiteralPath $base -Force|Where-Object Name -notmatch '^focus-hybrid-[a-f0-9]{32}$').Count){throw 'Legacy lab has not been cleaned'}
New-Item -ItemType Directory -Path $root|Out-Null
$root.Replace('\','/')
'@
if($remotePath -isnot [string] -or $remotePath -notmatch ('^[A-Za-z]:/[A-Za-z0-9_./-]+/OfficeInEditorLab/focus-hybrid-'+$Run+'$')){throw 'Unexpected focus deployment path'}
& scp -q -o BatchMode=yes -o ConnectTimeout=8 $zip ($Remote+':'+$remotePath+'/payload.zip')
if($LASTEXITCODE){throw 'Focus payload transfer failed'}
$remoteLaunch=@'
$archive=Join-Path $root 'payload.zip'
if((Get-FileHash -LiteralPath $archive).Hash -ne $ExpectedHash){throw 'Focus payload hash mismatch'}
Expand-Archive -LiteralPath $archive -DestinationPath $root
$user=(Get-CimInstance Win32_ComputerSystem).UserName
if(-not $user){throw 'Interactive desktop unavailable'}
$task='OfficeInEditor-FocusHybrid-'+$Run
$ps=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$args='-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy RemoteSigned -File "'+(Join-Path $root 'launch.ps1')+'" -Run '+$Run+' -App '+$app+' -PixelPolicy '+$pixelPolicy
if($UseSmoke){$args+=' -Smoke -Cycles '+$SmokeCycles+' -Scenario '+$SmokeScenario}
$action=New-ScheduledTaskAction -Execute $ps -Argument $args -WorkingDirectory $root
$principal=New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings=New-ScheduledTaskSettingsSet -DisallowHardTerminate -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $task -Action $action -Principal $principal -Settings $settings|Out-Null
Start-ScheduledTask -TaskName $task
@{run=$Run;app=$app;started=$true;mode=$(if($UseSmoke){'smoke'}else{'interactive'})}|ConvertTo-Json -Compress
'@
$smokeLiteral=if($Smoke){'$true'}else{'$false'}
Invoke-Remote ('$ExpectedHash = '''+$hash+"'`r`n"+'$UseSmoke = '+$smokeLiteral+[Environment]::NewLine+'$SmokeCycles = '+$Cycles+[Environment]::NewLine+'$SmokeScenario = '''+$Scenario+''''+[Environment]::NewLine+$remoteLaunch)
