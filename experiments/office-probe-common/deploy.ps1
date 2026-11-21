param([Parameter(Mandatory)][string]$Remote,[ValidateSet('P0','P0Visible','G1Open','G1Create','G1HtmlOpen','G1HtmlManual','G1HtmlExcel','G1HtmlPowerPoint')][string]$Mode='P0',[ValidateRange(10,90)][int]$TimeoutSec=45,[switch]$Loopback)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
if($Remote -notmatch '^[A-Za-z0-9_.-]+$' -or $Remote.StartsWith('-')){throw 'SSH alias required'}
if($Loopback -and $Mode -notlike 'G1Html*'){throw 'Loopback requires an HTML probe'}
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$run=[guid]::NewGuid().ToString('N')
$local=Join-Path $repo ('.local/office-probe-runs/'+$run)
New-Item -ItemType Directory -Path $local|Out-Null
$files=@((Join-Path $repo '.local/build/word-save-probe/WordSaveProbe.exe'),(Join-Path $PSScriptRoot 'supervise.ps1'))
if($Mode -like 'G1*'){$files+=@((Join-Path $repo '.local/build/legacy-trident/TridentDsoProbe.exe'),(Join-Path $repo '.local/build/legacy-trident/oie-legacy-framer.ocx'))}
foreach($file in $files){if(-not(Test-Path -LiteralPath $file)){throw 'Build all probes first'};Copy-Item -LiteralPath $file -Destination $local}
$fixtureNames = switch($Mode){'G1HtmlExcel'{('synthetic-'+$run+'.xlsx');'fixtures.json'} 'G1HtmlPowerPoint'{'synthetic.pptx';'fixtures.json'}}
if($fixtureNames){
 $fixtureArguments = @($local)
 if($Mode -eq 'G1HtmlExcel'){$fixtureArguments += $run}
 & python (Join-Path $repo 'experiments/editor-native-host/suite-fixtures.py') @fixtureArguments
 if($LASTEXITCODE -ne 0){throw 'Office fixture preparation failed'}
 $files += @($fixtureNames | ForEach-Object { Join-Path $local $_ })
}
$assets=@($files|ForEach-Object{[ordered]@{file=[IO.Path]::GetFileName($_);sha256=(Get-FileHash -LiteralPath $_).Hash}})
$archive=Join-Path $local 'payload.zip'
Compress-Archive -LiteralPath @($files|ForEach-Object{Join-Path $local ([IO.Path]::GetFileName($_))}) -DestinationPath $archive
$data=[ordered]@{run=$run;project='office-in-editor';purpose=$(if($Mode -like 'G1*'){'trident-g1'}else{'word-save-probe-p0'});mode=$Mode;timeout=$TimeoutSec;loopback=$Loopback.IsPresent;payloadSha256=(Get-FileHash -LiteralPath $archive).Hash;assets=$assets}
$payload=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($data|ConvertTo-Json -Depth 5 -Compress)))
$setup=@'
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$data=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__PAYLOAD__'))|ConvertFrom-Json
$parent=Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab'
$root=Join-Path $parent ('oie-verified-'+$data.run)
for($p=$parent;$p;$p=Split-Path -Parent $p){if((Test-Path -LiteralPath $p) -and ((Get-Item -LiteralPath $p -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Reparse point refused'}}
if(Test-Path -LiteralPath $root){throw 'Fresh run required'}
New-Item -ItemType Directory -Path $root|Out-Null
[IO.File]::WriteAllText((Join-Path $root 'ownership.json'),($data|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
$root
'@
$encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($setup.Replace('__PAYLOAD__',$payload)))
$response=ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
if($LASTEXITCODE -ne 0){throw 'Remote preparation failed'}
$root=([string]::Join('',[string[]]$response)).Trim()
if($root -notlike ('*\OfficeInEditorLab\oie-verified-'+$run) -or $root -match '["\r\n]'){throw 'Unexpected remote root'}
[IO.File]::WriteAllText((Join-Path $local 'remote.json'),(@{root=$root;run=$run;mode=$Mode}|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
scp -q -o BatchMode=yes $archive ($Remote+':'+($root.Replace('\','/')+'/payload.zip'))
if($LASTEXITCODE -ne 0){throw 'Payload copy failed'}
$runner=@'
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$data=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__PAYLOAD__'))|ConvertFrom-Json
$root=Join-Path (Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab') ('oie-verified-'+$data.run)
$archive=Join-Path $root 'payload.zip'
if((Get-FileHash -LiteralPath $archive).Hash -ne $data.payloadSha256){throw 'Payload digest mismatch'}
Expand-Archive -LiteralPath $archive -DestinationPath $root
foreach($asset in $data.assets){if((Get-FileHash -LiteralPath (Join-Path $root $asset.file)).Hash -ne $asset.sha256){throw 'Asset digest mismatch'}}
Remove-Item -LiteralPath $archive
# Interactive Task Scheduler owns the GUI process tree, not the SSH job.
# A disconnected SSH client cannot terminate a document-holding probe.
$taskName='OfficeInEditor-Probe-'+$data.run
$busy=@(Get-ScheduledTask -TaskName 'OfficeInEditor-Probe-*' -ErrorAction SilentlyContinue|Where-Object{$_.State -eq 'Running'})
if($busy.Count){throw 'An earlier Office probe is still running; recover it before opening another'}
$watcher=[IO.FileSystemWatcher]::new($root,'supervisor.json')
$watcher.EnableRaisingEvents=$true
$taskCreated=$false
try {
 $ps=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
 $arguments='-NoProfile -STA -ExecutionPolicy Bypass -File "'+(Join-Path $root 'supervise.ps1')+'" -RunRoot "'+$root+'" -Mode '+$data.mode+' -TimeoutSec '+$data.timeout
 if($data.loopback){$arguments+=' -Loopback'}
 $action=New-ScheduledTaskAction -Execute $ps -Argument $arguments -WorkingDirectory $root
 $principal=New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
 $settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DisallowHardTerminate -MultipleInstances IgnoreNew
 Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings|Out-Null
 $taskCreated=$true
 Start-ScheduledTask -TaskName $taskName
 if(-not(Test-Path -LiteralPath (Join-Path $root 'supervisor.json'))){
  $change=$watcher.WaitForChanged(([IO.WatcherChangeTypes]::Created -bor [IO.WatcherChangeTypes]::Renamed),($data.timeout+15)*1000)
  if($change.TimedOut){throw 'Supervisor result deadline reached; interactive task retained for recovery'}
 }
 $observed=[IO.File]::ReadAllText((Join-Path $root 'supervisor.json'))|ConvertFrom-Json
 $removed=$false
 if($observed.processExited){
  for($i=0;$i -lt 8;$i++){
   if((Get-ScheduledTask -TaskName $taskName).State -ne 'Running'){Unregister-ScheduledTask -TaskName $taskName -Confirm:$false;$removed=$true;break}
   Start-Sleep -Milliseconds 250
  }
 }
 [IO.File]::WriteAllText((Join-Path $root 'launcher.json'),(@{taskName=$taskName;taskRemoved=$removed;retainedForRecovery=(-not $removed -and -not $observed.manual);manualSession=[bool]$observed.manual;executionHost='interactive-task'}|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
 $observed|Add-Member -NotePropertyName interactiveTaskRemoved -NotePropertyValue $removed
 $observed|ConvertTo-Json -Depth 9 -Compress
 if(-not $observed.passed -and -not($observed.manual -and $observed.ready)){exit 1}
} finally {
 $watcher.Dispose()
 # No Stop-ScheduledTask or termination fallback: preserve unsaved documents.
 if($taskCreated -and -not(Test-Path -LiteralPath (Join-Path $root 'launcher.json'))){
  [IO.File]::WriteAllText((Join-Path $root 'launcher.json'),(@{taskName=$taskName;retainedForRecovery=$true}|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
 }
}
'@
$encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($runner.Replace('__PAYLOAD__',$payload)))
$response=ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
$status=$LASTEXITCODE
$plain=[string]::Join([Environment]::NewLine,[string[]]$response)
[IO.File]::WriteAllText((Join-Path $local 'supervisor.json'),$plain,[Text.UTF8Encoding]::new($false))
$observed=$plain|ConvertFrom-Json
[ordered]@{run=$run;mode=$Mode;application=$observed.result.application;ready=$observed.ready;passed=$observed.passed;exitCode=$observed.exitCode;lastStage=$observed.lastStage;preexistingWordPreserved=$observed.preexistingWordPreserved;preexistingOfficePreserved=$observed.preexistingOfficePreserved;retainedProbePid=$observed.retainedProbePid;interactiveTaskRemoved=$observed.interactiveTaskRemoved;error=$observed.result.errorType;errorHresult=$observed.result.errorHresult;reason=$observed.result.failureReason;documentVerified=$observed.result.documentVerified;nativeWordChild=$observed.result.nativeWordChildVerified;nativeOfficeChild=$observed.result.nativeOfficeChildVerified;documentWindowVisible=$observed.result.documentWindowVisible;periodicSaves=$observed.result.periodicSaveRequests;saveMechanism=$observed.result.periodicSaveMechanism;ownedProcessExited=$observed.result.ownedProcessExited;hostDpi=$observed.result.hostDpiAwareness;wordDpi=$observed.result.documentWindowDpiAwareness;nativeTrace=$observed.result.nativeTrace}|ConvertTo-Json -Depth 5 -Compress
Write-Host ('Local evidence: '+$local)
if($status -ne 0){throw 'Probe did not pass; evidence retained, no process was killed'}
