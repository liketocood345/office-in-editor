param([Parameter(Mandatory)][string]$RunRoot,[ValidateSet('P0','P0Visible','G1Open','G1Create','G1HtmlOpen','G1HtmlManual','G1HtmlExcel','G1HtmlPowerPoint')][string]$Mode='P0',[ValidateRange(10,90)][int]$TimeoutSec=45,[switch]$Loopback)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath($RunRoot).TrimEnd('\')
$owner=[IO.File]::ReadAllText((Join-Path $root 'ownership.json'))|ConvertFrom-Json
if($owner.project -ne 'office-in-editor' -or $owner.run -notmatch '^[a-f0-9]{32}$'){throw 'Invalid ownership'}
if($root -match '["\r\n]'){throw 'Unsupported argument characters'}
$result=Join-Path $root 'result.json'
if(Test-Path -LiteralPath $result){throw 'Fresh run required'}
$pre=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$appName=switch($Mode){'G1HtmlExcel'{'EXCEL'} 'G1HtmlPowerPoint'{'POWERPNT'} default{'WINWORD'}}
$appBefore=@(Get-Process -Name $appName -ErrorAction SilentlyContinue | ForEach-Object { '{0}:{1}' -f $_.Id,$_.StartTime.ToUniversalTime().Ticks })
$fixture=switch($Mode){'G1HtmlExcel'{('synthetic-'+$owner.run+'.xlsx')} 'G1HtmlPowerPoint'{'synthetic.pptx'} default{'synthetic.docx'}}
$wordProbe=Join-Path $root 'WordSaveProbe.exe'
if($fixture -eq 'synthetic.docx'){
$prepare=Start-Process -FilePath $wordProbe -ArgumentList @(('"'+$root+'"'),'--prepare') -WindowStyle Hidden -PassThru
[void]$prepare.Handle
if(-not $prepare.WaitForExit(5000)){throw 'Fixture preparation retained for diagnosis'}
$prepare.Refresh()
if($prepare.ExitCode -ne 0){throw 'Fixture preparation failed'}
}
$exe=$wordProbe
$arguments=@(('"'+$root+'"'))
if($Mode -eq 'P0Visible'){$arguments+='--visible'}
if($Mode -like 'G1*'){
 $exe=Join-Path $root 'TridentDsoProbe.exe'
 $arguments=@(('"--dso-doc='+(Join-Path $root $fixture)+'"'),('"--own-root='+$root+'"'))
 if($Mode -eq 'G1Create'){$arguments+='--create-new'}
 if($Mode -like 'G1Html*'){$arguments+='--html-object'}
 if($Loopback){$arguments+='--loopback'}
 if($Mode -eq 'G1HtmlManual'){$arguments+='--manual'}
}
$manualWatcher=$null
if($Mode -eq 'G1HtmlManual'){
 $manualWatcher=[IO.FileSystemWatcher]::new($root,'manual-ready.json')
 $manualWatcher.EnableRaisingEvents=$true
}
$windowStyle = if($Mode -like 'G1Html*'){'Normal'}else{'Hidden'}
$child=Start-Process -FilePath $exe -WorkingDirectory $root -ArgumentList $arguments -WindowStyle $windowStyle -RedirectStandardOutput (Join-Path $root 'stdout.txt') -RedirectStandardError (Join-Path $root 'stderr.txt') -PassThru
[void]$child.Handle
$clock=[Diagnostics.Stopwatch]::StartNew()
if($Mode -eq 'G1HtmlManual'){
 try {
  if(-not(Test-Path -LiteralPath (Join-Path $root 'manual-ready.json'))){
   $change=$manualWatcher.WaitForChanged(([IO.WatcherChangeTypes]::Created -bor [IO.WatcherChangeTypes]::Renamed),$TimeoutSec*1000)
   if($change.TimedOut){throw 'Manual window readiness deadline reached; task retained'}
  }
  $ready=[IO.File]::ReadAllText((Join-Path $root 'manual-ready.json'))|ConvertFrom-Json
  $summary=@{run=$owner.run;mode=$Mode;ready=($ready.nativeWordChildVerified -eq $true);manual=$true;processExited=$false;retainedProbePid=$child.Id;result=$ready}
  $pending=Join-Path $root 'supervisor.pending'
  [IO.File]::WriteAllText($pending,($summary|ConvertTo-Json -Depth 9),[Text.UTF8Encoding]::new($false))
  [IO.File]::Move($pending,(Join-Path $root 'supervisor.json'))
 } finally {$manualWatcher.Dispose()}
 # Keep the interactive task alive until the user saves and closes normally.
 # Waiting on a handle performs no periodic disk reads/writes.
 $child.WaitForExit()
 $child.Refresh()
 $probe=$null
 if(Test-Path -LiteralPath $result){$probe=[IO.File]::ReadAllText($result)|ConvertFrom-Json}
 $taskName='OfficeInEditor-Probe-'+$owner.run
 Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
 $final=@{run=$owner.run;mode=$Mode;passed=($child.ExitCode -eq 0 -and $probe.passed -eq $true);processExited=$true;exitCode=$child.ExitCode;interactiveTaskRemoved=$true;result=$probe}
 [IO.File]::WriteAllText((Join-Path $root 'supervisor-final.json'),($final|ConvertTo-Json -Depth 9),[Text.UTF8Encoding]::new($false))
 exit $(if($final.passed){0}else{1})
}
# Wait on a process handle, not a continuously written log. A timeout never
# kills Office or a host that might still hold an unsaved synthetic document.
while(-not $child.WaitForExit(1000) -and $clock.Elapsed.TotalSeconds -lt $TimeoutSec){}
$child.Refresh()
$probe=$null
if(Test-Path -LiteralPath $result){$probe=[IO.File]::ReadAllText($result)|ConvertFrom-Json}
$cleanupAccepted=$true
if($probe.ownedProcessExitDeferred -eq $true){
 $cleanupAccepted=$false
 if($child.HasExited -and $probe.application -eq 'excel' -and $probe.applicationIsOwned -eq $true){
  $office=Get-Process -Id ([int]$probe.applicationPid) -ErrorAction SilentlyContinue
  try {
   $cleanupAccepted= -not $office -or $office.StartTime.ToUniversalTime().Ticks -ne [long]$probe.applicationStartUtcTicks -or $office.WaitForExit(5000)
  } catch {
   $cleanupAccepted= -not (Get-Process -Id ([int]$probe.applicationPid) -ErrorAction SilentlyContinue)
  } finally { if($office){$office.Dispose()} }
 }
 $probe | Add-Member -NotePropertyName ownedProcessExited -NotePropertyValue $cleanupAccepted -Force
}
$checkpoint=$null
if(Test-Path -LiteralPath (Join-Path $root 'checkpoint.json')){$checkpoint=[IO.File]::ReadAllText((Join-Path $root 'checkpoint.json'))|ConvertFrom-Json}
$after=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$appAfter=@(Get-Process -Name $appName -ErrorAction SilentlyContinue | ForEach-Object { '{0}:{1}' -f $_.Id,$_.StartTime.ToUniversalTime().Ticks })
$summary=[ordered]@{run=$owner.run;mode=$Mode;passed=($child.HasExited -and $child.ExitCode -eq 0 -and $null -ne $probe -and $probe.passed -eq $true -and $probe.documentVerified -eq $true -and $cleanupAccepted);processExited=$child.HasExited;retainedProbePid=$(if(-not $child.HasExited){$child.Id}else{$null});exitCode=$(if($child.HasExited){$child.ExitCode}else{$null});timeout=(-not $child.HasExited);elapsedMs=$clock.ElapsedMilliseconds;lastStage=$checkpoint.stage;preexistingWordPreserved=(@($pre|Where-Object{$after -notcontains $_}).Count -eq 0);result=$probe}
$summary.preexistingOfficePreserved=(@($appBefore | Where-Object {$appAfter -notcontains $_}).Count -eq 0)
$pending=Join-Path $root 'supervisor.pending'
[IO.File]::WriteAllText($pending,($summary|ConvertTo-Json -Depth 9),[Text.UTF8Encoding]::new($false))
[IO.File]::Move($pending,(Join-Path $root 'supervisor.json'))
$summary|ConvertTo-Json -Depth 9 -Compress
if(-not $summary.passed){exit 1}
