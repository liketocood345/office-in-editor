param([Parameter(Mandatory)][string]$Remote)
$ErrorActionPreference='Stop'
if($Remote -notmatch '^[A-Za-z0-9_.-]+$' -or $Remote.StartsWith('-')){throw 'SSH alias required'}
$run=[guid]::NewGuid().ToString('N')
$rootLocal=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot ('../../.local/isolation-checks/'+$run)))
New-Item -ItemType Directory -Path $rootLocal|Out-Null
$data=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes((@{run=$run;supervisor=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'supervise.ps1'))}|ConvertTo-Json -Compress)))
$setup=@'
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$data=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__DATA__'))|ConvertFrom-Json
$root=Join-Path (Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab') ('isolation-'+$data.run)
if(Test-Path -LiteralPath $root){throw 'Fresh fixture required'}
New-Item -ItemType Directory -Path $root|Out-Null
[IO.File]::WriteAllText((Join-Path $root 'ownership.json'),(@{project='office-in-editor';run=$data.run;purpose='synthetic-isolation'}|ConvertTo-Json))
[IO.File]::WriteAllText((Join-Path $root 'supervise.ps1'),$data.supervisor)
# A synthetic stand-in never activates Office or creates a window. It lives
# beyond the supervisor timeout and writes a completion marker before exit.
$source=@"
using System; using System.IO; using System.Threading;
public static class DelayedSyntheticProbe {
 public static void Main(string[] args) {
  if(args.Length==2 && args[1]=="--prepare") return;
  Thread.Sleep(28000); File.WriteAllText(Path.Combine(args[0],"completed.marker"),"complete");
 }
}
"@
Add-Type -TypeDefinition $source -OutputAssembly (Join-Path $root 'WordSaveProbe.exe') -OutputType ConsoleApplication
$name='OfficeInEditor-Probe-'+$data.run
$before=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$watcher=[IO.FileSystemWatcher]::new($root,'supervisor.json');$watcher.EnableRaisingEvents=$true
try {
 $action=New-ScheduledTaskAction -Execute (Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe') -Argument ('-NoProfile -STA -ExecutionPolicy Bypass -File "'+(Join-Path $root 'supervise.ps1')+'" -RunRoot "'+$root+'" -TimeoutSec 10')
 $principal=New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
 $settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -DisallowHardTerminate -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
 Register-ScheduledTask -TaskName $name -Action $action -Principal $principal -Settings $settings|Out-Null
 Start-ScheduledTask -TaskName $name
 if(-not(Test-Path -LiteralPath (Join-Path $root 'supervisor.json'))){$event=$watcher.WaitForChanged(([IO.WatcherChangeTypes]::Created -bor [IO.WatcherChangeTypes]::Renamed),18000);if($event.TimedOut){throw 'Isolation fixture did not produce its timeout result'}}
 $j=[IO.File]::ReadAllText((Join-Path $root 'supervisor.json'))|ConvertFrom-Json
 $p=Get-Process -Id $j.retainedProbePid
 @{run=$data.run;root=$root;task=$name;pid=$p.Id;startTicks=$p.StartTime.ToUniversalTime().Ticks;before=$before;timedOut=$j.timeout}|ConvertTo-Json -Compress
} finally {$watcher.Dispose()}
'@
$encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($setup.Replace('__DATA__',$data)))
$response=ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
if($LASTEXITCODE){throw 'Isolation setup failed; inspect the owned fixture task'}
$plain=[string]::Join('',[string[]]$response)
[IO.File]::WriteAllText((Join-Path $rootLocal 'first-connection.json'),$plain)
$secondPayload=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($plain))
$check=@'
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$data=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__DATA__'))|ConvertFrom-Json
$p=Get-Process -Id $data.pid
if($p.StartTime.ToUniversalTime().Ticks -ne $data.startTicks -or $p.Path -ne (Join-Path $data.root 'WordSaveProbe.exe')){throw 'Synthetic process identity changed'}
$survived=$true
$exited=$p.WaitForExit(25000)
$removed=$false
if($exited){for($i=0;$i -lt 8;$i++){if((Get-ScheduledTask -TaskName $data.task).State -ne 'Running'){Unregister-ScheduledTask -TaskName $data.task -Confirm:$false;$removed=$true;break};Start-Sleep -Milliseconds 250}}
$after=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$result=@{passed=($data.timedOut -and $survived -and $exited -and $removed -and (Test-Path -LiteralPath (Join-Path $data.root 'completed.marker')) -and (($data.before -join ',') -eq ($after -join ',')));survivedSshDisconnect=$survived;exitedNormally=$exited;taskRemoved=$removed;officePidsUnchanged=(($data.before -join ',') -eq ($after -join ','))}
[IO.File]::WriteAllText((Join-Path $data.root 'isolation-result.json'),($result|ConvertTo-Json))
$result|ConvertTo-Json -Compress
if(-not $result.passed){exit 1}
'@
$encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($check.Replace('__DATA__',$secondPayload)))
$response=ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
$status=$LASTEXITCODE
$plain=[string]::Join([Environment]::NewLine,[string[]]$response)
[IO.File]::WriteAllText((Join-Path $rootLocal 'result.json'),$plain)
$plain
if($status){throw 'SSH disconnect isolation did not pass'}
