param()
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$exe=Join-Path $repo '.local/build/word-save-probe/WordSaveProbe.exe'
$root=Join-Path $repo ('.local/probe-contracts/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$before=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$checks=0
function Check([bool]$condition,[string]$name){if(-not $condition){throw $name};$script:checks++;Write-Output ('PASS '+$name)}
function Invoke-Probe([string]$folder,[string]$option){
 $args=@(('"'+$folder+'"'));if($option){$args+=$option}
 $p=Start-Process -FilePath $exe -ArgumentList $args -WindowStyle Hidden -RedirectStandardOutput (Join-Path $folder 'contract-stdout.txt') -RedirectStandardError (Join-Path $folder 'contract-stderr.txt') -PassThru
 if(-not $p.WaitForExit(5000)){throw 'Contract unexpectedly blocked; process retained'}
 $p.Refresh();return $p.ExitCode
}
# Bad metadata must fail before any Office activation.
$invalid=Join-Path $root 'invalid'
New-Item -ItemType Directory -Path $invalid|Out-Null
[IO.File]::WriteAllText((Join-Path $invalid 'ownership.json'),'{}')
Check ((Invoke-Probe $invalid '--prepare') -ne 0) 'invalid ownership is rejected'
Check (-not(Test-Path -LiteralPath (Join-Path $invalid 'synthetic.docx'))) 'invalid ownership cannot create a fixture'

$valid=Join-Path $root 'valid'
New-Item -ItemType Directory -Path $valid|Out-Null
$owner=@{project='office-in-editor';purpose='word-save-probe-p0';run=[guid]::NewGuid().ToString('N')}
[IO.File]::WriteAllText((Join-Path $valid 'ownership.json'),($owner|ConvertTo-Json))
Check ((Invoke-Probe $valid '--prepare') -eq 0) 'owned fixture preparation exits zero'
Check (Test-Path -LiteralPath (Join-Path $valid 'synthetic.json')) 'fixture identity manifest exists'
$hash=(Get-FileHash -LiteralPath (Join-Path $valid 'synthetic.docx')).Hash
Check ((Invoke-Probe $valid '--prepare') -ne 0) 'repeat preparation cannot overwrite a fixture'
Check ((Get-FileHash -LiteralPath (Join-Path $valid 'synthetic.docx')).Hash -eq $hash) 'rejected overwrite preserves original bytes'

[IO.File]::AppendAllText((Join-Path $valid 'synthetic.docx'),'tamper')
Check ((Invoke-Probe $valid '') -ne 0) 'modified fixture fails with a nonzero exit'
$result=[IO.File]::ReadAllText((Join-Path $valid 'result.json'))|ConvertFrom-Json
Check ($result.passed -eq $false) 'failed fixture cannot report success'
Check (@($result.stages).Count -eq 0) 'fixture rejection happens before COM phases'
Check ((Invoke-Probe $valid '--recover') -ne 0) 'recovery refuses a run without a proven owned application'
$recoveryFile=Get-ChildItem -LiteralPath $valid -Filter 'recovery-*.json'|Select-Object -First 1
$recovery=[IO.File]::ReadAllText($recoveryFile.FullName)|ConvertFrom-Json
Check ($recovery.passed -eq $false) 'rejected recovery never reports success'
$after=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
Check (($before -join ',') -eq ($after -join ',')) 'negative contracts start no Word process'
Write-Output ('Completed '+$checks+' contract checks.')
