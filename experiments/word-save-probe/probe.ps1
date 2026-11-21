param([Parameter(Mandatory)][string]$RunRoot,[string]$Marker,[switch]$Visible)
$ErrorActionPreference='Stop'
$exe=Join-Path $PSScriptRoot 'WordSaveProbe.exe'
if(-not(Test-Path -LiteralPath $exe)){$exe=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../.local/build/word-save-probe/WordSaveProbe.exe'))}
if(-not(Test-Path -LiteralPath $exe)){throw 'Build WordSaveProbe first'}
$arguments=@($RunRoot)
if($Visible){$arguments+='--visible'}
& $exe @arguments
exit $LASTEXITCODE
