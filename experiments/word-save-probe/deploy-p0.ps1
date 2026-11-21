param([Parameter(Mandatory)][string]$Remote,[ValidateRange(10,90)][int]$TimeoutSec=45,[ValidateSet('probe.ps1','probe-visible.ps1','probe-diag.ps1')][string]$ProbeFile='probe.ps1')
$ErrorActionPreference='Stop'
$mode=if($ProbeFile -eq 'probe.ps1'){'P0'}else{'P0Visible'}
& (Join-Path $PSScriptRoot '../office-probe-common/deploy.ps1') -Remote $Remote -Mode $mode -TimeoutSec $TimeoutSec
