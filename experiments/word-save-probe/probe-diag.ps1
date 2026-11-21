param([Parameter(Mandatory)][string]$RunRoot,[string]$Marker)
# Visible diagnostics share the same bounded compiled implementation.
& (Join-Path $PSScriptRoot 'probe.ps1') -RunRoot $RunRoot -Visible
exit $LASTEXITCODE
