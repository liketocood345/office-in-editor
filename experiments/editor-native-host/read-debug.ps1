param(
    [ValidatePattern('^[a-f0-9]{32}$')][string]$Run,
    [ValidatePattern('^[a-zA-Z0-9][a-zA-Z0-9.-]+$')][string]$SshHost = 'ollama-relay'
)
$ErrorActionPreference = 'Stop'
$localRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../.local/editor-native-runs'))
if (-not $Run) {
    $latest = Get-ChildItem -LiteralPath $localRoot -Directory |
        Where-Object { $_.Name -match '^[a-f0-9]{32}$' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'remote.json')) } |
        Sort-Object CreationTime -Descending | Select-Object -First 1
    if (-not $latest) { throw 'No deployed native editor run' }
    $Run = $latest.Name
}
$metadata = Get-Content -LiteralPath (Join-Path (Join-Path $localRoot $Run) 'remote.json') -Raw | ConvertFrom-Json
if ($metadata.run -ne $Run -or $metadata.ssh -ne $SshHost) { throw 'Native editor run does not match the SSH host' }
$body = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$root = Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\editor-slot-' + $Run)
$resultFile = Join-Path $root 'result.json'
$result = if (Test-Path -LiteralPath $resultFile) { Get-Content -LiteralPath $resultFile -Raw | ConvertFrom-Json }
$formatStages = @(foreach ($application in 'word','excel','powerpoint') {
    $file = Join-Path $root ($application + '-page-stage.json')
    if ((Test-Path -LiteralPath $file -PathType Leaf) -and (Get-Item -LiteralPath $file).Length -le 8192) {
        Get-Content -LiteralPath $file -Raw -Encoding UTF8 | ConvertFrom-Json
    }
})
$logs = @(Get-ChildItem -LiteralPath (Join-Path $root 'profile\logs') -Recurse -File -Filter '*Office In Editor*.log' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 2)
$events = @(foreach ($file in $logs) {
    foreach ($line in (Get-Content -LiteralPath $file.FullName -Tail 50)) {
        if ($line -match '^open=[a-f0-9]{8} stage=[a-zA-Z0-9_]+ elapsedMs=[0-9]+$') { $line }
        elseif ($line.StartsWith('{"nativeSurface":')) {
            $values = $line | ConvertFrom-Json
            [pscustomobject]@{ attempt=$values.attempt; nativeSurface=$values.nativeSurface }
        }
    }
})
[pscustomobject]@{ run=$Run; resultReady=($null -ne $result); result=$result; formatStages=$formatStages; events=$events } | ConvertTo-Json -Depth 10
'@
$script = '$Run = ''' + $Run + '''' + [Environment]::NewLine + $body
$encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($script))
& ssh -o BatchMode=yes -o ConnectTimeout=8 $SshHost "powershell.exe -NoProfile -EncodedCommand $encoded"
if ($LASTEXITCODE) { throw 'Native editor diagnostic read failed' }
