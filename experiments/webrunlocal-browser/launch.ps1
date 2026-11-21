param([Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$root = Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\webrunlocal-browser-' + $Run)
$taskName = 'OfficeInEditor-WebRunLocal-' + $Run
$profile = Join-Path $root 'edge-profile'
if ($root -ine $PSScriptRoot -or (Get-Process -Id $PID).SessionId -eq 0) { throw 'Unexpected browser lab session' }
$edge = @(
    (Join-Path ([Environment]::GetEnvironmentVariable('ProgramFiles(x86)')) 'Microsoft\Edge\Application\msedge.exe'),
    (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe')
) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
$step = 'preflight'
$status = [ordered]@{ run = $Run; browserExited = $false }
try {
    if (-not $edge) { throw 'Microsoft Edge unavailable' }
    $node = (Get-Command node.exe -ErrorAction Stop).Source
    $step = 'browser_probe'
    $probe = & $node (Join-Path $root 'probe.mjs') $edge $profile | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0 -or -not $probe) { throw 'Browser probe returned no result' }
    $status.probe = $probe
    $step = 'close_owned_browser'
    $edgeProcess = if ($probe.browserPid) { Get-Process -Id $probe.browserPid -ErrorAction SilentlyContinue }
    if ($edgeProcess -and $edgeProcess.Path -ieq $edge) {
        $record = Get-CimInstance Win32_Process -Filter ('ProcessId = ' + $edgeProcess.Id)
        if ($record.CommandLine -and $record.CommandLine.Contains($profile) -and $edgeProcess.CloseMainWindow()) {
            $status.browserExited = $edgeProcess.WaitForExit(5000)
        }
    } elseif (-not $edgeProcess) { $status.browserExited = $true }
    $step = 'complete'
} catch {
    $status.error = $_.Exception.GetType().Name
} finally {
    $status.stage = $step
    $temporary = Join-Path $root 'result.partial'
    [IO.File]::WriteAllText($temporary, ($status | ConvertTo-Json -Depth 10 -Compress), [Text.UTF8Encoding]::new($false))
    [IO.File]::Move($temporary, (Join-Path $root 'result.json'))
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
}
