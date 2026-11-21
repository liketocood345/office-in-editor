param(
    [Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9][a-zA-Z0-9._-]*$')][string]$Remote,
    [ValidatePattern('^[a-f0-9]{32}$')][string]$Run
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$readOnly = [bool]$Run
if (-not $Run) { $Run = [Guid]::NewGuid().ToString('N') }

function Invoke-Remote([string]$Body) {
    $prefix = '$ErrorActionPreference = ''Stop''; $ProgressPreference = ''SilentlyContinue''; $Run = ''' + $Run + "'`r`n" +
        '$base = Join-Path $env:LOCALAPPDATA ''OfficeInEditorLab''' + "`r`n" +
        '$root = Join-Path $base (''office-browser-'' + $Run)' + "`r`n"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($prefix + $Body))
    $result = & ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
    if ($LASTEXITCODE) { throw 'Remote browser probe command failed' }
    $result
}

if ($readOnly) {
    Invoke-Remote @'
$summary = @{ run = $Run; stages = @{} }
foreach ($stage in 'launch','activated','word','excel','powerpoint','result','finished','failure') {
    $record = Join-Path $root ($stage + '.json')
    if ((Test-Path -LiteralPath $record -PathType Leaf) -and (Get-Item -LiteralPath $record).Length -lt 32768) {
        $summary.stages[$stage] = Get-Content -LiteralPath $record -Raw -Encoding UTF8 | ConvertFrom-Json
    }
}
$summary | ConvertTo-Json -Depth 15 -Compress
'@
    return
}

$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$local = Join-Path $repo ('.local/office-browser-runs/' + $Run)
$probe = Join-Path $local 'probe'
New-Item -ItemType Directory -Path $probe -Force | Out-Null
foreach ($name in 'package.json','extension.cjs','page.js') {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $probe
}
$archive = Join-Path $local 'probe.zip'
Compress-Archive -LiteralPath @($probe, (Join-Path $PSScriptRoot 'launch.ps1')) -DestinationPath $archive
$digest = (Get-FileHash -LiteralPath $archive).Hash
$remotePath = Invoke-Remote @'
New-Item -ItemType Directory -Path $base -Force | Out-Null
if ((Get-Item -LiteralPath $base -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Redirected lab directory' }
New-Item -ItemType Directory -Path $root | Out-Null
$root.Replace('\','/')
'@
if ($remotePath -isnot [string] -or $remotePath -notmatch ('^[A-Za-z]:/[A-Za-z0-9_./-]+/OfficeInEditorLab/office-browser-' + $Run + '$')) {
    throw 'Unexpected transfer path'
}
& scp -q -o BatchMode=yes -o ConnectTimeout=8 $archive ($Remote + ':' + $remotePath + '/probe.zip')
if ($LASTEXITCODE) { throw 'Probe transfer failed' }
$start = @'
$archive = Join-Path $root 'probe.zip'
if ((Get-FileHash -LiteralPath $archive).Hash -ne $ExpectedHash) { throw 'Probe archive hash mismatch' }
Expand-Archive -LiteralPath $archive -DestinationPath $root
$user = (Get-CimInstance Win32_ComputerSystem).UserName
if (-not $user) { throw 'Interactive desktop user unavailable' }
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$executable = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$params = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy RemoteSigned -File "' + (Join-Path $root 'launch.ps1') + '" -Run ' + $Run
$action = New-ScheduledTaskAction -Execute $executable -Argument $params -WorkingDirectory $root
$settings = New-ScheduledTaskSettingsSet -DisallowHardTerminate -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$name = 'OfficeInEditor-Browser-' + $Run
Register-ScheduledTask -TaskName $name -Action $action -Principal $principal -Settings $settings | Out-Null
Start-ScheduledTask -TaskName $name
@{run=$Run; started=$true} | ConvertTo-Json -Compress
'@
Invoke-Remote ('$ExpectedHash = ''' + $digest + "'`r`n" + $start)
