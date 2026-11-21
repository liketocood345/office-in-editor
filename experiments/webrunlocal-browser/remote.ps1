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
        '$root = Join-Path $base (''webrunlocal-browser-'' + $Run)' + "`r`n"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($prefix + $Body))
    $result = & ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
    if ($LASTEXITCODE) { throw 'SSH browser experiment command failed' }
    $result
}

if ($readOnly) {
    Invoke-Remote @'
$result = Join-Path $root 'result.json'
if (Test-Path -LiteralPath $result -PathType Leaf) { Get-Content -LiteralPath $result -Raw -Encoding UTF8 }
else { @{run=$Run;stage='pending'} | ConvertTo-Json -Compress }
'@
    return
}

$remotePath = Invoke-Remote @'
New-Item -ItemType Directory -Path $base -Force | Out-Null
if ((Get-Item -LiteralPath $base -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Redirected lab folder' }
New-Item -ItemType Directory -Path $root | Out-Null
$root.Replace('\','/')
'@
if ($remotePath -isnot [string] -or $remotePath -notmatch ('^[A-Za-z]:/[A-Za-z0-9_./-]+/OfficeInEditorLab/webrunlocal-browser-' + $Run + '$')) {
    throw 'Unexpected experiment destination'
}
$probe = Join-Path $PSScriptRoot 'probe.mjs'
$launch = Join-Path $PSScriptRoot 'launch.ps1'
& scp -q -o BatchMode=yes -o ConnectTimeout=8 $probe $launch ($Remote + ':' + $remotePath + '/')
if ($LASTEXITCODE) { throw 'Browser probe copy failed' }
$probeHash = (Get-FileHash -LiteralPath $probe).Hash
$launchHash = (Get-FileHash -LiteralPath $launch).Hash
$configuration = '$ProbeHash = ''' + $probeHash + "'`r`n" + '$LaunchHash = ''' + $launchHash + "'`r`n"
$start = @'
$launcher = Join-Path $root 'launch.ps1'
if ((Get-FileHash -LiteralPath $launcher).Hash -ne $LaunchHash -or
    (Get-FileHash -LiteralPath (Join-Path $root 'probe.mjs')).Hash -ne $ProbeHash) { throw 'Browser probe checksum mismatch' }
$user = (Get-CimInstance Win32_ComputerSystem).UserName
if (-not $user) { throw 'Interactive desktop unavailable' }
$name = 'OfficeInEditor-WebRunLocal-' + $Run
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$executable = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$arguments = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy RemoteSigned -File "' + $launcher + '" -Run ' + $Run
$action = New-ScheduledTaskAction -Execute $executable -Argument $arguments -WorkingDirectory $root
$settings = New-ScheduledTaskSettingsSet -DisallowHardTerminate -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $name -Action $action -Principal $principal -Settings $settings | Out-Null
Start-ScheduledTask -TaskName $name
@{run=$Run;deployed=$true} | ConvertTo-Json -Compress
'@
Invoke-Remote ($configuration + $start)
