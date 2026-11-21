param(
    [Parameter(Mandatory)][string]$Remote,
    [switch]$TwoFiles,
    [ValidateSet('Edge','VSCode')][string]$Client = 'Edge',
    [switch]$VerifyExit
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
if ($Remote -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') { throw 'SSH alias required' }
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$id = [Guid]::NewGuid().ToString('N')
$secondId = if ($TwoFiles -or $Client -eq 'VSCode') { [Guid]::NewGuid().ToString('N') } else { '' }
$local = Join-Path $repo ('.local/word-browser-runs/' + $id)
New-Item -ItemType Directory -Path $local | Out-Null
$names = @('WordWorker.exe', 'launch.ps1', 'server.mjs', 'document-ports.mjs', 'index.html', 'app.js', 'style.css')
foreach ($name in $names) {
    $source = if ($name -eq 'WordWorker.exe') { Join-Path $repo '.local/build/word-browser-local/WordWorker.exe' }
        else { Join-Path $PSScriptRoot $name }
    Copy-Item -LiteralPath $source -Destination $local
}
$items = @($names | ForEach-Object { Join-Path $local $_ })
if ($Client -eq 'VSCode') {
    $lab = Join-Path $local 'vscode-lab'
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'vscode-lab') -Destination $lab -Recurse
    $media = Join-Path $lab 'media'
    New-Item -ItemType Directory -Path $media | Out-Null
    foreach ($name in @('index.html','app.js','style.css')) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $media
    }
    $items += $lab
}
$archive = Join-Path $local 'payload.zip'
Compress-Archive -LiteralPath $items -DestinationPath $archive
$digest = (Get-FileHash -LiteralPath $archive).Hash
function Invoke-Remote([string]$Body) {
    $code = '$Run = ''' + $id + '''' + [Environment]::NewLine + '$SecondRun = ''' + $secondId + '''' + [Environment]::NewLine +
        '$Client = ''' + $Client + '''' + [Environment]::NewLine + '$VerifyExit = $' + $VerifyExit.IsPresent.ToString().ToLowerInvariant() + [Environment]::NewLine +
        '$Digest = ''' + $digest + '''' + [Environment]::NewLine + $Body
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
    $result = & ssh -o BatchMode=yes -o ConnectTimeout=8 $Remote "powershell.exe -NoProfile -EncodedCommand $encoded"
    if ($LASTEXITCODE) { throw 'Remote command failed' }
    return $result
}
$remoteRoot = Invoke-Remote @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$parent = Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab'
New-Item -ItemType Directory -Path $parent -Force | Out-Null
if ((Get-Item -LiteralPath $parent -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Redirected lab parent' }
foreach ($id in @($Run, $SecondRun) | Where-Object { $_ }) {
    $root = Join-Path $parent ('word-web-' + $id)
    New-Item -ItemType Directory -Path $root | Out-Null
    $manifest = @{ project = 'office-in-editor'; purpose = 'word-browser-local'; run = $id; payloadSha256 = $Digest }
    [IO.File]::WriteAllText((Join-Path $root 'ownership.json'), ($manifest | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
}
Write-Output ((Join-Path $parent ('word-web-' + $Run)).Replace('\','/'))
'@
if ($remoteRoot -isnot [string] -or $remoteRoot -notmatch ('^[A-Za-z]:/[A-Za-z0-9_./-]+/OfficeInEditorLab/word-web-' + $id + '$')) {
    throw 'Invalid transfer directory'
}
[IO.File]::WriteAllText((Join-Path $local 'remote.json'), (@{ run = $id; secondRun = $secondId; root = $remoteRoot; remote = $Remote } | ConvertTo-Json))
& scp -q -o BatchMode=yes -o ConnectTimeout=8 $archive ($Remote + ':' + $remoteRoot + '/payload.zip')
if ($LASTEXITCODE) { throw 'Experiment transfer failed' }
Invoke-Remote @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$root = Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\word-web-' + $Run)
$archive = Join-Path $root 'payload.zip'
if ((Get-FileHash -LiteralPath $archive).Hash -ne $Digest) { throw 'Experiment archive mismatch' }
Expand-Archive -LiteralPath $archive -DestinationPath $root
$user = (Get-CimInstance Win32_ComputerSystem).UserName
if (-not $user) { throw 'No interactive desktop user' }
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$args = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy RemoteSigned -File "' + (Join-Path $root 'launch.ps1') + '" -Run ' + $Run
if ($SecondRun) { $args += ' -SecondRun ' + $SecondRun }
$args += ' -Client ' + $Client
if ($VerifyExit) { $args += ' -VerifyExit' }
$action = New-ScheduledTaskAction -Execute $ps -Argument $args -WorkingDirectory $root
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DisallowHardTerminate -ExecutionTimeLimit ([TimeSpan]::Zero)
$task = 'OfficeInEditor-WordWeb-' + $Run
Register-ScheduledTask -TaskName $task -Action $action -Principal $principal -Settings $settings | Out-Null
Start-ScheduledTask -TaskName $task
@{ run = $Run; secondRun = $SecondRun; client = $Client; started = $true } | ConvertTo-Json -Compress
'@
