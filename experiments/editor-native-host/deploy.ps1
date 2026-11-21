param([string]$SshHost = 'ollama-relay', [switch]$Suite, [switch]$WebPage)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
if ($SshHost -notmatch '^[a-zA-Z0-9.-]+$') { throw 'Invalid SSH alias' }
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$run = [guid]::NewGuid().ToString('N')
$local = Join-Path $repo ('.local/editor-native-runs/' + $run)
$extension = Join-Path $local 'extension'
New-Item -ItemType Directory -Path $extension -Force | Out-Null
foreach ($name in @('dist', 'assets', 'payload', 'package.json', 'LICENSE')) {
    Copy-Item -LiteralPath (Join-Path $repo $name) -Destination $extension -Recurse
}
# For the pinned experiment only. The public minimum version is unchanged until verified.
$manifestPath = Join-Path $extension 'package.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$manifest.engines.vscode = '^1.121.0'
[IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 60), [Text.UTF8Encoding]::new($false))
$bundle = Join-Path $local 'deployment.zip'
$items = @($extension, (Join-Path $PSScriptRoot 'run-lab.ps1'), (Join-Path $PSScriptRoot 'driver.cjs'))
if ($WebPage) {
    $marker = Join-Path $local 'web-page.flag'
    [IO.File]::WriteAllText($marker, $run)
    $items += $marker
}
if ($Suite) {
    & python (Join-Path $PSScriptRoot 'suite-fixtures.py') $local
    if ($LASTEXITCODE) { throw 'Synthetic suite fixture generation failed' }
    $items += @('synthetic.xlsx','synthetic.pptx','fixtures.json') | ForEach-Object { Join-Path $local $_ }
}
Compress-Archive -LiteralPath $items -DestinationPath $bundle
$digest = (Get-FileHash -LiteralPath $bundle).Hash

function Invoke-Remote([string]$Body) {
    $bodyWithRun = '$Run = ''' + $run + '''' + [Environment]::NewLine + '$Digest = ''' + $digest + '''' + [Environment]::NewLine + $Body
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($bodyWithRun))
    $result = & ssh -o BatchMode=yes -o ConnectTimeout=8 $SshHost "powershell.exe -NoProfile -EncodedCommand $encoded"
    if ($LASTEXITCODE) { throw 'Remote experiment command failed' }
    $result
}
$remote = Invoke-Remote @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$parent = Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab'
New-Item -ItemType Directory -Path $parent -Force | Out-Null
if ((Get-Item -LiteralPath $parent -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Unexpected experiment parent' }
$root = Join-Path $parent ('editor-slot-' + $Run)
New-Item -ItemType Directory -Path $root | Out-Null
Write-Output ($root.Replace('\','/'))
'@
if ($remote -isnot [string] -or $remote -notmatch ('^[A-Za-z]:/[A-Za-z0-9_./-]+/OfficeInEditorLab/editor-slot-' + $run + '$')) {
    throw 'Unexpected transfer directory'
}
$metadata = @{ run = $run; remote = $remote; ssh = $SshHost } | ConvertTo-Json
[IO.File]::WriteAllText((Join-Path $local 'remote.json'), $metadata)
& scp -o BatchMode=yes -o ConnectTimeout=8 $bundle ($SshHost + ':' + $remote + '/deployment.zip')
if ($LASTEXITCODE) { throw 'Experiment copy failed' }
Invoke-Remote @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$root = Join-Path (Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab') ('editor-slot-' + $Run)
$archive = Join-Path $root 'deployment.zip'
if ((Get-FileHash -LiteralPath $archive).Hash -ne $Digest) { throw 'Native experiment archive hash mismatch' }
Expand-Archive -LiteralPath $archive -DestinationPath $root
& (Join-Path $root 'run-lab.ps1') -Run $Run
'@
