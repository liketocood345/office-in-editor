param(
    [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run,
    [switch]$Interactive
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$parent = Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab'
$root = Join-Path $parent ('editor-slot-' + $Run)
if ($PSScriptRoot -ine $root) { throw 'Unexpected experiment directory' }
foreach ($directory in @($parent, $root)) {
    if ((Get-Item -LiteralPath $directory -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Redirected experiment directory' }
}
$taskName = 'OfficeInEditor-Slot-' + $Run
$utf8 = [Text.UTF8Encoding]::new($false)
$editor = Join-Path $root 'editor'
$profile = Join-Path $root 'profile'
$extensions = Join-Path $root 'extensions'
$shared = Join-Path $root 'shared'
$extension = Join-Path $root 'extension'
$exe = Join-Path $editor 'Code.exe'
$fixture = Join-Path $root 'synthetic.docx'

if ($Interactive) {
    if ((Get-Process -Id $PID).SessionId -eq 0) { throw 'Interactive desktop unavailable' }
    $env:OIE_NATIVE_LAB_ROOT = $root
    if (Test-Path -LiteralPath (Join-Path $root 'fixtures.json')) { $env:OIE_NATIVE_LAB_SUITE = '1' }
    $webPageFlag = Join-Path $root 'web-page.flag'
    if (Test-Path -LiteralPath $webPageFlag) {
        if ([IO.File]::ReadAllText($webPageFlag) -cne $Run) { throw 'Invalid browser experiment marker' }
        $env:OIE_NATIVE_LAB_HTTP = '1'
    }
    $arguments = @('--new-window', '--skip-welcome', '--skip-release-notes',
        ('--user-data-dir="' + $profile + '"'), ('--shared-data-dir="' + $shared + '"'),
        ('--extensions-dir="' + $extensions + '"'),
        ('--extensionDevelopmentPath="' + $extension + '"'))
    $process = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $editor -PassThru -RedirectStandardError (Join-Path $root 'editor-stderr.txt')
    [IO.File]::WriteAllText((Join-Path $root 'launch.json'), (@{ pid = $process.Id; run = $Run } | ConvertTo-Json), $utf8)
    return
}

$source = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code'
$fixtureManifest = Join-Path $root 'fixtures.json'
if (Test-Path -LiteralPath $fixtureManifest) {
    $hashes = Get-Content -LiteralPath $fixtureManifest -Raw | ConvertFrom-Json
    foreach ($name in @('synthetic.xlsx','synthetic.pptx')) {
        $file = Join-Path $root $name
        $item = Get-Item -LiteralPath $file
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $item.Length -gt 1048576 -or
            (Get-FileHash -LiteralPath $file).Hash -ine $hashes.$name) { throw 'Unexpected Office fixture' }
    }
}
$matches = @(Get-ChildItem -LiteralPath $source -Directory | Where-Object {
    Test-Path -LiteralPath (Join-Path $_.FullName 'resources\app\package.json')
})
if ($matches.Count -ne 1) { throw 'Expected one installed editor build' }
$relative = Join-Path $matches[0].Name 'resources\app'
$manifest = Get-Content -LiteralPath (Join-Path $source "$relative\package.json") -Raw | ConvertFrom-Json
if ($manifest.version -ne '1.121.0' -or $manifest.type -ne 'module' -or $manifest.main -ne './out/main.js') {
    throw 'The experiment only targets the verified editor build'
}
if ((Test-Path -LiteralPath $editor) -or (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)) {
    throw 'Experiment already prepared'
}
& robocopy.exe $source $editor /E /XJ /R:0 /W:0 /NFL /NDL /NJH /NJS /NP /XF 'unins*.exe' 'unins*.dat' | Out-Null
if ($LASTEXITCODE -ge 8) { throw 'Could not prepare the isolated editor' }
$out = Join-Path $editor "$relative\out"
$main = Join-Path $out 'main.js'
$assets = Join-Path $out 'oie-slot-lab'
New-Item -ItemType Directory -Path $assets, $shared, (Join-Path $profile 'User'), (Join-Path $extensions 'oie-local.slot-driver-0.0.1') -Force | Out-Null
$payload = Join-Path $extension 'payload\host'
foreach ($file in @('provider.mjs', 'oie-native-slot.node')) {
    Copy-Item -LiteralPath (Join-Path $payload $file) -Destination $assets
}
$sourceHash = (Get-FileHash -LiteralPath (Join-Path $source "$relative\out\main.js")).Hash
if ((Get-FileHash -LiteralPath $main).Hash -ne $sourceHash) { throw 'Editor copy mismatch' }
[IO.File]::WriteAllText($main, "import './oie-slot-lab/provider.mjs';" + [Environment]::NewLine + [IO.File]::ReadAllText($main), $utf8)
$settings = @{
    'telemetry.telemetryLevel' = 'off'
    'update.mode' = 'none'
    'extensions.autoUpdate' = $false
    'workbench.startupEditor' = 'none'
    'window.restoreWindows' = 'none'
}
[IO.File]::WriteAllText((Join-Path $profile 'User\settings.json'), ($settings | ConvertTo-Json), $utf8)
$driver = Join-Path $extensions 'oie-local.slot-driver-0.0.1'
$driverManifest = @{
    name = 'slot-driver'; version = '0.0.1'; publisher = 'oie-local'
    engines = @{ vscode = '^1.121.0' }; activationEvents = @('onStartupFinished'); main = './driver.cjs'
}
[IO.File]::WriteAllText((Join-Path $driver 'package.json'), ($driverManifest | ConvertTo-Json -Depth 4), $utf8)
Copy-Item -LiteralPath (Join-Path $root 'driver.cjs') -Destination $driver

# Generate the test document without reading an existing user's file.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($fixture, [IO.Compression.ZipArchiveMode]::Create)
try {
    $parts = [ordered]@{
        '[Content_Types].xml' = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        '_rels/.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        'word/document.xml' = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Office in Editor synthetic slot test</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
    }
    foreach ($part in $parts.GetEnumerator()) {
        $writer = [IO.StreamWriter]::new($zip.CreateEntry($part.Key).Open(), $utf8)
        try { $writer.Write($part.Value) } finally { $writer.Dispose() }
    }
} finally { $zip.Dispose() }

$user = (Get-CimInstance Win32_ComputerSystem).UserName
if (-not $user) { throw 'No signed-in desktop user' }
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$arguments = '-NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File "' + $PSCommandPath + '" -Run ' + $Run + ' -Interactive'
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$action = New-ScheduledTaskAction -Execute $ps -Argument $arguments -WorkingDirectory $root
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero)
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings | Out-Null
Start-ScheduledTask -TaskName $taskName
[pscustomobject]@{ started = $true; run = $Run; editorVersion = $manifest.version; upstreamEntrySha256 = $sourceHash } | ConvertTo-Json -Compress
