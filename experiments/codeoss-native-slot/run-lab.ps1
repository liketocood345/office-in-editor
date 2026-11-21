param([Parameter(Mandatory)][string]$LabRoot, [Parameter(Mandatory)][string]$Run)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$lab = [IO.Path]::GetFullPath($LabRoot)
$parent = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab')).TrimEnd('\')
if ((Split-Path -Parent $lab) -ine $parent -or (Split-Path -Leaf $lab) -ne ('codeoss-' + $Run.Substring(0,10))) {
    throw 'Unexpected lab path'
}
$marker = Get-Content -LiteralPath (Join-Path $lab 'ownership.json') -Raw | ConvertFrom-Json
if ($marker.run -ne $Run -or $marker.experiment -ne 'codeoss-native-slot') { throw 'Lab ownership mismatch' }
$codium = Join-Path $lab 'codium'
$exe = Join-Path $codium 'VSCodium.exe'
$outDir = Join-Path $codium 'resources\app\out'
$main = Join-Path $outDir 'main.js'
$mainOriginal = Join-Path $lab 'upstream-main.js'
$experimentAssets = Join-Path $outDir 'oie-lab'
$profile = Join-Path $lab 'profile'
$shared = Join-Path $lab 'shared-data'
$extensionDir = Join-Path $lab 'extensions'
$result = Join-Path $lab ('result-' + [guid]::NewGuid().ToString('N') + '.json')
$log = Join-Path $lab 'codium-stderr.txt'
$version = Get-Content -LiteralPath (Join-Path $codium 'resources\app\package.json') -Raw | ConvertFrom-Json
if ($version.version -ne '1.143.17232' -or $version.main -ne './out/main.js' -or $version.type -ne 'module') {
    throw 'Unexpected Code-OSS build'
}
if (Test-Path -LiteralPath $mainOriginal) { throw 'Previously patched lab copy; restore it before retrying' }
New-Item -ItemType Directory -Path $experimentAssets,$extensionDir,$shared,(Join-Path $profile 'User') -Force | Out-Null
foreach ($name in @('host-probe.mjs', 'oie-codeoss-slot.node')) {
    Copy-Item -LiteralPath (Join-Path $lab $name) -Destination $experimentAssets -Force
}
$utf8 = [Text.UTF8Encoding]::new($false)
[IO.File]::WriteAllText((Join-Path $profile 'User\settings.json'), (@{
    'telemetry.telemetryLevel' = 'off'; 'update.mode' = 'none'; 'window.restoreWindows' = 'none';
    'window.zoomLevel' = 0; 'workbench.startupEditor' = 'none'; 'extensions.autoUpdate' = $false
} | ConvertTo-Json), $utf8)
$fixture = Join-Path $lab 'synthetic-document.txt'
[IO.File]::WriteAllText($fixture, 'Synthetic editor fixture for the native slot experiment.', $utf8)
Copy-Item -LiteralPath $main -Destination $mainOriginal
$process = $null
try {
    $prefix = "import './oie-lab/host-probe.mjs';`n"
    [IO.File]::WriteAllText($main, $prefix + [IO.File]::ReadAllText($mainOriginal), $utf8)
    $env:OIE_CODEOSS_LAB_RESULT = $result
    $labExtension = Join-Path $lab 'lab-extension'
    $arguments = @('--new-window', '--skip-welcome', '--skip-release-notes',
        ('--user-data-dir="' + $profile + '"'), ('--shared-data-dir="' + $shared + '"'),
        ('--extensions-dir="' + $extensionDir + '"'), ('--extensionDevelopmentPath="' + $labExtension + '"'),
        ('"' + $fixture + '"'))
    $process = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $codium -RedirectStandardError $log -PassThru
    for ($i = 0; $i -lt 90; $i++) {
        if (Test-Path -LiteralPath $result) {
            Get-Content -LiteralPath $result -Raw
            return
        }
        if ($process.HasExited) { break }
        Start-Sleep -Milliseconds 500
    }
    if (Test-Path -LiteralPath $log) {
        Get-Content -LiteralPath $log -Tail 12 | Select-String -Pattern 'Error|failed|\[main\]' | ForEach-Object { $_.Line }
    }
    throw 'Code-OSS produced no result within the test deadline'
} finally {
    Remove-Item Env:OIE_CODEOSS_LAB_RESULT -ErrorAction SilentlyContinue
    if ($process -and -not $process.HasExited) {
        [void]$process.WaitForExit(1500)
        if (-not $process.HasExited -and $process.Path -ieq $exe) {
            [void]$process.CloseMainWindow()
            if (-not $process.WaitForExit(1500) -and $process.Path -ieq $exe) { Stop-Process -Id $process.Id -Force }
        }
    }
    Copy-Item -LiteralPath $mainOriginal -Destination $main -Force
    if ((Get-FileHash -LiteralPath $main).Hash -eq (Get-FileHash -LiteralPath $mainOriginal).Hash) {
        Remove-Item -LiteralPath $mainOriginal
    } else { throw 'The lab entry point was not restored' }
}
