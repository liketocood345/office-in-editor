param([Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$base = Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab'
$root = Join-Path $base ('office-browser-' + $Run)
$task = 'OfficeInEditor-Browser-' + $Run
$utf8 = [Text.UTF8Encoding]::new($false)
if ($PSScriptRoot -ine $root) { throw 'Unexpected isolated profile directory' }
foreach ($directory in @($base, $root)) {
    if ((Get-Item -LiteralPath $directory -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Redirected lab directory' }
}
if ((Get-Process -Id $PID).SessionId -eq 0) { throw 'Interactive desktop unavailable' }

function Write-Status($Name, $Data) {
    [IO.File]::WriteAllText((Join-Path $root ($Name + '.partial')), ($Data | ConvertTo-Json -Depth 5 -Compress), $utf8)
    [IO.File]::Move((Join-Path $root ($Name + '.partial')), (Join-Path $root ($Name + '.json')))
}

$stage = 'preflight'
try {
    $classes = @{}
    $programs = @{ word = 'Word.Document.12'; excel = 'Excel.Sheet.12'; powerpoint = 'PowerPoint.Show.12' }
    foreach ($office in $programs.Keys) {
        $key = Get-Item -LiteralPath ('Registry::HKEY_CLASSES_ROOT\' + $programs[$office] + '\CLSID') -ErrorAction SilentlyContinue
        $parsed = [Guid]::Empty
        if ($key -and [Guid]::TryParse($key.GetValue(''), [ref]$parsed)) { $classes[$office] = $parsed.ToString('D') }
    }
    Write-Status 'inputs' @{ run = $Run; classes = $classes }
    $editor = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\Code.exe'
    if (-not (Test-Path -LiteralPath $editor -PathType Leaf)) { throw 'Stock VS Code missing' }
    $profile = Join-Path $root 'profile'
    $settingsPath = Join-Path $profile 'User'
    $extensions = Join-Path $root 'extensions'
    $shared = Join-Path $root 'shared'
    New-Item -ItemType Directory -Path $settingsPath, $extensions, $shared | Out-Null
    $settings = @{
        'telemetry.telemetryLevel' = 'off'
        'update.mode' = 'none'
        'extensions.autoUpdate' = $false
        'workbench.startupEditor' = 'none'
        'window.restoreWindows' = 'none'
    }
    [IO.File]::WriteAllText((Join-Path $settingsPath 'settings.json'), ($settings | ConvertTo-Json), $utf8)
    $env:OIE_BROWSER_VSCODE_LAB = $root
    $arguments = @('--new-window', '--skip-welcome', '--skip-release-notes',
        ('--user-data-dir="' + $profile + '"'), ('--shared-data-dir="' + $shared + '"'),
        ('--extensions-dir="' + $extensions + '"'), ('--extensionDevelopmentPath="' + (Join-Path $root 'probe') + '"'))
    $stage = 'start_editor'
    $process = Start-Process -FilePath $editor -ArgumentList $arguments -WorkingDirectory $root -PassThru
    Write-Status 'launch' @{ run = $Run; stage = 'editor_started'; vscodeVersion = (Get-Item -LiteralPath $editor).VersionInfo.ProductVersion }
    $stage = 'wait_for_editor'
    $exited = $process.WaitForExit(55000)
    $resultReady = Test-Path -LiteralPath (Join-Path $root 'result.json') -PathType Leaf
    Write-Status 'finished' @{ run = $Run; editorExited = $exited; resultReady = $resultReady }
} catch {
    Write-Status 'failure' @{ run = $Run; stage = $stage; type = $_.Exception.GetType().Name }
    throw
} finally {
    Unregister-ScheduledTask -TaskName $task -Confirm:$false -ErrorAction SilentlyContinue
}
