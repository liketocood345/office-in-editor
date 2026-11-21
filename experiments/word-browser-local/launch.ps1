param(
    [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run,
    [ValidatePattern('^[a-f0-9]{32}$')][string]$SecondRun,
    [ValidateSet('Edge','VSCode')][string]$Client = 'Edge',
    [switch]$VerifyExit
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$base = Join-Path $env:LOCALAPPDATA 'OfficeInEditorLab'
$root = Join-Path $base ('word-web-' + $Run)
$secondRoot = if ($SecondRun) { Join-Path $base ('word-web-' + $SecondRun) }
if ($PSScriptRoot -ine $root -or (Get-Process -Id $PID).SessionId -eq 0) { throw 'Invalid interactive experiment context' }
foreach ($p in @($base, $root, $secondRoot) | Where-Object { $_ }) {
    if ((Get-Item -LiteralPath $p -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Unexpected experiment path' }
}
$utf8 = [Text.UTF8Encoding]::new($false)
$resultPath = Join-Path $root 'verification.json'
$readyPath = Join-Path $root 'web-ready.json'
$proofPath = Join-Path $root 'browser-proof.json'
$report = [ordered]@{ passed = $false; stage = 'startup'; run = $Run }
function Wait-Artifact([string]$Path, [int]$TimeoutMs) {
    $watcher = [IO.FileSystemWatcher]::new((Split-Path -Parent $Path), [IO.Path]::GetFileName($Path))
    try {
        $watcher.EnableRaisingEvents = $true
        if (-not (Test-Path -LiteralPath $Path)) {
            $event = $watcher.WaitForChanged(([IO.WatcherChangeTypes]::Renamed -bor [IO.WatcherChangeTypes]::Created), $TimeoutMs)
            if ($event.TimedOut) { throw 'Experiment artifact deadline reached' }
        }
        for ($i = 0; $i -lt 5; $i++) {
            try { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
            catch { if ($i -eq 4) { throw }; Start-Sleep -Milliseconds 100 }
        }
    } finally { $watcher.Dispose() }
}

function Test-TwoEditorWindows {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class OieEditorWindows {
    private delegate bool Visit(IntPtr window, IntPtr unused);
    [DllImport("user32.dll")] private static extern bool EnumWindows(Visit visit, IntPtr data);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int length);
    public static int Count(uint process) {
        int count = 0;
        EnumWindows((window, unused) => {
            uint pid;
            GetWindowThreadProcessId(window, out pid);
            if (pid != process || !IsWindowVisible(window)) return true;
            var name = new StringBuilder(128);
            GetClassName(window, name, name.Capacity);
            if (name.ToString() == "Chrome_WidgetWin_1") count++;
            return true;
        }, IntPtr.Zero);
        return count;
    }
}
'@
    function Wait-WindowCount([int]$Expected) {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        do {
            if ($editor.HasExited) { return $false }
            if ([OieEditorWindows]::Count($editor.Id) -eq $Expected) { return $true }
            Start-Sleep -Milliseconds 150
        } while ($timer.ElapsedMilliseconds -lt 6000)
        return $false
    }
    if (-not (Wait-WindowCount 1)) { return $false }
    $otherWindowArgs = @('--new-window','--skip-welcome','--skip-release-notes',
        ('--user-data-dir="' + $profile + '"'),('--extensions-dir="' + $extensions + '"'),
        ('--shared-data-dir="' + $shared + '"'))
    Start-Process -FilePath $code -ArgumentList $otherWindowArgs -WorkingDirectory $root -WindowStyle Normal
    if (-not (Wait-WindowCount 2)) { return $false }
    $editor.Refresh()
    if (-not $editor.CloseMainWindow() -or -not (Wait-WindowCount 1) -or $server.HasExited) { return $false }
    foreach ($endpoint in @($ready,$second)) {
        if ($endpoint.url -notmatch '^(http://localhost:\d{1,5})/#key=([a-f0-9]{64})$') { return $false }
        $result = Invoke-RestMethod -Uri ($Matches[1]+'/api/document') -Headers @{Authorization=('Bearer '+$Matches[2])} -TimeoutSec 5
        if (-not $result.ok) { return $false }
    }
    return $true
}

try {
    $node = (Get-Command node.exe -ErrorAction Stop).Source
    $previousWord = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object Id, StartTime)
    $serverArgs = @('"' + (Join-Path $root 'server.mjs') + '"', '"' + $root + '"')
    if ($secondRoot) { $serverArgs += '"' + $secondRoot + '"' }
    if ($Client -eq 'VSCode') {
        if (-not $secondRoot) { throw 'VS Code verification requires two fixtures' }
        $report.stage = 'prepare-fixtures'
        foreach ($fixtureRoot in @($root, $secondRoot)) {
            $prepare = Start-Process -FilePath (Join-Path $root 'WordWorker.exe') -ArgumentList @('"' + $fixtureRoot + '"','--prepare') -WindowStyle Hidden -PassThru
            [void]$prepare.Handle
            if (-not $prepare.WaitForExit(8000)) { throw 'Fixture preparation did not finish' }
            if ($prepare.ExitCode -ne 0) { throw 'Fixture preparation failed' }
        }
        $report.stage = 'wait-first-office-open'
        $code = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\Code.exe'
        if (-not (Test-Path -LiteralPath $code -PathType Leaf)) { throw 'Stock VS Code missing' }
        $profile = Join-Path $root 'vscode-profile'
        $extensions = Join-Path $root 'vscode-extensions'
        $shared = Join-Path $root 'vscode-shared'
        New-Item -ItemType Directory -Path (Join-Path $profile 'User'),$extensions,$shared -Force | Out-Null
        $config = @{ 'telemetry.telemetryLevel'='off'; 'update.mode'='none'; 'workbench.startupEditor'='none'; 'window.restoreWindows'='none' }
        [IO.File]::WriteAllText((Join-Path $profile 'User\settings.json'), ($config|ConvertTo-Json), $utf8)
        $env:OIE_VSCODE_WEB_LAB_ROOTS = ConvertTo-Json -InputObject @($root,$secondRoot) -Compress
        $codeArgs = @('--new-window','--skip-welcome','--skip-release-notes',
            ('--user-data-dir="' + $profile + '"'),('--extensions-dir="' + $extensions + '"'),
            ('--shared-data-dir="' + $shared + '"'),('--extensionDevelopmentPath="' + (Join-Path $root 'vscode-lab') + '"'))
        $editor = Start-Process -FilePath $code -ArgumentList $codeArgs -WorkingDirectory $root -WindowStyle Normal -PassThru
        [void]$editor.Handle
        [IO.File]::WriteAllText((Join-Path $root 'vscode-launch.json'), (@{pid=$editor.Id; run=$Run}|ConvertTo-Json -Compress), $utf8)
        $start = Wait-Artifact (Join-Path $root 'vscode-start.json') 32000
        if ($start.run -ne $Run -or $start.reason -ne 'first_office_open' -or $editor.HasExited) { throw 'Editor startup claim invalid' }
        $info = [Diagnostics.ProcessStartInfo]::new($node, ($serverArgs -join ' '))
        $info.WorkingDirectory = $root
        $info.UseShellExecute = $false
        $info.CreateNoWindow = $true
        $info.RedirectStandardInput = $true
        $info.EnvironmentVariables['OIE_WORD_WEB_SUPERVISED'] = '1'
        $server = [Diagnostics.Process]::Start($info)
        $server.StandardInput.AutoFlush = $true
    } else {
        $server = Start-Process -FilePath $node -ArgumentList $serverArgs -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $root 'server-error.txt')
    }
    $report.stage = 'word-ready'
    $ready = Wait-Artifact $readyPath 28000
    if ($ready.run -ne $Run -or $ready.serverPid -ne $server.Id -or $ready.url -notmatch '^http://localhost:\d{1,5}/#key=[a-f0-9]{64}$') { throw 'Unexpected loopback endpoint' }
    $report.stage = 'port-reservations'
    $ports = Wait-Artifact (Join-Path $root 'port-verification.json') 28000
    if (-not $ports.sameFileReused -or -not $ports.distinctPorts -or $ports.activeFiles -ne $(if ($secondRoot) { 2 } else { 1 })) {
        throw 'Document port reservations did not match'
    }
    $report.ports = @($ports.ports)
    $report.sameFileReusedPort = $ports.sameFileReused
    $report.distinctPorts = $ports.distinctPorts
    if ($Client -eq 'Edge') {
        $edgeCandidates = @(
            (Join-Path ([Environment]::GetEnvironmentVariable('ProgramFiles(x86)')) 'Microsoft\Edge\Application\msedge.exe'),
            (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
            (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe')
        )
        $edge = $edgeCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
        if (-not $edge) { throw 'Edge not found' }
        $profile = Join-Path $root 'edge-profile'
        $edgeArgs = @('--no-first-run', '--new-window', ('--user-data-dir="' + $profile + '"'), ('"' + $ready.url + '&verify=1"'))
    }
    if ($secondRoot) {
        $second = Wait-Artifact (Join-Path $secondRoot 'web-ready.json') 3000
        if ($second.run -ne $SecondRun -or $second.serverPid -ne $server.Id -or
            $second.url -notmatch '^http://localhost:\d{1,5}/#key=[a-f0-9]{64}$' -or
            ([Uri]$ready.url).Port -eq ([Uri]$second.url).Port) { throw 'Unexpected second endpoint' }
        if ($Client -eq 'Edge') { $edgeArgs += '"' + $second.url + '"' }
    }
    if ($Client -eq 'Edge') {
        $report.stage = 'edge-normal-tab'
        $browser = Start-Process -FilePath $edge -ArgumentList $edgeArgs -WorkingDirectory $root -PassThru
    } else { $report.stage = 'vscode-webview' }
    $proof = Wait-Artifact $proofPath 35000
    if (($Client -eq 'Edge' -and ($proof.browser -ne 'Edge' -or -not $proof.normalTab)) -or
        ($Client -eq 'VSCode' -and ($proof.browser -ne 'VS Code' -or -not $proof.webview)) -or
        -not $proof.editedViaWebForm -or -not $proof.wordSaved -or -not $proof.wordWindowHidden -or
        $proof.revision -ne 1 -or $proof.textSha256 -notmatch '^[a-f0-9]{64}$') { throw 'Incomplete browser proof' }
    if ($Client -eq 'VSCode') {
        $view = Wait-Artifact (Join-Path $root 'vscode-verification.json') 9000
        if (-not $view.passed) { throw 'Custom editor verification failed' }
        $report.twoRealTabs = $view.twoRealTabs
        $report.lazyStart = $view.lazyStart
        $report.reloadedAtSamePort = $view.initialRead.port -eq $view.reloadRead.port
        $report.otherFileUntouched = $view.secondaryRead.revision -eq 0
    }

    $report.stage = 'verify-local-docx'
    Add-Type -AssemblyName System.IO.Compression
    $file = [IO.FileStream]::new((Join-Path $root 'synthetic.docx'), [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $zip = [IO.Compression.ZipArchive]::new($file, [IO.Compression.ZipArchiveMode]::Read, $true)
        try {
            $part = $zip.GetEntry('word/document.xml')
            if (-not $part -or $part.Length -gt 1048576) { throw 'Unexpected document XML' }
            $reader = [IO.StreamReader]::new($part.Open(), $utf8)
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $namespace = [Xml.XmlNamespaceManager]::new($xml.NameTable)
            $namespace.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            $text = @($xml.SelectNodes('//w:t', $namespace) | ForEach-Object { $_.InnerText }) -join ''
            $hasher = [Security.Cryptography.SHA256]::Create()
            try { $hash = [BitConverter]::ToString($hasher.ComputeHash($utf8.GetBytes($text))).Replace('-', '') }
            finally { $hasher.Dispose() }
            if ($hash -ine $proof.textSha256) { throw 'Saved file text did not match the browser edit' }
        } finally { $zip.Dispose() }
    } finally { $file.Dispose() }
    $report.previousWordPreserved = $true
    foreach ($existing in $previousWord) {
        $process = Get-Process -Id $existing.Id -ErrorAction SilentlyContinue
        if (-not $process -or $process.StartTime -ne $existing.StartTime) { $report.previousWordPreserved = $false }
    }
    $report.savedDocxVerified = $true
    $report.browserNormalTab = $Client -eq 'Edge'
    $report.vscodeWebview = $Client -eq 'VSCode'
    $report.wordComSave = $true
    $report.wordWindowHidden = $true
    $report.stage = 'complete'
    $report.passed = $true
} catch {
    $report.error = $_.Exception.GetType().Name
    $report.hresult = '0x{0:X8}' -f $_.Exception.HResult
}
$pending = $resultPath + '.pending'
[IO.File]::WriteAllText($pending, ($report | ConvertTo-Json -Compress), $utf8)
[IO.File]::Move($pending, $resultPath)
if ($Client -eq 'VSCode' -and $editor) {
    $survivedFirstWindow = $null
    if ($VerifyExit -and $report.passed -and -not $editor.HasExited) {
        try { $survivedFirstWindow = Test-TwoEditorWindows }
        catch { $survivedFirstWindow = $false }
        [IO.File]::WriteAllText((Join-Path $root 'window-lifetime.json'), (@{backendSurvivedFirstWindow=$survivedFirstWindow} | ConvertTo-Json -Compress), $utf8)
        if ($survivedFirstWindow) {
            $editor.Refresh()
            [void]$editor.CloseMainWindow()
        }
    }
    # The stock VS Code main process owns every window in this isolated profile.
    $editor.WaitForExit()
    if ($server -and -not $server.HasExited) {
        $server.StandardInput.WriteLine('shutdown')
        $stopped = $server.WaitForExit(20000)
        $shutdown = Join-Path $root 'backend-shutdown.json'
        $clean = (Test-Path -LiteralPath $shutdown) -and (Get-Content -LiteralPath $shutdown -Raw | ConvertFrom-Json).closed
        [IO.File]::WriteAllText((Join-Path $root 'lifetime.json'), (@{
            editorExited=$true; backendStopped=$stopped; cleanWordClose=$clean; backendSurvivedFirstWindow=$survivedFirstWindow
        } | ConvertTo-Json -Compress), $utf8)
        if (-not $stopped) { $server.WaitForExit() } # Retain an unconfirmed Word session.
    }
} elseif ($server) { $server.WaitForExit() } # Keep the interactive process tree alive for manual editing.
