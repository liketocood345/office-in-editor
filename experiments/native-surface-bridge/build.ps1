param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $PSScriptRoot '..\..\.local\build\native-surface-bridge'
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Visual C++ Build Tools are required.' }
$install = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $install) { throw 'Visual C++ compiler not found.' }
$vcvars = Join-Path $install 'VC\Auxiliary\Build\vcvars64.bat'
$source = $PSScriptRoot
# Values are paths, reject command metacharacters before constructing cmd arguments.
foreach ($value in @($source, $OutputDirectory, $vcvars)) {
    if ($value -match '["%&|<>^!\r\n]') { throw 'Unsupported build path characters.' }
}
$batch = @"
@echo off
call "$vcvars" >nul
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /EHsc /W4 /WX /DUNICODE /D_UNICODE /LD "$source\bridge.cpp" /Fe:"$OutputDirectory\oie-native-surface.dll" /Fo:"$OutputDirectory\bridge.obj" /link user32.lib /IMPLIB:"$OutputDirectory\oie-native-surface.lib"
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /EHsc /W4 /WX /DUNICODE /D_UNICODE "$source\smoke.cpp" "$OutputDirectory\oie-native-surface.lib" /Fe:"$OutputDirectory\native-surface-smoke.exe" /Fo:"$OutputDirectory\smoke.obj" /link user32.lib
if errorlevel 1 exit /b 1
"@
$batchPath = Join-Path $OutputDirectory 'build.cmd'
[IO.File]::WriteAllText($batchPath, $batch, [Text.Encoding]::ASCII)
& $env:ComSpec /d /c $batchPath
if ($LASTEXITCODE -ne 0) { throw "Native bridge build failed: $LASTEXITCODE" }
Write-Output "Built native bridge in $OutputDirectory"
