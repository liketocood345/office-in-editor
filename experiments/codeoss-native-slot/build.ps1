$ErrorActionPreference = 'Stop'
$target = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\.local\build\codeoss-native-slot'))
New-Item -ItemType Directory -Path $target -Force | Out-Null
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$installation = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $installation) { throw 'MSVC x64 build tools are required' }
$vcvars = Join-Path $installation 'VC\Auxiliary\Build\vcvars64.bat'
$probe = Join-Path $PSScriptRoot 'slot_probe.cpp'
$bridge = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\native-surface-bridge\bridge.cpp'))
foreach ($item in @($target, $vcvars, $probe, $bridge)) {
    if ($item -match '["%&|<>^!\r\n]') { throw 'Unsupported build path characters' }
}
$batch = @"
@echo off
call "$vcvars" >nul
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /EHsc /W4 /WX /O2 /MT /DUNICODE /D_UNICODE /LD "$probe" "$bridge" /Fe:"$target\oie-codeoss-slot.node" /link user32.lib gdi32.lib dwmapi.lib
if errorlevel 1 exit /b 1
"@
$cmdFile = Join-Path $target 'build.cmd'
[IO.File]::WriteAllText($cmdFile, $batch, [Text.Encoding]::ASCII)
Push-Location $target
try {
    & $env:ComSpec /d /c $cmdFile
    if ($LASTEXITCODE -ne 0) { throw "Bridge build failed ($LASTEXITCODE)" }
} finally { Pop-Location }
