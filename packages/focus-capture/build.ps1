param([string]$OutDir = '')
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$out = if ($OutDir) { [IO.Path]::GetFullPath($OutDir) } else { Join-Path $repo 'payload/core' }
New-Item -ItemType Directory -Path $out -Force | Out-Null
$installer = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer'
$vs = & (Join-Path $installer 'vswhere.exe') -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $vs) { throw 'MSVC x64 is required' }
$vcvars = Join-Path $vs 'VC\Auxiliary\Build\vcvars64.bat'
$source = Join-Path $PSScriptRoot 'capture.cpp'
$temp = Join-Path $repo '.local/build/focus-capture'
foreach ($p in @($out, $vcvars, $source, $temp, $installer)) { if ($p -match '["%&|<>^!\r\n]') { throw 'Unsupported build path' } }
New-Item -ItemType Directory -Path $temp -Force | Out-Null
# vcvars64.bat calls vswhere by name; keep the installer folder on PATH for it.
$batch = @"
@echo off
set "PATH=$installer;%PATH%"
call "$vcvars" >nul
if errorlevel 1 exit /b 1
cl /nologo /std:c++17 /EHsc /MT /O2 /W3 /utf-8 /DUNICODE /D_UNICODE "$source" /Fo:"$temp\capture.obj" /Fe:"$out\oie-focus-capture.exe" /link d3d11.lib dxgi.lib dwmapi.lib windowsapp.lib runtimeobject.lib ole32.lib windowscodecs.lib user32.lib
if errorlevel 1 exit /b 1
"@
$file = Join-Path $temp 'build.cmd'
[IO.File]::WriteAllText($file, $batch, [Text.Encoding]::ASCII)
Push-Location $temp
try { & $env:ComSpec /d /c $file; if ($LASTEXITCODE) { throw 'Focus capture build failed' } } finally { Pop-Location }
