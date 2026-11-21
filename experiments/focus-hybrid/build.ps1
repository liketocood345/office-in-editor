$ErrorActionPreference='Stop'
$repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$out=Join-Path $repository '.local/build/focus-hybrid'
New-Item -ItemType Directory -Path $out -Force | Out-Null
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$vs=& $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if(-not $vs){throw 'MSVC x64 tools unavailable'}
$vcvars=Join-Path $vs 'VC\Auxiliary\Build\vcvars64.bat'
foreach($path in @($PSScriptRoot,$out,$vcvars)) { if($path -match '["%&|<>^!\r\n]'){throw 'Unsupported build path'} }
$batch=@"
@echo off
call "$vcvars" >nul
if errorlevel 1 exit /b 1
cl /nologo /std:c++17 /EHsc /MT /O2 /W3 /utf-8 "$PSScriptRoot\native.cpp" /Fo:"$out\native.obj" /Fe:"$out\native.exe" /link d3d11.lib dxgi.lib dwmapi.lib windowsapp.lib runtimeobject.lib ole32.lib oleaut32.lib windowscodecs.lib crypt32.lib user32.lib advapi32.lib
"@
$batchFile=Join-Path $out 'build.cmd'
[IO.File]::WriteAllText($batchFile,$batch,[Text.Encoding]::ASCII)
& $env:ComSpec /d /c $batchFile
if($LASTEXITCODE){throw 'Focus preview native build failed'}
