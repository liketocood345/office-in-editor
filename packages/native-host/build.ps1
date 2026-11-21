param()
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$out=Join-Path $repo 'payload/host'
New-Item -ItemType Directory -Path $out -Force|Out-Null
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vs=& $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if(-not $vs){throw 'MSVC x64 is required'}
$vcvars=Join-Path $vs 'VC/Auxiliary/Build/vcvars64.bat'
$source=Join-Path $PSScriptRoot 'slot-addon.cpp'
$bridge=Join-Path $repo 'experiments/native-surface-bridge/bridge.cpp'
foreach($p in @($out,$vcvars,$source,$bridge)){if($p -match '["%&|<>^!\r\n]'){throw 'Unsupported build path'}}
$batch=@"
@echo off
call "$vcvars" >nul
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /EHsc /W4 /WX /O2 /MT /DUNICODE /D_UNICODE /LD "$source" "$bridge" /Fe:"$out\oie-native-slot.node" /link /IMPLIB:oie-native-slot.lib user32.lib
if errorlevel 1 exit /b 1
"@
$temp=Join-Path $repo '.local/build/native-host'
New-Item -ItemType Directory -Path $temp -Force|Out-Null
$file=Join-Path $temp 'build.cmd'
[IO.File]::WriteAllText($file,$batch,[Text.Encoding]::ASCII)
Push-Location $temp
try {& $env:ComSpec /d /c $file;if($LASTEXITCODE){throw 'Native slot build failed'}} finally {Pop-Location}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'provider.mjs') -Destination $out
