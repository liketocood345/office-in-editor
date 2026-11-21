param()
$ErrorActionPreference='Stop'
$out=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../.local/build/word-save-probe'))
New-Item -ItemType Directory -Path $out -Force | Out-Null
$compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
& $compiler /nologo /target:exe /platform:x64 /optimize+ /r:System.Windows.Forms.dll /r:System.Web.Extensions.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:Microsoft.CSharp.dll ('/out:'+(Join-Path $out 'WordSaveProbe.exe')) (Join-Path $PSScriptRoot 'WordSaveProbe.cs') (Join-Path $PSScriptRoot 'RecoverOwnedDocument.cs') (Join-Path $PSScriptRoot '../office-probe-common/ProbeSafety.cs')
if($LASTEXITCODE){throw 'Word probe build failed'}
