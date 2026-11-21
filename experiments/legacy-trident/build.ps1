param([switch]$BuildDsoSample, [switch]$UseCachedSource)
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$out=Join-Path $repo '.local/build/legacy-trident'
New-Item -ItemType Directory -Path $out -Force|Out-Null
$compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
if(-not(Test-Path -LiteralPath $compiler)){throw '.NET Framework compiler is required'}
[xml]$basic=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'probe.manifest') -Raw
$fileNode=$basic.SelectSingleNode("/*[local-name()='assembly']/*[local-name()='file']")
[void]$fileNode.ParentNode.RemoveChild($fileNode)
$basicManifest=Join-Path $out 'trident-only.manifest'
$basic.Save($basicManifest)
$probeSources = @('TridentProbe.cs','DsoDocumentProbe.cs','ManualWordSession.cs','LoopbackPage.cs','../office-probe-common/ProbeSafety.cs') | ForEach-Object { Join-Path $PSScriptRoot $_ }
& $compiler /nologo /target:exe /platform:x86 /optimize+ /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:Microsoft.CSharp.dll ('/win32manifest:'+$basicManifest) ('/out:'+(Join-Path $out 'TridentProbe.exe')) @probeSources
if($LASTEXITCODE){throw 'Probe build failed'}
& $compiler /nologo /target:exe /platform:x86 /optimize+ /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:Microsoft.CSharp.dll ('/win32manifest:'+(Join-Path $PSScriptRoot 'probe.manifest')) ('/out:'+(Join-Path $out 'TridentDsoProbe.exe')) @probeSources
if($LASTEXITCODE){throw 'DSO probe build failed'}
if(-not $BuildDsoSample){Write-Output 'Both probe executables built; the optional OCX must be built separately with -BuildDsoSample.';exit 0}
$commit='224163f5a5acb692118a5be95e5e49fe750bfd50'
$headers=@{'User-Agent'='office-in-editor-legacy-probe'}
$source=Join-Path $out 'source'
$cache=Join-Path $out 'pinned-source'
if(-not $UseCachedSource){
 $tree=Invoke-RestMethod -Uri ('https://api.github.com/repos/shines77/DsoFramer/git/trees/'+$commit+'?recursive=1') -Headers $headers
 $selected=@($tree.tree|Where-Object {$_.type -eq 'blob' -and $_.path -like 'Source/*' -and $_.path -match '\.(cpp|h|idl|def|rc|ico|bmp)$'})
 $hashes=@()
 foreach($file in $selected){
  $target=Join-Path $cache $file.path.Substring(7)
  New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force|Out-Null
  Invoke-WebRequest -Uri ('https://raw.githubusercontent.com/shines77/DsoFramer/'+$commit+'/'+$file.path) -OutFile $target -UseBasicParsing
  $hashes+=@{path=$file.path;gitBlob=$file.sha}
 }
 foreach($name in @('eula.txt','LICENSE','README.md')){Invoke-WebRequest -Uri ('https://raw.githubusercontent.com/shines77/DsoFramer/'+$commit+'/'+$name) -OutFile (Join-Path $out ('reference-'+$name)) -UseBasicParsing}
 $hashes|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $out 'reference-source-hashes.json') -Encoding UTF8
} else {
 if(-not(Test-Path -LiteralPath (Join-Path $cache 'dsofauto.cpp'))){throw 'Pristine source cache absent; build once without UseCachedSource'}
 $hashes=[IO.File]::ReadAllText((Join-Path $out 'reference-source-hashes.json'))|ConvertFrom-Json
}
# Always verify the immutable original and reconstruct the build copy. Patches
# never accumulate across builds, and modified cached inputs are rejected.
foreach($file in $hashes){
 $relative=$file.path.Substring(7)
 $cached=Join-Path $cache $relative
 $bytes=[IO.File]::ReadAllBytes($cached)
 $prefix=[Text.Encoding]::ASCII.GetBytes(('blob '+$bytes.Length+[char]0))
 $combined=New-Object byte[] ($prefix.Length+$bytes.Length)
 [Array]::Copy($prefix,$combined,$prefix.Length);[Array]::Copy($bytes,0,$combined,$prefix.Length,$bytes.Length)
 $sha=[Security.Cryptography.SHA1]::Create()
 try{$blob=([BitConverter]::ToString($sha.ComputeHash($combined))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
 if($blob -ne $file.gitBlob){throw ('Source digest mismatch: '+$file.path)}
 $target=Join-Path $source $relative
 New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force|Out-Null
 [IO.File]::WriteAllBytes($target,$bytes)
}
# Private identities keep this experiment separate from globally installed controls.
$version=Join-Path $source 'version.h'
$encoding=[Text.Encoding]::GetEncoding(1252)
$text=$encoding.GetString([IO.File]::ReadAllBytes($version))
$text=$text.Replace('00460180-9E5E-11d5-B7C8-B8269041DD57','70B8EE70-2B79-4D45-9A77-AE395CFEA071').Replace('00460181-9E5E-11d5-B7C8-B8269041DD57','70B8EE71-2B79-4D45-9A77-AE395CFEA071').Replace('00460182-9E5E-11d5-B7C8-B8269041DD57','70B8EE72-2B79-4D45-9A77-AE395CFEA071').Replace('00460185-9E5E-11d5-B7C8-B8269041DD57','70B8EE75-2B79-4D45-9A77-AE395CFEA071').Replace('DSOFramer.FramerControl','OieLegacyProbe.FramerControl')
[IO.File]::WriteAllBytes($version,$encoding.GetBytes($text))
$utility=Join-Path $source 'utilities.cpp'
$text=$encoding.GetString([IO.File]::ReadAllBytes($utility))
$needle='hr = LoadTypeLib(pwszPath, &ptlib);'
if(-not $text.Contains($needle) -and -not $text.Contains('LoadTypeLibEx(pwszPath, REGKIND_NONE, &ptlib)')){throw 'Unexpected typelib loader source'}
$text=$text.Replace($needle,'hr = LoadTypeLibEx(pwszPath, REGKIND_NONE, &ptlib);')
$text=$text.Replace('RegisterTypeLib(ptlib, pwszPath, NULL);','((void)0); /* No automatic typelib registration in this private probe. */')
[IO.File]::WriteAllBytes($utility,$encoding.GetBytes($text))
$entry=Join-Path $source 'mainentry.cpp'
$text=$encoding.GetString([IO.File]::ReadAllBytes($entry))
$registrationStart=$text.IndexOf('STDAPI DllRegisterServer()')
if($registrationStart -lt 0){throw 'Unexpected registration source'}
$text=$text.Substring(0,$registrationStart)+"STDAPI DllRegisterServer() { return SELFREG_E_CLASS; }`r`nSTDAPI DllUnregisterServer() { return SELFREG_E_CLASS; }`r`n"
[IO.File]::WriteAllBytes($entry,$encoding.GetBytes($text))
$definition=Join-Path $source 'dsoframer.def'
$text=$encoding.GetString([IO.File]::ReadAllBytes($definition)).Replace('LIBRARY    dsoframer.ocx','LIBRARY    oie-legacy-framer.ocx')
[IO.File]::WriteAllBytes($definition,$encoding.GetBytes($text))
$entryText=$encoding.GetString([IO.File]::ReadAllBytes($entry)).Replace('DeleteObject(v_icoOffDocIcon)','DestroyIcon(v_icoOffDocIcon)')
[IO.File]::WriteAllBytes($entry,$encoding.GetBytes($entryText))
# Add bounded in-memory tracing to the private build only.
foreach($name in @('dso-probe-trace.h','dso-probe-trace.cpp')){Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $source $name) -Force}
$definitionText=$encoding.GetString([IO.File]::ReadAllBytes($definition))
if(-not $definitionText.Contains('OieProbeTrace')){[IO.File]::WriteAllBytes($definition,$encoding.GetBytes($definitionText+"`r`nOieProbeTrace=_OieProbeTrace@12`r`nOieProbeDocumentWindow=_OieProbeDocumentWindow@0`r`n"))}
foreach($name in @('dsofauto.cpp','dsofdocobj.cpp')){
 $target=Join-Path $source $name
 $text=$encoding.GetString([IO.File]::ReadAllBytes($target))
 if(-not $text.Contains('#include "dso-probe-trace.h"')){
  $text=$text.Replace('#include "dsoframer.h"','#include "dsoframer.h"'+"`r`n"+'#include "dso-probe-trace.h"')
  if($name -eq 'dsofauto.cpp'){
   $text=$text.Replace('TRACE1("CDsoFramerControl::Open(%S)\n", pwszDocument);','OieTrace(10, S_OK);'+"`r`n"+'TRACE1("CDsoFramerControl::Open(%S)\n", pwszDocument);')
   $text=$text.Replace('TRACE1("CDsoFramerControl::CreateNew(%S)\n", ProgIdOrTemplate);','OieTrace(11, S_OK);'+"`r`n"+'TRACE1("CDsoFramerControl::CreateNew(%S)\n", ProgIdOrTemplate);')
   $text=$text.Replace('hr = m_pDocObjFrame->CreateFromFile(pwszDocument, clsidAlt, &bopts);','hr = m_pDocObjFrame->CreateFromFile(pwszDocument, clsidAlt, &bopts); OieTrace(12, hr);')
   $text=$text.Replace('hr = m_pDocObjFrame->IPActivateView();','hr = m_pDocObjFrame->IPActivateView(); OieTrace(13, hr);')
  } else {
   $text=$text.Replace('if (FAILED(GetClassFile(pwszFile, &clsid)) && !(fLoadFromAltCLSID))','hr = GetClassFile(pwszFile, &clsid); OieTrace(20, hr);'+"`r`n"+'if (FAILED(hr) && !(fLoadFromAltCLSID))')
   $text=$text.Replace('FAILED(ValidateDocObjectServer(clsid)))','FAILED(hr = ValidateDocObjectServer(clsid)))')
   $text=$text.Replace('hr = pmkfile->BindToObject(pbctx, NULL, IID_IOleObject, (void**)&pole);','hr = pmkfile->BindToObject(pbctx, NULL, IID_IOleObject, (void**)&pole); OieTrace(21, hr);')
   $text=$text.Replace('else pole->SetClientSite(NULL);','else if (pole) pole->SetClientSite(NULL);')
   $text=$text.Replace('hr = pipfile->Load(pwszFile, pbndopts->grfMode);','hr = pipfile->Load(pwszFile, pbndopts->grfMode); OieTrace(22, hr);')
  }
  [IO.File]::WriteAllBytes($target,$encoding.GetBytes($text))
 }
}
# Fix a sample automation bug: absence of IOleInPlaceActiveObject must not
# prevent asking the loaded IOleObject for its document automation interface.
$automation=Join-Path $source 'dsofauto.cpp'
$text=$encoding.GetString([IO.File]::ReadAllBytes($automation))
$text=$text.Replace('if ((m_pDocObjFrame) && (punk = (IUnknown*)(m_pDocObjFrame->GetActiveObject())))','if (m_pDocObjFrame)'+"`r`n"+'    { punk = (IUnknown*)m_pDocObjFrame->GetActiveObject();'+"`r`n"+'      if (!punk) punk = (IUnknown*)m_pDocObjFrame->GetOleObject(); }'+"`r`n"+'    else punk = NULL;'+"`r`n"+'    if (punk)')
[IO.File]::WriteAllBytes($automation,$encoding.GetBytes($text))
$documentSource=Join-Path $source 'dsofdocobj.cpp'
$text=$encoding.GetString([IO.File]::ReadAllBytes($documentSource))
if(-not $text.Contains('OieTrace(30,')){
 $text=$text.Replace('ODS("CDsoDocObject::XOleDocumentSite::ActivateMe\n");','ODS("CDsoDocObject::XOleDocumentSite::ActivateMe\n"); OieTrace(30, S_OK);')
 $text=$text.Replace('pmsodoc->Release();','OieTrace(31, hr); pmsodoc->Release();')
 $text=$text.Replace('if (SUCCEEDED(hr = pView->UIActivate(TRUE)))','hr = pView->UIActivate(TRUE); OieTrace(32, hr);'+"`r`n"+'    if (SUCCEEDED(hr))')
 $text=$text.Replace('ODS("CDsoDocObject::XOleInPlaceFrame::SetActiveObject\n");','ODS("CDsoDocObject::XOleInPlaceFrame::SetActiveObject\n"); OieTrace(33, pIIPActiveObj ? S_OK : S_FALSE);')
}
$text=$text.Replace('pIIPActiveObj->GetWindow(&(pThis->m_hwndUIActiveObj));','pIIPActiveObj->GetWindow(&(pThis->m_hwndUIActiveObj)); OieRememberWindow(pThis->m_hwndUIActiveObj);')
[IO.File]::WriteAllBytes($documentSource,$encoding.GetBytes($text))
# Propagate layout/show failures rather than reporting a UI-active blank view.
$text=$encoding.GetString([IO.File]::ReadAllBytes($documentSource))
$text=$text.Replace('pView->SetRect(&(pThis->m_rcViewRect));','hr = pView->SetRect(&(pThis->m_rcViewRect)); OieTrace(34, hr); if (FAILED(hr)) return hr;')
$text=$text.Replace('pView->Show(TRUE);','hr = pView->Show(TRUE); OieTrace(35, hr); if (FAILED(hr)) return hr;')
[IO.File]::WriteAllBytes($documentSource,$encoding.GetBytes($text))
# Ask the server to start in-place activation before constructing a view on
# its behalf. Modern Word can return an inert view from direct CreateView.
$text=$encoding.GetString([IO.File]::ReadAllBytes($documentSource))
if(-not $text.Contains('hr = m_xOleDocumentSite.ActivateMe(NULL);')){throw 'Missing activation patch anchor'}
$text=$text.Replace('hr = m_xOleDocumentSite.ActivateMe(NULL);','RECT activationRect; GetClientRect(m_hwnd, &activationRect);'+"`r`n"+'        hr = m_pole->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, (IOleClientSite*)&m_xOleClientSite, 0, m_hwnd, &activationRect); OieTrace(40, hr);'+"`r`n"+'        if (hr != S_OK || !m_pipactive) { hr = m_xOleDocumentSite.ActivateMe(NULL); OieTrace(41, hr); }')
[IO.File]::WriteAllBytes($documentSource,$encoding.GetBytes($text))
$resource=Join-Path $source 'Res/dsoframer.rc'
$text=$encoding.GetString([IO.File]::ReadAllBytes($resource)).Replace('#include "afxres.h"','#include <winres.h>')
$tlb=(Join-Path $source 'Release/dsoframer.tlb').Replace('\','\\')
$text=[regex]::Replace($text,'(?m)^IDR_TYPELIB\s+TYPELIB\s+DISCARDABLE\s+"[^"]+"',('IDR_TYPELIB TYPELIB "'+$tlb+'"'))
[IO.File]::WriteAllBytes($resource,$encoding.GetBytes($text))
New-Item -ItemType Directory -Path (Join-Path $source 'Release') -Force|Out-Null
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vs=& $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if(-not $vs){throw 'Visual C++ x86 Build Tools are required'}
$vcvars=Join-Path $vs 'VC/Auxiliary/Build/vcvars32.bat'
foreach($p in @($source,$out,$vcvars)){if($p -match '["%&|<>^!\r\n]'){throw 'Unsupported build path'}}
$batch=@"
@echo off
call "$vcvars" >nul
if errorlevel 1 exit /b 1
pushd "$source"
midl /nologo /mktyplib203 /env win32 /tlb Release\dsoframer.tlb /h Lib\dsoframerlib.h /iid Lib\dsoframer_i.c Lib\dsoframer.idl
if errorlevel 1 exit /b 1
rc /nologo /I Res /fo Release\dsoframer.res Res\dsoframer.rc
if errorlevel 1 exit /b 1
cl /nologo /MT /EHsc /O1 /W3 /source-charset:.1252 /DWIN32 /DNDEBUG /D_WINDOWS /D_CRT_SECURE_NO_WARNINGS /LD classfactory.cpp dsofauto.cpp dsofcontrol.cpp dsofdocobj.cpp dsofprint.cpp mainentry.cpp utilities.cpp dso-probe-trace.cpp Lib\dsoframer_i.c Release\dsoframer.res /Fe:"$out\oie-legacy-framer.ocx" /FoRelease\ /link /DEF:dsoframer.def /IMPLIB:Release\oie-legacy-framer.lib oledlg.lib urlmon.lib ole32.lib oleaut32.lib uuid.lib user32.lib gdi32.lib advapi32.lib comdlg32.lib shell32.lib winspool.lib
if errorlevel 1 exit /b 1
popd
"@
$batchFile=Join-Path $out 'build-dso.cmd'
[IO.File]::WriteAllText($batchFile,$batch,[Text.Encoding]::ASCII)
& $env:ComSpec /d /c $batchFile
if($LASTEXITCODE){throw 'Legacy sample build failed'}
Write-Output 'Source-built experimental OCX and x86 probe are ready; no COM registration was performed.'
