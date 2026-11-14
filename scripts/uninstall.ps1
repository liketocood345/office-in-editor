[CmdletBinding(SupportsShouldProcess=$true)]
param([string]$InstallRoot='',[switch]$RemoveEditorAdapters,[switch]$RemoveOfficeRegistration,[switch]$RemoveOwnedBinaries)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'installation-common.ps1')
$productBase=[IO.Path]::GetFullPath((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\OfficeInEditor')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if(-not $InstallRoot){$InstallRoot=Join-Path $productBase 'current'}
$root=[IO.Path]::GetFullPath($InstallRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
if(-not $root.StartsWith($productBase+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Invalid owned product root.'}
Assert-OieSafePath $root;Assert-OieSafePath $PSScriptRoot
$manifestPath=Join-Path $root 'install-manifest.json'
if(-not(Test-Path -LiteralPath $manifestPath)){throw 'No ownership manifest; nothing was deleted.'}
$record=Get-Content -LiteralPath $manifestPath -Raw|ConvertFrom-Json
if($record.productId -ne 'office-in-editor'){throw 'Unknown installation ownership.'}
if($RemoveOwnedBinaries -and (-not $RemoveEditorAdapters -or -not $RemoveOfficeRegistration)){throw 'Explicitly remove all component registrations before shared binaries.'}
foreach($relative in @('core/OfficeInEditor.Core.exe','maintenance/uninstall-office.ps1')){
  $entry=@($record.files|Where-Object{$_.relativePath -eq $relative})|Select-Object -First 1
  $file=Resolve-OieOwnedChild $root $relative
  if(-not $entry -or -not(Test-Path -LiteralPath $file) -or (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.sha256){throw 'A maintenance component cannot be verified; it was not executed.'}
}
if(-not $PSCmdlet.ShouldProcess('Owned Office in Editor components','Remove selected plugin components')){return}
$remaining=New-Object 'System.Collections.Generic.List[object]'
foreach($hostRecord in $record.hosts){
  $cli=Get-OieCli $hostRecord.kind
  $present=if($cli){Test-OieExtensionPresent $cli $hostRecord.extensionId}else{$true}
  if($present -and $RemoveEditorAdapters -and $hostRecord.installedByThisProduct -and $cli){
    & $cli.Source --uninstall-extension $hostRecord.extensionId
    if($LASTEXITCODE -ne 0){throw 'Editor adapter removal failed; shared binaries were preserved.'}
    $present=Test-OieExtensionPresent $cli $hostRecord.extensionId
  }
  if($present){$remaining.Add($hostRecord)}
}
$record.hosts=@($remaining.ToArray());Save-OieManifest $manifestPath $record
if($RemoveOfficeRegistration -and $record.officeRegistered){
  & (Join-Path $root 'maintenance\uninstall-office.ps1') -InstallRoot $root
  $record.officeRegistered=$false;Save-OieManifest $manifestPath $record
}
if($RemoveOwnedBinaries){
  if($record.officeRegistered -or @($record.hosts).Count){throw 'An installed component still owns SharedCore; binaries were preserved.'}
  & (Join-Path $root 'core\OfficeInEditor.Core.exe') --shutdown-if-idle
  if($LASTEXITCODE -ne 0){throw 'A live host still owns SharedCore; binaries were preserved.'}
  foreach($entry in @($record.files)){
    $file=Resolve-OieOwnedChild $root $entry.relativePath
    if(-not(Test-Path -LiteralPath $file)){continue}
    if((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.sha256){throw 'Changed product file preserved; cleanup stopped.'}
    try{Remove-Item -LiteralPath $file -Force}catch{Write-Warning 'In-use owned file preserved; no host was closed.'}
  }
}
Write-Output 'Documents, references, recovery records and unrelated extensions were preserved; no background cleanup was scheduled.'
