[CmdletBinding(SupportsShouldProcess=$true)]
param([string]$PayloadRoot=(Join-Path $PSScriptRoot '..\payload'),[string]$VsixPath='',
  [ValidateSet('vscode','cursor')][string[]]$Hosts=@('vscode','cursor'),[string]$InstallRoot='',
  [switch]$SkipOfficeRegistration,[switch]$SkipEditors,[switch]$ValidateOnly)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'installation-common.ps1')
$productBase=[IO.Path]::GetFullPath((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\OfficeInEditor')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if(-not $InstallRoot){$InstallRoot=Join-Path $productBase 'current'}
$root=[IO.Path]::GetFullPath($InstallRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
if(-not $root.StartsWith($productBase+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Invalid user-owned product root.'}
Assert-OieSafePath $root;Assert-OieSafePath $PSScriptRoot
$manifestPath=Join-Path $root 'install-manifest.json'
$record=$null
if(Test-Path -LiteralPath $manifestPath){$record=Get-Content -LiteralPath $manifestPath -Raw|ConvertFrom-Json;if($record.productId -ne 'office-in-editor'){throw 'Unknown product ownership.'}}
elseif((Test-Path -LiteralPath $root) -and @(Get-ChildItem -LiteralPath $root -Force).Count){throw 'Refusing to overwrite an unowned nonempty product directory.'}
if($ValidateOnly){
  if(-not $record){throw 'No owned installation found.'}
  foreach($entry in $record.files){$file=Resolve-OieOwnedChild $root $entry.relativePath;if(-not(Test-Path -LiteralPath $file) -or (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.sha256){throw 'Installed component verification failed.'}}
  if($record.officeRegistered){& (Join-Path $PSScriptRoot 'install-office.ps1') -PayloadRoot (Join-Path $root 'office') -InstallRoot $root -ValidateOnly}
  [ordered]@{status='verified';registeredOffice=[bool]$record.officeRegistered;hostCount=@($record.hosts).Count}|ConvertTo-Json;return
}
Assert-OieSafePath $PayloadRoot
$plan=New-Object 'System.Collections.Generic.List[object]'
foreach($component in @('core','office')){
  $source=[IO.Path]::GetFullPath((Join-Path $PayloadRoot $component)).TrimEnd([IO.Path]::DirectorySeparatorChar)
  if(-not(Test-Path -LiteralPath $source)){throw ('Missing payload: '+$component)}
  Assert-OieSafePath $source
  $entries=@(Get-ChildItem -LiteralPath $source -Recurse -File)
  if($entries.Count -gt 512 -or ($entries|Measure-Object Length -Sum).Sum -gt 536870912){throw 'Payload budget exceeded.'}
  foreach($entry in $entries){
    Assert-OieSafePath $entry.FullName
    $relative=$component+'/'+$entry.FullName.Substring($source.Length).TrimStart([IO.Path]::DirectorySeparatorChar).Replace([IO.Path]::DirectorySeparatorChar,'/')
    $null=Resolve-OieOwnedChild $root $relative
    $plan.Add([ordered]@{relativePath=$relative;sha256=(Get-FileHash -LiteralPath $entry.FullName -Algorithm SHA256).Hash.ToLowerInvariant();source=$entry.FullName})
  }
}
foreach($name in @('installation-common.ps1','install-office.ps1','uninstall-office.ps1','uninstall.ps1')){
  $source=Join-Path $PSScriptRoot $name;Assert-OieSafePath $source
  if(-not(Test-Path -LiteralPath $source)){throw 'Missing maintenance entry point.'}
  $plan.Add([ordered]@{relativePath='maintenance/'+$name;sha256=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant();source=$source})
}
if(-not $SkipEditors){
  if(-not $VsixPath){$candidate=Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '..\artifacts') -Filter '*.vsix' -File -ErrorAction SilentlyContinue|Select-Object -First 1;if($candidate){$VsixPath=$candidate.FullName}}
  if(-not $VsixPath -or -not(Test-Path -LiteralPath $VsixPath)){throw 'Build the VSIX before installing editor adapters.'}
  Assert-OieSafePath $VsixPath
}
# Existing files may only be replaced when the previous durable manifest owns their exact bytes.
foreach($entry in $plan){
  $file=Resolve-OieOwnedChild $root $entry.relativePath
  if(Test-Path -LiteralPath $file){
    $actual=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant()
    $known=if($record){@($record.files|Where-Object{$_.relativePath -eq $entry.relativePath})|Select-Object -First 1}else{$null}
    $pending=if($record -and $record.transactionState -in @('installing','pending-repair','partial')){@($record.pendingFiles|Where-Object{$_.relativePath -eq $entry.relativePath})|Select-Object -First 1}else{$null}
    $ownedByManifest=($known -and $actual -eq $known.sha256) -or ($pending -and $actual -eq $pending.sha256)
    if($actual -ne $entry.sha256 -and -not $ownedByManifest){throw 'An existing file is not the verified product payload; it was preserved.'}
  }
}
if(-not $PSCmdlet.ShouldProcess('Current-user Office in Editor components','Install or repair verified plugin components')){return}
New-Item -Path $root -ItemType Directory -Force|Out-Null
if(-not $record){$record=[pscustomobject]@{productId='office-in-editor';version='0.1.0';files=@();hosts=@();officeRegistered=$false}}
$record.version=[Reflection.AssemblyName]::GetAssemblyName((Join-Path $PayloadRoot 'core/OfficeInEditor.Core.dll')).Version.ToString(3)
$record|Add-Member -NotePropertyName transactionState -NotePropertyValue 'installing' -Force
$record|Add-Member -NotePropertyName pendingFiles -NotePropertyValue @($plan.ToArray()|ForEach-Object{[ordered]@{relativePath=$_.relativePath;sha256=$_.sha256}}) -Force
# Durable ownership intent precedes copying. It survives interruption without unowned orphan directories.
Save-OieManifest $manifestPath $record
$results=New-Object 'System.Collections.Generic.List[object]'
try{
  foreach($entry in $plan){
    $file=Resolve-OieOwnedChild $root $entry.relativePath
    $same=(Test-Path -LiteralPath $file) -and (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() -eq $entry.sha256
    if(-not $same){New-Item -Path ([IO.Path]::GetDirectoryName($file)) -ItemType Directory -Force|Out-Null;Assert-OieSafePath $file;Copy-Item -LiteralPath $entry.source -Destination $file -Force}
  }
  $record.files=@($plan.ToArray()|ForEach-Object{[ordered]@{relativePath=$_.relativePath;sha256=$_.sha256}})
  Save-OieManifest $manifestPath $record
  if(-not $SkipOfficeRegistration){
    & (Join-Path $root 'maintenance\install-office.ps1') -PayloadRoot (Join-Path $root 'office') -InstallRoot $root
    $record.officeRegistered=$true;Save-OieManifest $manifestPath $record
    $results.Add([ordered]@{component='office';status='registered_for_next_host_load'})
  }
  if(-not $SkipEditors){foreach($kind in $Hosts){
    $cli=Get-OieCli $kind
    if(-not $cli){$results.Add([ordered]@{component=$kind;status='not_detected'});continue}
    $existing=@($record.hosts|Where-Object{$_.kind -eq $kind})|Select-Object -First 1
    $already=Test-OieExtensionPresent $cli 'liketocood345.office-in-editor'
    $owner=if($existing){[bool]$existing.installedByThisProduct}else{-not $already}
    $hostRecord=[pscustomobject]@{kind=$kind;extensionId='liketocood345.office-in-editor';installedByThisProduct=$owner;state='installing'}
    $record.hosts=@($record.hosts|Where-Object{$_.kind -ne $kind})+@($hostRecord);Save-OieManifest $manifestPath $record
    & $cli.Source --install-extension $VsixPath --force
    if($LASTEXITCODE -ne 0 -or -not(Test-OieExtensionPresent $cli 'liketocood345.office-in-editor')){$hostRecord.state='failed';$results.Add([ordered]@{component=$kind;status='failed'})}
    else{$hostRecord.state='installed';$results.Add([ordered]@{component=$kind;status='installed'})}
    Save-OieManifest $manifestPath $record
  }}
  $record.transactionState=if(@($results.ToArray()|Where-Object{$_.status -eq 'failed'}).Count){'partial'}else{'complete'}
  $record.pendingFiles=@();Save-OieManifest $manifestPath $record
}catch{
  $record.transactionState='pending-repair';Save-OieManifest $manifestPath $record
  throw ('Installation did not fully complete; owned state was preserved for repair. No Office process was closed. ' + $_.Exception.Message)
}
$results.ToArray()|ConvertTo-Json -Depth 6
Write-Output 'Installation finished without closing Office/editors. DSH is deferred in this release.'
