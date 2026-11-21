[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$PayloadRoot,
    [Parameter(Mandatory = $true)][string]$InstallRoot,
    [ValidateSet('Auto', 'x86', 'x64')][string]$OfficeBitness = 'Auto',
    [switch]$ValidateOnly
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'installation-common.ps1')
$addinClsid = '{6B4072D4-3EB1-4CAB-995D-4F922B19D4B1}'
$addinProgId = 'OfficeInEditor.Addin'
$addinClass = 'OfficeInEditor.OfficeAddin.OfficeAddin'
$ownerMarker = 'office-in-editor:addin:v1'
$productRoot = [IO.Path]::GetFullPath((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\OfficeInEditor'))
$stableRoot = [IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
if (-not $stableRoot.StartsWith($productRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'InstallRoot must be inside the current-user OfficeInEditor product directory.'
}
$installedDll = Join-Path $stableRoot 'office\OfficeInEditor.Addin.dll'
$payloadDll = Join-Path ([IO.Path]::GetFullPath($PayloadRoot)) 'OfficeInEditor.Addin.dll'
if (-not (Test-Path -LiteralPath $payloadDll -PathType Leaf) -or -not (Test-Path -LiteralPath $installedDll -PathType Leaf)) {
    throw 'Copy the Office payload into the stable office directory before registering.'
}
# Metadata and hashes only: never load/instantiate the COM class or contact a running Office process.
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($installedDll)
if ($assemblyName.Name -ne 'OfficeInEditor.Addin') { throw 'Unexpected add-in assembly.' }
if ((Get-FileHash -LiteralPath $payloadDll -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $installedDll -Algorithm SHA256).Hash) {
    throw 'Stable installed payload does not match the supplied Office payload.'
}
$codeBase = ([Uri]$installedDll).ToString()
if ($OfficeBitness -eq 'Auto') {
    foreach ($probeView in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
        $machine = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $probeView)
        try {
            $configuration = $machine.OpenSubKey('SOFTWARE\Microsoft\Office\ClickToRun\Configuration')
            if ($null -ne $configuration) {
                try { $platform = [string]$configuration.GetValue('Platform') } finally { $configuration.Dispose() }
                if ($platform -in @('x86', 'x64')) { $OfficeBitness = $platform; break }
            }
        } finally { $machine.Dispose() }
    }
    if ($OfficeBitness -eq 'Auto') { throw 'Office bitness was not detected. Specify -OfficeBitness x86 or x64.' }
}
$registryView = if ($OfficeBitness -eq 'x64') { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
$userRegistry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
$clsidPath = 'Software\Classes\CLSID\' + $addinClsid
$progIdPath = 'Software\Classes\' + $addinProgId
$officeApps = @('Word', 'Excel', 'PowerPoint')
try {
    $existing = $userRegistry.OpenSubKey($clsidPath)
    if ($null -ne $existing) {
        try { if ($existing.GetValue('OfficeInEditorOwner') -ne $ownerMarker) { throw 'The COM identifier is owned by another registration.' } }
        finally { $existing.Dispose() }
    }
    foreach ($ownedPath in @($progIdPath) + @($officeApps | ForEach-Object { 'Software\Microsoft\Office\' + $_ + '\Addins\' + $addinProgId })) {
        $ownedKey = $userRegistry.OpenSubKey($ownedPath)
        if ($null -ne $ownedKey) {
            try { if ($ownedKey.GetValue('OfficeInEditorOwner') -ne $ownerMarker) { throw 'An add-in registration is not owned by this product.' } }
            finally { $ownedKey.Dispose() }
        }
    }
    if ($ValidateOnly) {
        $server = $userRegistry.OpenSubKey($clsidPath + '\InprocServer32')
        if ($null -eq $server) { throw 'Add-in is not registered for the selected Office bitness.' }
        try {
            if ($server.GetValue('CodeBase') -notin @($codeBase,([Uri]$installedDll).AbsoluteUri) -or $server.GetValue('Class') -ne $addinClass -or
                $server.GetValue('Assembly') -ne $assemblyName.FullName) { throw 'COM registration does not match the stable installed payload.' }
        } finally { $server.Dispose() }
        foreach ($appName in $officeApps) {
            $appKey = $userRegistry.OpenSubKey('Software\Microsoft\Office\' + $appName + '\Addins\' + $addinProgId)
            if ($null -eq $appKey) { throw 'One or more Office add-in registrations are missing.' }
            try { if ($appKey.GetValue('LoadBehavior') -ne 3) { throw 'Office add-in loading is not enabled.' } }
            finally { $appKey.Dispose() }
        }
        [pscustomobject]@{ Validated = $true; OfficeBitness = $OfficeBitness; ActivatedOffice = $false }
        return
    }
    if (-not $PSCmdlet.ShouldProcess('Current-user Office add-in registry keys', 'Register stable OfficeInEditor managed COM add-in')) { return }
    $classKey = $userRegistry.CreateSubKey($clsidPath)
    try { $classKey.SetValue('', 'Office in Editor'); $classKey.SetValue('OfficeInEditorOwner', $ownerMarker) }
    finally { $classKey.Dispose() }
    foreach ($serverPath in @(($clsidPath + '\InprocServer32'), ($clsidPath + '\InprocServer32\' + $assemblyName.Version.ToString()))) {
        $server = $userRegistry.CreateSubKey($serverPath)
        try {
            $server.SetValue('', 'mscoree.dll'); $server.SetValue('ThreadingModel', 'Apartment')
            $server.SetValue('Class', $addinClass); $server.SetValue('Assembly', $assemblyName.FullName)
            $server.SetValue('RuntimeVersion', 'v4.0.30319'); $server.SetValue('CodeBase', $codeBase)
        } finally { $server.Dispose() }
    }
    $progLink = $userRegistry.CreateSubKey($clsidPath + '\ProgId')
    try { $progLink.SetValue('', $addinProgId) } finally { $progLink.Dispose() }
    $category = $userRegistry.CreateSubKey($clsidPath + '\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}')
    $category.Dispose()
    $program = $userRegistry.CreateSubKey($progIdPath)
    try { $program.SetValue('', 'Office in Editor'); $program.SetValue('OfficeInEditorOwner', $ownerMarker) }
    finally { $program.Dispose() }
    $programClass = $userRegistry.CreateSubKey($progIdPath + '\CLSID')
    try { $programClass.SetValue('', $addinClsid) } finally { $programClass.Dispose() }
    foreach ($appName in $officeApps) {
        $appKey = $userRegistry.CreateSubKey('Software\Microsoft\Office\' + $appName + '\Addins\' + $addinProgId)
        try {
            $appKey.SetValue('OfficeInEditorOwner', $ownerMarker)
            $appKey.SetValue('FriendlyName', 'Office in Editor')
            $appKey.SetValue('Description', 'Add selected Office content to an explicitly registered chat. No automatic message submission.')
            $appKey.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
            $appKey.SetValue('CommandLineSafe', 0, [Microsoft.Win32.RegistryValueKind]::DWord)
        } finally { $appKey.Dispose() }
    }
    $disabled = Clear-OieDisabledItems -SkipIfOfficeRunning
    [pscustomobject]@{
      Registered = $true
      OfficeBitness = $OfficeBitness
      ActivatedOffice = $false
      ExistingOfficeSessions = 'unchanged; add-in loads on a subsequent launch'
      DisabledItemsCleared = [int]$disabled.Cleared
      DisabledItemsSkipped = [bool]$disabled.Skipped
    }
} finally { $userRegistry.Dispose() }
