[CmdletBinding(SupportsShouldProcess = $true)]
param([Parameter(Mandatory = $true)][string]$InstallRoot, [switch]$ValidateOnly)
$ErrorActionPreference = 'Stop'
$addinClsid = '{6B4072D4-3EB1-4CAB-995D-4F922B19D4B1}'
$addinProgId = 'OfficeInEditor.Addin'
$ownerMarker = 'office-in-editor:addin:v1'
$productRoot = [IO.Path]::GetFullPath((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\OfficeInEditor'))
$stableRoot = [IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
if (-not $stableRoot.StartsWith($productRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'InstallRoot must be inside the current-user OfficeInEditor product directory.'
}
$addinUri = [Uri](Join-Path $stableRoot 'office\OfficeInEditor.Addin.dll')
$expectedCodeBases = @($addinUri.ToString(),$addinUri.AbsoluteUri)
$removals = 0
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    $userRegistry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        $classPath = 'Software\Classes\CLSID\' + $addinClsid
        $classKey = $userRegistry.OpenSubKey($classPath)
        if ($null -eq $classKey) { continue }
        try { if ($classKey.GetValue('OfficeInEditorOwner') -ne $ownerMarker) { throw 'COM registration is not owned by this product; nothing was removed.' } }
        finally { $classKey.Dispose() }
        $server = $userRegistry.OpenSubKey($classPath + '\InprocServer32')
        if ($null -eq $server) { throw 'Incomplete COM registration requires explicit inspection; nothing was removed.' }
        try { if ($server.GetValue('CodeBase') -notin $expectedCodeBases -or $server.GetValue('Class') -ne 'OfficeInEditor.OfficeAddin.OfficeAddin') {
                throw 'COM registration belongs to a different installation; nothing was removed.'
            } } finally { $server.Dispose() }
        $paths = @('Word', 'Excel', 'PowerPoint') | ForEach-Object { 'Software\Microsoft\Office\' + $_ + '\Addins\' + $addinProgId }
        $paths = @($paths) + @('Software\Classes\' + $addinProgId, $classPath)
        # Validate the complete bounded set before removing any key in this view.
        foreach ($entryPath in $paths) {
            $entry = $userRegistry.OpenSubKey($entryPath)
            if ($null -ne $entry) {
                try { if ($entry.GetValue('OfficeInEditorOwner') -ne $ownerMarker) { throw 'A registry key has unknown ownership; nothing was removed.' } }
                finally { $entry.Dispose() }
            }
        }
        if (-not $ValidateOnly -and $PSCmdlet.ShouldProcess('Owned current-user Office add-in registry keys', 'Unregister OfficeInEditor without unloading or closing Office')) {
            foreach ($entryPath in $paths) { $userRegistry.DeleteSubKeyTree($entryPath, $false) }
            $removals++
        }
    } finally { $userRegistry.Dispose() }
}
[pscustomobject]@{ UnregisteredViews = $removals; ValidatedOnly = [bool]$ValidateOnly; ActivatedOffice = $false; RemovedFiles = $false; ExistingOfficeSessions = 'unchanged' }
