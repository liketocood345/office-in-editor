param()
$ErrorActionPreference='Stop'
$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
$settings=Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
if($identity.User.Value -notmatch '-500$' -or $settings.EnableLUA -ne 0){
    throw 'Use the product current-user installer for this account.'
}
$owner='office-in-editor:focus-lab-com:v1'
$className='OfficeInEditor.OfficeAddin.OfficeAddin'
$clsid='{6B4072D4-3EB1-4CAB-995D-4F922B19D4B1}'
$progId='OfficeInEditor.Addin'
$current=Join-Path $env:LOCALAPPDATA 'Programs\OfficeInEditor\current'
$source=Join-Path $current 'office\OfficeInEditor.Addin.dll'
$manifest=Get-Content -LiteralPath (Join-Path $current 'install-manifest.json') -Raw|ConvertFrom-Json
$file=@($manifest.files|Where-Object relativePath -eq 'office/OfficeInEditor.Addin.dll')[0]
$sha=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant()
if($manifest.productId -ne 'office-in-editor' -or -not $manifest.officeRegistered -or $file.sha256 -ne $sha){
    throw 'Installed Office add-in does not match the product manifest.'
}
$assembly=[Reflection.AssemblyName]::GetAssemblyName($source)
if($assembly.Name -ne 'OfficeInEditor.Addin'){throw 'Unexpected COM assembly.'}
$programFiles=[Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)
$stable=Join-Path $programFiles ('OfficeInEditor\LabCom\'+$sha)
$dll=Join-Path $stable 'OfficeInEditor.Addin.dll'
foreach($candidate in $programFiles,(Split-Path (Split-Path $stable)),(Split-Path $stable),$stable,$dll){
    if((Test-Path -LiteralPath $candidate) -and ((Get-Item -LiteralPath $candidate -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){
        throw 'Redirected machine add-in path.'
    }
}
$classPath='SOFTWARE\Classes\CLSID\'+$clsid
$programPath='SOFTWARE\Classes\'+$progId
$machine=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
function Set-Key([string]$Path,[hashtable]$Values){
    $key=$machine.CreateSubKey($Path)
    try{foreach($name in $Values.Keys){$key.SetValue($name,$Values[$name])}}finally{$key.Dispose()}
}
try{
    foreach($path in $classPath,$programPath){
        $key=$machine.OpenSubKey($path)
        if($key){try{if($key.GetValue('OfficeInEditorOwner') -ne $owner){throw 'The machine COM key has another owner.'}}finally{$key.Dispose()}}
    }
    New-Item -ItemType Directory -Path $stable -Force|Out-Null
    if(-not (Test-Path -LiteralPath $dll)){Copy-Item -LiteralPath $source -Destination $dll}
    if((Get-FileHash -LiteralPath $dll).Hash.ToLowerInvariant() -ne $sha){throw 'Machine COM copy checksum mismatch.'}
    $server=@{''='mscoree.dll';ThreadingModel='Apartment';Class=$className;Assembly=$assembly.FullName;
        RuntimeVersion='v4.0.30319';CodeBase=([Uri]$dll).ToString()}
    Set-Key $classPath @{''='Office in Editor';OfficeInEditorOwner=$owner}
    Set-Key ($classPath+'\InprocServer32') $server
    Set-Key ($classPath+'\InprocServer32\'+$assembly.Version) $server
    Set-Key ($classPath+'\ProgId') @{''=$progId}
    Set-Key ($classPath+'\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}') @{}
    Set-Key $programPath @{''='Office in Editor';OfficeInEditorOwner=$owner}
    Set-Key ($programPath+'\CLSID') @{''=$clsid}
}finally{$machine.Dispose()}
$instance=[Activator]::CreateInstance([type]::GetTypeFromProgID($progId,$true))
try{
    $markup=[string]$instance.GetCustomUI('Microsoft.Word.Document')
    if(-not $markup.Contains('OieCaptureOnMouse') -or -not $markup.Contains('OieContext0')){throw 'Unexpected COM Ribbon markup.'}
    @{machineClassActivated=$true;capturePolicyAvailable=$true;addToChatMarkup=$true}|ConvertTo-Json -Compress
}finally{if([Runtime.InteropServices.Marshal]::IsComObject($instance)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($instance)}}
