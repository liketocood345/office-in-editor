param([Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run,[switch]$Smoke,
      [ValidateSet('word','excel','powerpoint')][string]$App='word',
      [ValidateSet('always','switch','mouse')][string]$PixelPolicy='always', [ValidateRange(0,40)][int]$Cycles=0,
      [ValidateSet('click','raise','raise-baseline')][string]$Scenario='click')
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$root=Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\focus-hybrid-'+$Run)
$task='OfficeInEditor-FocusHybrid-'+$Run
$utf8=[Text.UTF8Encoding]::new($false)
if($root -ine $PSScriptRoot -or (Get-Process -Id $PID).SessionId -eq 0){throw 'Unexpected focus lab context'}
$status=[ordered]@{ run=$Run; app=$App; codeExited=$false; officeExited=$false; stage='preflight' }
$officeProcess=$null
$editor=$null
$stem='synthetic-'+$Run
$format=switch($App){'word'{'docx'} 'excel'{'xlsx'} 'powerpoint'{'pptx'}}
$source=Join-Path $root ($stem+'.'+$format)
$sourceHash=$null
function Read-SaveEvidence {
    $result=[ordered]@{sourceExists=(Test-Path -LiteralPath $source -PathType Leaf);readable=$false;hashChanged=$false;markerPresent=$false}
    if(-not $result.sourceExists -or -not $sourceHash){return $result}
    try {
        $stream=[IO.File]::Open($source,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        try {
            $result.hashChanged=(Get-FileHash -InputStream $stream -Algorithm SHA256).Hash -ne $sourceHash
            $stream.Position=0
            $package=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Read,$true)
            try {
                $entryPath=switch($App){'word'{'word/document.xml'} 'excel'{'xl/worksheets/sheet1.xml'} 'powerpoint'{'ppt/slides/slide1.xml'}}
                $entry=$package.GetEntry($entryPath)
                if(-not $entry -or $entry.Length -gt 1048576){return $result}
                $reader=[IO.StreamReader]::new($entry.Open())
                try {$xml=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
                $body=$xml.DocumentElement.InnerText
                if($App -eq 'excel'){
                    $cell=$xml.SelectSingleNode("//*[local-name()='c' and @r='E2']")
                    $body=if($cell){$cell.InnerText}else{''}
                    if($cell -and $cell.GetAttribute('t') -eq 's'){
                        $strings=$package.GetEntry('xl/sharedStrings.xml')
                        if(-not $strings -or $strings.Length -gt 1048576){return $result}
                        $reader=[IO.StreamReader]::new($strings.Open())
                        try {$shared=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
                        $body=$shared.DocumentElement.ChildNodes.Item([int]$body).InnerText
                    }
                }
                $result.markerPresent=$body.Contains('OIE_FOCUS_SAVED')
                $result.readable=$true
            }finally{$package.Dispose()}
        }finally{$stream.Dispose()}
    }catch{
        $result.readError=$_.Exception.GetType().Name
    }
    return $result
}
try {
    $editorExe=Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\Code.exe'
    if(-not (Test-Path -LiteralPath $editorExe -PathType Leaf)){throw 'VS Code unavailable'}
    Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
    $fixture=[IO.Compression.ZipFile]::Open((Join-Path $root ($stem+'.docx')),[IO.Compression.ZipArchiveMode]::Create)
    try {
        $parts=@{
            '[Content_Types].xml'='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
            '_rels/.rels'='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
            'word/document.xml'='<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Focus hybrid lab synthetic document</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
        }
        foreach($part in $parts.GetEnumerator()){
            $writer=[IO.StreamWriter]::new($fixture.CreateEntry($part.Key).Open(),$utf8)
            try{$writer.Write($part.Value)}finally{$writer.Dispose()}
        }
    }finally{$fixture.Dispose()}
    $sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    $core=Join-Path $env:LOCALAPPDATA 'Programs\OfficeInEditor\current\core\OfficeInEditor.Core.exe'
    if(Test-Path -LiteralPath $core -PathType Leaf){
        $env:OFFICE_IN_EDITOR_TRACE='1'
        & $core --ensure-broker | Out-Null
        $status.brokerReady=$LASTEXITCODE -eq 0
    }
    $status.stage='office_open'
    $opened=& (Join-Path $root 'open-fixture.ps1') -Run $Run -App $App
    $officeProcess=Get-Process -Id $opened.officePid
    $profile=Join-Path $root 'editor-profile'
    $extensions=Join-Path $root 'extensions'
    New-Item -ItemType Directory -Path (Join-Path $profile 'User'),$extensions | Out-Null
    $settings=@{'telemetry.telemetryLevel'='off';'update.mode'='none';'extensions.autoUpdate'=$false;
        'window.restoreWindows'='none';'workbench.startupEditor'='none';'window.title'='${activeEditorShort}${separator}Office Focus Lab'}
    [IO.File]::WriteAllText((Join-Path $profile 'User\settings.json'),($settings|ConvertTo-Json),$utf8)
    $env:OIE_FOCUS_LAB=$root
    $os=Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $uac=Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA
    $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
    $principal=[Security.Principal.WindowsPrincipal]::new($identity)
    $fingerprints=[ordered]@{}
    foreach($item in @(@('native',(Join-Path $root 'native.exe')),
                      @('extension',(Join-Path $root 'extension\extension.cjs')),
                      @('mcp',(Join-Path $root 'extension\mcp.cjs')),
                      @('addin',(Join-Path $env:LOCALAPPDATA 'Programs\OfficeInEditor\current\office\OfficeInEditor.Addin.dll')))){
        if(Test-Path -LiteralPath $item[1] -PathType Leaf){$fingerprints[$item[0]]=(Get-FileHash -LiteralPath $item[1]).Hash.Substring(0,16).ToLowerInvariant()}
    }
    $officeVersion=try{$officeProcess.MainModule.FileVersionInfo.FileVersion}catch{$null}
    $environment=[ordered]@{run=$Run;app=$App;pixelPolicy=$PixelPolicy;cycles=$Cycles;scenario=$Scenario;
        osBuild=[string]$os.CurrentBuild+'.'+[string]$os.UBR;officeVersion=$officeVersion;
        vscodeVersion=(Get-Item -LiteralPath $editorExe).VersionInfo.ProductVersion;
        uacEnabled=([int]$uac.EnableLUA -ne 0);elevated=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);
        build=$fingerprints}
    [IO.File]::WriteAllText((Join-Path $root 'environment.json'),($environment|ConvertTo-Json -Depth 4 -Compress),$utf8)
    $status.stage='editor_open'
    $arguments=@('--new-window','--skip-welcome','--skip-release-notes','--force-renderer-accessibility',('--user-data-dir="'+$profile+'"'),
        ('--extensions-dir="'+$extensions+'"'),('--shared-data-dir="'+(Join-Path $root 'shared')+'"'),
        ('--extensionDevelopmentPath="'+(Join-Path $root 'extension')+'"'))
    $editor=Start-Process -FilePath $editorExe -ArgumentList $arguments -WorkingDirectory $root -PassThru
    [void]$editor.Handle
    $input=@{run=$Run;app=$App;officePid=$officeProcess.Id;codePid=$editor.Id;stem=$stem;smoke=[bool]$Smoke;pixelPolicy=$PixelPolicy;cycles=$Cycles;scenario=$Scenario}
    [IO.File]::WriteAllText((Join-Path $root 'inputs.partial'),($input|ConvertTo-Json -Compress),$utf8)
    [IO.File]::Move((Join-Path $root 'inputs.partial'),(Join-Path $root 'inputs.json'))
    $status.stage='editor_wait'
    if($Smoke){$status.codeExited=$editor.WaitForExit(55000+4500*$Cycles)}
    else{$editor.WaitForExit();$status.codeExited=$true}
    $status.resultReady=Test-Path -LiteralPath (Join-Path $root 'result.json') -PathType Leaf
} catch {
    $status.error=$_.Exception.GetType().Name
    $status.errorId=($_.FullyQualifiedErrorId -split ',')[0]
    $status.line=$_.InvocationInfo.ScriptLineNumber
} finally {
    if($editor -and -not $status.codeExited){
        $editor.Refresh()
        $status.codeExited=$editor.HasExited -or ($editor.CloseMainWindow() -and $editor.WaitForExit(7000))
    }
    if($officeProcess -and (-not $editor -or $status.codeExited)){
        $officeProcess.Refresh()
        if($officeProcess.HasExited){$status.officeExited=$true}
        elseif($officeProcess.MainWindowTitle.StartsWith($stem,[StringComparison]::OrdinalIgnoreCase) -and $officeProcess.CloseMainWindow()){
            $status.officeExited=$officeProcess.WaitForExit(7000)
        }
    }
    $status.sourceSave=Read-SaveEvidence
    $status.fixtureLockPresent=Test-Path -LiteralPath (Join-Path $root ('~$'+$stem.Substring(2)+'.'+$format))
    [IO.File]::WriteAllText((Join-Path $root 'finished.partial'),($status|ConvertTo-Json -Compress),$utf8)
    [IO.File]::Move((Join-Path $root 'finished.partial'),(Join-Path $root 'finished.json'))
    if(-not $editor -or $status.codeExited){Unregister-ScheduledTask -TaskName $task -Confirm:$false -ErrorAction SilentlyContinue}
}
