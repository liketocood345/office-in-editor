param([Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{32}$')][string]$Run,
      [Parameter(Mandatory)][ValidateSet('word','excel','powerpoint')][string]$App,[switch]$AsJson)
$ErrorActionPreference='Stop'
$root=Join-Path $env:LOCALAPPDATA ('OfficeInEditorLab\focus-hybrid-'+$Run)
if($PSScriptRoot -ine $root -or (Get-Process -Id $PID).SessionId -eq 0){throw 'Unexpected focus lab context'}
$stem='synthetic-'+$Run
$appInfo=switch($App){
    'word' {@{name='WINWORD';ext='docx';switches=@('/n','/q')}}
    'excel' {@{name='EXCEL';ext='xlsx';switches=@('/x')}}
    'powerpoint' {@{name='POWERPNT';ext='pptx';switches=@('/O')}}
}
$source=Join-Path $root ($stem+'.'+$appInfo.ext)
$key=Get-Item -LiteralPath ('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+$appInfo.name+'.EXE')
$exe=$key.GetValue('')
if(-not (Test-Path -LiteralPath $source -PathType Leaf) -or -not (Test-Path -LiteralPath $exe -PathType Leaf)){
    throw 'Office executable or test fixture is unavailable'
}
$sessionId=(Get-Process -Id $PID).SessionId
$argsForOffice=@($appInfo.switches)+('"'+$source+'"')
[void](Start-Process -FilePath $exe -ArgumentList $argsForOffice -WorkingDirectory $root -PassThru)
for($attempt=0;$attempt -lt 85;$attempt++){
    $hostProcess=Get-Process -Name $appInfo.name -ErrorAction SilentlyContinue | Where-Object {
        $_.SessionId -eq $sessionId -and $_.MainWindowTitle.StartsWith($stem,[StringComparison]::OrdinalIgnoreCase)
    } | Select-Object -First 1
    if($hostProcess){
        $opened=[pscustomobject]@{app=$App;officePid=$hostProcess.Id;stem=$stem}
        if($AsJson){$opened|ConvertTo-Json -Compress}else{$opened}
        return
    }
    Start-Sleep -Milliseconds 200
}
throw 'Synthetic Office document window did not open'
