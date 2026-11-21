function Assert-OieSafePath {
  param([Parameter(Mandatory=$true)][string]$Path)
  $current=[IO.Path]::GetFullPath($Path)
  while($current){
    if(Test-Path -LiteralPath $current){$item=Get-Item -LiteralPath $current -Force;if(($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw 'A product path crosses a reparse point; no change was made.'}}
    $parent=[IO.Path]::GetDirectoryName($current.TrimEnd([IO.Path]::DirectorySeparatorChar))
    if(-not $parent -or $parent -eq $current){break};$current=$parent
  }
}
function Resolve-OieOwnedChild {
  param([string]$Root,[string]$Relative)
  if(-not $Relative -or [IO.Path]::IsPathRooted($Relative) -or $Relative.Contains(':') -or @($Relative -split '[\\/]'|Where-Object{$_ -eq '..'}).Count){throw 'Invalid ownership manifest path.'}
  $base=[IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar)
  $candidate=[IO.Path]::GetFullPath((Join-Path $base $Relative))
  if(-not $candidate.StartsWith($base+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Ownership path escaped its product root.'}
  Assert-OieSafePath $candidate;return $candidate
}
function Save-OieManifest {
  param([string]$Path,[object]$Record)
  Assert-OieSafePath $Path;$text=$Record|ConvertTo-Json -Depth 16
  if([Text.Encoding]::UTF8.GetByteCount($text) -gt 1048576){throw 'Installation metadata budget exceeded.'}
  $next=$Path+'.next';Assert-OieSafePath $next
  [IO.File]::WriteAllText($next,$text,(New-Object Text.UTF8Encoding($false)))
  if(Test-Path -LiteralPath $Path){
    $backup=$Path+'.bak';Assert-OieSafePath $backup
    [IO.File]::Replace($next,$Path,$backup)
    if(Test-Path -LiteralPath $backup){Remove-Item -LiteralPath $backup -Force}
  }else{Move-Item -LiteralPath $next -Destination $Path}
}
function Get-OieCli {
  param([string]$Kind)
  $name=if($Kind -eq 'vscode'){'code'}elseif($Kind -eq 'cursor'){'cursor'}else{throw 'Unknown editor adapter.'}
  return Get-Command $name -ErrorAction SilentlyContinue
}
function Test-OieOfficeRunning {
  return [bool](Get-Process -Name @('WINWORD','EXCEL','POWERPNT') -ErrorAction SilentlyContinue)
}

function Clear-OieDisabledItems {
  param([switch]$SkipIfOfficeRunning)
  if ($SkipIfOfficeRunning -and (Test-OieOfficeRunning)) {
    return [pscustomobject]@{ Cleared = 0; Skipped = $true; Reason = 'Office is running; DisabledItems unchanged' }
  }
  $cleared = 0
  $needle = 'officeineditor.addin'
  foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    $hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
      $office = $hkcu.OpenSubKey('Software\Microsoft\Office', $true)
      if ($null -eq $office) { continue }
      try {
        foreach ($versionName in $office.GetSubKeyNames()) {
          $version = $office.OpenSubKey($versionName, $true)
          if ($null -eq $version) { continue }
          try {
            foreach ($appName in @('Word','Excel','PowerPoint','Common')) {
              $disabled = $version.OpenSubKey("$appName\Resiliency\DisabledItems", $true)
              if ($null -eq $disabled) { continue }
              try {
                foreach ($valueName in @($disabled.GetValueNames())) {
                  $raw = $disabled.GetValue($valueName)
                  if ($raw -isnot [byte[]]) { continue }
                  $unicode = [Text.Encoding]::Unicode.GetString([byte[]]$raw)
                  $ascii = [Text.Encoding]::ASCII.GetString([byte[]]$raw)
                  if ($unicode.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -lt 0 -and
                      $ascii.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                  $disabled.DeleteValue($valueName, $false)
                  $cleared++
                }
              } finally { $disabled.Dispose() }
            }
          } finally { $version.Dispose() }
        }
      } finally { $office.Dispose() }
    } finally { $hkcu.Dispose() }
  }
  return [pscustomobject]@{ Cleared = $cleared; Skipped = $false }
}

function Test-OieExtensionPresent {
  param([object]$Cli,[string]$Id)
  if(-not $Cli){throw 'The editor CLI is not available to verify ownership.'}
  $ids=@(& $Cli.Source --list-extensions)
  if($LASTEXITCODE -ne 0){throw 'The editor extension inventory could not be verified.'}
  return $ids -contains $Id
}
