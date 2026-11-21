# Post-hang diagnostic: dump Win32 windows and the UIA tree of the probe's own WINWORD process.
# Read-only. Never clicks, never closes, never kills. Writes into -RunRoot (or stdout).
param(
  [Parameter(Mandatory = $true)][string]$RunRoot,
  [int]$ProcId = 0
)
$ErrorActionPreference = 'Continue'
$outPath = Join-Path $RunRoot 'windows-dump.log'
function Log([string]$msg) {
  try { [IO.File]::AppendAllText($outPath, ((Get-Date).ToString('HH:mm:ss.fff') + ' ' + $msg + [Environment]::NewLine), [Text.Encoding]::UTF8) } catch { }
  Write-Output $msg
}

if (-not ('WinDiag' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
public static class WinDiag {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder sb, int max);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder sb, int max);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
  [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
  static string Text(IntPtr h) { StringBuilder sb = new StringBuilder(512); GetWindowTextW(h, sb, 512); return sb.ToString(); }
  static string Cls(IntPtr h) { StringBuilder sb = new StringBuilder(256); GetClassNameW(h, sb, 256); return sb.ToString(); }
  public static List<string> TopWindows(uint[] pids) {
    List<string> rows = new List<string>();
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (Array.IndexOf(pids, pid) >= 0) {
        rows.Add(string.Format("top hwnd=0x{0:X} pid={1} visible={2} enabled={3} owner=0x{4:X} cls=[{5}] title=[{6}]",
          h.ToInt64(), pid, IsWindowVisible(h), IsWindowEnabled(h), GetWindow(h, 4).ToInt64(), Cls(h), Text(h)));
      }
      return true;
    }, IntPtr.Zero);
    return rows;
  }
  public static List<string> ChildWindows(IntPtr parent, int maxDepth, int depth) {
    List<string> rows = new List<string>();
    if (depth > maxDepth) return rows;
    EnumChildWindows(parent, delegate(IntPtr h, IntPtr l) {
      string pad = new string('.', depth);
      rows.Add(string.Format("kid{0} hwnd=0x{1:X} parent=0x{2:X} visible={3} enabled={4} cls=[{5}] title=[{6}]",
        pad, h.ToInt64(), GetParent(h).ToInt64(), IsWindowVisible(h), IsWindowEnabled(h), Cls(h), Text(h)));
      return true;
    }, IntPtr.Zero);
    return rows;
  }
}
'@
}

$procs = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
if ($procs.Count -eq 0) { Log 'no WINWORD process'; exit 0 }
$targets = if ($ProcId -gt 0) { @($procs | Where-Object { $_.Id -eq $ProcId }) } else { $procs }
foreach ($p in $targets) { Log ("winword pid=$($p.Id) start=$($p.StartTime.ToString('HH:mm:ss'))") }
$pids = @($targets | ForEach-Object { [uint32]$_.Id })
$rows = [WinDiag]::TopWindows($pids)
foreach ($r in $rows) { Log ('win32 ' + $r) }
foreach ($r in $rows) {
  if ($r -match 'hwnd=(0x[0-9A-F]+)') {
    $h = [IntPtr][Convert]::ToInt64($Matches[1].Substring(2), 16)
    foreach ($k in [WinDiag]::ChildWindows($h, 3, 1)) { Log ('win32 ' + $k) }
  }
}

try {
  Add-Type -AssemblyName UIAutomationClient
  Add-Type -AssemblyName UIAutomationTypes
  $root = [System.Windows.Automation.AutomationElement]::RootElement
  foreach ($id in $pids) {
    $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$id)
    $wins = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
    Log ("uia pid=$id top-level=$($wins.Count)")
    foreach ($w in $wins) {
      Log ("uia top: [" + $w.Current.ClassName + "] '" + $w.Current.Name + "' offscreen=" + $w.Current.IsOffscreen)
      $all = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
      Log ("uia descendants: " + $all.Count)
      $i = 0
      foreach ($d in $all) {
        $i++
        if ($i -gt 400) { Log 'uia dump: truncated at 400'; break }
        $n = $d.Current.Name
        $c = $d.Current.ClassName
        $interesting = ($n -and $n.Trim().Length -gt 0) -or ($c -match 'NUIDialog|#32770|NetUI|MsoCommandBar|CtrlNotifySink|Button')
        if (-not $interesting) { continue }
        Log ("uia d: [" + $c + "] '" + $n + "' type=" + $d.Current.ControlType.ProgrammaticName + " id=" + $d.Current.AutomationId + " offscreen=" + $d.Current.IsOffscreen + " enabled=" + $d.Current.IsEnabled)
      }
    }
  }
} catch { Log ('uia error: ' + $_.Exception.Message) }
Log 'dump done'
