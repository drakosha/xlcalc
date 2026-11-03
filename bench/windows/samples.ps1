# The samples of the Number Format list on the Home tab of Excel for Windows: Excel has a ribbon only in a session on a
# screen, so bench/windows/samples.ts runs this as a task there (schtasks /IT). The job (JSON, UTF-8) lists books beside
# it, each with the cells to read on its sheet S. Each book is opened (read only); for each cell the cell is selected
# through the object model, Excel taken to the front, the box's list opened by UI Automation (its Expand, from another
# thread, a click on it when no window comes), and the list's window read through MSAA, as UI Automation shows only
# its frame: each item's name and description, the description being the sample under the name. A screenshot of the
# list goes beside the output (shots\<book>-<cell>.png). The list is closed by Escape before the next call to Excel.
# Output: a JSON object { version, ui, screen, answers: [per book [{ box, text, code, items: [[name, sample]], size }]] }
# written as UTF-8 to $out.
param([Parameter(Mandatory = $true)][string]$job, [Parameter(Mandatory = $true)][string]$out)
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, UIAutomationClient, UIAutomationTypes
. (Join-Path $PSScriptRoot 'front.ps1')
Add-Type -ReferencedAssemblies Accessibility -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Accessibility;

// The list items (MSAA role 34) of a window by name and description, depth first, and the visible top-level windows
// of a process by Win32 alone.
public static class XlList {
  [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object o);
  [DllImport("oleacc.dll")] static extern int AccessibleChildren(IAccessible p, int start, int count, [Out] object[] children, out int got);
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  static Guid IID = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");

  static string S(Func<object> f) { try { object o = f(); return o == null ? "" : o.ToString(); } catch { return ""; } }

  static void Walk(IAccessible a, object id, int d, List<string[]> list) {
    if (d > 20) return;
    object r = null;
    try { r = a.get_accRole(id); } catch { }
    if (r is int && (int)r == 34) list.Add(new string[] { S(() => a.get_accName(id)), S(() => a.get_accDescription(id)) });
    if (!(id is int && (int)id == 0)) return;
    int n = 0;
    try { n = a.accChildCount; } catch { return; }
    if (n == 0) return;
    var c = new object[n];
    int got;
    if (AccessibleChildren(a, 0, n, c, out got) != 0) return;
    for (int i = 0; i < got; i++) {
      var x = c[i] as IAccessible;
      if (x != null) Walk(x, 0, d + 1, list); else Walk(a, c[i], d + 1, list);
    }
  }

  public static List<string[]> Items(IntPtr h) {
    object o;
    var list = new List<string[]>();
    // OBJID_CLIENT
    if (AccessibleObjectFromWindow(h, 0xFFFFFFFC, ref IID, out o) != 0) return list;
    Walk((IAccessible)o, 0, 0, list);
    return list;
  }

  public static List<long> Windows(int pid) {
    var r = new List<long>();
    EnumWindows((h, l) => {
      uint p;
      GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h)) r.Add(h.ToInt64());
      return true;
    }, IntPtr.Zero);
    return r;
  }
}
'@
$UIA = [System.Windows.Automation.AutomationElement]
$en = [Globalization.CultureInfo]'en-US'
$books = [IO.File]::ReadAllText($job, [Text.Encoding]::UTF8) | ConvertFrom-Json
$dir = Split-Path $job
$shots = Join-Path $dir 'shots'
New-Item -ItemType Directory -Force $shots | Out-Null
function Code($cell) { [System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::GetProperty, $null, $cell, $null, $null, $en, $null) }
# the windows of Excel's process other than those of the start
function Others { @([XlList]::Windows($script:proc) | Where-Object { $script:baseline -notcontains $_ }) }
function Escape {
  [XlScreen.Win]::keybd_event(0x1B, 0, 0, [UIntPtr]::Zero)
  [XlScreen.Win]::keybd_event(0x1B, 0, 2, [UIntPtr]::Zero)
}
# the list's window, opened by $open, once it shows its items (four seconds at most), or $null
function Opened($open) {
  & $open
  for ($i = 0; $i -lt 40; $i++) {
    Start-Sleep -Milliseconds 100
    $w = @(Others)
    if ($w.Count) {
      Start-Sleep -Milliseconds 600
      return $w[0]
    }
  }
  $null
}
function ExpandAsync($e) {
  $ps = [powershell]::Create()
  [void]$ps.AddScript({ param($x) $x.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand() }).AddArgument($e)
  [void]$ps.BeginInvoke()
}
function ClickOn($e) {
  $r = $e.Current.BoundingRectangle
  [void][XlScreen.Win]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
  Start-Sleep -Milliseconds 200
  [XlScreen.Win]::mouse_event(0x2, 0, 0, 0, [UIntPtr]::Zero)
  [XlScreen.Win]::mouse_event(0x4, 0, 0, 0, [UIntPtr]::Zero)
}
function Shot($h, $name) {
  $r = $UIA::FromHandle([IntPtr]$h).Current.BoundingRectangle
  if ($r.IsEmpty -or $r.Width -lt 2) { return '' }
  $bmp = New-Object System.Drawing.Bitmap([int]$r.Width, [int]$r.Height)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen([int]$r.X, [int]$r.Y, 0, 0, $bmp.Size)
  $bmp.Save((Join-Path $shots "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
  '{0}x{1}' -f [int]$r.Width, [int]$r.Height
}
$res = [ordered]@{ answers = New-Object System.Collections.ArrayList }
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  $xl.WindowState = -4137
  $res.version = [string]$xl.Version + ' build ' + [string]$xl.Build
  $res.ui = [int]$xl.LanguageSettings.LanguageID(2)
  $res.screen = [string][System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Size
  $hwnd = [int64]$xl.Hwnd
  $script:proc = (Get-Process EXCEL | Where-Object { [int64]$_.MainWindowHandle -eq $hwnd } | Select-Object -First 1).Id
  foreach ($b in $books) {
    $wb = $xl.Workbooks.Open((Join-Path $dir $b.file), 0, $true)
    $ws = $wb.Worksheets.Item('S')
    [void]$ws.Activate()
    Start-Sleep 2
    $byId = New-Object System.Windows.Automation.PropertyCondition($UIA::AutomationIdProperty, 'NumberFormatGallery')
    $box = $UIA::FromHandle([IntPtr]$hwnd).FindFirst([System.Windows.Automation.TreeScope]::Descendants, $byId)
    if (-not $box) { throw 'no Number Format box (NumberFormatGallery) in the window' }
    $res.box = $box.Current.Name
    $cells = New-Object System.Collections.ArrayList
    foreach ($at in @($b.cells)) {
      $cell = $ws.Range($at)
      [void]$cell.Select()
      $text = [string]$cell.Text
      $code = [string](Code $cell)
      Front $hwnd
      Start-Sleep -Milliseconds 700
      $shown = ''
      try { $shown = [string]$box.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value } catch {}
      $script:baseline = @([XlList]::Windows($script:proc))
      $w = Opened { ExpandAsync $box }
      $how = 'expand'
      if (-not $w) { $w = Opened { ClickOn $box }; $how = 'click' }
      $entry = [ordered]@{ cell = $at; box = $shown; text = $text; code = $code; how = $how; items = @(); size = '' }
      if ($w) {
        $entry.items = @([XlList]::Items([IntPtr]$w) | ForEach-Object { , @($_[0], $_[1]) })
        $entry.size = Shot $w ('{0}-{1}' -f [IO.Path]::GetFileNameWithoutExtension($b.file), $at)
      }
      # the list closed before the next call to Excel
      for ($k = 0; $k -lt 5 -and @(Others).Count; $k++) { Front $hwnd; Escape; Start-Sleep -Milliseconds 400 }
      if (@(Others).Count) { throw "the list stayed open after $at" }
      if (-not $w) { $entry.error = 'the list did not open' }
      [void]$cells.Add($entry)
    }
    [void]$res.answers.Add($cells)
    $wb.Close($false)
  }
} catch { $res.error = [string]$_.Exception.Message + ' ' + [string]$_.InvocationInfo.PositionMessage } finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding $false))
}
