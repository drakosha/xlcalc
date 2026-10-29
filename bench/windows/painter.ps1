# The Format Painter of Excel for Windows: the painter works only in a session on a screen, so bench/windows/paste.ts
# runs this as a task there (schtasks /IT). The job (JSON, UTF-8) lists books beside it, each with its source, its
# target as selected (the target and one more column or row) and the key that shrinks the selection onto it, and the
# file to save. Each book is opened in an Excel of the task's own, the source selected, the Home tab's Format Painter
# button found by UI Automation and toggled on, the target sheet activated and the selection made through the object
# model, which must leave the painter on, then Excel in front and Shift with the key typed, as the painter applies at a
# change of the selection by the keyboard; the button must be off after it. With more targets the button is
# double-clicked with the mouse, which keeps the painter on, each target taken the same way and Escape turning it off.
# The book is saved as xlsx and closed.
# Output: a JSON object { version, answers: [per book { error }] } written as UTF-8 to $out.
param([Parameter(Mandatory = $true)][string]$job, [Parameter(Mandatory = $true)][string]$out)
Add-Type -AssemblyName System.Windows.Forms, UIAutomationClient, UIAutomationTypes
$UIA = [System.Windows.Automation.AutomationElement]
# the window's Format Painter button: an element of that name with a toggle
function Painter($hwnd) {
  $byName = New-Object System.Windows.Automation.PropertyCondition($UIA::NameProperty, 'Format Painter')
  $b = @($UIA::FromHandle([IntPtr]$hwnd).FindAll([System.Windows.Automation.TreeScope]::Descendants, $byName)) | Where-Object { $_.GetSupportedPatterns() -contains [System.Windows.Automation.TogglePattern]::Pattern } | Select-Object -First 1
  if (-not $b) { throw 'no Format Painter button in the window' }
  $b
}
function Toggle($b) { $b.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern) }
function IsOn($t) { [int]$t.Current.ToggleState -eq 1 }
# a double click of the left button at the middle of the element, which UI Automation's toggle does not make
function DoubleClick($b) {
  $r = $b.Current.BoundingRectangle
  [void][XlPaint.Win]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
  for ($i = 0; $i -lt 2; $i++) {
    [XlPaint.Win]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero)
    [XlPaint.Win]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 60
  }
}
# the target sheet active, the selection made through the object model, which must leave the painter on, and Shift
# with the key typed (again where it was lost, the selection as it was: the first book of a run lost it once)
function Take($xl, $wb, $t, $to, $key) {
  $wt = $wb.Worksheets.Item($to[0])
  [void]$wt.Activate()
  [void]$wt.Range($to[1]).Select()
  Start-Sleep -Milliseconds 500
  if (-not (IsOn $t)) { throw 'the Format Painter applied at a selection of the object model' }
  $sel = $xl.Selection.Address()
  for ($k = 0; $k -lt 2 -and $xl.Selection.Address() -eq $sel; $k++) {
    Front $xl.Hwnd
    [System.Windows.Forms.SendKeys]::SendWait($(if ($key -eq 'up') { '+{UP}' } else { '+{LEFT}' }))
    Start-Sleep -Milliseconds 500
  }
  $sel
}
Add-Type -Namespace XlPaint -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, uint f, UIntPtr x);
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, UIntPtr e);
'@
# Excel in front: Windows lets a process take the foreground once it sent the last input, so Shift goes down and up
# first; a window that keeps the foreground (Windows's Get Started did) stops the step
function Front($hwnd) {
  $h = [IntPtr]$hwnd
  for ($i = 0; $i -lt 3; $i++) {
    [XlPaint.Win]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
    [XlPaint.Win]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)
    [void][XlPaint.Win]::SetForegroundWindow($h)
    Start-Sleep -Milliseconds 400
    if ([XlPaint.Win]::GetForegroundWindow() -eq $h) { return }
  }
  throw "Excel did not come to the front: $([XlPaint.Win]::GetForegroundWindow()) is, not $hwnd"
}
$books = [IO.File]::ReadAllText($job, [Text.Encoding]::UTF8) | ConvertFrom-Json
$dir = Split-Path $job
$res = [ordered]@{ answers = New-Object System.Collections.ArrayList }
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  $xl.WindowState = -4137
  $res.version = [string]$xl.Version
  foreach ($b in $books) {
    $wb = $xl.Workbooks.Open((Join-Path $dir $b.file), 0, $false)
    $err = ''
    try {
      $ws = $wb.Worksheets.Item($b.from[0])
      [void]$ws.Activate()
      [void]$ws.Range($b.from[1]).Select()
      Front $xl.Hwnd
      $button = Painter $xl.ActiveWindow.Hwnd
      $t = Toggle $button
      $more = @($b.more | Where-Object { $_ })
      if ($more.Count) { DoubleClick $button } else { $t.Toggle() }
      Start-Sleep -Milliseconds 800
      if (-not (IsOn $t)) { throw 'the Format Painter did not turn on' }
      $sel = Take $xl $wb $t $b.to $b.key
      foreach ($m in $more) {
        Start-Sleep -Milliseconds 500
        if (-not (IsOn $t)) { throw "the Format Painter turned off before $($m.to[1])" }
        $sel = Take $xl $wb $t $m.to $m.key
      }
      if ($more.Count) {
        Start-Sleep -Milliseconds 500
        Front $xl.Hwnd
        [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
      }
      for ($i = 0; $i -lt 30 -and (IsOn $t); $i++) { Start-Sleep -Milliseconds 100 }
      if (IsOn $t) { throw "the Format Painter stayed on, $sel then $($xl.Selection.Address()) selected" }
    } catch { $err = 'paste: ' + [string]$_.Exception.Message + ' at line ' + [string]$_.InvocationInfo.ScriptLineNumber }
    [void]$wb.Worksheets.Item(1).Activate()
    [void]$wb.Worksheets.Item(1).Range('A1').Select()
    $wb.SaveAs((Join-Path $dir $b.out), 51)
    $wb.Close($false)
    [void]$res.answers.Add([ordered]@{ error = $err })
  }
} catch { $res.error = [string]$_.Exception.Message + ' ' + [string]$_.InvocationInfo.PositionMessage } finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
}
