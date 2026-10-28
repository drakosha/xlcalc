# The Format Painter of Excel for Windows: the painter works only in a session on a screen, so bench/windows/paste.ts
# runs this as a task there (schtasks /IT). The job (JSON, UTF-8) lists books beside it, each with its source, its
# target as selected (the target and one more column or row) and the key that shrinks the selection onto it, and the
# file to save. Each book is opened in an Excel of the task's own, the source selected, the Home tab's Format Painter
# button found by UI Automation and toggled on, the target sheet activated and the selection made through the object
# model, which must leave the painter on, then Excel in front and Shift with the key typed, as the painter applies at a
# change of the selection by the keyboard; the button must be off after it. The book is saved as xlsx and closed.
# Output: a JSON object { version, answers: [per book { error }] } written as UTF-8 to $out.
param([Parameter(Mandatory = $true)][string]$job, [Parameter(Mandatory = $true)][string]$out)
Add-Type -AssemblyName System.Windows.Forms, UIAutomationClient, UIAutomationTypes
$UIA = [System.Windows.Automation.AutomationElement]
# the window's Format Painter button: an element of that name with a toggle
function Painter($hwnd) {
  $byName = New-Object System.Windows.Automation.PropertyCondition($UIA::NameProperty, 'Format Painter')
  $b = @($UIA::FromHandle([IntPtr]$hwnd).FindAll([System.Windows.Automation.TreeScope]::Descendants, $byName)) | Where-Object { $_.GetSupportedPatterns() -contains [System.Windows.Automation.TogglePattern]::Pattern } | Select-Object -First 1
  if (-not $b) { throw 'no Format Painter button in the window' }
  $b.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
}
function IsOn($t) { [int]$t.Current.ToggleState -eq 1 }
Add-Type -Namespace XlPaint -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
[DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
[DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
[DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
[DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, uint f, UIntPtr x);
'@
$shell = New-Object -ComObject WScript.Shell
# Excel in front, by the first way Windows lets a background process take the foreground: after an input of its own,
# by the process's activation, by a restore, by joining the input of the window in front; the way named in the answer
function Front($hwnd) {
  $h = [IntPtr]$hwnd
  $tries = @(
    { [XlPaint.Win]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero); [XlPaint.Win]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero); [void][XlPaint.Win]::SetForegroundWindow($h) },
    { $p = 0; [void][XlPaint.Win]::GetWindowThreadProcessId($h, [ref]$p); [void]$shell.AppActivate([int]$p) },
    { [void][XlPaint.Win]::ShowWindow($h, 6); Start-Sleep -Milliseconds 300; [void][XlPaint.Win]::ShowWindow($h, 3) },
    {
      $p = 0
      $fg = [XlPaint.Win]::GetWindowThreadProcessId([XlPaint.Win]::GetForegroundWindow(), [ref]$p)
      $me = [XlPaint.Win]::GetCurrentThreadId()
      [void][XlPaint.Win]::AttachThreadInput($me, $fg, $true)
      [void][XlPaint.Win]::BringWindowToTop($h)
      [void][XlPaint.Win]::SetForegroundWindow($h)
      [void][XlPaint.Win]::AttachThreadInput($me, $fg, $false)
    }
  )
  for ($i = 0; $i -lt $tries.Count; $i++) {
    & $tries[$i]
    Start-Sleep -Milliseconds 400
    if ([XlPaint.Win]::GetForegroundWindow() -eq $h) { $script:front = $i; return }
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
    $script:front = -1
    try {
      $ws = $wb.Worksheets.Item($b.from[0])
      [void]$ws.Activate()
      [void]$ws.Range($b.from[1]).Select()
      Front $xl.Hwnd
      $t = Painter $xl.ActiveWindow.Hwnd
      $t.Toggle()
      Start-Sleep -Milliseconds 800
      if (-not (IsOn $t)) { throw 'the Format Painter did not turn on' }
      $wt = $wb.Worksheets.Item($b.to[0])
      [void]$wt.Activate()
      [void]$wt.Range($b.to[1]).Select()
      Start-Sleep -Milliseconds 500
      if (-not (IsOn $t)) { throw 'the Format Painter applied at a selection of the object model' }
      Front $xl.Hwnd
      [System.Windows.Forms.SendKeys]::SendWait($(if ($b.key -eq 'up') { '+{UP}' } else { '+{LEFT}' }))
      Start-Sleep 1
      if (IsOn $t) { throw 'the Format Painter stayed on' }
    } catch { $err = 'paste: ' + [string]$_.Exception.Message + ' at line ' + [string]$_.InvocationInfo.ScriptLineNumber }
    [void]$wb.Worksheets.Item(1).Activate()
    [void]$wb.Worksheets.Item(1).Range('A1').Select()
    $wb.SaveAs((Join-Path $dir $b.out), 51)
    $wb.Close($false)
    [void]$res.answers.Add([ordered]@{ error = $err; front = $script:front })
  }
} catch { $res.error = [string]$_.Exception.Message + ' ' + [string]$_.InvocationInfo.PositionMessage } finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
}
