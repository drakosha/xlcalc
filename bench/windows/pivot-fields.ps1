# The PivotTable Fields list of Excel for Windows driven by UI Automation: the list is there only in a session on a
# screen, so bench/windows/pivot-fields.ts runs this as a task there (schtasks /IT, screen.ts). The job (JSON, UTF-8)
# lists books beside it, each with its source (R1C1 text on D) and the fields to place in turn, by name. Each book is
# opened, a PivotTable made at P!A3 over the source through the object model, its cell selected so that Excel shows the
# list, and each field's box in the list checked by UI Automation (its Toggle pattern), or, for one to the columns
# (`to`), the field's menu opened with the mouse's right button and its Add to Column Labels invoked; then the book saved. No call of the object
# model is made while a window of Excel's other than the list is open (a modal window blocks every call): such a window
# is read and closed with Escape, and the book's step fails; a book with `answer` has it answered (`yes` with Return,
# `no` with Escape) and goes on, its boxes clicked with the mouse then, so that no call waits on the window. Before the
# PivotTable is made, the book's `cells` (values, formulas, number formats), `merge` and `table` are put on P through
# the object model. Each wait has its own limit, and a watchdog stops this script's Excel when the whole job runs past
# its limit (a minute and 40 seconds a book). With `discover` the elements of Excel's windows that name a field are listed instead.
# Output: { version, answers: [per book { error, dialog, found }] } as UTF-8 JSON in $out.
param([Parameter(Mandatory = $true)][string]$job, [Parameter(Mandatory = $true)][string]$out)
Add-Type -AssemblyName System.Windows.Forms, UIAutomationClient, UIAutomationTypes
. (Join-Path $PSScriptRoot 'front.ps1')
$UIA = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]
$All = [System.Windows.Automation.Condition]::TrueCondition
$Toggle = [System.Windows.Automation.TogglePattern]::Pattern
$Invoke = [System.Windows.Automation.InvokePattern]::Pattern

# the top windows of Excel's process: its own and any other (a dialog, a floating pane)
function TopWindows {
  $byPid = New-Object System.Windows.Automation.PropertyCondition($UIA::ProcessIdProperty, $proc)
  @($UIA::RootElement.FindAll($Scope::Children, $byPid))
}
# windows of the process that wait for an answer: any top window but Excel's own and the field list floating
function Modals {
  @(TopWindows | Where-Object { $_.Current.NativeWindowHandle -ne $hwnd -and $_.Current.Name -notmatch 'PivotTable Fields' })
}
function Texts($e) {
  $byText = New-Object System.Windows.Automation.PropertyCondition($UIA::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text)
  "$($e.Current.Name): " + ((@($e.FindAll($Scope::Descendants, $byText)) | ForEach-Object { $_.Current.Name } | Where-Object { $_ }) -join ' ')
}
# a modal window read and answered with `key` (Escape without it; at most three tries); true when one was there
function Closed($key = '{ESC}') {
  $m = Modals
  if (-not $m.Count) { return $false }
  $script:dialog += (($m | ForEach-Object { Texts $_ }) -join ' | ')
  for ($k = 0; $k -lt 3 -and (Modals).Count; $k++) {
    [System.Windows.Forms.SendKeys]::SendWait($key)
    Start-Sleep 1
  }
  $true
}
# a click of the left button at the middle of an element
function Click($e) {
  $r = $e.Current.BoundingRectangle
  [void][XlScreen.Win]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
  Start-Sleep -Milliseconds 300
  [XlScreen.Win]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
  [XlScreen.Win]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
}
# the elements of Excel's windows named one of `names`, every one when `names` is empty
function Named($names) {
  foreach ($w in TopWindows) {
    foreach ($e in @($w.FindAll($Scope::Descendants, $All))) {
      $n = $e.Current.Name
      if ($names.Count -eq 0 -or $names -contains $n) { $e }
    }
  }
}
# the field's box in the list, waited for up to `seconds`
function FieldBox($name, $seconds) {
  $until = (Get-Date).AddSeconds($seconds)
  do {
    $b = @(Named @($name) | Where-Object { $_.GetSupportedPatterns() -contains $Toggle }) | Select-Object -First 1
    if ($b) { return $b }
    Start-Sleep -Milliseconds 500
  } while ((Get-Date) -lt $until)
  throw "no box of the field $name in the field list"
}
# the menu item named `name` of the menu Excel opened, waited for up to `seconds`
function MenuItem($name, $seconds) {
  $until = (Get-Date).AddSeconds($seconds)
  do {
    $m = @(Named @($name) | Where-Object { $_.GetSupportedPatterns() -contains $Invoke }) | Select-Object -First 1
    if ($m) { return $m }
    Start-Sleep -Milliseconds 500
  } while ((Get-Date) -lt $until)
  throw "no menu item $name"
}
# the field's menu opened by the right button on its name, right of its box, and its item to the columns invoked
function ToColumns($box) {
  $r = $box.Current.BoundingRectangle
  [void][XlScreen.Win]::SetCursorPos([int]($r.X + $r.Width + 30), [int]($r.Y + $r.Height / 2))
  Start-Sleep -Milliseconds 300
  [XlScreen.Win]::mouse_event(0x0008, 0, 0, 0, [UIntPtr]::Zero)
  [XlScreen.Win]::mouse_event(0x0010, 0, 0, 0, [UIntPtr]::Zero)
  (MenuItem 'Add to Column Labels' 10).GetCurrentPattern($Invoke).Invoke()
}
function Describe($e) {
  $r = $e.Current.BoundingRectangle
  "$($e.Current.Name) | $($e.Current.ControlType.ProgrammaticName) | $($e.Current.ClassName) | $(@($e.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName }) -join ',') | $([int]$r.X),$([int]$r.Y),$([int]$r.Width),$([int]$r.Height)"
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
  $hwnd = [int64]$xl.Hwnd
  $proc = (Get-Process EXCEL | Where-Object { [int64]$_.MainWindowHandle -eq $hwnd } | Select-Object -First 1).Id
  $limit = 60 + 40 * @($books).Count
  Start-Process powershell -WindowStyle Hidden -ArgumentList "-NoProfile -Command `"Start-Sleep $limit; if (-not (Test-Path '$out')) { Stop-Process -Id $proc -Force }`""
  foreach ($b in $books) {
    $script:dialog = ''
    $found = @()
    $err = ''
    $wb = $xl.Workbooks.Open((Join-Path $dir $b.file), 0, $false)
    try {
      $ws = $wb.Worksheets.Item('P')
      foreach ($c in @($b.cells)) {
        if ($null -eq $c) { continue }
        $rg = $ws.Range([string]$c.cell)
        if ($null -ne $c.format) { $rg.NumberFormat = [string]$c.format }
        if ($null -ne $c.formula) { $rg.Formula = [string]$c.formula } elseif ($null -ne $c.value) { $rg.Value2 = $c.value }
      }
      if ($b.merge) { [void]$ws.Range([string]$b.merge).Merge() }
      if ($b.table) { [void]$ws.ListObjects.Add(1, $ws.Range([string]$b.table), [Type]::Missing, 1) }
      $pc = $wb.PivotCaches().Create(1, [string]$b.source, 8)
      [void]$pc.CreatePivotTable($ws.Range('A3'), 'PT1', $true, 8)
      [void]$ws.Activate()
      [void]$ws.Range('A3').Select()
      $wb.ShowPivotTableFieldList = $true
      Front $xl.Hwnd
      Start-Sleep 2
      if ($b.discover) {
        $found = @(Named @($b.fields | ForEach-Object { [string]$_.name }) | ForEach-Object { Describe $_ })
        # the areas of the list, where a field is dragged to
        $found += @(Named @('Columns', 'Rows', 'Values', 'Filters', 'Column Labels', 'Row Labels', 'Report Filter', 'Drag fields between areas below:') | ForEach-Object { Describe $_ })
      }
      else {
        foreach ($f in @($b.fields)) {
          Front $xl.Hwnd
          $box = FieldBox ([string]$f.name) 15
          if ($f.to -eq 'columns') { ToColumns $box } elseif ($b.answer) { Click $box } else { $box.GetCurrentPattern($Toggle).Toggle() }
          Start-Sleep 3
          if ($b.answer) { [void](Closed $(if ($b.answer -eq 'yes') { '{ENTER}' } else { '{ESC}' })) }
          elseif (Closed) { throw "a window came after placing $($f.name)" }
        }
      }
    } catch { $err = [string]$_.Exception.Message + ' at line ' + [string]$_.InvocationInfo.ScriptLineNumber }
    if (-not (Closed)) {
      try { $wb.SaveAs((Join-Path $dir $b.out), 51) } catch { $err += ' save: ' + [string]$_.Exception.Message }
    }
    if (-not (Closed)) { try { $wb.Close($false) } catch { $err += ' close: ' + [string]$_.Exception.Message } }
    [void]$res.answers.Add([ordered]@{ error = $err; dialog = $script:dialog; found = $found })
  }
} catch { $res.error = [string]$_.Exception.Message + ' ' + [string]$_.InvocationInfo.PositionMessage } finally {
  if (-not (Closed)) { $xl.Quit() }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
}
