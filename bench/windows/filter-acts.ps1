# The filter-acts sets on the screen: Excel has its buttons only in a session on a screen, so bench/windows/filter-acts.ts
# runs this as a task there (schtasks /IT). The job (JSON, UTF-8) lists books beside it, each with its steps (`actSpec`).
# Each book is opened and computed, the sheet read, each step taken and the sheet read again, and the book saved as xlsx.
# A column's button is Alt+Down on its header cell, with O for Sort Largest to Smallest; Data > Filter is Ctrl+Shift+L with
# the cell or the range selected; the rest goes through the object model. A file Excel opens repaired is that book's error.
# Output: a JSON object { version, answers: [per book { error, dialog, opened, after }] } written as UTF-8 to $out.
param([Parameter(Mandatory = $true)][string]$job, [Parameter(Mandatory = $true)][string]$out)
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'front.ps1')
# Data > Filter: Ctrl+Shift+L by the keyboard's own calls, as SendKeys loses the modifiers of this one and types an l
function FilterKey {
  Front $xl.Hwnd
  foreach ($k in 0x11, 0x10, 0x4C) { [XlScreen.Win]::keybd_event([byte]$k, 0, 0, [UIntPtr]::Zero) }
  foreach ($k in 0x4C, 0x10, 0x11) { [XlScreen.Win]::keybd_event([byte]$k, 0, 2, [UIntPtr]::Zero) }
  Start-Sleep -Milliseconds 1500
  # a window Excel brought waits for an answer, and no call goes through meanwhile
  for ($i = 0; $i -lt 3; $i++) {
    $ready = $false
    try { $ready = [bool]$xl.Ready } catch {}
    if ($ready) { return }
    $script:dialog += 'a window after Data > Filter; '
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep 1
  }
}
# until Excel takes calls again: a sort or a filter on the screen may still be working
function WaitReady {
  for ($i = 0; $i -lt 40; $i++) {
    try { if ([bool]$xl.Ready) { return } } catch {}
    Start-Sleep -Milliseconds 500
  }
  throw 'Excel did not become ready'
}
function FilterOf($ws, $b) { if ($b.table) { $ws.ListObjects.Item(1).AutoFilter } else { $ws.AutoFilter } }
function HeaderCell($ws, $b, $col) { $ws.Range($col + [string](FilterOf $ws $b).Range.Row) }
function Step($ws, $b, $s) {
  $names = $s.PSObject.Properties.Name
  if ($names -contains 'criteria') {
    $rg = (FilterOf $ws $b).Range
    $field = $ws.Range($s.column + '1').Column - $rg.Column + 1
    [void]$rg.AutoFilter([int]$field, [string]$s.criteria)
  } elseif ($names -contains 'sort') {
    if ($s.via -eq 'button') {
      [void](HeaderCell $ws $b $s.sort).Select()
      # the button's list is open between the two: Excel takes no call meanwhile, Front among them
      Front $xl.Hwnd
      [System.Windows.Forms.SendKeys]::SendWait('%{DOWN}')
      Start-Sleep -Milliseconds 1500
      [System.Windows.Forms.SendKeys]::SendWait('o')
      Start-Sleep -Milliseconds 2000
    } else {
      $f = FilterOf $ws $b
      $key = $ws.Range($s.sort + [string]$f.Range.Row + ':' + $s.sort + [string]($f.Range.Row + $f.Range.Rows.Count - 1))
      $so = if ($b.table) { $ws.ListObjects.Item(1).Sort } else { $f.Sort }
      [void]$so.SortFields.Clear()
      [void]$so.SortFields.Add($key, 0, 2)
      [void]$so.Apply()
    }
  } elseif ($names -contains 'off') {
    if ($s.off -eq 'button') {
      [void]$ws.Range('A1').Select()
      FilterKey
    } else { $ws.AutoFilterMode = $false }
  } elseif ($names -contains 'over') {
    $rg = $ws.Range($s.over)
    if ($s.via -eq 'button') {
      [void]$rg.Select()
      FilterKey
    } elseif ($s.via -eq 'field') { [void]$rg.AutoFilter(1) } else { [void]$rg.AutoFilter() }
  } elseif ($names -contains 'set') {
    $ws.Range($s.at).Formula = [string]$s.set
  } elseif ($names -contains 'totals') {
    $t = $ws.ListObjects.Item(1)
    $t.ShowTotals = $true
    $t.ListColumns.Item(3).TotalsCalculation = 1
  }
}
function State($ws, $b) {
  $range = ''
  if ($b.table) { $range = $ws.ListObjects.Item(1).AutoFilter.Range.Address($false, $false) }
  elseif ($ws.AutoFilterMode) { $range = $ws.AutoFilter.Range.Address($false, $false) }
  $hidden = New-Object System.Collections.ArrayList
  $column = New-Object System.Collections.ArrayList
  $heights = @{}
  $total = $null
  for ($r = 1; $r -le [int]$b.rows; $r++) {
    $row = $ws.Rows.Item($r)
    if ($row.Hidden) { [void]$hidden.Add($r) }
    if (@($b.watch) -contains $r) { $heights[[string]$r] = [double]$row.RowHeight }
    $c = $ws.Range("C$r")
    if ($r -gt 1) { [void]$column.Add([string]$c.Text) }
    if (-not $total -and [string]$ws.Range("A$r").Text -ceq $b.label) { $total = [ordered]@{ row = $r; text = [string]$c.Text; formula = [string]$c.Formula } }
  }
  [ordered]@{ range = [string]$range; filterMode = [bool]$ws.FilterMode; hidden = $hidden; heights = $heights; standard = [double]$ws.StandardHeight; total = $total; column = $column }
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
    $script:dialog = ''
    $err = ''
    $opened = $null
    $after = New-Object System.Collections.ArrayList
    $wb = $null
    try {
      $wb = $xl.Workbooks.Open((Join-Path $dir $b.file), 0, $false)
      if ([string]$xl.Caption -match 'Repaired') { throw "Excel repaired the file: $($xl.Caption)" }
      $ws = $wb.Worksheets.Item('S')
      [void]$ws.Activate()
      $xl.CalculateFull()
      $opened = State $ws $b
      foreach ($s in @($b.steps)) {
        Step $ws $b $s
        WaitReady
        [void]$after.Add((State $ws $b))
      }
      $wb.SaveAs((Join-Path $dir $b.out), 51)
    } catch { $err = 'step: ' + [string]$_.Exception.Message + ' at line ' + [string]$_.InvocationInfo.ScriptLineNumber }
    try { if ($wb) { $wb.Close($false) } } catch { $err += ' close: ' + [string]$_.Exception.Message }
    [void]$res.answers.Add([ordered]@{ error = $err; dialog = $script:dialog; opened = $opened; after = $after })
  }
} catch { $res.error = [string]$_.Exception.Message + ' ' + [string]$_.InvocationInfo.PositionMessage } finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding $false))
}
