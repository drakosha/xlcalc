# Ctrl+U on cells of each underline in Excel for Windows: Excel takes keys only in a session on a screen, so this runs
# as a task there (schtasks /IT, bench/windows/screen.ts). A new workbook: each cell of column B given an underline
# through the object model (none, single, double, single accounting, double accounting), selected, Excel taken to the
# front, Ctrl+U typed by its virtual keys (any layout) twice, Font.Underline read after each; then B12:B13 (double and
# single) selected together. The job file is not read. Output: { version, ui, answers: [[label, before, after 1, after 2]] }
# as UTF-8 in $out, Font.Underline's numbers: -4142 none, 2 single, -4119 double, 4 single accounting, 5 double accounting.
param([Parameter(Mandatory = $true)][string]$job, [Parameter(Mandatory = $true)][string]$out)
. (Join-Path $PSScriptRoot 'front.ps1')
function CtrlU {
  [XlScreen.Win]::keybd_event(0x11, 0, 0, [UIntPtr]::Zero)
  [XlScreen.Win]::keybd_event(0x55, 0, 0, [UIntPtr]::Zero)
  [XlScreen.Win]::keybd_event(0x55, 0, 2, [UIntPtr]::Zero)
  [XlScreen.Win]::keybd_event(0x11, 0, 2, [UIntPtr]::Zero)
  Start-Sleep -Milliseconds 800
}
$res = [ordered]@{ answers = New-Object System.Collections.ArrayList }
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  $xl.WindowState = -4137
  $res.version = [string]$xl.Version + ' build ' + [string]$xl.Build
  $res.ui = [int]$xl.LanguageSettings.LanguageID(2)
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Range('B1:B13').Value2 = 1234.5
  $hwnd = [int64]$xl.Hwnd
  $cases = @(@('none', 'B2', -4142), @('single', 'B4', 2), @('double', 'B6', -4119), @('single accounting', 'B8', 4), @('double accounting', 'B10', 5))
  foreach ($c in $cases) {
    $r = $ws.Range($c[1])
    $r.Font.Underline = $c[2]
    [void]$ws.Activate(); [void]$r.Select()
    Front $hwnd
    $before = [int]$r.Font.Underline
    CtrlU; $one = [int]$r.Font.Underline
    CtrlU; $two = [int]$r.Font.Underline
    [void]$res.answers.Add(@($c[0], $before, $one, $two))
  }
  $ws.Range('B12').Font.Underline = -4119
  $ws.Range('B13').Font.Underline = 2
  [void]$ws.Range('B12:B13').Select()
  Front $hwnd
  CtrlU
  [void]$res.answers.Add(@('double B12 and single B13 together', 'B12 -4119, B13 2', "B12 $([int]$ws.Range('B12').Font.Underline), B13 $([int]$ws.Range('B13').Font.Underline)", ''))
  $wb.Close($false)
} catch { $res.error = [string]$_.Exception.Message + ' ' + [string]$_.InvocationInfo.PositionMessage } finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
}
