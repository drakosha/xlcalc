# The ribbon's Number Format box of Excel for Windows, read by UI Automation: Excel has a screen only in a session on
# one, so bench/windows/ribbon.ts runs this as a task there (schtasks /IT). The job (JSON, UTF-8) lists books beside it,
# each with the cells to read on its sheet S. Each book is opened (links not updated, read only); A1 takes the text
# format as a control cell, as on the Mac (`excel-ui.applescript`): before each cell A1 is selected and the box must
# show A1's category, then the cell is selected and the box read once it has changed and stays, with the code Excel
# keeps (NumberFormat read by an en-US call, as COM reads a code by the call's locale) and NumberFormatLocal.
# Output: a JSON object { version, ui, box, answers: [per book [[box, code, local] per cell]] } written as UTF-8 to $out.
param([Parameter(Mandatory = $true)][string]$job, [Parameter(Mandatory = $true)][string]$out)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$A = [System.Windows.Automation.AutomationElement]
$en = [Globalization.CultureInfo]'en-US'
$books = [IO.File]::ReadAllText($job, [Text.Encoding]::UTF8) | ConvertFrom-Json
$dir = Split-Path $job
function Code($cell) { [System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::GetProperty, $null, $cell, $null, $null, $en, $null) }
# the window's Number Format box: its combo box of that name
function FindBox($hwnd) {
  $byType = New-Object System.Windows.Automation.PropertyCondition($A::ControlTypeProperty, [System.Windows.Automation.ControlType]::ComboBox)
  $script:box = @($A::FromHandle([IntPtr]$hwnd).FindAll([System.Windows.Automation.TreeScope]::Descendants, $byType)) | Where-Object { $_.Current.Name -match 'Number Format|Числовой формат' } | Select-Object -First 1
  if (-not $script:box) { throw 'no Number Format box in the window' }
}
function Now { [string]$script:box.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value }
# what the box shows once it has changed from $was (two seconds at most) and stays so
function After($was) {
  $v = $was
  for ($i = 0; $i -lt 20 -and $v -eq $was; $i++) { Start-Sleep -Milliseconds 100; $v = Now }
  for ($i = 0; $i -lt 10; $i++) { Start-Sleep -Milliseconds 300; $again = Now; if ($again -eq $v) { break }; $v = $again }
  $v
}
# A1 selected until the box shows $ctl, three seconds at most, the box looked for again twice
function Control($ws, $ctl, $hwnd) {
  for ($k = 0; $k -lt 3; $k++) {
    [void]$ws.Range('A1').Select()
    for ($i = 0; $i -lt 30; $i++) { Start-Sleep -Milliseconds 100; if ((Now) -eq $ctl) { return } }
    FindBox $hwnd
  }
  throw "the Number Format box does not follow the selection: it shows $(Now), not $ctl"
}
$res = [ordered]@{ answers = New-Object System.Collections.ArrayList }
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  $xl.WindowState = -4137
  $res.version = [string]$xl.Version
  $res.ui = [int]$xl.LanguageSettings.LanguageID(2)
  foreach ($b in $books) {
    $wb = $xl.Workbooks.Open((Join-Path $dir $b.file), 0, $true)
    $ws = $wb.Worksheets.Item('S')
    [void]$ws.Activate()
    $hwnd = $xl.ActiveWindow.Hwnd
    Start-Sleep 2
    FindBox $hwnd
    $res.box = $script:box.Current.Name
    $ws.Range('A1').NumberFormat = '@'
    # A1's category: what the box changes to from the empty A2's
    [void]$ws.Range('A2').Select()
    Start-Sleep 1
    $was = Now
    [void]$ws.Range('A1').Select()
    $ctl = After $was
    if ($ctl -eq $was) { throw "the Number Format box does not follow the selection: A1 and A2 show $ctl" }
    $cells = New-Object System.Collections.ArrayList
    foreach ($a in @($b.cells)) {
      Control $ws $ctl $hwnd
      $cell = $ws.Range($a)
      [void]$cell.Select()
      [void]$cells.Add(@((After $ctl), [string](Code $cell), [string]$cell.NumberFormatLocal))
    }
    [void]$res.answers.Add($cells)
    $wb.Close($false)
  }
} catch { $res.error = [string]$_.Exception.Message + ' ' + [string]$_.InvocationInfo.PositionMessage } finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
}
