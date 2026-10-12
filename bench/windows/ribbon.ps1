# The ribbon's Number Format box of Excel for Windows, read by UI Automation: Excel has a screen only in a logged-on
# session, so bench/windows/ribbon.ts runs this as a task there (schtasks /IT). For each code of the cases (a JSON
# array of strings, en-US syntax) cell A1 takes it as NumberFormat, by an en-US call (COM reads a code by the call's
# locale), B1 is selected and then A1, and the box's value is read with the code read back and NumberFormatLocal.
# Output: a JSON object { version, ui, box, answers: [[box value, code, local] or { refused }] } written as UTF-8 to $out.
param([Parameter(Mandatory = $true)][string]$cases, [Parameter(Mandatory = $true)][string]$out)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$A = [System.Windows.Automation.AutomationElement]
$en = [Globalization.CultureInfo]'en-US'
$codes = [IO.File]::ReadAllText($cases, [Text.Encoding]::UTF8) | ConvertFrom-Json
function Code($cell) { [System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::GetProperty, $null, $cell, $null, $null, $en, $null) }
$res = [ordered]@{ answers = @() }
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $true
  $xl.WindowState = -4137
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Range('A1').Value2 = 1.5
  $ws.Range('B1').Value2 = 'x'
  [void]$ws.Activate()
  Start-Sleep 4
  $byType = New-Object System.Windows.Automation.PropertyCondition($A::ControlTypeProperty, [System.Windows.Automation.ControlType]::ComboBox)
  # the box is the combo box of the Home tab that is neither the font nor its size
  $boxes = @($A::FromHandle([IntPtr]$xl.Hwnd).FindAll([System.Windows.Automation.TreeScope]::Descendants, $byType))
  $box = $boxes[2]
  $res.version = [string]$xl.Version
  $res.ui = [int]$xl.LanguageSettings.LanguageID(2)
  $res.box = $box.Current.Name
  $res.boxes = @($boxes | ForEach-Object { $_.Current.Name })
  foreach ($c in $codes) {
    $cell = $ws.Range('A1')
    try {
      [void][System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::SetProperty, $null, $cell, @([string]$c), $null, $en, $null)
    } catch {
      $res.answers += , @{ refused = [string]$_.Exception.InnerException.Message }
      continue
    }
    [void]$ws.Range('B1').Select()
    Start-Sleep -Milliseconds 700
    [void]$ws.Range('A1').Select()
    Start-Sleep -Milliseconds 1500
    $v = $box.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value
    $res.answers += , @([string]$v, [string](Code $cell), [string]$cell.NumberFormatLocal)
  }
  $wb.Close($false)
} catch { $res.error = [string]$_.Exception.Message } finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
}
