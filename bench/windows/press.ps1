# The ribbon's Home tab buttons of Excel for Windows pressed by UI Automation: Excel has a ribbon only in a session on
# a screen, so bench/windows/ui.ts runs this as a task there (schtasks /IT). The job (JSON, UTF-8) lists books beside
# it, each with the range to select on its sheet S (its first cell the active one), the buttons by their names in the
# interface's language, whether Return follows them (AutoSum on one cell opens its editing) and the file to save. Each
# book is opened, the range selected through the object model, Excel taken to the front, each button invoked in turn
# (the first element of its name with the Invoke pattern), Return typed after them, and the book saved as xlsx.
# Output: a JSON object { version, answers: [per book { error }] } written as UTF-8 to $out; a button not found names
# the buttons the window has.
param([Parameter(Mandatory = $true)][string]$job, [Parameter(Mandatory = $true)][string]$out)
Add-Type -AssemblyName System.Windows.Forms, UIAutomationClient, UIAutomationTypes
. (Join-Path $PSScriptRoot 'front.ps1')
$UIA = [System.Windows.Automation.AutomationElement]
$Invoke = [System.Windows.Automation.InvokePattern]::Pattern
function Press($hwnd, $name) {
  $all = @($UIA::FromHandle([IntPtr]$hwnd).FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition))
  $b = $all | Where-Object { $_.Current.Name -eq $name -and $_.GetSupportedPatterns() -contains $Invoke } | Select-Object -First 1
  if (-not $b) {
    $names = $all | Where-Object { $_.GetSupportedPatterns() -contains $Invoke } | ForEach-Object { $_.Current.Name } | Where-Object { $_ } | Select-Object -Unique
    throw "no button $name; the window's buttons: $($names -join ' | ')"
  }
  $b.GetCurrentPattern($Invoke).Invoke()
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
      $ws = $wb.Worksheets.Item('S')
      [void]$ws.Activate()
      [void]$ws.Range($b.range).Select()
      Start-Sleep 1
      foreach ($n in @($b.buttons)) {
        Front $xl.Hwnd
        Press $xl.ActiveWindow.Hwnd $n
        Start-Sleep 1
      }
      if ($b.enter) {
        Front $xl.Hwnd
        [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
        Start-Sleep 1
      }
    } catch { $err = 'press: ' + [string]$_.Exception.Message + ' at line ' + [string]$_.InvocationInfo.ScriptLineNumber }
    $wb.SaveAs((Join-Path $dir $b.out), 51)
    $wb.Close($false)
    [void]$res.answers.Add([ordered]@{ error = $err })
  }
} catch { $res.error = [string]$_.Exception.Message + ' ' + [string]$_.InvocationInfo.PositionMessage } finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
}
