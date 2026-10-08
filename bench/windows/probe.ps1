# One probe run: UI language and regional format, Excel via COM, a few formulas, answers as JSON.
# The first probe of the machine, kept as it ran: the Office keys it sets do not change Excel's interface language,
# which follows the Windows display language (host.ps1 -do locale writes it).
param([string]$ui = 'en-us', [string]$culture = 'en-US', [string]$out = 'C:\xlprobe\out.json')
$ErrorActionPreference = 'Stop'
$lcid = @{ 'en-us' = 1033; 'ru-ru' = 1049 }[$ui]
$pol = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\LanguageResources'
New-Item -Force -Path $pol | Out-Null
Set-ItemProperty $pol -Name UILanguage -Value $lcid -Type DWord
Set-ItemProperty $pol -Name HelpLanguage -Value $lcid -Type DWord
$lr = 'HKCU:\Software\Microsoft\Office\16.0\Common\LanguageResources'
Set-ItemProperty $lr -Name FollowSystemUI -Value 'Off'
Set-ItemProperty $lr -Name UILanguageTag -Value $ui
Set-Culture $culture
$formulas = @('=CODE("Ж")', '=CODE("中")', '=CODE("é")', '=CHAR(200)', '=CHAR(198)', '=IF("TRUE",1,2)', '=LEN(TRUE)', '=1/3&""', '=ROWS(SEQUENCE(3))', '=TEXT(0.5,"0.0")')
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$res = [ordered]@{ ui = $xl.LanguageSettings.LanguageID(2); version = $xl.Version; build = $xl.Build; decimal = $xl.International(3); list = $xl.International(5); cases = @() }
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($i = 0; $i -lt $formulas.Count; $i++) { $ws.Cells.Item($i + 1, 1).Formula2 = $formulas[$i] }
$xl.CalculateFull()
for ($i = 0; $i -lt $formulas.Count; $i++) {
  $c = $ws.Cells.Item($i + 1, 1)
  $v = $c.Value2
  $res.cases += [ordered]@{ f = $formulas[$i]; value = $(if ($v -is [int] -and $v -lt -2146820000) { "err:$v" } elseif ($v -is [string]) { [string]$v + " (" + (([int[]][char[]]$v) -join ',') + ")" } else { $v }); text = $c.Text }
}
$wb.Close($false); $xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
[IO.File]::WriteAllText($out, ($res | ConvertTo-Json -Depth 4), [Text.UTF8Encoding]::new($false))
