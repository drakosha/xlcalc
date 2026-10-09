# The Windows machine's side of the harness (bench/windows/remote.ts): its settings, the lock, cleaning up Excel.
#   -do env               the settings a run depends on, and Excel's as it starts with them (JSON)
#   -do locale -arg TAG   interface language, regional format and system locale for TAG (en-US, ru-RU); a new system
#                         locale needs a restart, which the caller does
#   -do ui -arg TAG       interface language and regional format for TAG alone, the system locale (code page) as it is
#   -do lock -arg OWNER   takes the lock (C:\xlcalc\lock), prints "taken" or the owner holding it
#   -do unlock -arg OWNER drops the lock when OWNER holds it
#   -do clean             stops the Excel processes outside interactive sessions (left by a job cut short), removes the
#                         jobs' files and makes sure of the folders Excel needs to open files over SSH
param([Parameter(Mandatory = $true)][string]$do, [string]$arg = '')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$lock = 'C:\xlcalc\lock'

# Excel processes started over SSH run in session 0; an interactive (RDP, console) user's Excel is left alone
function Ours { @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq 0 }) }

switch ($do) {
  'env' {
    Add-Type -AssemblyName System.Drawing
    $o = [ordered]@{
      ui = [string](Get-ItemProperty 'HKCU:\Control Panel\Desktop' -Name PreferredUILanguages -ErrorAction SilentlyContinue).PreferredUILanguages
      culture = (Get-Culture).Name
      system = (Get-WinSystemLocale).Name
      codepage = [Text.Encoding]::Default.CodePage
      dpi = [Drawing.Graphics]::FromHwnd([IntPtr]::Zero).DpiX
      excelRunning = (Ours).Count
    }
    $xl = New-Object -ComObject Excel.Application
    try {
      $o.excelUI = $xl.LanguageSettings.LanguageID(2)
      $o.version = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').VersionToReport
      $o.build = $xl.Build
      $o.decimal = $xl.International(3)
      $o.thousands = [int][char]$xl.International(4)
      $o.list = $xl.International(5)
      $o.books = $xl.Workbooks.Count
    } finally {
      $xl.Quit()
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    $o | ConvertTo-Json -Compress
  }
  { $_ -in 'locale', 'ui' } {
    # Excel takes the Windows display language of the logon: PreferredUILanguages is written directly, as the
    # override alone stays pending until an interactive logon, while a new SSH logon reads the value written
    New-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name PreferredUILanguages -PropertyType MultiString -Value @($arg) -Force | Out-Null
    Remove-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name PreferredUILanguagesPending -ErrorAction SilentlyContinue
    Set-WinUILanguageOverride -Language $arg
    Set-Culture $arg
    if ($do -eq 'locale' -and (Get-WinSystemLocale).Name -ne $arg) { Set-WinSystemLocale $arg; 'restart' } else { 'ok' }
  }
  'lock' {
    try {
      New-Item -ItemType Directory -Path $lock | Out-Null
      Set-Content -Path "$lock\owner" -Value $arg -Encoding ASCII
      'taken'
    } catch {
      $owner = Get-Content "$lock\owner" -ErrorAction SilentlyContinue
      $since = (Get-Item $lock).CreationTimeUtc.ToString('o')
      "busy $owner since $since"
    }
  }
  'unlock' {
    if ((Get-Content "$lock\owner" -ErrorAction SilentlyContinue) -eq $arg) { Remove-Item -Recurse -Force $lock; 'released' } else { 'not ours' }
  }
  'clean' {
    # Excel started over SSH looks for the system profile's desktop and opens no file without it ("cannot access the file")
    New-Item -ItemType Directory -Force C:\Windows\System32\config\systemprofile\Desktop, C:\Windows\SysWOW64\config\systemprofile\Desktop | Out-Null
    $p = Ours
    $p | Stop-Process -Force
    Remove-Item -Recurse -Force C:\xlcalc\work\* -ErrorAction SilentlyContinue
    "stopped $($p.Count)"
  }
  default { throw "unknown -do $do" }
}
