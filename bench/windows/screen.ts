/**
 * A script of Excel run in the session of the machine's user on its screen, as a task (`schtasks /IT`): Excel over SSH
 * has no screen, and the ribbon, UI Automation and the keyboard need one. The session must be on the screen (log on
 * through Remote Desktop first).
 */
import { execFileSync } from 'node:child_process'
import { readdirSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { HOST, ssh } from './remote.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))

export const scp = (...args: string[]) => execFileSync('scp', ['-q', '-o', 'BatchMode=yes', ...args])

/** Whether a session is on the machine's screen: quser names its state in English when its output is UTF-8, Active for one on a screen, Disc for one left. */
export const onScreen = () => /\bActive\b/.test(ssh('[Console]::OutputEncoding = [Text.Encoding]::UTF8; quser 2>&1 | Out-String'))

/**
 * Runs `script` (a file beside this one) as the task `task` on the screen with the files of `dir` in `C:\xlcalc\<task>`,
 * its arguments `-job job.json -out out.json` there, until it writes `out.json` or `seconds` pass; the task's Excel
 * stopped after it where it stays without a window, the user's own Excel left alone.
 */
export async function runOnScreen(task: string, script: string, dir: string, seconds: number): Promise<void> {
  if (!onScreen()) throw new Error(`${task}: no session on a screen on the machine (log on through Remote Desktop first)`)
  const work = `C:\\xlcalc\\${task}`
  ssh(`Remove-Item -Recurse -Force ${work} -ErrorAction SilentlyContinue; New-Item -ItemType Directory ${work} | Out-Null`)
  scp(join(HERE, script), `${HOST}:C:/xlcalc/bin/`)
  scp(...readdirSync(dir).map((f) => join(dir, f)), `${HOST}:C:/xlcalc/${task}/`)
  const started = ssh('(Get-Date).ToString("o")').trim()
  ssh(
    `schtasks /Create /TN xl${task} /TR 'powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\\xlcalc\\bin\\${script} -job ${work}\\job.json -out ${work}\\out.json' /SC ONCE /ST 23:59 /IT /F | Out-Null; schtasks /Run /TN xl${task} | Out-Null`,
  )
  try {
    const until = Date.now() + seconds * 1000
    while (ssh(`Test-Path ${work}\\out.json`).trim() !== 'True') {
      if (Date.now() > until) throw new Error(`${task}: the task wrote no answer in time`)
      await new Promise((r) => setTimeout(r, 5000))
    }
  } finally {
    ssh(`schtasks /Delete /TN xl${task} /F | Out-Null`)
    // the task's Excel may stay after it quits (it did once in two runs of the ribbon): stopped when it has no window;
    // the session first, as the start of another session's process may be denied; none found is no failure
    ssh(`Start-Sleep 10; Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -ne 0 -and $_.StartTime -ge [datetime]'${started}' -and -not $_.MainWindowTitle } | Stop-Process -Force; exit 0`)
  }
}
