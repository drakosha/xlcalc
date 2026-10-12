/**
 * The ribbon's Number Format box of Excel for Windows (`ribbon.ps1`), which Excel shows only on a screen: Excel over SSH
 * runs without one, so the step runs as a task in a logged-on session of the machine's user (`schtasks /IT`; log on
 * through Remote Desktop first) and reads the box by UI Automation. `npx tsx bench/windows/run.ts ribbon [locale]`
 * sends the codes of `RIBBON_CODES` and stores, for each, the box's value, the code read back and its local form in
 * `test/excel/win/ribbon/[<tag>/]box.json`; `test/format-category-win.test.ts` compares `numberFormatCategory` with
 * it. Excel takes the whole run, about two seconds a code, in front of the session.
 */
import { execFileSync } from 'node:child_process'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { localeDir } from '../coverage/locales.js'
import { secondsSince, writeAnswers } from '../coverage/stored.js'
import { HOST, metaOf, ssh, type Tag, type WinEnv } from './remote.js'
import { winFile } from './functions.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))

const ruble = '[$\u20BD-419]'

/** The ruble as ru-RU's own formats write it, quoted after the amount (ru-RU's Currency item, its built-in money formats). */
const rub = '"\u20BD"'

/**
 * The codes the box is read for: each category's forms in en-US syntax, those Excel for Mac and Windows part on, dates
 * in both orders and Excel for Mac's own short date, negatives with a minus as ru-RU's built-in formats write them, the
 * ruble bracketed and quoted, ru-RU's built-in money formats (7, 8, 43, 44).
 */
export const RIBBON_CODES: readonly string[] = [
  'General', '0', '0.00', '0.000', '#,##0', '#,##0.00', '0.0', '0.00;[Red]0.00', '[Red]0.00', '#,##0.00_);[Red](#,##0.00)',
  '#,##0.00;-#,##0.00', '#,##0.00;[Red]-#,##0.00',
  '0%', '0.00%', '0.0E+00', '0.00E+00', '0E+00', '# ?/?', '# ??/??', '# ?/4', '@',
  'm/d/yyyy', 'm/d/yy', 'dd/mm/yyyy', 'dd.mm.yyyy', 'dd/mm/yyyy;@', 'd-mmm-yy', 'mmm-yy', 'd mmmm yyyy;@', '[$-F800]dddd, mmmm dd, yyyy',
  '[$-F400]h:mm:ss AM/PM', 'h:mm', 'h:mm;@', 'h:mm:ss;@', 'dd/mm/yyyy h:mm',
  '"$"#,##0.00', '$#,##0.00', `#,##0.00 ${ruble}`, `#,##0.00\\ ${ruble}`, `#,##0.00\\ ${ruble};[Red]\\-#,##0.00\\ ${ruble}`, '[$€-407]#,##0.00',
  `_-* #,##0.00\\ ${ruble}_-;\\-* #,##0.00\\ ${ruble}_-;_-* "-"??\\ ${ruble}_-;_-@_-`, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  `#,##0.00\\ ${rub}`, `#,##0.00 ${rub};-#,##0.00 ${rub}`, `#,##0.00 ${rub};[Red]-#,##0.00 ${rub}`,
  '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-', `_-* #,##0.00 ${rub}_-;-* #,##0.00 ${rub}_-;_-* "-"?? ${rub}_-;_-@_-`,
  '00000', '000\\-00\\-0000',
]

export interface StoredRibbonBox {
  meta: Record<string, string | number>
  codes: readonly string[]
  answers: ({ box: string; code: string; local: string } | { refused: string })[]
}

export async function stepRibbon(tag: Tag, e: WinEnv): Promise<void> {
  const t0 = Date.now()
  // quser names a session's state in English when its output is UTF-8: Active for one on a screen, Disc for one left
  if (!/\bActive\b/.test(ssh('[Console]::OutputEncoding = [Text.Encoding]::UTF8; quser 2>&1 | Out-String'))) throw new Error('ribbon: no session on a screen on the machine (log on through Remote Desktop first)')
  const dir = mkdtempSync(join(tmpdir(), 'xlwin-ribbon-'))
  writeFileSync(join(dir, 'codes.json'), JSON.stringify(RIBBON_CODES))
  execFileSync('scp', ['-q', '-o', 'BatchMode=yes', join(HERE, 'ribbon.ps1'), join(dir, 'codes.json'), `${HOST}:C:/xlcalc/bin/`])
  ssh(
    "Remove-Item C:\\xlcalc\\bin\\ribbon-out.json -ErrorAction SilentlyContinue; schtasks /Create /TN xlribbon /TR 'powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\\xlcalc\\bin\\ribbon.ps1 -cases C:\\xlcalc\\bin\\codes.json -out C:\\xlcalc\\bin\\ribbon-out.json' /SC ONCE /ST 23:59 /IT /F | Out-Null; schtasks /Run /TN xlribbon | Out-Null",
  )
  try {
    const until = Date.now() + (60 + 4 * RIBBON_CODES.length) * 1000
    while (ssh('Test-Path C:\\xlcalc\\bin\\ribbon-out.json').trim() !== 'True') {
      if (Date.now() > until) throw new Error('ribbon: the task wrote no answer in time')
      await new Promise((r) => setTimeout(r, 3000))
    }
  } finally {
    ssh('schtasks /Delete /TN xlribbon /F | Out-Null')
  }
  execFileSync('scp', ['-q', '-o', 'BatchMode=yes', `${HOST}:C:/xlcalc/bin/ribbon-out.json`, join(dir, 'out.json')])
  const raw = JSON.parse(readFileSync(join(dir, 'out.json'), 'utf8').replace(/^\uFEFF/, '')) as { version: string; ui: number; box: string; error?: string; answers: (string[] | { refused: string })[] }
  if (raw.error) throw new Error(`ribbon: ${raw.error}`)
  if (!/Number Format|Числовой формат/i.test(raw.box)) throw new Error(`ribbon: the box read is ${JSON.stringify(raw.box)}, not the Number Format box`)
  const answers = raw.answers.map((a) => (Array.isArray(a) ? { box: a[0], code: a[1], local: a[2] } : a))
  const stored: StoredRibbonBox = { meta: { ...metaOf(e, tag), box: raw.box, date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, codes: RIBBON_CODES, answers }
  writeAnswers(winFile(`ribbon/${localeDir(tag)}box.json`), stored)
  console.log(`ribbon ${tag}: ${answers.length} codes, ${answers.filter((a) => 'refused' in a).length} refused`)
}
