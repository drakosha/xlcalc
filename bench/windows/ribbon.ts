/**
 * The ribbon's Number Format box of Excel for Windows over the ribbon cases of a ui set (`bench/coverage/ui/<set>.ts`),
 * in the books the Mac step writes (`caseFile`). Excel shows the ribbon only on a screen and Excel over SSH runs
 * without one, so `ribbon.ps1` runs as a task in the session of the machine's user (`schtasks /IT`), which must be on
 * the screen (log on through Remote Desktop first), and reads the box by UI Automation. `npx tsx bench/windows/run.ts
 * ribbon <set> [locale]` stores, for each cell, the box, the code Excel keeps (read in en-US) and its local form in
 * `test/excel/win/ribbon/[<tag>/]<set>.json`; `test/format-category-win.test.ts` compares `numberFormatCategory` with
 * it. The cases of `choose`, an item of the list applied, are the Mac's alone. Excel takes the whole run in front of the
 * session, about a second and a half a cell.
 */
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { localeDir } from '../coverage/locales.js'
import { secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { caseFile } from '../coverage/ui.js'
import { readCells, type UiCase, type UiSet } from '../coverage/ui/cases.js'
import { HOST, metaOf, type Tag, type WinEnv } from './remote.js'
import { winFile } from './functions.js'
import { runOnScreen, scp } from './screen.js'

/** What the box showed for a cell, the code Excel keeps for its format (en-US) and that code in the locale. */
export interface RibbonBox {
  box: string
  code: string
  local: string
}

export interface StoredRibbon {
  meta: Record<string, string | number>
  cases: UiCase[]
  answers: Record<string, Record<string, RibbonBox>>
}

/** The cases of a ui set's module that read the box. */
export const ribbonCases = (set: UiSet): UiCase[] => set.buildUiCases().filter((c) => c.act.do === 'ribbon')

export async function stepRibbon(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const t0 = Date.now()
  const cases = ribbonCases(await setModule<UiSet>('ui', set))
  const dir = mkdtempSync(join(tmpdir(), 'xlwin-ribbon-'))
  for (const c of cases) writeFileSync(join(dir, `${c.id}.xlsx`), await caseFile(c, `${set} ${c.id}`))
  writeFileSync(join(dir, 'job.json'), JSON.stringify(cases.map((c) => ({ file: `${c.id}.xlsx`, cells: readCells(c) }))))
  const cells = cases.reduce((n, c) => n + readCells(c).length, 0)
  await runOnScreen('ribbon', 'ribbon.ps1', dir, 60 + 10 * cases.length + 4 * cells)
  scp(`${HOST}:C:/xlcalc/ribbon/out.json`, join(dir, 'out.json'))
  const raw = JSON.parse(readFileSync(join(dir, 'out.json'), 'utf8').replace(/^\uFEFF/, '')) as { version: string; ui: number; box: string; error?: string; answers: string[][][] }
  if (raw.error) throw new Error(`ribbon: ${raw.error}`)
  if (raw.answers.length !== cases.length) throw new Error(`ribbon: ${raw.answers.length} books answered of ${cases.length}`)
  const answers = Object.fromEntries(cases.map((c, i) => [c.id, Object.fromEntries(readCells(c).map((x, j) => [x, { box: raw.answers[i][j][0], code: raw.answers[i][j][1], local: raw.answers[i][j][2] }]))]))
  const stored: StoredRibbon = { meta: { ...metaOf(e, tag), box: raw.box, date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, cases, answers }
  writeAnswers(winFile(`ribbon/${localeDir(tag)}${set}.json`), stored)
  console.log(`ribbon ${set} ${tag}: ${cases.length} books, ${Object.values(answers).reduce((n, a) => n + Object.keys(a).length, 0)} cells`)
}
