/**
 * The books Excel for Windows saved for a pivot set (`test/excel/win/pivot/<set>/`), opened in Excel for Mac as they
 * are (no recalculation, no refresh, `excel-pivot.applescript`): what each cell of the pivot tables' areas and of the
 * watched ranges shows, and the book Excel for Mac saved. The answers go to `test/excel/pivot/<set>.json` with the
 * saved books beside them, each case with the cells where the Mac shows other text than Windows (`differs`).
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CellValue } from '../../src/contract.js'
import { cellsOf } from './case-book.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { localeMeta } from './locales.js'
import type { PivotAnswer, PivotCase } from './pivot/cases.js'
import { fileAnswers, runFileCases, secondsSince } from './stored.js'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const HERE = fileURLToPath(new URL('.', import.meta.url))

/** What the Mac shows in a case's cells (`P!A3`: [text, value]), and where its text is not Windows's (`P!A3`: [Mac, Windows]). */
export interface PivotMacAnswer {
  cells: Record<string, [text: string, value: string]>
  differs?: Record<string, [mac: string, windows: string]>
}

/** What Windows showed in a cell: the text of a pivot table's cell, or a watched cell's value as Excel's General shows it. */
function windowsText(v: CellValue): string {
  if (v === null) return ''
  if (typeof v === 'object') return { DIV0: '#DIV/0!', NA: '#N/A' }[v.error as string] ?? `#${v.error}!`
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return typeof v === 'number' ? String(Number(v.toPrecision(10))) : v
}

/** Whether the Mac's text says what Windows's does: the same text, or the same number to General's digits. */
const same = (mac: string, win: string) => mac === win || (mac !== '' && win !== '' && Number.isFinite(Number(mac)) && Math.abs(Number(mac) - Number(win)) <= 1e-9 * Math.max(1, Math.abs(Number(win))))

export async function stepPivotMac(set: string): Promise<void> {
  const winFile = join(ROOT, `test/excel/win/pivot/${set}.json`)
  const win = JSON.parse(readFileSync(winFile, 'utf8')) as { cases: PivotCase[]; answers: Record<string, PivotAnswer> }
  const cases = win.cases.filter((c) => win.answers[c.id])
  // each case's cells: the last phase's areas, then the watched cells, with what Windows showed in each
  const expected = new Map<string, Map<string, string>>()
  for (const c of cases) {
    const last = win.answers[c.id].phases.at(-1)!
    const cells = new Map<string, string>()
    for (const [pt, area] of Object.entries(last.pivots)) {
      if ('error' in area) continue
      const sheet = pt === 'PT1' ? 'P' : (c.ops.find((o) => 'create' in o && o.create === pt) as { at: string } | undefined)?.at.split('!')[0] ?? 'P'
      const { cols, addrs } = cellsOf(area.area)
      addrs.forEach((a, k) => cells.set(`${sheet}!${a}`, area.text[Math.floor(k / cols)][k % cols]))
    }
    for (const [key, v] of Object.entries(last.cells ?? {})) cells.set(key, windowsText(v))
    expected.set(c.id, cells)
  }
  const files = new Map(cases.map((c) => [c.id, new Uint8Array(readFileSync(join(ROOT, `test/excel/win/pivot/${set}/${c.id}.xlsx`)))]))
  const sent = Object.fromEntries(cases.map((c) => [c.id, `${win.answers[c.id].phases.length} ${[...expected.get(c.id)!.keys()].join(';')}`]))
  const store = fileAnswers<PivotCase, PivotMacAnswer>(join(ROOT, `test/excel/pivot/${set}.json`), cases, sent)
  const t0 = Date.now()
  await ensureExcel('en-US')
  const save = () => store.save({ version: excelVersion(), locale: localeMeta('en-US'), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, { cases })
  await runFileCases(`pivot ${set} (Mac)`, store, store.todo, 'FPIVOT_ONLY', async (c) => {
    const want = expected.get(c.id)!
    const keys = [...want.keys()]
    const r = await runSavingScript(`pivot-${c.id}`, join(HERE, 'excel-pivot.applescript'), files.get(c.id)!, [keys.join(';')])
    if ('rejected' in r) return r
    const got = JSON.parse(r.out) as [string, string][]
    const cells = Object.fromEntries(keys.map((k, i) => [k, got[i]]))
    const differs = Object.fromEntries(keys.flatMap((k, i) => (same(got[i][0], want.get(k)!) ? [] : [[k, [got[i][0], want.get(k)!]]])))
    return { answer: { cells, ...(Object.keys(differs).length ? { differs } : {}) }, saved: r.saved }
  }, save, t0)
}
