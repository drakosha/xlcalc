/**
 * The xlsx tests of IronCalc in Excel for Mac: each file of `ironcalc/cases.json`, from a checkout of IronCalc
 * (`IRONCALC_ROOT`, at the cases' commit), opened as it is, computed fully, and each case's cell read by a sheet the
 * script adds (`excel-ironcalc.applescript`); the book is closed unsaved. The answers go to `ironcalc/mac.json`, beside
 * the value each file stored and xlcalc's; Windows' are `ironcalc/win.json` (`bench/windows/ironcalc.ts`).
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ensureExcel, excelVersion, runScriptWatched, typedValue } from './excel.js'

const HERE = fileURLToPath(new URL('./ironcalc', import.meta.url))
const SCRIPT = fileURLToPath(new URL('./excel-ironcalc.applescript', import.meta.url))

interface Case {
  id: string
  file: string
  cell: string
}

/** A cell's reference in a formula: its sheet quoted. */
const refOf = (cell: string) => {
  const [sheet, a] = cell.split('!')
  return `'${sheet.replace(/'/g, "''")}'!${a}`
}

/** Excel's value as the cases keep values: `{ t, v }`, `t` one of number, string, boolean, e, empty. */
function asCase(type: number, num: unknown, text: unknown): { t: string; v?: unknown } {
  if (type === -5) return { t: 'empty' }
  const v = typedValue(type, num, text)
  return { t: ({ n: 'number', s: 'string', b: 'boolean' } as Record<string, string>)[v.t] ?? v.t, v: v.v }
}

export async function stepIronCalc(): Promise<void> {
  const root = process.env.IRONCALC_ROOT
  if (!root || !existsSync(join(root, 'xlsx/tests'))) throw new Error('IRONCALC_ROOT: a checkout of IronCalc at the commit of bench/coverage/ironcalc/cases.json')
  await ensureExcel('en-US')
  const { cases } = JSON.parse(readFileSync(join(HERE, 'cases.json'), 'utf8')) as { cases: Case[] }
  const out = join(HERE, 'mac.json')
  const answers: Record<string, unknown> = existsSync(out) ? (JSON.parse(readFileSync(out, 'utf8')) as { answers: Record<string, unknown> }).answers : {}
  const rejected: string[] = []
  const t0 = Date.now()
  const save = () => writeFileSync(out, JSON.stringify({ meta: { version: excelVersion(), platform: 'Mac', locale: 'en-US', seconds: Math.round((Date.now() - t0) / 1000) }, answers, ...(rejected.length ? { rejected } : {}) }, null, 1) + '\n')
  for (const f of [...new Set(cases.filter((c) => !answers[c.id]).map((c) => c.file))]) {
    const mine = cases.filter((c) => c.file === f)
    const got = await runScriptWatched(SCRIPT, join(root, 'xlsx/tests', f), mine.length, [mine.map((c) => refOf(c.cell)).join('\n')])
    if (got === 'rejected') rejected.push(f)
    else {
      // one cell's range comes as its value alone
      const [t, nums, strs] = (JSON.parse(got) as unknown[]).map((x) => (Array.isArray(x) ? x : [[x]])) as unknown[][][]
      mine.forEach((c, i) => (answers[c.id] = asCase(Number(t[i][0]), nums[i][0], strs[i][0])))
    }
    save()
    console.log(`${f}: ${mine.length} cells`)
  }
  save()
}
