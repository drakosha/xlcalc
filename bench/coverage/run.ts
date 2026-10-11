/**
 * The formulajs coverage measurement against Excel, and Excel's answers for the tests. Steps (each writes its JSON, see `resultFile`):
 *
 *   npx tsx bench/coverage/run.ts cases      the cases and the book (bench/coverage/out/book.xlsx)
 *   npx tsx bench/coverage/run.ts fj         formulajs through the adapter and bare
 *   npx tsx bench/coverage/run.ts lo         LibreOffice en-US
 *   npx tsx bench/coverage/run.ts excel en   Excel en-US (restarts Excel with the English interface and locale)
 *   npx tsx bench/coverage/run.ts excel ru   Excel ru-RU (restarts Excel without arguments)
 *   npx tsx bench/coverage/run.ts excel T    Excel in the locale T (a tag such as de-DE) into excel-T.json
 *   npx tsx bench/coverage/run.ts report     the comparison, the summary and the report's tables
 *   npx tsx bench/coverage/run.ts all        all of the above; Excel is left running without arguments
 *   npx tsx bench/coverage/run.ts probes     Excel en-US probes outside the measurement (probes.ts) into test/excel/probes.json
 *   npx tsx bench/coverage/run.ts probes X   the set sets/X.ts into test/excel/cases/X.json (test: test/excel-sets.test.ts)
 *   npx tsx bench/coverage/run.ts probes X ru   the same set in Excel ru-RU into test/excel/cases/ru-RU/X.json (test: test/excel-locales-sets.test.ts)
 *   npx tsx bench/coverage/run.ts probes X T    the same set in Excel in the locale T into test/excel/cases/T/X.json
 *   npx tsx bench/coverage/run.ts legacy X   the set legacy/X.ts written as formulas of the old kind into test/excel/legacy/X.json, with each formula as Excel 365 shows it
 *   npx tsx bench/coverage/run.ts struct X   the structure operations of struct/X.ts into test/excel/struct/X.json (test: test/structure-excel.test.ts)
 *   npx tsx bench/coverage/run.ts spill X    the spill cases of spill/X.ts into test/excel/spill/X.json (test: test/spill-excel.test.ts)
 *   npx tsx bench/coverage/run.ts display X [ru|T]   cells shown by display/X.ts into test/excel/display/[ru-RU/ or T/]X.json (test: test/display-excel.test.ts)
 *   npx tsx bench/coverage/run.ts xlsx X     the xlsx subpath through Excel (roundtrip | saved | colors | formats | intersect) into test/excel/xlsx/ (test: test/xlsx-excel.test.ts)
 *   npx tsx bench/coverage/run.ts workbook X   book-level features of workbook/X.ts into test/excel/workbook/X.json (test: test/workbook-excel.test.ts)
 *   npx tsx bench/coverage/run.ts tables X [ru|T]   tables and structured references of tables/X.ts into test/excel/tables/[ru-RU/ or T/]X.json (test: test/tables-excel.test.ts)
 *   npx tsx bench/coverage/run.ts parts X    notes, hyperlinks and frozen panes of parts/X.ts into test/excel/parts/X.json and X/ (test: test/sheet-parts-excel.test.ts)
 *   npx tsx bench/coverage/run.ts entry X [ru|T]    text typed into cells, entry/X.ts, into test/excel/entry/[ru-RU/ or T/]X.json (test: test/entry-excel.test.ts)
 *   npx tsx bench/coverage/run.ts cond X     conditional formats of cond/X.ts (in the set's locale) into test/excel/cond/[T/]X.json and X/ (test: test/cond-excel.test.ts)
 *   npx tsx bench/coverage/run.ts valid X [ru]    data validation and outlines of valid/X.ts into test/excel/valid/[ru-RU/]X.json and X/ (test: test/validation-excel.test.ts)
 *   npx tsx bench/coverage/run.ts hidden X   hidden sheets of hidden/X.ts into test/excel/hidden/X.json and X/ (test: test/hidden-sheets-excel.test.ts)
 *   npx tsx bench/coverage/run.ts fit X [ru|T]    column widths fitted and widened by typing, fit/X.ts, into test/excel/fit/[ru-RU/ or T/]X.json (test: test/fit-excel.test.ts)
 *   npx tsx bench/coverage/run.ts lines X   formats of rows and columns and the cells under them, lines/X.ts, into test/excel/lines/X.json and X/ (test: test/lines-excel.test.ts)
 *   npx tsx bench/coverage/run.ts ui X [ru|T]    the formula bar, a paste and Control+Return in Excel's window (Excel in front), ui/X.ts, into test/excel/ui/[ru-RU/ or T/]X.json (test: test/ui-excel.test.ts)
 *
 * A locale is `en`, `ru` or a tag (`locales.ts`); en-US's answers are at the top of their directory, another locale's
 * in a directory of its tag.
 * FCOV_ONLY=ABS,SUM limits the functions (a trial run): everything is then written to out/.
 * Excel computes in parts of up to 400 cases; after each part Excel is checked for a modal window.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { inCaseOrder, loadStored, setModule, took, writeAnswers } from './stored.js'
import { functionParts, inParts, writeBook } from './book.js'
import { buildCases, type Case, type CaseSet } from './cases.js'
import { declineRepair, dialogText, ensureExcel, excelVersion, halving, runExcel, runExcelWatched, withExcelLock } from './excel.js'
import type { FjResult, Val } from './formulajs.js'
import { libreOfficeVersion, runLibreOffice } from './libreoffice.js'
import { localeDir, localeMeta, measurementFile, parseTag } from './locales.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from './probes.js'
import { writeReport } from './report.js'
import { stepStruct } from './struct.js'
import { stepSpill } from './spill.js'
import { stepDisplay } from './display.js'
import { stepXlsx } from './xlsx.js'
import { stepWorkbook } from './workbook.js'
import { stepTables } from './tables.js'
import { stepParts } from './parts.js'
import { stepEntry } from './entry.js'
import { stepHidden } from './hidden.js'
import { stepCond } from './cond.js'
import { stepValid } from './valid.js'
import { stepFit } from './fit.js'
import { stepLines } from './lines.js'
import { stepUi } from './ui.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const ONLY = process.env.FCOV_ONLY?.split(',').filter(Boolean)
const OUT = join(HERE, 'out')
const BOOK = join(OUT, 'book.xlsx')
/** The measurement files the tests read, the cases and Excel's answers in each locale; the rest (formulajs, LibreOffice, the summary) only feed the report. */
const tested = (name: string) => name === 'cases.json' || name.startsWith('excel-')
/** A measurement file: under test/excel/coverage if the tests read it, else under bench/coverage/data; all in out/ on a trial run. */
const resultFile = (name: string) => join(ONLY ? OUT : tested(name) ? join(ROOT, 'test/excel/coverage') : join(HERE, 'data'), name)
const REPORT = ONLY ? join(OUT, 'report.md') : join(ROOT, 'docs/coverage-formulajs.md')

export interface Column<T = Val> {
  meta: Record<string, string | number | boolean>
  values: Record<string, T>
}

const save = (name: string, data: unknown) => writeAnswers(resultFile(name), data)
const load = <T>(name: string): T => JSON.parse(readFileSync(resultFile(name), 'utf8')) as T

function caseSet(): CaseSet {
  const set = buildCases()
  return ONLY ? { ...set, cases: set.cases.filter((c) => ONLY.includes(c.fn)) } : set
}

async function stepCases() {
  const set = caseSet()
  save('cases.json', set)
  mkdirSync(OUT, { recursive: true })
  await writeBook(set.cases, BOOK)
  console.log(`${set.cases.length} cases, ${new Set(set.cases.map((c) => c.fn)).size} functions, book ${BOOK}`)
}

const cases = (): Case[] => load<CaseSet>('cases.json').cases

/** formulajs in a child process: a case that hangs (over 5 s) is killed and marked. */
async function stepFj() {
  const cs = cases()
  const file = resultFile('cases.json')
  const out: Record<string, FjResult> = {}
  const t0 = Date.now()
  let start = 0
  while (start < cs.length) {
    start = await new Promise<number>((resolve) => {
      const child = spawn(process.execPath, ['--max-old-space-size=2048', '--import', 'tsx', join(HERE, 'fj-worker.ts'), file, String(start)], { stdio: ['ignore', 'pipe', 'ignore'] })
      let next = start
      let buf = ''
      let timer: NodeJS.Timeout
      let killed = false
      const arm = () => {
        clearTimeout(timer)
        timer = setTimeout(() => {
          killed = true
          child.kill('SIGKILL')
        }, 5_000)
      }
      arm()
      child.stdout.on('data', (d: Buffer) => {
        buf += d.toString('utf8')
        let nl: number
        while ((nl = buf.indexOf('\n')) >= 0) {
          const { i, r } = JSON.parse(buf.slice(0, nl)) as { i: number; r: FjResult }
          buf = buf.slice(nl + 1)
          out[cs[i].id] = r
          next = i + 1
          arm()
        }
      })
      child.on('exit', () => {
        clearTimeout(timer)
        if (next < cs.length) {
          const fail: Val = { t: 'other', v: killed ? 'hang' : 'crash' }
          out[cs[next].id] = { adapter: fail, bare: fail }
          console.log(`${killed ? 'hangs' : 'crashed'}: ${cs[next].id} ${cs[next].formula}`)
          next++
        }
        resolve(next)
      })
    })
  }
  const pkg = JSON.parse(readFileSync(join(ROOT, 'node_modules/@formulajs/formulajs/package.json'), 'utf8')) as { version: string }
  save('formulajs.json', { meta: { formulajs: pkg.version, node: process.version, tz: Intl.DateTimeFormat().resolvedOptions().timeZone, seconds: (Date.now() - t0) / 1000 }, values: out })
  console.log(`formulajs: ${cs.length} cases in ${((Date.now() - t0) / 1000).toFixed(1)} s`)
}

const byId = (cs: Case[], vals: Val[]): Record<string, Val> => Object.fromEntries(cs.map((c, i) => [c.id, vals[i]]))

async function stepLo() {
  const cs = cases()
  if (!existsSync(BOOK)) throw new Error('no book: run the cases step first')
  const t0 = Date.now()
  const vals = runLibreOffice(BOOK, cs.length)
  save('libreoffice.json', { meta: { version: libreOfficeVersion(), locale: 'en-US', seconds: (Date.now() - t0) / 1000 }, values: byId(cs, vals) })
  console.log(`LibreOffice: ${((Date.now() - t0) / 1000).toFixed(1)} s`)
}

/** What a legacy set stores besides the values: each case's formula as Excel 365 shows it (`formula2`) and its row. */
export interface LegacyAnswers {
  formulas: Record<string, string>
  rows: Record<string, number>
}

interface ExcelColumn extends Column {
  meta: Column['meta'] & { parts: number; done: number; stopped?: string; dynamic: true }
}

/**
 * Excel in parts: before and after each part Excel is checked for a window that is not a book (a modal
 * dialog). If there is one, the run stops with the dialog's text and the part's number; the parts done are
 * in the file already. A new run goes on from the first part not done (FCOV_FRESH=1 starts over).
 */
async function stepExcel(tag: string) {
  const cs = cases()
  const name = measurementFile(tag)
  const file = resultFile(name)
  await ensureExcel(tag)
  const ps = functionParts(cs)
  const stored = loadStored<ExcelColumn>(file)
  const col: ExcelColumn = stored ?? { meta: { version: excelVersion(), locale: localeMeta(tag), seconds: 0, parts: ps.length, done: 0, dynamic: true }, values: {} }
  delete col.meta.stopped
  const stop = (k: number, why: string) => {
    col.meta.stopped = `part ${k + 1} of ${ps.length} (${ps[k][0].id} … ${ps[k][ps[k].length - 1].id}): ${why}`
    save(name, col)
    throw new Error(`Excel ${tag}: ${col.meta.stopped}`)
  }
  mkdirSync(OUT, { recursive: true })
  for (let k = 0; k < ps.length; k++) {
    const part = ps[k]
    if (part.every((c) => c.id in col.values)) continue
    const before = dialogText()
    if (before) stop(k, `before the part Excel has a window that is not a book: ${before}`)
    const book = join(OUT, `part-${tag}-${k + 1}.xlsx`)
    await writeBook(part, book)
    const t0 = Date.now()
    let vals: Val[] | null = null
    let err = ''
    try {
      vals = runExcel(book, part.length)
    } catch (e) {
      err = e instanceof Error ? e.message.split('\n')[0] : String(e)
    }
    const after = dialogText()
    if (after) stop(k, `Excel modal window: ${after}${err ? `; read error: ${err}` : ''}`)
    if (!vals) stop(k, `read error: ${err}`)
    col.values = inCaseOrder(cs, { ...col.values, ...byId(part, vals!) })
    col.meta.seconds = Math.round((Number(col.meta.seconds) + (Date.now() - t0) / 1000) * 10) / 10
    col.meta.done = ps.filter((p) => p.every((c) => c.id in col.values)).length
    save(name, col)
    console.log(`Excel ${tag}: part ${k + 1}/${ps.length}, ${part.length} cases, ${took(t0)} s`)
  }
}

/**
 * Probes in Excel in parts of PART, with the same check for modal windows before and after each part. Without
 * a set name these are `probes.ts` and `test/excel/probes.json` (en-US only); with `<set>`, the generator
 * `sets/<set>.ts` (exporting `buildProbes()` and, when it needs data sheets, `SHEETS`) and the answers in
 * `test/excel/cases/<set>.json` for en-US, `test/excel/cases/<tag>/<set>.json` for another locale (`ru-RU/`). A file
 * per set lets parallel branches add their cases without touching shared files.
 */
async function stepProbes(set: string | undefined, tag: string, legacy = false) {
  if (set === undefined && tag !== 'en-US') throw new Error('the probes of probes.ts are computed in en-US only; a set in another locale: probes <set> <locale>')
  const mod = set ? await setModule<{ buildProbes: () => Probe[]; SHEETS?: SheetCells }>(legacy ? 'legacy' : 'sets', set) : { buildProbes, SHEETS: PROBE_SHEETS }
  const ps = mod.buildProbes()
  const sheets = mod.SHEETS ?? {}
  const stem = `${set ?? 'probes'}-${tag}`
  const target = !set ? join(ROOT, 'test/excel/probes.json') : join(ROOT, `test/excel/${legacy ? 'legacy' : 'cases'}/${localeDir(tag)}${set}.json`)
  // going on: answers whose formula is unchanged are taken from the file (FCOV_FRESH=1 starts over)
  const prev = loadStored<LegacyAnswers & { cases: Probe[]; values: Record<string, Val> }>(target)
  const was = new Map(prev?.cases.map((c) => [c.id, c.formula]) ?? [])
  const values: Record<string, Val> = {}
  // a legacy set: each case's formula as Excel 365 shows it and its row (implicit intersection takes the row)
  const formulas: Record<string, string> = {}
  const rows: Record<string, number> = {}
  for (const p of ps)
    if (prev?.values[p.id] && was.get(p.id) === p.formula) {
      values[p.id] = prev.values[p.id]
      if (legacy) [formulas[p.id], rows[p.id]] = [prev.formulas[p.id], prev.rows[p.id]]
    }
  const rejected: Probe[] = []
  const t0 = Date.now()
  const save = () => {
    const meta = { version: excelVersion(), locale: localeMeta(tag), seconds: (Date.now() - t0) / 1000, ...(legacy ? { legacy: true } : { dynamic: true }) }
    const more = legacy ? { formulas: inCaseOrder(ps, formulas), rows: inCaseOrder(ps, rows) } : {}
    writeAnswers(target, { meta, cases: ps, values: inCaseOrder(ps, values), ...more, ...(rejected.length ? { rejected: rejected.map((p) => p.id) } : {}) })
  }
  // a book Excel does not read is halved down to the formula at fault
  const run = async (cases: Probe[], label: string): Promise<boolean> => {
    if (declineRepair()) console.log('closed a leftover book repair window')
    const before = dialogText()
    if (before) throw new Error(`probes ${label}: Excel has a window that is not a book: ${before}`)
    const book = join(OUT, `${stem}-${label}.xlsx`)
    // the read script expects a column of two cells at least: a single case gets a pad
    const sent = cases.length === 1 ? [...cases, { id: 'pad', group: 'pad', formula: '0' }] : cases
    await writeBook(sent, book, sheets, legacy)
    const got = await runExcelWatched(book, sent.length, legacy)
    if (got === 'rejected') return false
    cases.forEach((p, i) => {
      values[p.id] = got.values[i]
      if (legacy) [formulas[p.id], rows[p.id]] = [got.formulas![i], i + 1]
    })
    return true
  }
  const refused = (p: Probe) => {
    rejected.push(p)
    console.log(`Excel does not load the formula: ${p.id} ${p.formula}`)
  }
  await ensureExcel(tag)
  mkdirSync(OUT, { recursive: true })
  for (const [k, all] of inParts(ps).entries()) {
    const part = all.filter((p) => !values[p.id])
    if (!part.length) continue
    await halving(part, String(k + 1), run, refused)
    save()
    console.log(`probes ${stem}: part ${k + 1}, ${part.length} cases, ${took(t0)} s`)
  }
  save()
  if (rejected.length) throw new Error(`Excel does not load ${rejected.length} formulas, remove them from the generator: ${rejected.map((p) => `${p.id} ${p.formula}`).join('; ')}`)
}

const step = process.argv[2]
/** The locale a step names (`excel <locale>`, `probes|display|tables|entry|valid|fit <set> <locale>`), en-US when none: read before Excel's lock is taken. */
const locale = parseTag(process.argv[step === 'excel' ? 3 : 4])
const t0 = Date.now()
if (step === 'cases') await stepCases()
else if (step === 'fj') await stepFj()
else if (step === 'lo') await stepLo()
else if (step === 'excel') await withExcelLock(() => stepExcel(locale))
else if (step === 'report') writeReport(resultFile, REPORT)
else if (step === 'probes') await withExcelLock(() => stepProbes(process.argv[3], locale))
else if (step === 'legacy') await withExcelLock(() => stepProbes(process.argv[3], 'en-US', true))
else if (step === 'struct') await withExcelLock(() => stepStruct(process.argv[3]))
else if (step === 'spill') await withExcelLock(() => stepSpill(process.argv[3]))
else if (step === 'display') await withExcelLock(() => stepDisplay(process.argv[3], locale))
else if (step === 'xlsx') await withExcelLock(() => stepXlsx(process.argv[3]))
else if (step === 'workbook') await withExcelLock(() => stepWorkbook(process.argv[3]))
else if (step === 'tables') await withExcelLock(() => stepTables(process.argv[3], locale))
else if (step === 'parts') await withExcelLock(() => stepParts(process.argv[3]))
else if (step === 'entry') await withExcelLock(() => stepEntry(process.argv[3], locale))
else if (step === 'cond') await withExcelLock(() => stepCond(process.argv[3]))
else if (step === 'valid') await withExcelLock(() => stepValid(process.argv[3], locale))
else if (step === 'hidden') await withExcelLock(() => stepHidden(process.argv[3]))
else if (step === 'lines') await withExcelLock(() => stepLines(process.argv[3]))
else if (step === 'fit') await withExcelLock(() => stepFit(process.argv[3], locale))
else if (step === 'ui') await withExcelLock(() => stepUi(process.argv[3], locale))
else if (step === 'all') {
  await stepCases()
  await stepFj()
  await stepLo()
  await withExcelLock(async () => {
    await stepExcel('en-US')
    await stepExcel('ru-RU')
  })
  writeReport(resultFile, REPORT)
} else {
  console.error('step: cases | fj | lo | excel en|ru|<tag> | report | all | probes [set [ru|<tag>]] | legacy set | struct set | spill set | display set [ru|<tag>] | workbook set | tables set [ru|<tag>] | parts set | entry set [ru|<tag>] | cond set | valid set [ru] | hidden set | fit set [ru|<tag>] | lines set | ui set [ru|<tag>] | xlsx roundtrip|saved|colors|formats|intersect')
  process.exit(2)
}
console.log(`done in ${((Date.now() - t0) / 1000).toFixed(1)} s`)
