/**
 * The formulajs coverage measurement against Excel, and Excel's answers for the tests. Steps (each writes its JSON, see `resultFile`):
 *
 *   npx tsx bench/coverage/run.ts cases      the cases and the book (bench/coverage/out/book.xlsx)
 *   npx tsx bench/coverage/run.ts fj         formulajs through the adapter and bare
 *   npx tsx bench/coverage/run.ts lo         LibreOffice en-US
 *   npx tsx bench/coverage/run.ts excel en   Excel en-US (restarts Excel with the English interface and locale)
 *   npx tsx bench/coverage/run.ts excel ru   Excel ru-RU (restarts Excel without arguments)
 *   npx tsx bench/coverage/run.ts report     the comparison, the summary and the report's tables
 *   npx tsx bench/coverage/run.ts all        all of the above; Excel is left running without arguments
 *   npx tsx bench/coverage/run.ts probes     Excel en-US probes outside the measurement (probes.ts) into test/excel/probes.json
 *   npx tsx bench/coverage/run.ts probes X   the set sets/X.ts into test/excel/cases/X.json (test: test/excel-sets.test.ts)
 *   npx tsx bench/coverage/run.ts probes X ru   the same set in Excel ru-RU into test/excel/cases/ru-RU/X.json (test: test/excel-ru-probes.test.ts)
 *   npx tsx bench/coverage/run.ts struct X   the structure operations of struct/X.ts into test/excel/struct/X.json (test: test/structure-excel.test.ts)
 *   npx tsx bench/coverage/run.ts spill X    the spill cases of spill/X.ts into test/excel/spill/X.json (test: test/spill-excel.test.ts)
 *   npx tsx bench/coverage/run.ts display X [ru]   cells shown by display/X.ts into test/excel/display/[ru-RU/]X.json (test: test/display-excel.test.ts)
 *
 * FCOV_ONLY=ABS,SUM limits the functions (a trial run): everything is then written to out/.
 * Excel computes in parts of up to 400 cases; after each part Excel is checked for a modal window.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeBook } from './book.js'
import { buildCases, type Case, type CaseSet } from './cases.js'
import { declineRepair, dialogText, ensureExcel, excelVersion, runExcel, runExcelWatched, withExcelLock, type Locale } from './excel.js'
import type { FjResult, Val } from './formulajs.js'
import { libreOfficeVersion, runLibreOffice } from './libreoffice.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from './probes.js'
import { writeReport } from './report.js'
import { stepStruct } from './struct.js'
import { stepSpill } from './spill.js'
import { stepDisplay } from './display.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const ONLY = process.env.FCOV_ONLY?.split(',').filter(Boolean)
const OUT = join(HERE, 'out')
const BOOK = join(OUT, 'book.xlsx')
/** The measurement files the tests read; the rest (formulajs, LibreOffice, the summary) only feed the report. */
const TESTED = ['cases.json', 'excel-en.json', 'excel-ru.json']
/** A measurement file: under test/excel/coverage if the tests read it, else under bench/coverage/data; all in out/ on a trial run. */
const resultFile = (name: string) => join(ONLY ? OUT : TESTED.includes(name) ? join(ROOT, 'test/excel/coverage') : join(HERE, 'data'), name)
const REPORT = ONLY ? join(OUT, 'report.md') : join(ROOT, 'docs/coverage-formulajs.md')

export interface Column<T = Val> {
  meta: Record<string, string | number | boolean>
  values: Record<string, T>
}

const save = (name: string, data: unknown) => {
  mkdirSync(dirname(resultFile(name)), { recursive: true })
  writeFileSync(resultFile(name), JSON.stringify(data, null, 1) + '\n')
}
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

/** The parts of an Excel run: whole functions in order, at most PART cases per book. */
const PART = 400

function parts(cs: Case[]): Case[][] {
  const out: Case[][] = []
  let cur: Case[] = []
  for (let i = 0; i < cs.length; ) {
    let j = i
    while (j < cs.length && cs[j].fn === cs[i].fn) j++
    if (cur.length && cur.length + (j - i) > PART) {
      out.push(cur)
      cur = []
    }
    cur.push(...cs.slice(i, j))
    i = j
  }
  if (cur.length) out.push(cur)
  return out
}

interface ExcelColumn extends Column {
  meta: Column['meta'] & { parts: number; done: number; stopped?: string; dynamic: true }
}

/** The locale in an answer file's metadata. */
const LOCALE_META: Record<Locale, string> = { en: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', ru: 'ru-RU (the system language and locale)' }

/**
 * Excel in parts: before and after each part Excel is checked for a window that is not a book (a modal
 * dialog). If there is one, the run stops with the dialog's text and the part's number; the parts done are
 * in the file already. A new run goes on from the first part not done (FCOV_FRESH=1 starts over).
 */
async function stepExcel(locale: Locale) {
  const cs = cases()
  const name = `excel-${locale}.json`
  const file = resultFile(name)
  await ensureExcel(locale)
  const ps = parts(cs)
  const stored = existsSync(file) && !process.env.FCOV_FRESH ? load<ExcelColumn>(name) : null
  const col: ExcelColumn = stored ?? { meta: { version: excelVersion(), locale: LOCALE_META[locale], seconds: 0, parts: ps.length, done: 0, dynamic: true }, values: {} }
  delete col.meta.stopped
  const stop = (k: number, why: string) => {
    col.meta.stopped = `part ${k + 1} of ${ps.length} (${ps[k][0].id} … ${ps[k][ps[k].length - 1].id}): ${why}`
    save(name, col)
    throw new Error(`Excel ${locale}: ${col.meta.stopped}`)
  }
  mkdirSync(OUT, { recursive: true })
  for (let k = 0; k < ps.length; k++) {
    const part = ps[k]
    if (part.every((c) => c.id in col.values)) continue
    const before = dialogText()
    if (before) stop(k, `before the part Excel has a window that is not a book: ${before}`)
    const book = join(OUT, `part-${locale}-${k + 1}.xlsx`)
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
    Object.assign(col.values, byId(part, vals!))
    col.meta.seconds = Math.round((Number(col.meta.seconds) + (Date.now() - t0) / 1000) * 10) / 10
    col.meta.done = ps.filter((p) => p.every((c) => c.id in col.values)).length
    save(name, col)
    console.log(`Excel ${locale}: part ${k + 1}/${ps.length}, ${part.length} cases, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}

/**
 * Probes in Excel in parts of PART, with the same check for modal windows before and after each part. Without
 * a set name these are `probes.ts` and `test/excel/probes.json` (en-US only); with `<set>`, the generator
 * `sets/<set>.ts` (exporting `buildProbes()` and, when it needs data sheets, `SHEETS`) and the answers in
 * `test/excel/cases/<set>.json` for en-US, `test/excel/cases/ru-RU/<set>.json` for ru-RU. A file per set lets
 * parallel branches add their cases without touching shared files.
 */
async function stepProbes(set: string | undefined, locale: Locale) {
  if (set !== undefined && !/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  if (set === undefined && locale !== 'en') throw new Error('the probes of probes.ts are computed in en-US only; a set in ru-RU: probes <set> ru')
  const mod: { buildProbes: () => Probe[]; SHEETS?: SheetCells } = set ? await import(`./sets/${set}.js`) : { buildProbes, SHEETS: PROBE_SHEETS }
  const ps = mod.buildProbes()
  const sheets = mod.SHEETS ?? {}
  const tag = `${set ?? 'probes'}-${locale}`
  const target = !set ? join(ROOT, 'test/excel/probes.json') : join(ROOT, `test/excel/cases/${locale === 'en' ? '' : 'ru-RU/'}${set}.json`)
  // going on: answers whose formula is unchanged are taken from the file (FCOV_FRESH=1 starts over)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as { cases: Probe[]; values: Record<string, Val> }) : null
  const was = new Map(prev?.cases.map((c) => [c.id, c.formula]) ?? [])
  const values: Record<string, Val> = {}
  for (const p of ps) if (prev?.values[p.id] && was.get(p.id) === p.formula) values[p.id] = prev.values[p.id]
  const rejected: Probe[] = []
  const t0 = Date.now()
  const save = () => {
    const meta = { version: excelVersion(), locale: LOCALE_META[locale], seconds: (Date.now() - t0) / 1000, dynamic: true }
    writeFileSync(target, JSON.stringify({ meta, cases: ps, values, ...(rejected.length ? { rejected: rejected.map((p) => p.id) } : {}) }, null, 1) + '\n')
  }
  // a book Excel does not read is halved down to the formula at fault
  const run = async (cases: Probe[], label: string): Promise<void> => {
    if (declineRepair()) console.log('closed a leftover book repair window')
    const before = dialogText()
    if (before) throw new Error(`probes ${label}: Excel has a window that is not a book: ${before}`)
    const book = join(OUT, `${tag}-${label}.xlsx`)
    // the read script expects a column of two cells at least: a single case gets a pad
    const sent = cases.length === 1 ? [...cases, { id: 'pad', group: 'pad', formula: '0' }] : cases
    await writeBook(sent, book, sheets)
    const vals = await runExcelWatched(book, sent.length)
    if (vals !== 'rejected') return void cases.forEach((p, i) => (values[p.id] = vals[i]))
    if (cases.length === 1) {
      rejected.push(cases[0])
      return void console.log(`Excel does not load the formula: ${cases[0].id} ${cases[0].formula}`)
    }
    const mid = cases.length >> 1
    await run(cases.slice(0, mid), `${label}a`)
    await run(cases.slice(mid), `${label}b`)
  }
  await ensureExcel(locale)
  mkdirSync(OUT, { recursive: true })
  mkdirSync(dirname(target), { recursive: true })
  for (let k = 0; k * PART < ps.length; k++) {
    const part = ps.slice(k * PART, (k + 1) * PART).filter((p) => !values[p.id])
    if (!part.length) continue
    await run(part, String(k + 1))
    save()
    console.log(`probes ${tag}: part ${k + 1}, ${part.length} cases, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
  save()
  if (rejected.length) throw new Error(`Excel does not load ${rejected.length} formulas, remove them from the generator: ${rejected.map((p) => `${p.id} ${p.formula}`).join('; ')}`)
}

const step = process.argv[2]
const t0 = Date.now()
if (step === 'cases') await stepCases()
else if (step === 'fj') await stepFj()
else if (step === 'lo') await stepLo()
else if (step === 'excel') await withExcelLock(() => stepExcel(process.argv[3] as Locale))
else if (step === 'report') writeReport(resultFile, REPORT)
else if (step === 'probes') await withExcelLock(() => stepProbes(process.argv[3], process.argv[4] === 'ru' ? 'ru' : 'en'))
else if (step === 'struct') await withExcelLock(() => stepStruct(process.argv[3]))
else if (step === 'spill') await withExcelLock(() => stepSpill(process.argv[3]))
else if (step === 'display') await withExcelLock(() => stepDisplay(process.argv[3], process.argv[4] === 'ru' ? 'ru' : 'en'))
else if (step === 'all') {
  await stepCases()
  await stepFj()
  await stepLo()
  await withExcelLock(async () => {
    await stepExcel('en')
    await stepExcel('ru')
  })
  writeReport(resultFile, REPORT)
} else {
  console.error('step: cases | fj | lo | excel en|ru | report | all | probes [set [ru]] | struct set | spill set')
  process.exit(2)
}
console.log(`done in ${((Date.now() - t0) / 1000).toFixed(1)} s`)
