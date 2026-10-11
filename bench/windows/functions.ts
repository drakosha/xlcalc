/**
 * Formula cases in Excel for Windows: the coverage measurement, the probes, a case set or a legacy set, each in the
 * very books the Mac harness writes (`bench/coverage/book.ts`: the data sheet, the set's sheets, the case formulas as
 * dynamic array formulas, or of the old kind for a legacy set). Excel computes a book fully and the case column is
 * read through COM in one call (`Value2`: errors come as codes, so no wrapper columns are read). Parts of up to 400
 * cases, up to 40 books to a job; a book Excel does not open, opens repaired (its case formulas gone) or hangs on is halved
 * down to the formulas at fault, which are stored as `rejected`. A rerun keeps the answers of unchanged formulas.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { functionParts, inParts, writeBook } from '../coverage/book.js'
import type { CaseSet } from '../coverage/cases.js'
import type { ErrCode } from '../coverage/data.js'
import type { Val } from '../coverage/formulajs.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from '../coverage/probes.js'
import { localeDir, measurementFile } from '../coverage/locales.js'
import { inCaseOrder, loadStored, secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { again, inJobs, metaOf, type Answer, type Book, type Tag, type WinEnv } from './remote.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../..')

/** Where the Windows answers are stored, beside the Mac's in `test/excel/`. */
export const winFile = (rel: string) => join(ROOT, 'test/excel/win', rel)

/** Excel's error codes (`CVErr` numbers) as COM gives them, 0x800A0000 + code; codes past `#CALC!` are `ERROR`, as the Mac's ERROR.TYPE reading has them. */
const XL_ERRORS: Record<number, ErrCode> = { 2000: 'NULL', 2007: 'DIV0', 2015: 'VALUE', 2023: 'REF', 2029: 'NAME', 2036: 'NUM', 2042: 'NA', 2043: 'GETTING_DATA', 2045: 'SPILL', 2050: 'CALC' }
const COM_ERROR_BASE = -2146828288

/** A cell as `xl.ps1` writes it: null blank, a number, a string, a boolean, `{ e }` an error code, `{ other }` anything else. */
export type WinCell = null | number | string | boolean | { e: number } | { other: string }

export function valOf(c: WinCell): Val | null {
  if (c === null) return null
  if (typeof c === 'number') return { t: 'n', v: c }
  if (typeof c === 'string') return { t: 's', v: c }
  if (typeof c === 'boolean') return { t: 'b', v: c }
  if ('e' in c) return { t: 'e', v: XL_ERRORS[c.e - COM_ERROR_BASE] ?? 'ERROR' }
  return { t: 'other', v: c.other }
}

/**
 * A legacy set's books as the Mac's run had them: implicit intersection takes a formula's row, so each case goes to the
 * row it stood in there (a new book where the rows start again, `rows` of the Mac's answers).
 */
function macBooks(cases: Probe[], set: string): Probe[][] | null {
  const file = join(ROOT, `test/excel/legacy/${set}.json`)
  if (!existsSync(file)) return null
  const rows = (JSON.parse(readFileSync(file, 'utf8')) as { rows: Record<string, number> }).rows
  const out: Probe[][] = []
  for (const c of cases) {
    if (!rows[c.id]) return null
    const last = out.at(-1)?.at(-1)
    if (!last || rows[c.id] !== rows[last.id] + 1) {
      if (rows[c.id] !== 1) return null
      out.push([])
    }
    out.at(-1)!.push(c)
  }
  return out
}

/** What a run of formula cases stores. */
export interface FormulaAnswers {
  values: Record<string, Val>
  /** A legacy set: each case's formula as Excel 365 shows it (`Formula2`) and its row. */
  formulas: Record<string, string>
  rows: Record<string, number>
  /** Cases whose formula Excel did not take. */
  rejected: string[]
}

/** A run's cases and where their answers go. */
interface Target {
  cases: Probe[]
  sheets: SheetCells
  file: string
  /** The measurement stores values only (its cases are in `test/excel/coverage/cases.json`). */
  valuesOnly?: true
  legacy?: true
}

/** `measurement`, `probes`, a set of `sets/` or `legacy/<set>`, in a locale (the probes and the legacy sets in en-US only). */
export async function target(name: string, tag: Tag): Promise<Target> {
  const en = tag === 'en-US'
  if (name === 'measurement') {
    const set = JSON.parse(readFileSync(join(ROOT, 'test/excel/coverage/cases.json'), 'utf8')) as CaseSet
    const cases = set.cases.map((c) => ({ id: c.id, group: c.fn, fn: c.fn, formula: c.formula }))
    return { cases, sheets: {}, file: winFile(`coverage/${measurementFile(tag)}`), valuesOnly: true }
  }
  if (name === 'probes') {
    if (!en) throw new Error('the probes are computed in en-US only, as on the Mac')
    return { cases: buildProbes(), sheets: PROBE_SHEETS, file: winFile('probes.json') }
  }
  const [dir, set] = name.startsWith('legacy/') ? ['legacy', name.slice(7)] : ['sets', name]
  if (dir === 'legacy' && !en) throw new Error('the legacy sets are computed in en-US only, as on the Mac')
  const mod = await setModule<{ buildProbes: () => Probe[]; SHEETS?: SheetCells }>(dir, set)
  return { cases: mod.buildProbes(), sheets: mod.SHEETS ?? {}, file: winFile(`${dir === 'legacy' ? 'legacy' : 'cases'}/${localeDir(tag)}${set}.json`), ...(dir === 'legacy' ? { legacy: true as const } : {}) }
}

/** A part of a target's cases. */
type Item = { r: Run; part: Probe[] }

/** A target in a run: its answers so far, and how they are stored. */
interface Run {
  name: string
  t: Target
  got: FormulaAnswers
  save: () => void
}

/** A target's stored answers that still hold: those of unchanged formulas (FCOV_FRESH=1 keeps none). */
function storedOf(t: Target): FormulaAnswers {
  const prev = loadStored<Partial<FormulaAnswers> & { cases?: Probe[] }>(t.file)
  const got: FormulaAnswers = { values: {}, formulas: {}, rows: {}, rejected: [] }
  if (!prev) return got
  // the measurement's file has no cases: its formulas are those of cases.json, so an answer stays while that file does
  const was = new Map((prev.cases ?? (t.valuesOnly ? t.cases : [])).map((c) => [c.id, c.formula]))
  for (const c of t.cases) {
    if (was.get(c.id) !== c.formula) continue
    if (prev.values?.[c.id]) got.values[c.id] = prev.values[c.id]
    else if (prev.rejected?.includes(c.id)) got.rejected.push(c.id)
    if (t.legacy && prev.formulas?.[c.id] !== undefined) [got.formulas[c.id], got.rows[c.id]] = [prev.formulas[c.id], prev.rows![c.id]]
  }
  return got
}

/** A part's book: computed fully, then whether every case cell still holds its formula, the values and, for a legacy set, the formulas as Excel 365 shows them. */
async function bookOf(r: Run, part: Probe[], file: string): Promise<Book> {
  await writeBook(part, file, r.t.sheets, !!r.t.legacy)
  const range = { sheet: 'C', range: `A1:A${part.length}` }
  return { file, steps: [{ do: 'full' }, { do: 'hasFormula', ...range }, { do: 'values', ...range }, ...(r.t.legacy ? [{ do: 'formulas2', ...range }] : [])] }
}

/** A part's answer into its run; else the part again, halved down to the case Excel does not take. */
function take(r: Run, part: Probe[], a: Answer, cutShort: boolean): Probe[][] {
  if (a && 'answers' in a && a.answers[1] === true) {
    const [, , values, f2] = a.answers as [null, boolean, WinCell[][], string[][]?]
    part.forEach((p, i) => {
      r.got.values[p.id] = valOf(values[i][0]) ?? { t: 'other', v: 'blank' }
      if (r.t.legacy) [r.got.formulas[p.id], r.got.rows[p.id]] = [f2![i][0], i + 1]
    })
    return []
  }
  return again(part, a, cutShort, (p) => {
    r.got.rejected.push(p.id)
    console.log(`${r.name}: Excel does not take ${p.id} ${p.formula}: ${a ? ('rejected' in a ? a.rejected : 'the book was repaired') : 'it hangs'}`)
  })
}

/**
 * The targets' cases without a stored answer for the same formula, in jobs of up to 40 books across targets
 * (a job costs some 30 s of SSH and Excel's start); the targets' files are written after every job.
 */
export async function stepFunctions(names: string[], tag: Tag, e: WinEnv): Promise<void> {
  const t0 = Date.now()
  const runs: Run[] = []
  const queue: Item[] = []
  for (const name of names) {
    const t = await target(name, tag)
    const got = storedOf(t)
    const r: Run = {
      name: `${name}-${tag}`,
      t,
      got,
      save: () => {
        const meta = { ...metaOf(e, tag), seconds: secondsSince(t0), ...(t.legacy ? { legacy: true } : { dynamic: true }) }
        const rejected = t.cases.filter((c) => got.rejected.includes(c.id)).map((c) => c.id)
        const values = inCaseOrder(t.cases, got.values)
        const body = t.valuesOnly ? { meta, values } : { meta, cases: t.cases, values, ...(t.legacy ? { formulas: inCaseOrder(t.cases, got.formulas), rows: inCaseOrder(t.cases, got.rows) } : {}) }
        writeAnswers(t.file, { ...body, ...(rejected.length ? { rejected } : {}) })
      },
    }
    runs.push(r)
    const todo = t.cases.filter((c) => !got.values[c.id] && !got.rejected.includes(c.id))
    // a legacy set run whole goes into the Mac's books, so each formula intersects at its Mac row
    const books = t.legacy && todo.length === t.cases.length ? macBooks(todo, name.slice('legacy/'.length)) : null
    if (t.legacy && !books) console.log(`${name}: not in the Mac's books (some answers kept, or the Mac's rows do not fit)`)
    // the books the Mac harness makes: the measurement's of whole functions, the others' parts in order
    queue.push(...(books ?? (t.valuesOnly ? functionParts(todo) : inParts(todo))).map((part) => ({ r, part })))
  }
  await inJobs(
    `functions-${tag}`,
    queue,
    (it, file) => bookOf(it.r, it.part, file),
    (it, a, cutShort) => take(it.r, it.part, a, cutShort).map((part) => ({ r: it.r, part })),
    () => runs.forEach((r) => r.save()),
  )
  for (const r of runs) {
    r.save()
    console.log(`${r.name}: ${Object.keys(r.got.values).length} of ${r.t.cases.length} answered${r.got.rejected.length ? `, ${r.got.rejected.length} not taken` : ''}`)
  }
}
