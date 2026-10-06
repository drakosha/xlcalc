/**
 * The coverage measurement's cases computed by this engine and compared with Microsoft Excel. The cases
 * (`docs/coverage-formulajs/cases.json`), the data sheet they read (`bench/coverage/data.ts`) and Excel's
 * answers (`docs/coverage-formulajs/excel-en.json`: Excel 16.113 for Mac, en-US locale and English UI) are
 * stored in the repository, so neither Excel nor LibreOffice is needed. Comparison follows the measurement
 * (`bench/coverage/compare.ts`): numbers within |a−b| ≤ 1e-9·max(1,|a|), strings, booleans and error codes
 * exactly.
 */
import { readFileSync } from 'node:fs'
import { compare } from '../bench/coverage/compare.js'
import { DATA_SHEET, dataCells, type Cell } from '../bench/coverage/data.js'
import type { Val } from '../bench/coverage/formulajs.js'
import type { BookInput, CellInput, CellValue } from '../src/contract.js'
import { compute } from '../src/index.js'

export interface ExcelCase {
  id: string
  fn: string
  kind: 'typical' | 'edge'
  tag: string
  formula: string
}

const data = (file: string) => JSON.parse(readFileSync(new URL(`../docs/coverage-formulajs/${file}`, import.meta.url), 'utf8'))

export const CASES: ExcelCase[] = (data('cases.json') as { cases: ExcelCase[] }).cases
export const EXCEL: Record<string, Val> = (data('excel-en.json') as { values: Record<string, Val> }).values

const CASE_SHEET = 'C'
const ERROR_LITERAL: Record<string, string> = { DIV0: '#DIV/0!', NA: '#N/A', VALUE: '#VALUE!', NUM: '#NUM!', REF: '#REF!', NAME: '#NAME?' }

/** A cell of the measurement's data sheet as engine input: errors as error literals, dates as serial numbers. */
function dataInput(c: Cell): CellInput {
  if (typeof c !== 'object') return { value: c }
  if ('error' in c) return { formula: ERROR_LITERAL[c.error] }
  if ('date' in c) return { value: c.date }
  return { value: c.text }
}

/** The formula as this engine reads it: Excel's storage prefixes of newer functions removed. */
export const engineFormula = (f: string) => f.replace(/_xlfn\.(_xlws\.)?/g, '')

/** The measurement's data sheet, as in every book sent to Excel. */
export function dataSheet(): BookInput['sheets'][number] {
  const d: Record<string, CellInput> = {}
  for (const [addr, v] of dataCells()) d[addr] = dataInput(v)
  return { name: DATA_SHEET, cells: d }
}

export function excelBook(cases: readonly ExcelCase[]): BookInput {
  const c: Record<string, CellInput> = {}
  cases.forEach((x, i) => (c[`A${i + 1}`] = { formula: engineFormula(x.formula) }))
  return { sheets: [dataSheet(), { name: CASE_SHEET, cells: c }] }
}

export function toVal(v: CellValue): Val {
  if (v === null) return { t: 'n', v: 0 }
  if (typeof v === 'number') return { t: 'n', v }
  if (typeof v === 'string') return { t: 's', v }
  if (typeof v === 'boolean') return { t: 'b', v }
  return { t: 'e', v: v.error as never }
}

export interface Outcome {
  c: ExcelCase
  excel: Val
  engine: Val
  ok: boolean
  /** Why a case in KNOWN fails. */
  why?: string
}

/**
 * Functions compared with the measurement's soft tolerance |a−b| ≤ 1e-6·max(1,|a|), and why: Excel's own
 * iteration stops short of the root, which the engine finds exactly (`docs/coverage-formulajs.md`).
 */
export const SOFT: Readonly<Record<string, string>> = {
  XIRR: 'Excel stops within its 1e-8 relative tolerance of the root (the measurement matched it only softly)',
}

/** A case matches Excel: within the usual tolerance, or the soft one for the functions in SOFT. */
export function matches(fn: string | undefined, excel: Val, engine: Val): boolean {
  const v = compare(excel, engine)
  return v.ok || (fn !== undefined && fn in SOFT && v.soft === true)
}

/** Reasons of known differences; each is a line of the README's "Known differences from Excel". */
export const REASONS = {
  rateZeroRoot:
    '`RATE` with a rate of exactly 0 as the answer gives 0 or a value within 1e-10 of it; Excel answers its own rounding noise up to about 2e-9 or, from some guesses, `#NUM!`.',
} as const

/**
 * Cases where this engine knowingly differs from Excel, by `<source>:<case id>` (`measurement` or a case
 * set's name): the engine's answer the tests pin instead of Excel's. An entry whose case matches Excel fails
 * the tests, so the list stays current.
 */
export const KNOWN: Readonly<Record<string, { reason: keyof typeof REASONS; engine: Val }>> = {
  'findate:RATE#20': { reason: 'rateZeroRoot', engine: { t: 'n', v: 4.17387402722136e-11 } },
  'findate:RATE#23': { reason: 'rateZeroRoot', engine: { t: 'n', v: 2.1421887871288e-10 } },
  'findate:RATE#24': { reason: 'rateZeroRoot', engine: { t: 'n', v: 1.22749297268431e-10 } },
}

/** A case passes: it matches Excel, or it is in KNOWN, differs from Excel and gives exactly the pinned answer. */
export function verdict(key: string, fn: string | undefined, excel: Val, engine: Val): { ok: boolean; why?: string } {
  const known = KNOWN[key]
  const same = matches(fn, excel, engine)
  if (!known) return { ok: same }
  if (same) return { ok: false, why: 'matches Excel, remove it from KNOWN' }
  return engine.t === known.engine.t && engine.v === known.engine.v ? { ok: true } : { ok: false, why: `KNOWN pins ${show(known.engine)}` }
}

/** Every case computed in one book and compared with Excel en-US. */
export function runExcelCases(cases: readonly ExcelCase[] = CASES): Outcome[] {
  const r = compute(excelBook(cases), { clock: { now: '2026-10-05T12:00:00Z', seed: 1 } })
  const cells = r.sheets.find((s) => s.name === CASE_SHEET)!.cells
  return cases.map((c, i) => {
    const excel = EXCEL[c.id]
    const engine = toVal(cells[`A${i + 1}`].value)
    const v = verdict(`measurement:${c.id}`, c.fn, excel, engine)
    return { c, excel, engine, ok: v.ok, why: v.why }
  })
}

export const show = (v: Val) => (v.t === 'e' ? `#${v.v}` : JSON.stringify(v.v))
