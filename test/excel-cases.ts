/**
 * The coverage measurement's cases computed by this engine and compared with Microsoft Excel. The cases
 * (`docs/coverage-formulajs/cases.json`), the data sheet they read (`bench/coverage/data.ts`) and Excel's
 * answers (`docs/coverage-formulajs/excel-en.json`: Excel 16.113 for Mac, en-US locale and English UI, dynamic
 * array formulas as Excel 365 writes them) are stored in the repository, so neither Excel nor LibreOffice is needed. Comparison follows the measurement
 * (`bench/coverage/compare.ts`): numbers within |a−b| ≤ 1e-9·max(1,|a|), strings, booleans and error codes
 * exactly.
 */
import { readFileSync } from 'node:fs'
import { compare } from '../bench/coverage/compare.js'
import { DATA_SHEET, dataCells, type Cell } from '../bench/coverage/data.js'
import type { Val } from '../bench/coverage/formulajs.js'
import type { BookInput, CellInput, CellValue } from '../src/contract.js'
import { compute } from '../src/index.js'
import { OPEN } from './excel-open.js'

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

/** Case formulas in column A as in Excel's book: a neighbour in B (Excel's wrappers) and a value under the last, so an array answer is `#SPILL!`. */
export function caseCells(formulas: readonly string[]): Record<string, CellInput> {
  const c: Record<string, CellInput> = {}
  formulas.forEach((f, i) => {
    c[`A${i + 1}`] = { formula: f }
    c[`B${i + 1}`] = { value: 0 }
  })
  c[`A${formulas.length + 1}`] = { value: 'end' }
  return c
}

export function excelBook(cases: readonly ExcelCase[]): BookInput {
  return { sheets: [dataSheet(), { name: CASE_SHEET, cells: caseCells(cases.map((x) => engineFormula(x.formula))) }] }
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
  'F.INV': "Excel loses up to 7 digits when a degree of freedom is 10^9 or more; the engine's quantile checks out with mpmath",
  'F.INV.RT': "Excel loses up to 7 digits when a degree of freedom is 10^9 or more; the engine's quantile checks out with mpmath",
  FINV: "Excel loses up to 7 digits when a degree of freedom is 10^9 or more; the engine's quantile checks out with mpmath",
  'T.INV': "Excel loses up to 7 digits at 10^9 degrees of freedom and more (it inverts through the beta function near 1); the engine's quantile checks out with mpmath",
  'T.INV.2T': "Excel loses up to 7 digits at 10^9 degrees of freedom and more (it inverts through the beta function near 1); the engine's quantile checks out with mpmath",
  TINV: "Excel loses up to 7 digits at 10^9 degrees of freedom and more (it inverts through the beta function near 1); the engine's quantile checks out with mpmath",
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
  gammaInvConvergence:
    '`GAMMA.INV`, `CHISQ.INV`, `CHISQ.INV.RT` (and `GAMMAINV`, `CHIINV`) give the quantile where Excel\'s iteration fails with `#NUM!`: shapes from about 5·10^9 (10^10 degrees of freedom) and probabilities near 1 with shapes under 0.03; Excel\'s own `GAMMA.DIST` at the engine\'s answer gives the probability back.',
  binomInvTie:
    '`BINOM.INV` where alpha equals the probability of no success, (1 − p)^n, exactly: Excel decides by its own rounding of that power, the engine by exp(n·ln(1 − p)); they part on `BINOM.INV(3,0.25,0.421875)`, 0 in Excel and 1 here.',
  lookupPastEnd: '`LOOKUP` with a result range shorter than the lookup range is `#REF!` past its end; Excel reads on past the end of the range.',
  textDigits:
    'A number that Excel computes by another algorithm can differ from the engine\'s in the 15th significant digit, within the tests\' relative tolerance of 1e-9; turned into text (`&`, `TEXTJOIN`) the last digit then differs (`TEXTJOIN(";",FALSE,BESSELK(10,3))` is `2.72527006358477E-05` in Excel and `2.72527006358476E-05` here).',
  sortbyOrders:
    '`SORTBY` with an array or a range of several cells as a sort order is `#VALUE!`; Excel answers an array the size of the orders (the first sorted value, then blanks, or an error where a sort fails).',
  linestConstantColumn:
    '`LINEST` with a constant x variable besides the intercept: both give it coefficient 0 and the same fit, but Excel\'s statistics count it as a variable in some way the engine does not reproduce (`LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE)`: degrees of freedom 1 in Excel, 2 here, and the F statistic and the intercept\'s standard error differ).',
  rowsOfIfs: '`ROWS` and `COLUMNS` of an `IFS` that answers an error: Excel counts 1 (`ROWS(IFS("x",1,TRUE,2))`), the engine answers the error.',
  lastBits:
    '`FTEST` of a sample with itself is 1 − 3·2^-53 in Excel and 1 − 2^-52 here: the last bits of a number Excel computes by another algorithm can differ, which shows only when the difference is magnified (`(1-FTEST(A1:A8,A1:A8))*2^52` is 1.5 in Excel, 1 here).',
  omittedArgument: 'An omitted argument between commas (`TEXTSPLIT("a,b",",",,,1)`) is a syntax error here (`#ERROR!`); Excel takes it as not given.',
  excelContext:
    '`WEEKNUM` with an invalid return type is `#NUM!`, as Excel answers it alone; inside one generated book of the random check Excel 16.113 gave a number for `DAY(WEEKNUM(40977,1E+3)<>2.5)` (1, on two runs).',
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
  'measurement:CHIINV#7': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232617.1965 } },
  'measurement:CHISQ.INV#14': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000208709.154 } },
  'measurement:CHISQ.INV.RT#15': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232617.1965 } },
  'measurement:GAMMAINV#8': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 19999701973.2484 } },
  'stat:CHISQ.INV#27': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232618.5678 } },
  'stat:CHISQ.INV.RT#27': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 9990232502.22995 } },
  'stat:GAMMA.INV#5': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 1.50908414769475 } },
  'stat:GAMMA.INV#16': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 9990052413.58303 } },
  'stat:GAMMA.INV#17': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000052439.8096 } },
  'stat:GAMMA.INV#26': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 0.265052550251588 } },
  'stat:GAMMA.INV#27': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 0.934242187465308 } },
  'stat:GAMMA.INV#31': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 5000116309.28388 } },
  'stat:GAMMA.INV#33': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 5999959379.86929 } },
  'stat:GAMMAINV#5': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 1.50908414769475 } },
  'stat:GAMMAINV#16': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 9990052413.58303 } },
  'stat:GAMMAINV#17': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000052439.8096 } },
  'stat:BINOM.INV#30': { reason: 'binomInvTie', engine: { t: 'n', v: 1 } },
  'fuzz:1#697': { reason: 'lookupPastEnd', engine: { t: 'e', v: 'REF' } },
  'fuzz:1#234': { reason: 'textDigits', engine: { t: 's', v: '2.72527006358476E-05' } },
  'fuzz:1#403': { reason: 'textDigits', engine: { t: 's', v: '4.18382560777941E-23' } },
  'fuzz:1#947': { reason: 'textDigits', engine: { t: 's', v: '0.00803735952120296' } },
  'fuzz:1#1327': { reason: 'textDigits', engine: { t: 's', v: '0.1963093102686' } },
  'fuzz:1#185': { reason: 'sortbyOrders', engine: { t: 'e', v: 'VALUE' } },
  'fuzz:1#232': { reason: 'sortbyOrders', engine: { t: 'e', v: 'VALUE' } },
  'fuzz:1#426': { reason: 'sortbyOrders', engine: { t: 'e', v: 'VALUE' } },
  'fuzz-rules:LINEST#41': { reason: 'linestConstantColumn', engine: { t: 'n', v: 0.474341649025257 } },
  'fuzz-rules:LINEST#42': { reason: 'linestConstantColumn', engine: { t: 'n', v: 56.3333333333333 } },
  'fuzz-rules:LINEST#43': { reason: 'linestConstantColumn', engine: { t: 'n', v: 2 } },
  'fuzz:1#1263': { reason: 'excelContext', engine: { t: 'e', v: 'NUM' } },
  'fuzz:2#472': { reason: 'textDigits', engine: { t: 's', v: '2048.47532711546' } },
  'fuzz:2#411': { reason: 'sortbyOrders', engine: { t: 'e', v: 'NUM' } },
  'fuzz:2#1577': { reason: 'sortbyOrders', engine: { t: 'e', v: 'VALUE' } },
  'fuzz:2#165': { reason: 'rowsOfIfs', engine: { t: 'e', v: 'VALUE' } },
  'near-integer:literal#7': { reason: 'lastBits', engine: { t: 'n', v: 1 } },
  'near-integer:sample:TEXTSPLIT#1': { reason: 'omittedArgument', engine: { t: 'e', v: 'ERROR' } },
  'near-integer:sample:TEXTSPLIT#2': { reason: 'omittedArgument', engine: { t: 'e', v: 'ERROR' } },
}

/** A case passes: it matches Excel, or it is in KNOWN (or OPEN), differs from Excel and gives exactly the pinned answer. */
export function verdict(key: string, fn: string | undefined, excel: Val, engine: Val): { ok: boolean; why?: string } {
  const known = KNOWN[key] ?? OPEN[key]
  const same = matches(fn, excel, engine)
  if (!known) return { ok: same }
  if (same) return { ok: false, why: `matches Excel, remove it from ${key in KNOWN ? 'KNOWN' : 'OPEN'}` }
  return engine.t === known.engine.t && engine.v === known.engine.v ? { ok: true } : { ok: false, why: `KNOWN pins ${show(known.engine)}` }
}

/** The clock of every case book: the volatile functions' cases do not depend on it. */
export const CLOCK = { now: '2026-10-05T12:00:00Z', seed: 1 }

/** Every case computed in one book and compared with Excel en-US. */
export function runExcelCases(cases: readonly ExcelCase[] = CASES): Outcome[] {
  const r = compute(excelBook(cases), { clock: CLOCK })
  const cells = r.sheets.find((s) => s.name === CASE_SHEET)!.cells
  return cases.map((c, i) => {
    const excel = EXCEL[c.id]
    const engine = toVal(cells[`A${i + 1}`].value)
    const v = verdict(`measurement:${c.id}`, c.fn, excel, engine)
    return { c, excel, engine, ok: v.ok, why: v.why }
  })
}

export const show = (v: Val) => (v.t === 'e' ? `#${v.v}` : JSON.stringify(v.v))
