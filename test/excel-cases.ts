/**
 * The coverage measurement's cases computed by this engine and compared with Microsoft Excel. The cases
 * (`test/excel/coverage/cases.json`), the data sheet they read (`bench/coverage/data.ts`) and Excel's
 * answers (`test/excel/coverage/excel-en.json`: Excel 16.113 for Mac, en-US locale and English UI, dynamic
 * array formulas as Excel 365 writes them) are stored in the repository, so neither Excel nor LibreOffice is needed. Comparison follows the measurement
 * (`bench/coverage/compare.ts`): numbers within |a−b| ≤ 1e-9·max(1,|a|), strings, booleans and error codes
 * exactly.
 */
import { readFileSync } from 'node:fs'
import { compare, samePin } from '../bench/coverage/compare.js'
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

const data = (file: string) => JSON.parse(readFileSync(new URL(`./excel/coverage/${file}`, import.meta.url), 'utf8'))

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

/** The formula as this engine reads it: Excel's storage prefixes of newer functions and of LET's names removed. */
export const engineFormula = (f: string) => f.replace(/_xlfn\.(_xlws\.)?|_xlpm\./g, '')

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

/** Reasons of known differences; each is a line of docs/limitations.md's "Known differences from Excel". */
export const REASONS = {
  rateZeroRoot:
    'When the result of `RATE` is exactly 0, it returns 0 or a value within 1e-10 of it; Excel returns its own rounding noise of up to about 2e-9, or, from some guesses, `#NUM!`.',
  gammaInvConvergence:
    '`GAMMA.INV`, `CHISQ.INV`, `CHISQ.INV.RT` (and `GAMMAINV`, `CHIINV`) return the quantile where Excel\'s iteration fails with `#NUM!`: shapes from about 5·10^9 (10^10 degrees of freedom), and probabilities near 1 with shapes under 0.03; Excel\'s own `GAMMA.DIST` at the engine\'s result gives back the probability.',
  binomInvTie:
    '`BINOM.INV` where alpha exactly equals the probability of no success, (1 − p)^n: Excel decides by its own rounding of that power, the engine by exp(n·ln(1 − p)); they differ on `BINOM.INV(3,0.25,0.421875)`, which is 0 in Excel and 1 here.',
  textDigits:
    'A number that Excel calculates with a different algorithm can differ from the engine\'s in the 15th significant digit, within the tests\' relative tolerance of 1e-9; converted to text (`&`, `TEXTJOIN`), the last digit then differs (`TEXTJOIN(";",FALSE,BESSELK(10,3))` is `2.72527006358477E-05` in Excel and `2.72527006358476E-05` here).',
  linestConstantColumn:
    '`LINEST` with a constant x variable besides the intercept: both give it a coefficient of 0 and the same fit, but Excel\'s statistics count it as a variable in a way the engine does not reproduce (`LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE)`: 1 degree of freedom in Excel and 2 here, and the F statistic and the intercept\'s standard error differ).',
  lastBits:
    '`FTEST` of a sample with itself is 1 − 3·2^-53 in Excel and 1 − 2^-52 here: the last bits of a number Excel calculates with a different algorithm can differ, which shows only when the difference is magnified (`(1-FTEST(A1:A8,A1:A8))*2^52` is 1.5 in Excel and 1 here).',
  letFileName:
    'A LET name that looks like an R1C1 row or column reference (`R2`, `C5`, `RC5`) or a boolean (`TRUE`): an xlsx file writes LET names with the prefix `_xlpm.`, so Excel loads `LET(_xlpm.TRUE,1,_xlpm.TRUE)` (1); formula text has no prefix, and there Excel rejects such a word as a LET name, as the engine does (`LET(R2,1,R2)` typed), while it otherwise accepts a cell address (`LET(A1,1,A1)` is 1).',
  codePage:
    '`CHAR` and `CODE` take codes 128–255 from the Windows code page of the workbook\'s locale (Windows-1252 in en-US, Windows-1251 in ru-RU), as Excel for Windows does; Excel for Mac uses Mac Roman and Mac Cyrillic (`CHAR(200)` is `È` here and `»` in Excel for Mac en-US, and `CODE("Ж")` in ru-RU is 198 here and 134 there).',
  outsideCode: '`CODE` of a character outside the code page of the workbook\'s locale is 63 (`?`), as in Excel for Windows; Excel for Mac returns 95 (`_`): `CODE("中")`, and `CODE("Ж")` in en-US.',
  macMonthCase:
    'Excel for Mac writes month names in ru-RU in lowercase (`TEXT` with `ММММ`, `МММММ` or `[$-419]mmmm`, and cells: `январь`, `я`); the engine capitalizes them, as Excel for Windows does (`Январь`, `Я`).',
  momentRounding:
    '`VAR`, `STDEV`, `DEVSQ`, `CORREL`, `RSQ`, `SLOPE`, `STEYX` and related functions, including the database ones, are calculated here from exact sums; Excel calculates them in double precision around a mean computed by plain summation, and where the spread is small relative to the mean, or the residual relative to the variation, it loses the digits of the result: for 10,000 values 1E12 + 0, 0.01, 0.02, `VAR.P` is 1.598E-04 in Excel and 6.680E-05 here, the exact value; `CORREL` of them with their steps 0, 1, 2 is 0.647 in Excel and 1 here; and `STEYX` of 1, 2.00000001, 3 on 1, 2, 3 is 0 in Excel and 8.16E-09 here.',
  momentRange:
    'Where the squares or products of the values fall outside the range of doubles, Excel returns what is left of them and the engine the exact result: the squares of 1E-200 and 2E-200 are 0 in Excel, so `STDEV.P` of them is 0 (5E-201 here, within the tests\' tolerance), and `CORREL` with 1, 2 and `SLOPE` of 1, 2 on them are `#DIV/0!` (1 and 1E+200 here); a sum or a product beyond 1.8E+308 is `#NUM!` in Excel even where the result fits (`AVERAGE` of 1E+308 three times and its variance 0, and the covariance of 1E+308, −1E+308, 1E-200 with 0, 0, 1E+308, which is 2.2E+107 here).',
  lookupReadZone:
    '`LOOKUP` counts among its references every cell it may read, as far as its lookup array extends, whatever value it looks up: a formula among them, the `LOOKUP` itself included, is calculated first, and a `LOOKUP` within that range is a cycle here and is not calculated (it keeps its value, 0 in a freshly calculated workbook), whereas Excel, which follows the cells `LOOKUP` actually reads, calculates it (`E10` = `LOOKUP(5,A1:A6,B10:D10)` reads `F10` in Excel and keeps 0 here, since `E10` lies within `B10:G10`; the same applies next to a spill that could not spill, and to the right of a rectangle from which `INDEX` takes a column).',
  linestMean:
    '`LINEST` and `TREND` over y far from 0 relative to its spread return Excel\'s coefficients, but Excel\'s intercept and regression sum of squares follow the mean it computes by plain summation: for 10,000 values 1E12 + 0, 0.01, 0.02 on their steps 0, 1, 2, the slope is 0.0100098 in both (0.010009765625 exactly), the intercept is 1E12 − 0.0096 in Excel and 1E12 here, and the regression sum of squares is 1.598 in Excel and 0.668 here (the exact value, `DEVSQ` of the values), so `TREND` at step 3 is 1E12 + 0.0204 in Excel and 1E12 + 0.0300 here.',
  growthRange:
    '`GROWTH` far from 0 in x is `#NUM!` in Excel even where the fitted value fits in a double: for 100·1.1^k rounded to cents on x = 1E10 + k, k = 0…11, the value at 1E10 + 12 is 313.84 here, where b·m^x has the b of `LOGEST` equal to 0 and m^x beyond the largest double.',
  macLongS:
    'Excel for Mac reads the long s (`ſ`) as `s`: `falſe` is FALSE in a formula, both typed and as text (`IF("falſe",1,2)` is 2), `=ſum(1,2)` typed is `SUM`, and `ſ` matches the name `S` (a LET name, a defined name, a table, a sheet); the engine follows Excel for Windows, where `ſ` is a separate letter (`falſe` is a name, so `#NAME?`, and the text `"falſe"` is not a boolean).',
  excelContext:
    '`WEEKNUM` with an invalid return type is `#NUM!`, as Excel returns when it stands alone; inside one generated workbook of the random check, Excel 16.113 returned a number for `DAY(WEEKNUM(40977,1E+3)<>2.5)` (1, in two runs).',
} as const

/** Known differences that hold in en-US and in ru-RU alike: CODE of characters outside the code page, inverse gamma's convergence. */
export const BOTH_KNOWN: Readonly<Record<string, { reason: keyof typeof REASONS; engine: Val }>> = {
  'char-code:CODE#142': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#143': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#144': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#145': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#146': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#147': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'char-code:CODE#148': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'measurement:CHIINV#7': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232617.1965 } },
  'measurement:CHISQ.INV#14': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000208709.154 } },
  'measurement:CHISQ.INV.RT#15': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 10000232617.1965 } },
  'measurement:GAMMAINV#8': { reason: 'gammaInvConvergence', engine: { t: 'n', v: 19999701973.2484 } },
}

/**
 * Cases where this engine knowingly differs from Excel, by `<source>:<case id>` (`measurement` or a case
 * set's name): the engine's answer the tests pin instead of Excel's. An entry whose case matches Excel fails
 * the tests, so the list stays current.
 */
export const KNOWN: Readonly<Record<string, { reason: keyof typeof REASONS; engine: Val }>> = {
  'findate:RATE#20': { reason: 'rateZeroRoot', engine: { t: 'n', v: 4.17387402722136e-11 } },
  'findate:RATE#23': { reason: 'rateZeroRoot', engine: { t: 'n', v: 2.1421887871288e-10 } },
  'findate:RATE#24': { reason: 'rateZeroRoot', engine: { t: 'n', v: 1.22749297268431e-10 } },
  'measurement:CHAR#3': { reason: 'codePage', engine: { t: 's', v: 'È' } },
  'measurement:CODE#3': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'measurement:CODE#5': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'measurement:CODE#6': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  'measurement:CODE#12': { reason: 'outsideCode', engine: { t: 'n', v: 63 } },
  ...BOTH_KNOWN,
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
  'fuzz:1#234': { reason: 'textDigits', engine: { t: 's', v: '2.72527006358476E-05' } },
  'fuzz:1#403': { reason: 'textDigits', engine: { t: 's', v: '4.18382560777941E-23' } },
  'fuzz:1#947': { reason: 'textDigits', engine: { t: 's', v: '0.00803735952120296' } },
  'fuzz:1#1327': { reason: 'textDigits', engine: { t: 's', v: '0.1963093102686' } },
  'fuzz-rules:LINEST#41': { reason: 'linestConstantColumn', engine: { t: 'n', v: 0.474341649025257 } },
  'fuzz-rules:LINEST#42': { reason: 'linestConstantColumn', engine: { t: 'n', v: 56.3333333333333 } },
  'fuzz-rules:LINEST#43': { reason: 'linestConstantColumn', engine: { t: 'n', v: 2 } },
  'fuzz:1#1263': { reason: 'excelContext', engine: { t: 'e', v: 'NUM' } },
  'fuzz:2#472': { reason: 'textDigits', engine: { t: 's', v: '2048.47532711546' } },
  'near-integer:literal#7': { reason: 'lastBits', engine: { t: 'n', v: 1 } },
  'let:let:file#5': { reason: 'letFileName', engine: { t: 'e', v: 'ERROR' } },
  'name-case-held:bool#1': { reason: 'macLongS', engine: { t: 'e', v: 'NAME' } },
  'name-case-held:bool#2': { reason: 'macLongS', engine: { t: 'e', v: 'NAME' } },
  'name-case-held:bool#3': { reason: 'macLongS', engine: { t: 'e', v: 'NAME' } },
  'name-case-held:bool#4': { reason: 'macLongS', engine: { t: 'e', v: 'NAME' } },
  'name-case-held:bool#5': { reason: 'macLongS', engine: { t: 'e', v: 'ERROR' } },
  'name-case-held:bool#6': { reason: 'macLongS', engine: { t: 'e', v: 'NAME' } },
  'name-case-held:let#2': { reason: 'macLongS', engine: { t: 'e', v: 'NAME' } },
  'name-case-held:text#1': { reason: 'macLongS', engine: { t: 'e', v: 'VALUE' } },
  'text-format:date#78': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'text-format:date#79': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'text-format:date#82': { reason: 'macMonthCase', engine: { t: 's', v: "Февраль" } },
  'text-format:date#83': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'text-format:date#84': { reason: 'macMonthCase', engine: { t: 's', v: "Январь" } },
  'lookup-own-held:LOOKUP#5': { reason: 'lookupReadZone', engine: { t: 'n', v: 0 } },
  'lookup-own-held:LOOKUP#7': { reason: 'lookupReadZone', engine: { t: 'n', v: 0 } },
  'lookup-own-held:LOOKUP#8': { reason: 'lookupReadZone', engine: { t: 'n', v: 0 } },
  'lookup-own-held:LOOKUP#9': { reason: 'lookupReadZone', engine: { t: 'n', v: 0 } },
  'stat-edges-held:VAR.P#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.0000668002774232626 } },
  'stat-edges-held:VAR.P#2': { reason: 'momentRounding', engine: { t: 'n', v: 0.0000668002774232626 } },
  'stat-edges-held:VAR.S#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.0000668069581190745 } },
  'stat-edges-held:STDEV.P#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.00817314366833611 } },
  'stat-edges-held:STDEV.P#2': { reason: 'momentRounding', engine: { t: 'n', v: 0.00817314366833611 } },
  'stat-edges-held:STDEV.S#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.00817355235617137 } },
  'stat-edges-held:DEVSQ#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.668002774232626 } },
  'stat-edges-held:DEVSQ#2': { reason: 'momentRounding', engine: { t: 'n', v: 0.668002774232626 } },
  'stat-edges-held:CORREL#1': { reason: 'momentRounding', engine: { t: 'n', v: 1 } },
  'stat-edges-held:PEARSON#1': { reason: 'momentRounding', engine: { t: 'n', v: 1 } },
  'stat-edges-held:RSQ#1': { reason: 'momentRounding', engine: { t: 'n', v: 1 } },
  'stat-edges-held:SLOPE#2': { reason: 'momentRounding', engine: { t: 'n', v: 99.9024390243902 } },
  'stat-edges-held:STEYX#1': { reason: 'momentRounding', engine: { t: 'n', v: 0 } },
  'stat-edges-held:DVAR#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.0000668069581190745 } },
  'stat-edges-held:DVARP#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.0000668002774232626 } },
  'stat-edges-held:DSTDEV#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.00817355235617137 } },
  'stat-edges-held:DSTDEVP#1': { reason: 'momentRounding', engine: { t: 'n', v: 0.00817314366833611 } },
  'stat-edges-held:CORREL#3': { reason: 'momentRange', engine: { t: 'n', v: 1 } },
  'stat-edges-held:PEARSON#3': { reason: 'momentRange', engine: { t: 'n', v: 1 } },
  'stat-edges-held:RSQ#3': { reason: 'momentRange', engine: { t: 'n', v: 1 } },
  'stat-edges-held:SLOPE#6': { reason: 'momentRange', engine: { t: 'n', v: 1e+200 } },
  'stat-edges-held:AVERAGE#6': { reason: 'momentRange', engine: { t: 'n', v: 1e+308 } },
  'stat-edges-held:VAR.P#6': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:VAR.S#5': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:STDEV.P#6': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:STDEV.S#5': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:DEVSQ#6': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:DAVERAGE#3': { reason: 'momentRange', engine: { t: 'n', v: 1e+308 } },
  'stat-edges-held:DVAR#3': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:COVARIANCE.P#2': { reason: 'momentRange', engine: { t: 'n', v: 2.22222222222222e+107 } },
  'stat-edges-held:COVARIANCE.S#4': { reason: 'momentRange', engine: { t: 'n', v: 3.33333333333333e+107 } },
  'stat-edges-held:CORREL#4': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:PEARSON#4': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:RSQ#4': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:SLOPE#7': { reason: 'momentRange', engine: { t: 'n', v: 0 } },
  'stat-edges-held:INTERCEPT#2': { reason: 'momentRange', engine: { t: 'n', v: 3.33333333333333e+307 } },
  'stat-edges-held:FORECAST.LINEAR#4': { reason: 'momentRange', engine: { t: 'n', v: 3.33333333333333e+307 } },
  'stat-edges-held:STEYX#2': { reason: 'momentRounding', engine: { t: 'n', v: 8.16496575965492e-9 } },
  'stat-edges-held:STEYX#3': { reason: 'momentRounding', engine: { t: 'n', v: 8.16496575965492e-9 } },
  'fit-offset-held:LINEST#1': { reason: 'linestMean', engine: { t: 'n', v: 0 } },
  'fit-offset-held:LINEST#4': { reason: 'linestMean', engine: { t: 'n', v: 0 } },
  'fit-offset-held:LINEST#11': { reason: 'linestMean', engine: { t: 'n', v: 0.668002774232626 } },
  'fit-offset-held:TREND#1': { reason: 'linestMean', engine: { t: 'n', v: 0.030029296875 } },
  'fit-offset-held:TREND#2': { reason: 'linestMean', engine: { t: 'n', v: 0.030029296875 } },
  'fit-offset-held:GROWTH#2': { reason: 'growthRange', engine: { t: 'n', v: 313.83972099924 } },
  'fit-offset-held:GROWTH#3': { reason: 'growthRange', engine: { t: 'n', v: 285.309168434641 } },
}

/** A case passes: it matches Excel, or it is in KNOWN (or OPEN), differs from Excel and gives the pinned answer (`samePin`). */
export function verdict(key: string, fn: string | undefined, excel: Val, engine: Val): { ok: boolean; why?: string } {
  const known = KNOWN[key] ?? OPEN[key]
  const same = matches(fn, excel, engine)
  if (!known) return { ok: same }
  if (same) return { ok: false, why: `matches Excel, remove it from ${key in KNOWN ? 'KNOWN' : 'OPEN'}` }
  return samePin(known.engine, engine) ? { ok: true } : { ok: false, why: `KNOWN pins ${show(known.engine)}` }
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
