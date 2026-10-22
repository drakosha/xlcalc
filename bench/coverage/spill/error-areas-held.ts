/**
 * Held out: array formulas of the old kind over N1:O3 whose formula fails to prepare (a sheet not in the book, an unknown
 * function, a syntax error) or fails at run time (the control), as a file brings them; what the cells of the area show,
 * what formulas reading them see, whether a dynamic array beside the area spills into it, and whether a value is taken
 * in part of the area. The engine's answers predicted before Excel ran them (`error-areas-held.predicted.json`, from
 * `predict.ts`), with the ids Excel answered otherwise in `MISSED`. A literal past the largest number (`1E400`) as a file
 * brings it, alone and inside a function, is read beside them.
 */
import { existsSync, readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { CellInput } from '../../../src/contract.js'
import { one, type CellSpec, type SpillCase } from './cases.js'

const FILE = new URL('./error-areas-held.predicted.json', import.meta.url)
const PREDICTED = (existsSync(FILE) ? JSON.parse(readFileSync(FILE, 'utf8')) : {}) as Record<string, Record<string, Val>>

/** The ids whose prediction Excel did not confirm. */
export const MISSED: readonly string[] = [
  // Excel shows the error in every cell of the area and spills nothing into it, refusing the edit too; the engine had the error in N1 alone
  ...['sheet', 'function'].flatMap((k) => [`${k}-area`, `${k}-spill-into`, `${k}-type-into`]),
  // Excel did not load the book: a syntax error, and a literal past the largest number, reject the whole file
  'syntax-area',
  'syntax-spill-into',
  'syntax-type-into',
  'literal-alone',
  'literal-inside',
]

const DATA: Record<string, CellSpec> = { A1: 1, A2: 2, B1: 10, B2: 20, B3: 30, C1: 100, C2: 200, C3: 300, D1: 1000, D2: 2000, D3: 3000 }
const arr = (formula: string): CellInput => ({ formula, array: 'N1:O3' })

/** Formulas reading the area: blanks, values, errors, one cell. */
const READERS: Record<string, CellSpec> = { I1: '=COUNTBLANK(N1:O3)', I2: '=COUNTA(N1:O3)', I3: '=ISERROR(O2)', I4: '=ISBLANK(O2)', I5: '=SUMPRODUCT(--ISERROR(N1:O3))' }

const FORMULAS: [string, string, string][] = [
  ['sheet', 'a sheet not in the book', 'TRANSPOSE(Nope!B2:D3)'],
  ['function', 'an unknown function', 'FOO(A1:B3)'],
  ['syntax', 'a parenthesis left open', 'SUM(A1:B3'],
  ['runtime', 'the control: an error at run time', 'TRANSPOSE(A1:B2)/0'],
]

const held = (id: string, about: string, cells: Record<string, CellSpec>, ranges: string[], extra: Partial<SpillCase> = {}): SpillCase =>
  one(id, about, { ...DATA, ...cells }, ranges, { ...extra, ...(PREDICTED[id] ? { predicted: PREDICTED[id] } : {}) })

export function buildSpillCases(): SpillCase[] {
  const out: SpillCase[] = []
  for (const [k, about, f] of FORMULAS) {
    out.push(held(`${k}-area`, `${about} over N1:O3: the area and formulas reading it`, { N1: arr(f), ...READERS }, ['N1:O3', 'I1:I5']))
    out.push(held(`${k}-spill-into`, `${about} over N1:O3, a dynamic array at M2 spilling across its middle row`, { N1: arr(f), M2: '=SEQUENCE(1,3)' }, ['M1:O3']))
    out.push(held(`${k}-type-into`, `${about} over N1:O3, a value typed into O2`, { N1: arr(f) }, ['N1:O3'], { edits: [{ sheet: 'S', cell: 'O2', value: 5 }] }))
  }
  out.push(held('literal-alone', 'a literal past the largest number, alone', { F1: '=1E400', F2: '=-1E400' }, ['F1:F2']))
  out.push(held('literal-inside', 'a literal past the largest number inside a function', { F1: '=ISNUMBER(1E400)', F2: '=IFERROR(1E400,"e")' }, ['F1:F2']))
  return out
}
