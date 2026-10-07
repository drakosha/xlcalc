/**
 * The case books Excel saved (`xlsx saved`): each part of a case set, as the harness sends it to Excel (`book.ts`),
 * here as the book it stands for: the set's own sheets, the data sheet `D` with its formulas as written (`1/0`,
 * `CHAR(9)&"abc"&CHAR(10)`), and the case sheet `C` with the cases in column A and Excel's wrappers beside them.
 * Formulas lose their file prefixes by a plain replacement, not by the reader's code.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { CASE_SHEET } from '../book.js'
import { DATA_SHEET, dataCells } from '../data.js'
import type { Probe, SheetCells } from '../probes.js'

/** A formula without the file's prefixes of newer functions and of LET's names. */
export const plainFormula = (f: string) => f.replace(/_xlfn\.(_xlws\.)?|_xlpm\./g, '')

/** The parts of a set as `xlsx saved` sends them: 400 cases each, `<set>-<k>`. */
export const PART = 400

export function sentBook(probes: Probe[], sheets: SheetCells): BookInput {
  const own = Object.entries(sheets).map(([name, cells]) => ({
    name,
    cells: Object.fromEntries(Object.entries(cells).map(([a, v]): [string, CellInput] => [a, typeof v === 'object' ? { formula: plainFormula(v.formula) } : { value: v }])),
  }))
  const d: Record<string, CellInput> = {}
  for (const [addr, v] of dataCells()) d[addr] = typeof v !== 'object' ? { value: v } : 'date' in v ? { value: v.date } : { formula: v.formula }
  const c: Record<string, CellInput> = {}
  probes.forEach((p, i) => {
    const r = i + 1
    c[`A${r}`] = { formula: plainFormula(p.formula) }
    c[`B${r}`] = { formula: `IF(ISERROR(A${r}),ERROR.TYPE(A${r}),IF(ISNUMBER(A${r}),0,IF(ISLOGICAL(A${r}),IF(A${r},-1,-2),IF(ISTEXT(A${r}),-3,-4))))` }
    c[`C${r}`] = { formula: `IF(ISNUMBER(A${r}),A${r},0)` }
    c[`D${r}`] = { formula: `IF(ISTEXT(A${r}),A${r},"")` }
    c[`E${r}`] = { value: p.id }
  })
  c[`A${probes.length + 1}`] = { value: 'end' }
  return { sheets: [...own, { name: DATA_SHEET, cells: d }, { name: CASE_SHEET, cells: c }] }
}
