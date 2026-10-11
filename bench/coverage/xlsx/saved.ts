/**
 * The case books Excel saved (`xlsx saved`): each part of a case set, as the harness sends it to Excel (`book.ts`),
 * here as the book it stands for: the set's own sheets, the data sheet `D` with its formulas as written (`1/0`,
 * `CHAR(9)&"abc"&CHAR(10)`), and the case sheet `C` with the cases in column A and Excel's wrappers beside them.
 * Formulas lose their file prefixes by a plain replacement, not by the reader's code.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { CASE_SHEET, READ_COLUMNS } from '../book.js'
import { DATA_SHEET, dataCells } from '../data.js'
import type { Probe, SheetCells } from '../probes.js'

/** A formula without the file's prefixes of newer functions and of LET's names. */
export const plainFormula = (f: string) => f.replace(/_xlfn\.(_xlws\.)?|_xlpm\./g, '')

/** The parts of a set as `xlsx saved` sends them: PART cases each, `<set>-<k>`. */
export { PART } from '../book.js'

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
    c[`B${r}`] = { formula: READ_COLUMNS.type(`A${r}`) }
    c[`C${r}`] = { formula: READ_COLUMNS.num(`A${r}`) }
    c[`D${r}`] = { formula: READ_COLUMNS.text(`A${r}`) }
    c[`E${r}`] = { value: p.id }
  })
  c[`A${probes.length + 1}`] = { value: 'end' }
  return { sheets: [...own, { name: DATA_SHEET, cells: d }, { name: CASE_SHEET, cells: c }] }
}
