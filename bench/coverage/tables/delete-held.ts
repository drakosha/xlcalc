/**
 * Held out: rows deleted that take one table whole and part of another, predicted by the engine before Excel took them
 * (`delete-held.predicted.json`, from `predict.ts`): a table below or beside the other gone with some of the other's
 * data rows, its last rows, its totals row, all its data, its first rows; the other's header row among the rows; a
 * table without a header row cut; columns deleted the same way. Formulas over both tables are on Other, which no
 * deletion touches, and below them on S.
 */
import type { StructureOp } from '../../../src/contract.js'
import { book, on, sales, table, type CellSpec, type TableCase } from './cases.js'

/** Formulas over Sales (column A) and over the other table `t` (column B) on Other. */
const over = (t: string, x: string, y: string, cells: string): Record<string, CellSpec> => ({
  A1: '=SUM(Sales[Qty])',
  A2: '=ROWS(Sales[#All])',
  A3: '=SUM(Sales[[Qty]:[Price]])',
  A4: '=COLUMNS(Sales[#All])',
  A5: '=SUM(S!B2:B5)',
  A6: '=Sales[[#Headers],[Price]]',
  B1: `=SUM(${t}[${y}])`,
  B2: `=ROWS(${t}[#All])`,
  B3: `=SUM(${t}[[${x}]:[${y}]])`,
  B4: `=SUM(S!${cells})`,
  B5: `=${t}[[#Headers],[${x}]]`,
})

/** A formula over each table on S, below them. */
const under = (t: string, y: string): Record<string, CellSpec> => ({ A14: '=SUM(Sales[Price])', B14: `=SUM(${t}[${y}])` })
const READ = [...on('S', 'A1:H14'), ...on('Other', 'A1:B6')]

const rows = (row: number, count: number): StructureOp => ({ op: 'deleteRows', sheet: 'S', row, count })

/** Sales on S!A1:D5; another table below it (A8) or beside it (F at `at`), with columns x and y, or a and b. */
const below = (at = 'A8') => table('Below', at, ['x', 'y'], [[1, 2], [3, 4]])
const side = (at: string, n = 1) => table('Side', at, ['a', 'b'], Array.from({ length: n }, (_, i) => [i + 1, 10 * (i + 1)]))

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  // the other table, its two columns and its cells
  const add = (id: string, about: string, op: StructureOp, [t, x, y, cells]: [string, string, string, string], ...tables: ReturnType<typeof table>[]) =>
    cases.push({ id, about, book: book(['S', under(t, y), ...tables], ['Other', over(t, x, y, cells)]), steps: [{ op }], read: READ })
  const totals = sales()
  totals.table.range = 'A1:D6'
  totals.table.totals = true
  totals.table.columns[1].totals = { function: 'sum' }
  totals.cells.B6 = { formula: 'SUBTOTAL(109,Sales[Qty])' }

  add('below-last-rows', "the table below gone with Sales's last two data rows and the rows between", rows(4, 7), ['Below', 'x', 'y', 'A8:B10'], sales(), below())
  add('below-all-data', "the table below gone with all of Sales's data rows: Sales keeps one row", rows(2, 9), ['Below', 'x', 'y', 'A8:B10'], sales(), below())
  add('below-totals', "the table below gone with Sales's totals row alone", rows(6, 5), ['Below', 'x', 'y', 'A8:B10'], totals, below())
  add('below-header', "Sales gone with the header row and the first data row of the table below: a header row without its table", rows(1, 9), ['Below', 'x', 'y', 'A8:B10'], sales(), below())
  const bare = table('Below', 'A7', ['x', 'y'], [[1, 2], [3, 4], [5, 6]], { headers: false })
  add('below-no-header', 'Sales gone with the first two rows of a table below it without a header row', rows(1, 8), ['Below', 'x', 'y', 'A7:B9'], sales(), bare)
  add('side-inside', "a table beside Sales gone with two of Sales's data rows", rows(3, 2), ['Side', 'a', 'b', 'F3:G4'], sales(), side('F3'))
  add('side-first', "a table beside Sales gone with Sales's first two data rows", rows(2, 2), ['Side', 'a', 'b', 'F2:G3'], sales(), side('F2'))
  add('side-past-end', "a table beside Sales, reaching past its end, gone with Sales's last two data rows", rows(4, 3), ['Side', 'a', 'b', 'F4:G6'], sales(), side('F4', 2))
  add('side-header', "a table beside Sales gone with Sales's header row and first two data rows", rows(1, 3), ['Side', 'a', 'b', 'F1:G3'], sales(), side('F1', 2))
  const cols: StructureOp = { op: 'deleteColumns', sheet: 'S', column: 'D', count: 4 }
  add('columns-side', "a table right of Sales gone with Sales's last column, the calculated one", cols, ['Side', 'a', 'b', 'F1:G3'], sales(), side('F1', 2))
  return cases
}
