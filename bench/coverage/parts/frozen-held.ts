/**
 * Held out after `struct-held`: frozen panes when lines above their split are deleted past it or inserted several at
 * once, predicted by the engine before Excel ran them (`frozen-held.predicted.json`). `struct-held` found that deleting
 * every frozen row or column leaves the first one frozen; the engine now keeps the split at one line at least.
 */
import type { BookInput, CellInput, StructureOp } from '../../../src/contract.js'
import type { PartsCase, PartsCases } from './cases.js'

/** S: values in A1:E12, rows 1:3 and columns A:B frozen; T: its first row and column frozen. */
function frozenBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 12; r++) for (const c of 'ABCDE') cells[`${c}${r}`] = { value: `${c}${r}` }
  return {
    sheets: [
      { name: 'S', cells, frozen: { rows: 3, columns: 2 } },
      { name: 'T', cells: { A1: { value: 1 } }, frozen: { rows: 1, columns: 1 } },
    ],
  }
}

export function buildPartsCases(): PartsCases {
  const cases: PartsCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'frozen', op })
  add('delete-rows-1-6', { op: 'deleteRows', sheet: 'S', row: 1, count: 6 })
  add('delete-rows-2-10', { op: 'deleteRows', sheet: 'S', row: 2, count: 9 })
  add('delete-rows-3-4', { op: 'deleteRows', sheet: 'S', row: 3, count: 2 })
  add('delete-cols-a-d', { op: 'deleteColumns', sheet: 'S', column: 'A', count: 4 })
  add('delete-col-b', { op: 'deleteColumns', sheet: 'S', column: 'B' })
  add('insert-rows-1x3', { op: 'insertRows', sheet: 'S', row: 1, count: 3 })
  add('insert-cols-b-x2', { op: 'insertColumns', sheet: 'S', column: 'B', count: 2 })
  add('t-delete-row-1', { op: 'deleteRows', sheet: 'T', row: 1 })
  add('t-delete-col-a', { op: 'deleteColumns', sheet: 'T', column: 'A' })
  return { books: { frozen: frozenBook() }, cases }
}
