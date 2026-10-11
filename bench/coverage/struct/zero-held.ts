/**
 * Held out: rows inserted under a row of height 0, predicted by the engine before Excel ran them (`zero-held.predicted.json`,
 * from `predict.ts`). A row of height 0 (written `ht="0"`, which Excel takes as hidden), one with an orange fill and the
 * format `0.00` too, a span of two, one also marked hidden; beside them a row hidden at height 30, and a row inserted at
 * a row of height 0 rather than under it. The engine gives an inserted row the entry of the row above it but `hidden`,
 * so under a row of height 0 the new row has height 0 and is hidden.
 */
import type { StructCase, StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'

const f = (formula: string): StoredCell => ({ formula })

/** S: values in A1:B16, the sums in D1:D2; rows 3, 6 (orange, `0.00`), 9:10 and 13 (also hidden) at height 0, 15 hidden at 30. T: a sum over S. */
function zeroBook(): StoredBook {
  const cells: Record<string, StoredCell> = { D1: f('SUM(A1:A16)'), D2: f('SUBTOTAL(109,A1:A16)') }
  for (let r = 1; r <= 16; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { value: r * 10 } })
  return {
    sheets: [
      { name: 'S', cells, rows: { 3: { height: 0 }, 6: { height: 0, style: 0, format: '0.00' }, '9:10': { height: 0 }, 13: { height: 0, hidden: true }, 15: { height: 30, hidden: true } } },
      { name: 'T', cells: { A1: f('SUM(S!A1:A16)') } },
    ],
    styles: [{ fill: { color: '#FFC000' } }],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const insert = (id: string, row: number, count?: number) => cases.push({ id, book: 'zero', op: { op: 'insertRows', sheet: 'S', row, ...(count ? { count } : {}) } })
  insert('under-3', 4)
  insert('under-3x2', 4, 2)
  insert('at-3', 3)
  insert('under-6-styled', 7)
  insert('inside-9-10', 10)
  insert('under-9-10', 11)
  insert('under-13-hidden', 14)
  insert('under-15-hidden-30', 16)
  return { books: { zero: zeroBook() }, cases }
}
