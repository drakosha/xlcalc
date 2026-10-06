/**
 * Two spills over one cell: which one Excel lets spill. A column spill from row 1 and a row spill from column
 * A cross at one cell; the anchors' places vary, the rectangles are read in either order, and one pair sits on
 * a sheet of its own after another sheet. In some pairs one anchor reads the other.
 */
import { formatAddress } from '../../../src/model.js'
import { one, sheet, type SpillCase } from './cases.js'

export function buildSpillCases(): SpillCase[] {
  const out: SpillCase[] = []
  // a column of 5 from (1, c) and a row of 5 from (r, A): they cross at (r, c)
  for (const c of [2, 3, 4])
    for (const r of [2, 3, 4]) {
      const col = formatAddress(0, c - 1)
      const row = formatAddress(r - 1, 0)
      out.push(one(`cross-${col}-${row}`, `a column from ${col}, a row from ${row}`, { [col]: '=SEQUENCE(5)', [row]: '=SEQUENCE(1,5)' }, ['A1:F6']))
    }
  out.push(one('cross-C1-A2-row-first', 'the row anchor read before the column anchor', { C1: '=SEQUENCE(3)', A2: '=SEQUENCE(1,3)' }, ['A2:D4', 'A1:D1']))
  out.push(one('cross-B1-A3-col-first', 'the column anchor read before the row anchor', { B1: '=SEQUENCE(4)', A3: '=SEQUENCE(1,4)' }, ['B1:B5', 'A1:A5', 'C1:E5']))
  out.push(one('cross-far', 'a column from E1 and a row from A4', { E1: '=SEQUENCE(6)', A4: '=SEQUENCE(1,6)' }, ['A1:G7']))
  out.push(one('cross-long-row', 'a short column from C1, a long row from A2', { C1: '=SEQUENCE(2)', A2: '=SEQUENCE(1,8)' }, ['A1:I3']))
  out.push(one('cross-long-col', 'a long column from C1, a short row from A2', { C1: '=SEQUENCE(8)', A2: '=SEQUENCE(1,3)' }, ['A1:D9']))
  // one anchor reads the other: Excel computes the one read first
  for (const [col, row] of [['B1', 'A3'], ['D1', 'A4'], ['C1', 'A2']]) {
    out.push(one(`dep-${col}-${row}-row-reads`, `the row from ${row} reads the column anchor ${col}`, { [col]: '=SEQUENCE(5)', [row]: `=SEQUENCE(1,5)+0*${col}` }, ['A1:F6']))
    out.push(one(`dep-${col}-${row}-col-reads`, `the column from ${col} reads the row anchor ${row}`, { [col]: `=SEQUENCE(5)+0*${row}`, [row]: '=SEQUENCE(1,5)' }, ['A1:F6']))
  }
  out.push({
    id: 'cross-second-sheet',
    about: 'the pair of cross-B1-A3 on a second sheet',
    book: { sheets: [sheet('P', { A1: 1 }), sheet('S', { B1: '=SEQUENCE(4)', A3: '=SEQUENCE(1,4)' })] },
    read: [{ sheet: 'S', range: 'A1:E5' }],
  })
  return out
}
