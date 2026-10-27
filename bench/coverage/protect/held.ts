/**
 * Protection held out, with whole lines: a column and a row unlocked as `writeXlsx` writes them, Excel unlocking a
 * column, rows inserted below an unlocked row, a column copied and a row's formats cleared, and a file `writeKept` wrote
 * after a column's width changed (its `<cols>` written again). Predicted before Excel ran it (`held.predicted.json`).
 */
import type { BookInput } from '../../../src/contract.js'
import type { ProtectCase, ProtectCases } from './cases.js'

/** S: B unlocked, row 4 unlocked and hidden, a few values; protected with what Excel allows by default. */
const lined = (protection?: BookInput['sheets'][number]['protection']): BookInput => ({
  styles: [{ protection: { locked: false } }, { protection: { locked: false, hidden: true } }],
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 1 }, B1: { value: 2 }, C1: { value: 3 }, B2: { formula: 'A1+C1' }, A4: { value: 4 } },
      columns: { B: { style: 0, width: 12 } },
      rows: { 4: { style: 1 } },
      ...(protection ? { protection } : {}),
    },
  ],
})

const READ: ProtectCase['read'] = [['S', 'A1:E6']]

export function buildProtectCases(): ProtectCases {
  const cases: ProtectCase[] = [
    { id: 'lines-written', book: 'guarded', read: READ },
    { id: 'column-unlocked', book: 'plain', steps: [{ lock: { sheet: 'S', range: 'D:D' }, locked: false, hidden: false }], read: READ },
    { id: 'rows-below', book: 'plain', steps: [{ insertRows: { sheet: 'S', range: '5:6' } }], read: [['S', 'A3:C7']] },
    { id: 'column-copied', book: 'plain', steps: [{ copy: { sheet: 'S', range: 'B:B' }, to: { sheet: 'S', range: 'E:E' } }], read: READ },
    { id: 'row-cleared', book: 'plain', steps: [{ clearFormats: { sheet: 'S', range: '4:4' } }], read: READ },
    { id: 'kept-width', book: 'guarded', kept: { layout: [{ sheet: 'S', columns: 'B', set: { width: 20 } }], cells: [{ sheet: 'S', cell: 'B1', input: { value: 9 } }] }, read: READ },
  ]
  return { books: { guarded: lined({}), plain: lined() }, cases }
}
