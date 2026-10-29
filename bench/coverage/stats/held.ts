/**
 * Workbook Statistics held out after `fit`, predicted by the engine before Excel ran it (`held.predicted.json`): the end
 * of sheet from notes alone, data on two edges, a merged area, hidden rows and columns, the sheet's last cell, a
 * table's header alone; arrays over two dimensions, a spill over a cell with a note alone; a thread without notes; and
 * the counts of a workbook whose active sheet is its last.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { StatsCase } from './cases.js'

const one = (cells: Record<string, CellInput>, more: Partial<BookInput['sheets'][number]> = {}): BookInput => ({ sheets: [{ name: 'S', cells, ...more }] })

export function buildStatsCases(): StatsCase[] {
  return [
    { id: 'note-far', book: one({ A1: { value: 1 }, H30: { note: { text: 'far' } } }), sheet: 'S' },
    { id: 'two-edges', book: one({ B9: { value: 1 }, E2: { formula: '""' } }), sheet: 'S' },
    { id: 'merged-wide', book: one({ A1: { value: 'm' } }, { merges: ['A1:D4'] }), sheet: 'S' },
    { id: 'hidden-lines', book: one({ A1: { value: 1 }, C8: { value: 2 }, F3: { value: 3 } }, { rows: { '8': { hidden: true } }, columns: { F: { hidden: true } } }), sheet: 'S' },
    { id: 'last-cell', book: one({ A1: { value: 1 }, XFD1048576: { value: 2 } }), sheet: 'S' },
    { id: 'table-header', book: one({ C3: { value: 'a' }, D3: { value: 'b' } }, { tables: [{ name: 'Empty', range: 'C3:D8', columns: [{ name: 'a' }, { name: 'b' }] }] }), sheet: 'S' },
    { id: 'arrays-2d', book: one({ A1: { formula: 'SEQUENCE(100,30)' }, AH2: { formula: 'ROW(AH2:AI3)+COLUMN(AH2:AI3)', array: 'AH2:AI3' } }), sheet: 'S' },
    { id: 'spill-over-note', book: one({ A1: { formula: 'SEQUENCE(4)' }, A3: { note: { text: 'under' } } }), sheet: 'S' },
    { id: 'thread-only', book: one({ B2: { value: 1 } }), sheet: 'S', threads: [{ sheet: 'S', cell: 'D6', comments: [{ author: 'A', text: 'only' }] }] },
    {
      id: 'last-active',
      book: {
        sheets: [
          { name: 'A', cells: { A1: { value: 'x' }, B1: { value: 'y' }, A2: { value: 1 } }, tables: [{ name: 'One', range: 'A1:B2', columns: [{ name: 'x' }, { name: 'y' }] }] },
          { name: 'B', hidden: true, cells: { A1: { formula: 'SEQUENCE(3,3)' } } },
          { name: 'C', cells: { C2: { formula: 'A!A2*2' }, E4: { note: { text: 'n' } } } },
        ],
      },
      sheet: 'C',
    },
  ]
}
