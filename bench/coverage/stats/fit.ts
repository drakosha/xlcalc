/**
 * Workbook Statistics, the first set, predicted before Excel ran it (`fit.predicted.json`): what a cell with data is
 * (constants of each kind, an error, empty text, a formula answering empty text, a cell with a style alone, a hyperlink alone, a
 * spill, one blocked, a large one, a legacy array formula, a merged area), where the sheet ends (a formatted cell, a
 * row and a column of their own past the data, a table), an empty sheet, hidden sheets, notes and threads of comments.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { StatsCase } from './cases.js'

const one = (cells: Record<string, CellInput>, more: Partial<BookInput['sheets'][number]> = {}): BookInput => ({ sheets: [{ name: 'S', cells, ...more }] })

export function buildStatsCases(): StatsCase[] {
  return [
    { id: 'empty', book: one({}), sheet: 'S' },
    {
      id: 'kinds',
      book: one({ A1: { value: 1 }, A2: { value: 'a' }, A3: { value: true }, A4: { formula: '#N/A' }, A5: { value: { date: '2025-01-31' } }, B1: { value: '' }, B2: { formula: '""' }, B3: { formula: 'Z99' }, B4: { formula: '1/0' } }),
      sheet: 'S',
    },
    { id: 'styled', book: { styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1 }, H20: { style: 0 } } }] }, sheet: 'S' },
    { id: 'row-own', book: one({ B2: { value: 1 } }, { rows: { '30': { height: 30 } } }), sheet: 'S' },
    { id: 'column-own', book: one({ B2: { value: 1 } }, { columns: { K: { width: 20 } } }), sheet: 'S' },
    { id: 'link-only', book: one({ A1: { value: 1 }, C3: { link: { url: 'https://example.com/' } } }), sheet: 'S' },
    { id: 'spill', book: one({ A1: { formula: 'SEQUENCE(3,2)' } }), sheet: 'S' },
    { id: 'spill-blocked', book: one({ A1: { formula: 'SEQUENCE(3)' }, A3: { value: 'x' } }), sheet: 'S' },
    { id: 'spill-large', book: one({ A1: { formula: 'SEQUENCE(5000)' } }), sheet: 'S' },
    { id: 'legacy-array', book: one({ B2: { formula: 'ROW(B2:B4)*2', array: 'B2:B4' } }), sheet: 'S' },
    { id: 'merged', book: one({ A1: { value: 'm' }, D4: { value: 1 } }, { merges: ['A1:C2'] }), sheet: 'S' },
    {
      id: 'tables',
      book: one({ A1: { value: 'a' }, B1: { value: 'b' }, A2: { value: 1 }, E1: { value: 'c' } }, { tables: [{ name: 'Left', range: 'A1:B5', columns: [{ name: 'a' }, { name: 'b' }] }, { name: 'Right', range: 'E1:E3', columns: [{ name: 'c' }] }] }),
      sheet: 'S',
    },
    {
      id: 'sheets',
      book: {
        sheets: [
          { name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1+T!A1' } } },
          { name: 'T', cells: { A1: { value: 2 }, A2: { formula: 'SEQUENCE(2)' }, C3: { note: { text: 'n' } } } },
          { name: 'H', hidden: true, cells: { A1: { value: 3 }, B2: { formula: 'A1*2' }, D4: { note: { text: 'h' } } } },
        ],
      },
      sheet: 'T',
    },
    {
      id: 'notes',
      book: one({ A1: { value: 1, note: { text: 'on a value' } }, B2: { note: { text: 'alone', author: 'Ann' } }, C3: { value: 'x', note: { text: 'shown', shown: true } } }),
      sheet: 'S',
      threads: [{ sheet: 'S', cell: 'E5', comments: [{ author: 'A', text: 'thread' }, { author: 'B', text: 'reply' }] }],
    },
  ]
}
