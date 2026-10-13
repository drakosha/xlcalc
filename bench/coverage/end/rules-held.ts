/**
 * Held out, predicted before Excel ran it by `getEnd` of the day. The questions: where `End` goes from a cell inside a
 * block, at its edge and from an empty cell, to the next cell with a value and to the sheet's edge (row 1,048,576,
 * column XFD, A1); blocks one empty cell apart; which cells count as empty (an empty text from a formula, an empty
 * text in the file and an apostrophe alone typed, a cell with a look alone, also of the sheet's looks, a number format
 * alone, a note alone, a hyperlink alone, an error, zero, a space, FALSE, a value a format hides); the cells of a spill,
 * its anchor, a spill blocked, an old-kind array and a spill of the whole column; merged areas from each side and from
 * inside; rows and columns hidden by hand and rows a filter hid (whether `End` goes through them and stops on them),
 * and every line hidden past some row and column.
 */
import type { CellInput, EndDirection } from '../../../src/contract.js'
import { collector, type EndBook, type EndCase } from './cases.js'

/** Cells holding their address as text. */
const named = (...xs: string[]): Record<string, CellInput> => Object.fromEntries(xs.map((a) => [a, { value: a }]))
/** The cells of a column from row `a` to row `b`. */
const run = (col: string, a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => `${col}${a + i}`)

/** The kinds of cell whose emptiness is asked, by column; each column holds a value at rows 1, 10, 20, 22 and 23, the kind at 5 and 21. */
const KINDS: Record<string, CellInput> = {
  B: { formula: '""' },
  C: { value: '' },
  // typed as an apostrophe alone once Excel opened the book (`typed`)
  D: { value: '' },
  E: { style: 1 },
  F: { numberFormat: '0.00' },
  // the sheet's looks (`looks`) are over G1:G30
  G: {},
  H: { formula: 'NA()' },
  I: { value: 0 },
  J: { value: ' ' },
  K: { formula: 'Z1000' },
  L: { note: { text: 'a note' } },
  M: { value: false },
  N: { link: { url: 'https://example.com/' } },
  O: { value: 5, numberFormat: ';;;' },
}

const kinds = (): Record<string, CellInput> =>
  Object.fromEntries(
    Object.entries(KINDS).flatMap(([col, kind]) => [
      ...[1, 10, 20, 22, 23].map((r) => [`${col}${r}`, { value: `${col}${r}` }]),
      ...(Object.keys(kind).length ? [5, 21].map((r) => [`${col}${r}`, kind]) : []),
    ]),
  )

/** Rows and columns hidden by hand, or the rows a filter hides: 5, 12 and 20 to 22. */
const HIDDEN_ROWS = [5, 12, 20, 21, 22]

export const BOOKS: Record<string, EndBook> = {
  // blocks of a column and of a row, one empty cell apart and further; A1:A10 beside B1:B500 as the sheet component met it
  blocks: {
    cells: {
      ...Object.fromEntries(run('A', 1, 10).map((a, i) => [a, { value: i + 1 }])),
      ...Object.fromEntries(run('B', 1, 500).map((a, i) => [a, { value: i + 1 }])),
      ...named(...run('D', 2, 4), ...run('D', 6, 8), 'D10', 'D12', 'D13', 'A700', 'B700', 'C700', 'E700', 'G700', 'H700'),
    },
  },
  kinds: {
    cells: kinds(),
    styles: [{}, { font: { bold: true } }],
    looks: [{ range: 'G1:G30', style: 1 }],
    typed: { D5: "'", D21: "'" },
  },
  // spills down and across, a spill blocked, an old-kind array, a spill with an empty text, a spill of the whole column
  spill: {
    cells: {
      A1: { formula: 'SEQUENCE(5)' },
      A7: { value: 'x' },
      C3: { formula: 'SEQUENCE(1,3)' },
      G3: { value: 'x' },
      H1: { formula: 'SEQUENCE(4)' },
      H3: { value: 'block' },
      H8: { value: 'x' },
      J1: { formula: 'SEQUENCE(3)', array: 'J1:J3' },
      J5: { value: 'x' },
      L1: { formula: 'IF(SEQUENCE(4)=2,"",1)' },
      L6: { value: 'x' },
      N1: { formula: 'P:P*0' },
    },
  },
  // merged areas across, down and both, with a value and empty
  merged: {
    cells: named('A2', 'B2', 'E2', 'F2', 'H2', 'B5', 'G5', 'B7', 'B8', 'B11', 'B13', 'D8', 'D12', 'E12', 'G7', 'G11'),
    merges: ['B2:D2', 'B5:D5', 'B8:B10', 'D8:E10', 'G8:H9'],
  },
  // rows 5, 12 and 20 to 22 and columns M, P and Q hidden by hand
  hidden: {
    cells: named(
      ...run('A', 1, 10).filter((a) => a !== 'A5'),
      'B1', 'B12', 'B15',
      ...run('C', 6, 12), 'C16',
      'D19', 'D23', 'D24', 'D26',
      'K12', 'L12', 'N12',
      'K30', 'L30', 'N30', 'O30',
      'K31', 'P31', 'S31',
      'K32', 'L32', 'M32',
    ),
    rows: Object.fromEntries(HIDDEN_ROWS.map((r) => [String(r), { hidden: true }])),
    columns: { M: { hidden: true }, 'P:Q': { hidden: true } },
  },
  // the same rows hidden by a filter over A1:A30 keeping `keep`
  filter: {
    cells: {
      A1: { value: 'key' },
      ...Object.fromEntries(run('A', 2, 30).map((a, i) => [a, { value: HIDDEN_ROWS.includes(i + 2) ? 'drop' : 'keep' }])),
      ...named(...run('B', 1, 10).filter((a) => a !== 'B5'), 'C1', 'C12', 'C15', ...run('D', 6, 12), 'D16', 'E19', 'E23', 'E24', 'E26'),
    },
    filter: { range: 'A1:A30', field: 1, keep: 'keep' },
  },
  // every row from 20 and every column from K hidden, a value in a hidden row and in a hidden column
  tail: {
    cells: named('A1', 'A2', 'A3', 'A30', 'C1', 'D1', 'M1'),
    rows: { '20:1048576': { hidden: true } },
    columns: { 'K:XFD': { hidden: true } },
  },
  // values at the sheet's edges only
  edges: {
    cells: named('A1048576', 'XFD1'),
  },
}

export function buildEndCases(): EndCase[] {
  const { out, add } = collector()
  const go = (group: string, book: string, list: string) => {
    for (const x of list.split(' ')) {
      const [cell, d] = x.split(':')
      add(group, book, cell, ({ u: 'up', d: 'down', l: 'left', r: 'right' } as Record<string, EndDirection>)[d])
    }
  }
  go('blocks', 'blocks', 'A1:d A5:d A10:d A10:u A1:u A11:d A11:u B1:d B500:d B700:d B700:u D1:d D2:d D4:d D5:d D5:u D8:d D10:d D13:d D1048576:d D1048576:u E5:d E5:u')
  go('blocks', 'blocks', 'A700:r C700:r E700:r G700:r H700:r XFD700:r XFD700:l H700:l D700:l D700:r F700:r A700:l A3:r B3:r D3:r XFD3:l')
  for (const col of Object.keys(KINDS)) go('kinds', 'kinds', `${col}1:d ${col}20:d`)
  go('spill', 'spill', 'A1:d A3:d A5:d A5:u A6:u A3:u A3:r C3:r D3:r E3:r G3:l E3:l H1:d H2:d H3:u H3:d J1:d J3:d L1:d L2:u')
  go('spill', 'spill', 'N1:d N1048576:u N500:u L1:r N1:r N5:l XFD7:l P7:l')
  go('merged', 'merged', 'A2:r B2:r C2:r D2:r E2:l F2:l G2:l H2:l A5:r B5:r D5:r G5:l F5:l')
  go('merged', 'merged', 'B7:d B8:d B9:d B11:u B12:u B13:u D12:u E12:u D7:d E7:d C9:r G7:d')
  go('hidden', 'hidden', 'A1:d A10:u B1:d B11:d B15:u C6:d D19:d D24:u K12:r K30:r O30:l K31:r S31:l K32:r')
  go('filter', 'filter', 'A1:d A30:u A30:d B1:d B10:u C1:d C15:u D6:d E19:d E24:u')
  go('tail', 'tail', 'A3:d B5:d A19:d B19:d D1:r J1:r A1048576:u XFD1:l')
  go('edges', 'edges', 'A1:d A1048576:u A1048576:d A1:r XFD1:l XFD1:r XFD1:u A1:u A1:l XFD1048576:r XFD1048576:d XFD1048576:u XFD1048576:l B2:d B2:r')
  return out
}
