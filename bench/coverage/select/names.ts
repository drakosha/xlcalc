/**
 * Create Names from Selection (`Range.CreateNames`), predicted by the engine before Excel ran it
 * (`names.predicted.json`): labels of every kind in a top row (a reference, a number, a boolean, a date, a formatted
 * number, an error, a formula's text, spaces, punctuation, Cyrillic, a name twice in two cases, too long), labels in the
 * left column and the corner, labels on each side alone and on all four, a range of labels alone, an empty corner.
 */
import type { CellInput } from '../../../src/contract.js'
import type { SelectBook, SelectCase } from './cases.js'

const LABELS: CellInput[] = [
  { value: 'Q1' }, { value: '2025 total' }, { value: 'Цена, руб.' }, { value: 'a-b' }, { value: 'TRUE' }, { value: true }, { value: 'R2C3' }, { value: 'x y z' },
  { value: 2024 }, { value: 1.5 }, { formula: '"Sum"&1' }, { value: 45000, numberFormat: 14 }, { value: 2, numberFormat: '0.00' }, { formula: 'NA()' },
  { value: ' lead' }, { value: 'A1B' }, { value: 'x'.repeat(300) }, { value: '\\path' }, { value: 'a?b' }, { value: '.dot' }, { value: 'Total' }, { value: 'total' }, { value: 'R' },
]

const COLS = 'BCDEFGHIJKLMNOPQRSTUVWXYZ'

/** Labels in row 1 from B, in column A from row 2 and in A1; numbers in B2:X4. */
function labelled(corner?: CellInput): SelectBook {
  const cells: Record<string, CellInput> = { A2: { value: 'north' }, A3: { value: 'south side' }, A4: { value: '3rd' } }
  if (corner) cells.A1 = corner
  LABELS.forEach((x, i) => {
    cells[`${COLS[i]}1`] = x
    for (let r = 2; r <= 4; r++) cells[`${COLS[i]}${r}`] = { value: r * 10 + i }
  })
  return { cells }
}

/** Numbers in B2:C3 with labels on every side: B1:C1 above, B4:C4 below, A2:A3 to the left, D2:D3 to the right, and the four corners. */
const edges: SelectBook = {
  cells: {
    A1: { value: 'nw' }, B1: { value: 'top1' }, C1: { value: 'top2' }, D1: { value: 'ne' },
    A2: { value: 'left1' }, B2: { value: 1 }, C2: { value: 2 }, D2: { value: 'right1' },
    A3: { value: 'left2' }, B3: { value: 3 }, C3: { value: 4 }, D3: { value: 'right2' },
    A4: { value: 'sw' }, B4: { value: 'bottom1' }, C4: { value: 'bottom2' }, D4: { value: 'se' },
  },
}

export const BOOKS: Record<string, SelectBook> = { labels: labelled({ value: 'Corner' }), cornerless: labelled(), edges }

export function buildSelectCases(): SelectCase[] {
  return [
    { id: 'top', book: 'labels', range: 'B1:X4', labels: { top: true } },
    { id: 'left', book: 'labels', range: 'A2:D4', labels: { left: true } },
    { id: 'top-left', book: 'labels', range: 'A1:D4', labels: { top: true, left: true } },
    { id: 'top-left-empty-corner', book: 'cornerless', range: 'A1:D4', labels: { top: true, left: true } },
    { id: 'labels-alone', book: 'labels', range: 'B1:D1', labels: { top: true } },
    { id: 'bottom', book: 'edges', range: 'B2:C4', labels: { bottom: true } },
    { id: 'right', book: 'edges', range: 'B2:D3', labels: { right: true } },
    { id: 'bottom-right', book: 'edges', range: 'B2:D4', labels: { bottom: true, right: true } },
    { id: 'all', book: 'edges', range: 'A1:D4', labels: { top: true, left: true, bottom: true, right: true } },
  ]
}
