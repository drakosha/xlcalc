/**
 * How cut and paste adjusts range references, systematically: T holds every reference `S!Ai:Aj` with
 * 1 ≤ i ≤ j ≤ 12 (cell (i, j) of T, single cells on the diagonal) and two-column ones `S!Ai:Bj` below; each
 * case moves a block of S along the column, across it or to another sheet.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'

const col = (i: number) => String.fromCharCode(64 + i)

function windowBook(): BookInput {
  const s: Record<string, CellInput> = {}
  for (let r = 1; r <= 14; r++) {
    s[`A${r}`] = { value: r }
    s[`B${r}`] = { value: r * 10 }
  }
  const t: Record<string, CellInput> = {}
  for (let i = 1; i <= 12; i++)
    for (let j = i; j <= 12; j++) {
      t[`${col(j)}${i}`] = { formula: i === j ? `S!A${i}` : `SUM(S!A${i}:A${j})` }
      t[`${col(j)}${i + 20}`] = { formula: `SUM(S!A${i}:B${j})` }
    }
  return { sheets: [{ name: 'S', cells: s }, { name: 'T', cells: t }, { name: 'U', cells: {} }] }
}

export function buildStructCases(): StructCases {
  const moves: [string, string, string][] = [
    ['A4', 'S', 'A8'],
    ['A8', 'S', 'A4'],
    ['A4', 'S', 'A13'],
    ['A1', 'S', 'A5'],
    ['A12', 'S', 'A6'],
    ['A4:A5', 'S', 'A7'],
    ['A4:A6', 'S', 'A5'],
    ['A6:A8', 'S', 'A5'],
    ['A3:A10', 'S', 'A5'],
    ['A10:A12', 'S', 'A2'],
    ['A4:A6', 'S', 'A20'],
    ['A4:A6', 'S', 'C4'],
    ['A4:A6', 'S', 'C10'],
    ['A4:A6', 'U', 'A4'],
    ['A4:B6', 'S', 'A8'],
    ['A4:B4', 'S', 'A9'],
    ['B4:B6', 'S', 'C4'],
    ['B4:B6', 'S', 'A4'],
    ['A1:A3', 'S', 'A12'],
    ['A12', 'S', 'A1'],
  ]
  const cases: StructCase[] = moves.map(([from, sheet, cell]) => ({ id: `move-${from}-${sheet}!${cell}`, book: 'window', op: { op: 'moveRange', from: { sheet: 'S', range: from }, to: { sheet, cell } } }))
  return { books: { window: windowBook() }, cases }
}
