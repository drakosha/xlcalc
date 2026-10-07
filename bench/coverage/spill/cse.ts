/**
 * Array formulas of the old kind (Ctrl+Shift+Enter; `run.ts spill cse`), written as a file stores them: the formula
 * on the top-left cell of its area, `t="array"` over the area, no dynamic array metadata. The array fills the area
 * and never spills: a smaller one repeated when it has one row or column, else #N/A past it, a larger one cut, one
 * value everywhere; no implicit intersection, a single-cell one too; readers of the area, a spill meeting it, and
 * edits of part of it, which Excel refuses.
 */
import type { CellInput } from '../../../src/contract.js'
import { one, type CellSpec, type SpillCase } from './cases.js'

/** Data on S: A1:A3 1…3, B1:B3 10…30, A5:C5 a row. */
const DATA: Record<string, CellSpec> = { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30, A5: 5, B5: 6, C5: 7 }

/** An array formula of the old kind over `area`. */
const arr = (formula: string, area: string): CellInput => ({ formula, array: area })

export function buildSpillCases(): SpillCase[] {
  return [
    one('fit', 'an array of the area size', { ...DATA, E1: arr('A1:A3*2', 'E1:E3') }, ['E1:E3']),
    one('larger', 'an area larger than the array: #N/A past it, a column repeated across', { ...DATA, E1: arr('A1:A3*2', 'E1:E5'), F1: arr('A1:A3', 'F1:G3') }, ['E1:G5']),
    one('smaller', 'an area smaller than the array: cut', { ...DATA, E1: arr('A1:A3*2', 'E1:E2'), F1: arr('A1:C5', 'F1:G2') }, ['E1:G3']),
    one('row', 'a row repeated down, #N/A past its columns', { ...DATA, E1: arr('A5:C5', 'E1:G3'), E5: arr('A5:C5', 'E5:H6') }, ['E1:H6']),
    one('scalar', 'one value fills the area', { ...DATA, E1: arr('5', 'E1:F2'), E5: arr('SUM(A1:A3)', 'E5:F6') }, ['E1:F6']),
    one('block', 'a 2×2 array over 3×3: #N/A in the third row and column', { ...DATA, E1: arr('SEQUENCE(2,2)', 'E1:G3') }, ['E1:G3']),
    one('position', 'ROW and COLUMN in an area: the top-left cell', { ...DATA, E1: arr('ROW()', 'E1:E3'), F1: arr('COLUMN()', 'F1:H1'), G3: arr('ROW(A1:A3)', 'G3:G5') }, ['E1:H5']),
    one('single', 'one cell: computed as an array, its top-left value, no intersection', { ...DATA, E1: arr('SUM(A1:A3*B1:B3)', 'E1'), E3: arr('A1:A3', 'E3'), E4: arr('INDEX(A1:B3,0,2)', 'E4'), E2: arr('SEQUENCE(3)', 'E2') }, ['E1:E5']),
    one('readers', 'formulas reading the area, a spill reference to it', { ...DATA, E1: arr('A1:A3*2', 'E1:E3'), G1: '=SUM(E1:E3)', G2: '=E2', G3: '=E1#', G4: '=ROWS(E1#)' }, ['E1:G4']),
    one('collision', 'a spill meeting the area', { ...DATA, E1: arr('A1:A3', 'E1:E3'), D2: '=SEQUENCE(1,3)' }, ['D1:G3']),
    one('transpose', 'TRANSPOSE of a column over a row', { ...DATA, E1: arr('TRANSPOSE(A1:A3)', 'E1:G1') }, ['E1:G1']),
    one('edits', 'part of the area refused (a value, clearing the top-left cell, a formula), a precedent, the whole area cleared', { ...DATA, E1: arr('A1:A3*2', 'E1:E3') }, ['E1:E3'], {
      edits: [
        { sheet: 'S', cell: 'E2', value: 5 },
        { sheet: 'S', cell: 'E1', value: null },
        { sheet: 'S', cell: 'E3', value: { formula: '1' } },
        { sheet: 'S', cell: 'A1', value: 10 },
        { sheet: 'S', cell: 'E1:E3', value: null },
        { sheet: 'S', cell: 'E2', value: 7 },
      ],
    }),
  ]
}
