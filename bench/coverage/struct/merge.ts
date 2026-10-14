/**
 * Merge Cells and Unmerge Cells (Excel for Windows, `Range.Merge`, `Range.UnMerge` with alerts off): values in several
 * cells of the range, an empty top-left cell, formulas in the range and reading it, cells of their own look, merged
 * areas met, contained and split, whole rows and columns, a table and an array formula of the old kind; the look of
 * merged cells split again (`before`). Predicted by the engine before Excel ran it (`merge.predicted.json`).
 */
import type { StructureOp } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'

const f = (formula: string): StoredCell => ({ formula })

/**
 * S: numbers and text in A1:C6, A3 and B3 filled, B4 `0.00`; formulas over them in H; formulas in B8:C8; an empty A10 and
 * A11 beside values; merged areas E2:F3 and E10:F10; a table's cells in B13:C16 (made by `before`); an array formula of
 * the old kind over E13:E15; values in rows 20:21 and column J. T: formulas over S.
 */
function mergeBook(): StoredBook {
  const cells: Record<string, StoredCell> = {}
  for (let r = 1; r <= 6; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { value: r * 10 }, [`C${r}`]: { value: `c${r}` } })
  Object.assign(cells, {
    B1: { value: 10, style: 0 },
    A3: { value: 3, style: 1 },
    B3: { value: 30, style: 2 },
    B4: { value: 40, format: '0.00' },
    H1: f('A1'),
    H2: f('B1'),
    H3: f('B2+C3'),
    H4: f('SUM(A2:C3)'),
    H5: f('COUNTA(A1:C6)'),
    H6: f('C5&"!"'),
    H7: f('B4'),
    H8: f('J5'),
    B8: f('A1*2'),
    C8: f('B8+1'),
    D8: { value: 5 },
    E2: { value: 'm' },
    E10: { value: 'n' },
    B10: { value: 'x' },
    B11: { value: 7 },
    C11: { value: 8 },
    B13: { value: 'h1' },
    C13: { value: 'h2' },
    B14: { value: 1 },
    B15: { value: 2 },
    B16: { value: 3 },
    C14: { value: 4 },
    C15: { value: 5 },
    C16: { value: 6 },
    E13: { formula: 'B14:B16*2', array: 'E13:E15' },
    G13: f('SUM(E13:E15)'),
    A20: { value: 1 },
    C20: { value: 3 },
    B21: { value: 2 },
    J1: { value: 'j' },
    J5: { value: 5 },
  })
  return {
    sheets: [
      { name: 'S', cells, merges: ['E2:F3', 'E10:F10'] },
      { name: 'T', cells: { A1: f('S!B1'), A2: f('SUM(S!A1:C2)'), A3: f('S!C11') } },
    ],
    styles: [{ font: { bold: true } }, { fill: { color: '#FFC000' } }, { font: { bold: true }, fill: { color: '#92D050' } }],
  }
}

const TABLE: StructureOp = { op: 'addTable', sheet: 'S', table: { name: 'Sales', range: 'B13:C16', columns: [{ name: 'h1' }, { name: 'h2' }] } }

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const add = (id: string, op: StructureOp, more: Partial<StructCase> = {}) => cases.push({ id, book: 'merge', op, ...more })
  const merge = (range: string, flags: { across?: true; center?: true } = {}): StructureOp => ({ op: 'mergeCells', sheet: 'S', range, ...flags })
  const unmerge = (range: string): StructureOp => ({ op: 'unmergeCells', sheet: 'S', range })
  add('merge-A1-B2', merge('A1:B2'))
  add('merge-A1-C1-center', merge('A1:C1', { center: true }))
  add('merge-A1-C3-across', merge('A1:C3', { across: true }))
  add('merge-A5-C6-across-center', merge('A5:C6', { across: true, center: true }))
  add('merge-A3-B4-styled', merge('A3:B4'))
  add('unmerge-A3-after-merge', unmerge('A3'), { before: merge('A3:B4') })
  add('unmerge-B1-after-center', unmerge('B1'), { before: merge('A1:C1', { center: true }) })
  add('unmerge-B5-B6-after-across-center', unmerge('B5:B6'), { before: merge('A5:C6', { across: true, center: true }) })
  add('merge-A10-B10-empty-top', merge('A10:B10'))
  add('merge-A10-C11-empty-top', merge('A10:C11'))
  add('merge-B8-C8-formulas', merge('B8:C8'))
  add('merge-A8-D8-empty-top', merge('A8:D8'))
  add('merge-D2-E2-meets', merge('D2:E2'))
  add('merge-D1-G4-covers', merge('D1:G4'))
  add('merge-D2-F3-across-over', merge('D2:F3', { across: true }))
  add('merge-E3-F3-inside', merge('E3:F3'))
  add('unmerge-E2', unmerge('E2'))
  add('unmerge-F3', unmerge('F3'))
  add('unmerge-D1-G12-two', unmerge('D1:G12'))
  add('unmerge-A1-plain', unmerge('A1'))
  add('merge-A1-single', merge('A1'))
  add('merge-A1-single-center', merge('A1', { center: true }))
  add('merge-rows-20-21', merge('20:21'), { extra: [{ sheet: 'S', cell: 'XFD21' }] })
  add('merge-columns-J-K', merge('J:K'), { extra: [{ sheet: 'S', cell: 'K1048576' }] })
  add('merge-in-table', merge('B14:C14'), { before: TABLE })
  add('merge-meets-table', merge('A14:B14'), { before: TABLE })
  add('merge-array-part', merge('E13:F13'))
  add('merge-array-whole', merge('E13:F15'))
  return { books: { merge: mergeBook() }, cases }
}
