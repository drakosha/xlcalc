/**
 * Held out after `merge`, predicted by the engine before Excel ran it (`merge-held.predicted.json`): an empty top-left
 * cell taking the first value or formula of the area, which row by row and which column by column would part; formulas
 * reading the cell it comes from; the look and the number format it takes; across, centred, over a merged area met and
 * the same area again; a cell of a look alone; a formula giving empty text; a spill's anchor and a cell of its spill.
 * Formulas are dynamic array formulas, as the spill needs.
 */
import type { StructureOp } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'

const f = (formula: string): StoredCell => ({ formula })

/**
 * S: an empty filled A1 under B1 (bold) and A2; a formula in B4 beside a percentage; empty A6 and B7 among values; a
 * percentage beside an empty A9; a spill from D12 and its sum; a cell of a look alone in A16, beside a number; a formula
 * giving empty text in A18; a merged area E20:F21; a merged area B30:B31 beside values; a filled B40 alone. T: a formula
 * over S!B1.
 */
function heldBook(): StoredBook {
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { style: 1 },
          B1: { value: 'r', style: 0 },
          A2: { value: 'c' },
          B2: { value: 'd' },
          H1: f('B1'),
          H2: f('A2'),
          B4: f('A10+1'),
          C4: { value: 0.5, format: '0%' },
          H4: f('B4*2'),
          A10: { value: 7 },
          B6: { value: 'p' },
          C6: { value: 'q' },
          A7: { value: 'z' },
          C7: { value: 'w' },
          B9: { value: 0.25, format: '0%' },
          D12: f('SEQUENCE(3)'),
          G12: f('SUM(D12#)'),
          A16: { style: 1 },
          B16: { value: 3, format: '0.00' },
          A18: f('""'),
          B18: { value: 'v' },
          E20: { value: 'e' },
          A30: { value: 'a' },
          C31: { value: 'k' },
          B40: { style: 1 },
        },
        merges: ['E20:F21', 'B30:B31'],
      },
      { name: 'T', cells: { A1: f('S!B1') } },
    ],
    styles: [{ font: { bold: true } }, { fill: { color: '#FFC000' } }],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const add = (id: string, op: StructureOp, more: Partial<StructCase> = {}) => cases.push({ id, book: 'held', op, ...more })
  const merge = (range: string, flags: { across?: true; center?: true } = {}): StructureOp => ({ op: 'mergeCells', sheet: 'S', range, ...flags })
  const unmerge = (range: string): StructureOp => ({ op: 'unmergeCells', sheet: 'S', range })
  add('merge-A1-B2-order', merge('A1:B2'))
  add('merge-A1-B2-order-center', merge('A1:B2', { center: true }))
  add('merge-A1-A2-down', merge('A1:A2'))
  add('merge-B1-B2-top-holds', merge('B1:B2'))
  add('unmerge-A1-after-empty-top', unmerge('A1'), { before: merge('A1:B2') })
  add('merge-A4-C4-formula', merge('A4:C4'))
  add('merge-A6-C7', merge('A6:C7'))
  add('merge-A6-C7-across', merge('A6:C7', { across: true }))
  add('merge-rows-6-7', merge('6:7'), { extra: [{ sheet: 'S', cell: 'XFD7' }] })
  add('merge-A9-B9-center', merge('A9:B9', { center: true }))
  add('merge-A16-B16-look-alone', merge('A16:B16'))
  add('merge-A18-B18-empty-text', merge('A18:B18'))
  add('merge-C12-D12-spill-anchor', merge('C12:D12'))
  add('merge-D13-E13-spill-cell', merge('D13:E13'))
  add('merge-A30-C30-across-grows', merge('A30:C30', { across: true }))
  add('merge-E20-F21-again-center', merge('E20:F21', { center: true }))
  add('unmerge-column-E', unmerge('E:E'))
  add('merge-A40-B40-empty', merge('A40:B40'))
  return { books: { held: heldBook() }, cases, dynamic: true, look: ['A25:J41'] }
}
