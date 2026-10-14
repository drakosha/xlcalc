/**
 * Held out after `merge-held`, predicted by the revised rule before Excel ran it (`merge-third-held.predicted.json`):
 * the cell moved into an empty top-left cell read by a name, by absolute and mixed references, by a range it leaves and
 * a range it stays in, by a formula it holds itself; a top-left formula reading a cell the merge clears; a cell of a
 * spill as the first value; whole columns where the first value row by row and column by column part; across and
 * centred rows each moving their own; cells merged then split again.
 */
import type { StructureOp } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'

const f = (formula: string): StoredCell => ({ formula })

/**
 * S: an empty A1 before B1 (filled, read by a name, absolute and mixed references, ranges) and C1; a top-left A3
 * reading B3; B5 reading C5 after an empty A5; a spill from D12 beside an empty C13; values in rows 16:17 with empty
 * A16 and A17; an empty column J with K5 and J7 under it; T: formulas over S!B1.
 */
function thirdBook(): StoredBook {
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          B1: { value: 2, style: 0 },
          C1: { value: 3 },
          B2: { value: 20 },
          H1: f('$B$1*10'),
          H2: f('B$1+$B1'),
          H3: f('SUM(B1:B2)'),
          H4: f('SUM(A1:C1)'),
          H5: f('nb'),
          H6: f('C1'),
          A3: f('B3+1'),
          B3: { value: 30 },
          H7: f('B3'),
          B5: f('C5*2'),
          C5: { value: 50 },
          H8: f('B5'),
          D12: f('SEQUENCE(3)'),
          H12: f('D13'),
          B16: { value: 'u' },
          C16: { value: 'v' },
          C17: { value: 'w' },
          K5: { value: 'k' },
          J7: { value: 'j' },
          H9: f('K5'),
          H10: f('J7'),
        },
      },
      { name: 'T', cells: { A1: f('S!B1'), A2: f('SUM(S!B1:B2)'), A3: f('S!$B$1') } },
    ],
    names: [{ name: 'nb', ref: "'S'!$B$1" }],
    styles: [{ fill: { color: '#FFC000' } }],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const add = (id: string, op: StructureOp, more: Partial<StructCase> = {}) => cases.push({ id, book: 'third', op, ...more })
  const merge = (range: string, flags: { across?: true; center?: true } = {}): StructureOp => ({ op: 'mergeCells', sheet: 'S', range, ...flags })
  const unmerge = (range: string): StructureOp => ({ op: 'unmergeCells', sheet: 'S', range })
  add('merge-A1-C1-moves-B1', merge('A1:C1'))
  add('merge-A1-B2-moves-B1', merge('A1:B2'))
  add('merge-A1-C1-center', merge('A1:C1', { center: true }))
  add('unmerge-A1-after-move', unmerge('A1:C1'), { before: merge('A1:C1') })
  add('merge-A3-B3-top-formula', merge('A3:B3'))
  add('merge-A5-C5-moves-formula', merge('A5:C5'))
  add('merge-C13-D13-spill-cell', merge('C13:D13'))
  add('merge-columns-J-K', merge('J:K'), { extra: [{ sheet: 'S', cell: 'K1048576' }] })
  add('merge-A16-C17-across', merge('A16:C17', { across: true }))
  add('merge-A16-C17-across-center', merge('A16:C17', { across: true, center: true }))
  add('merge-A16-C17', merge('A16:C17'))
  add('merge-B1-C2-top-holds', merge('B1:C2'))
  return { books: { third: thirdBook() }, cases, dynamic: true }
}
