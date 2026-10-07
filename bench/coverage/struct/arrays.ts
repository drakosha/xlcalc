/**
 * Array formulas of the old kind under structure operations: rows and columns inserted or deleted before, inside and
 * over the area, the area moved and copied whole or in part, a paste over part of it, and a formula reading it.
 */
import type { BookInput, StructureOp } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'

function book(): BookInput {
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, A4: { value: 4 }, A5: { value: 5 },
          C2: { formula: 'A1:A3*2', array: 'C2:C4' },
          D2: { formula: 'TRANSPOSE(A1:A2)', array: 'D2:E2' },
          G2: { formula: 'SUM(C2:C4)' },
        },
      },
      { name: 'T', cells: { A1: { formula: 'SUM(S!C2:C4)' } } },
    ],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'arrays', op })
  add('insert-row-above', { op: 'insertRows', sheet: 'S', row: 1 })
  add('insert-row-inside', { op: 'insertRows', sheet: 'S', row: 3 })
  add('insert-row-below', { op: 'insertRows', sheet: 'S', row: 5 })
  add('delete-row-inside', { op: 'deleteRows', sheet: 'S', row: 3 })
  add('delete-rows-over', { op: 'deleteRows', sheet: 'S', row: 2, count: 3 })
  add('insert-column-inside', { op: 'insertColumns', sheet: 'S', column: 'E' })
  add('insert-column-before', { op: 'insertColumns', sheet: 'S', column: 'C' })
  add('delete-column-part', { op: 'deleteColumns', sheet: 'S', column: 'E' })
  add('delete-column-whole', { op: 'deleteColumns', sheet: 'S', column: 'C' })
  add('move-whole', { op: 'moveRange', from: { sheet: 'S', range: 'C2:C4' }, to: { sheet: 'S', cell: 'J2' } })
  add('move-part', { op: 'moveRange', from: { sheet: 'S', range: 'C2:C3' }, to: { sheet: 'S', cell: 'J2' } })
  add('move-other-sheet', { op: 'moveRange', from: { sheet: 'S', range: 'D2:E2' }, to: { sheet: 'T', cell: 'C1' } })
  add('copy-whole', { op: 'copyRange', from: { sheet: 'S', range: 'C2:C4' }, to: { sheet: 'S', range: 'J2' } })
  add('copy-part', { op: 'copyRange', from: { sheet: 'S', range: 'C3' }, to: { sheet: 'S', range: 'J3' } })
  add('paste-over-part', { op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'C3' } })
  add('paste-over-whole', { op: 'copyRange', from: { sheet: 'S', range: 'A1:A3' }, to: { sheet: 'S', range: 'C2' } })
  add('move-onto-part', { op: 'moveRange', from: { sheet: 'S', range: 'A5' }, to: { sheet: 'S', cell: 'C4' } })
  return { books: { arrays: book() }, cases }
}
