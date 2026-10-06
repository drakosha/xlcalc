/**
 * The sheet's edges: references to the last row and column and to the whole sheet when rows or columns
 * are inserted or deleted; then the two operations Excel refuses because data would leave the sheet
 * (last in the set: if Excel showed a window there, the other answers are stored already).
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'

const f = (formula: string): CellInput => ({ formula })

/** E: one value; F: references to E's edges; a name on the last row. */
function edgeBook(): BookInput {
  return {
    sheets: [
      { name: 'E', cells: { A1: { value: 1 }, B2: { value: 2 } } },
      {
        name: 'F',
        cells: {
          A1: f('E!A1048576'),
          A2: f('SUM(E!A5:A1048576)'),
          A3: f('SUM(E!A1:A1048576)'),
          A4: f('SUM(E!1:1048576)'),
          A5: f('SUM(E!A:XFD)'),
          A6: f('E!XFD1'),
          A7: f('SUM(E!XFA1:XFD1)'),
          A8: f('SUM(E!B1:XFD1)'),
          A9: f('E!A1048575'),
          A10: f('SUM(E!A1048570:A1048576)'),
          A11: f('SUM(E!1048575:1048576)'),
          A12: f('SUM(E!XFC:XFD)'),
          A13: f('E!$A$1048576'),
          A14: f('SUM(E!B2:C3)'),
          A15: f('SUM(E!A1:XFD1048576)'),
        },
      },
    ],
    names: [{ name: 'last', ref: "'E'!$A$1048576" }],
  }
}

/** A value in the last row and in the last column: inserting rows or columns would push it off the sheet. */
function fullBook(): BookInput {
  return {
    sheets: [
      { name: 'E', cells: { A1: { value: 1 }, A1048576: { value: 2 }, XFD1: { value: 3 } } },
      { name: 'F', cells: { A1: f('E!A1048576+E!XFD1'), A2: f('E!A1') } },
    ],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = [
    { id: 'edge-insert-rows-1', book: 'edge', op: { op: 'insertRows', sheet: 'E', row: 1 } },
    { id: 'edge-insert-rows-1048570x2', book: 'edge', op: { op: 'insertRows', sheet: 'E', row: 1048570, count: 2 } },
    { id: 'edge-insert-cols-A', book: 'edge', op: { op: 'insertColumns', sheet: 'E', column: 'A' } },
    { id: 'edge-insert-cols-XFAx2', book: 'edge', op: { op: 'insertColumns', sheet: 'E', column: 'XFA', count: 2 } },
    { id: 'edge-delete-rows-1048576', book: 'edge', op: { op: 'deleteRows', sheet: 'E', row: 1048576 } },
    { id: 'edge-delete-rows-1', book: 'edge', op: { op: 'deleteRows', sheet: 'E', row: 1 } },
    { id: 'edge-delete-cols-XFD', book: 'edge', op: { op: 'deleteColumns', sheet: 'E', column: 'XFD' } },
    { id: 'edge-delete-cols-A', book: 'edge', op: { op: 'deleteColumns', sheet: 'E', column: 'A' } },
    { id: 'full-insert-rows-5', book: 'full', op: { op: 'insertRows', sheet: 'E', row: 5 }, extra: [{ sheet: 'E', cell: 'A1048576' }, { sheet: 'E', cell: 'XFD1' }] },
    { id: 'full-insert-cols-C', book: 'full', op: { op: 'insertColumns', sheet: 'E', column: 'C' }, extra: [{ sheet: 'E', cell: 'A1048576' }, { sheet: 'E', cell: 'XFD1' }] },
  ]
  return { books: { edge: edgeBook(), full: fullBook() }, cases }
}
