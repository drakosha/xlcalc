/**
 * Held out after `struct-held`: the rules fitted there, on other books, predicted before Excel ran them
 * (`struct-held-2.predicted.json`): lines inserted right after rules (several, and columns), a move of cells with rules
 * onto other rules and of plain cells onto part of a rule, a move to another sheet, lines inserted into collapsed groups
 * with summaries below and above, rows deleted from a collapsed group, a collapsed group copied whole.
 */
import type { BookInput, CellInput, StructureOp } from '../../../src/contract.js'
import type { ValidCase, ValidCases } from './cases.js'

/**
 * S: values in A1:A12; a custom rule relative to A over B2:B4, a list over C3, whole numbers over D2:E3, a list over F2:F6;
 * rows 8:10 grouped and collapsed under row 11; columns H:I grouped and collapsed beside J. U: the same group of rows with
 * its summary above.
 */
function book(): BookInput {
  const cells: Record<string, CellInput> = { B2: { value: 5 }, C3: { value: 'a' }, D2: { value: 3 }, F3: { value: 'q' }, H1: { value: 'h' }, J1: { value: 'sum' }, A11: { value: 'sum' } }
  for (let r = 1; r <= 12; r++) cells[`A${r}`] = { value: r }
  return {
    sheets: [
      {
        name: 'S',
        cells,
        validations: [
          { range: 'B2:B4', type: 'custom', formula1: 'B2>A2' },
          { range: 'C3', type: 'list', formula1: '"a,b"' },
          { range: 'D2:E3', type: 'whole', operator: 'greaterThan', formula1: '0' },
          { range: 'F2:F6', type: 'list', formula1: '"p,q"', errorStyle: 'warning' },
        ],
        rows: { '8:10': { outlineLevel: 1, hidden: true }, 11: { collapsed: true } },
        columns: { 'H:I': { outlineLevel: 1, hidden: true }, J: { collapsed: true } },
      },
      { name: 'T', cells: { A1: { value: 1 } }, validations: [{ range: 'B1:B3', type: 'decimal', operator: 'lessThan', formula1: '$A$1' }] },
      { name: 'U', cells: { A1: { value: 'u' } }, rows: { 2: { collapsed: true }, '3:5': { outlineLevel: 1, hidden: true } }, outline: { summaryBelow: false } },
    ],
  }
}

export function buildValidCases(): ValidCases {
  const cases: ValidCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'struct', op })
  const S = 'S'
  add('insert-rows-5x2', { op: 'insertRows', sheet: S, row: 5, count: 2 })
  add('insert-row-4', { op: 'insertRows', sheet: S, row: 4 })
  add('insert-col-f', { op: 'insertColumns', sheet: S, column: 'F' })
  add('insert-cols-d', { op: 'insertColumns', sheet: S, column: 'D', count: 2 })
  add('move-rule-onto-rule', { op: 'moveRange', from: { sheet: S, range: 'C3' }, to: { sheet: S, cell: 'F4' } })
  add('move-plain-onto-part', { op: 'moveRange', from: { sheet: S, range: 'A1:A2' }, to: { sheet: S, cell: 'F3' } })
  add('move-rules-onto-rules', { op: 'moveRange', from: { sheet: S, range: 'B2:C4' }, to: { sheet: S, cell: 'E2' } })
  add('move-to-t', { op: 'moveRange', from: { sheet: S, range: 'B2:B3' }, to: { sheet: 'T', cell: 'D2' } })
  add('insert-row-9-collapsed', { op: 'insertRows', sheet: S, row: 9 })
  add('insert-row-11-collapsed', { op: 'insertRows', sheet: S, row: 11 })
  add('insert-row-8-collapsed', { op: 'insertRows', sheet: S, row: 8 })
  add('insert-col-i-collapsed', { op: 'insertColumns', sheet: S, column: 'I' })
  add('insert-row-4-above', { op: 'insertRows', sheet: 'U', row: 4 })
  add('delete-row-9-collapsed', { op: 'deleteRows', sheet: S, row: 9 })
  add('copy-rows-8-11', { op: 'copyRange', from: { sheet: S, range: '8:11' }, to: { sheet: S, range: '20:23' } })
  return { books: { struct: book() }, cases, window: 'A1:L24' }
}
