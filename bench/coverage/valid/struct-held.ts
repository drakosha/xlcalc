/**
 * Held out: data validation and outlines under structure operations, predicted by the engine before Excel ran them
 * (`struct-held.predicted.json`, from `predict.ts`). The engine moves a rule's ranges as references move (growing with
 * lines inserted after their first line), rewrites the references of its formulas as a cell's, writes them again for the
 * new first cell, takes rules along with a move, pastes them with a copy (a pasted cell without a rule clears the
 * target's), and gives inserted lines the outline level of the line before them but not its collapsed mark.
 */
import type { BookInput, CellInput, StructureOp } from '../../../src/contract.js'
import type { ValidCase, ValidCases } from './cases.js'

/**
 * S: values 1 to 12 in A; a literal list over B2:B6, whole numbers between $H$1 and $H$2 over C2:C6, a custom rule
 * relative to A over D2:D6, a list of the sheet L over E2:E3 and E5, a text length over F3, a whole column G; rows 8:10
 * grouped (9 twice) with row 11 their summary, columns J:K grouped with L their summary. L: the list's items. T: a list of
 * S's cells.
 */
function book(): BookInput {
  const cells: Record<string, CellInput> = { H1: { value: 0 }, H2: { value: 10 }, B2: { value: 'Yes' }, B4: { value: 'Nope' }, C3: { value: 4 }, C4: { value: 12 }, D3: { value: 1 }, D4: { value: 9 }, E2: { value: 'b' }, F3: { value: 'abc' }, G1: { value: 'g' }, J1: { value: 'j' }, K1: { value: 'k' }, L1: { value: 'sum' } }
  for (let r = 1; r <= 12; r++) cells[`A${r}`] = { value: r }
  return {
    sheets: [
      {
        name: 'S',
        cells,
        validations: [
          { range: 'B2:B6', type: 'list', formula1: '"Yes,No"' },
          { range: 'C2:C6', type: 'whole', formula1: '$H$1', formula2: '$H$2' },
          { range: 'D2:D6', type: 'custom', formula1: 'D2<A2' },
          { range: 'E2:E3 E5', type: 'list', formula1: "'L'!$A$1:$A$4" },
          { range: 'F3', type: 'textLength', operator: 'lessThanOrEqual', formula1: '3' },
          { range: 'G:G', type: 'any', prompt: 'column G' },
        ],
        rows: { 8: { outlineLevel: 1 }, 9: { outlineLevel: 2 }, 10: { outlineLevel: 1 } },
        columns: { 'J:K': { outlineLevel: 1 } },
      },
      { name: 'L', cells: { A1: { value: 'a' }, A2: { value: 'b' }, A3: { value: 'c' }, A4: { value: 'd' } } },
      { name: 'T', cells: { A1: { value: 2 } }, validations: [{ range: 'A1:A3', type: 'list', formula1: "'S'!$A$2:$A$4" }] },
    ],
  }
}

export function buildValidCases(): ValidCases {
  const cases: ValidCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'struct', op })
  const S = 'S'
  add('insert-row-1', { op: 'insertRows', sheet: S, row: 1 })
  add('insert-row-2', { op: 'insertRows', sheet: S, row: 2 })
  add('insert-row-4', { op: 'insertRows', sheet: S, row: 4 })
  add('insert-row-7', { op: 'insertRows', sheet: S, row: 7 })
  add('insert-row-8', { op: 'insertRows', sheet: S, row: 8 })
  add('insert-row-9', { op: 'insertRows', sheet: S, row: 9 })
  add('insert-row-11', { op: 'insertRows', sheet: S, row: 11 })
  add('insert-rows-3x2', { op: 'insertRows', sheet: S, row: 3, count: 2 })
  add('delete-row-2', { op: 'deleteRows', sheet: S, row: 2 })
  add('delete-row-4', { op: 'deleteRows', sheet: S, row: 4 })
  add('delete-rows-2-6', { op: 'deleteRows', sheet: S, row: 2, count: 5 })
  add('delete-rows-8-10', { op: 'deleteRows', sheet: S, row: 8, count: 3 })
  add('delete-row-11', { op: 'deleteRows', sheet: S, row: 11 })
  add('insert-col-a', { op: 'insertColumns', sheet: S, column: 'A' })
  add('insert-col-k', { op: 'insertColumns', sheet: S, column: 'K' })
  add('insert-col-l', { op: 'insertColumns', sheet: S, column: 'L' })
  add('delete-col-d', { op: 'deleteColumns', sheet: S, column: 'D' })
  add('delete-col-a', { op: 'deleteColumns', sheet: S, column: 'A' })
  add('move-b2-c3', { op: 'moveRange', from: { sheet: S, range: 'B2:C3' }, to: { sheet: S, cell: 'I20' } })
  add('move-plain-onto-rule', { op: 'moveRange', from: { sheet: S, range: 'A1' }, to: { sheet: S, cell: 'C4' } })
  add('move-d2-d3', { op: 'moveRange', from: { sheet: S, range: 'D2:D3' }, to: { sheet: S, cell: 'D20' } })
  add('move-list-source', { op: 'moveRange', from: { sheet: 'L', range: 'A1:A4' }, to: { sheet: 'L', cell: 'C5' } })
  add('copy-b2', { op: 'copyRange', from: { sheet: S, range: 'B2' }, to: { sheet: S, range: 'I2:I4' } })
  add('copy-plain-onto-rule', { op: 'copyRange', from: { sheet: S, range: 'A1' }, to: { sheet: S, range: 'B3' } })
  add('copy-d2-d3', { op: 'copyRange', from: { sheet: S, range: 'D2:D3' }, to: { sheet: S, range: 'M20' } })
  add('copy-to-t', { op: 'copyRange', from: { sheet: S, range: 'C2:D3' }, to: { sheet: 'T', range: 'C5' } })
  add('rename-l', { op: 'renameSheet', sheet: 'L', name: 'Lists' })
  add('rename-s', { op: 'renameSheet', sheet: S, name: 'Data' })
  add('delete-l', { op: 'deleteSheet', sheet: 'L' })
  add('move-rows-8-11', { op: 'moveRange', from: { sheet: S, range: '8:11' }, to: { sheet: S, cell: 'A20' } })
  add('copy-rows-8-9', { op: 'copyRange', from: { sheet: S, range: '8:9' }, to: { sheet: S, range: '25:26' } })
  return { books: { struct: book() }, cases, window: 'A1:N30' }
}
