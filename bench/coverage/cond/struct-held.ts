/**
 * Held out: conditional formats under structure operations, predicted by the engine before Excel ran them
 * (`struct-held.predicted.json`, from `predict.ts`). The engine changes a rule's ranges as ranges change (inserted lines
 * inside grow them, deleted ones shrink them, a rule with nothing left goes), moves the part of a rule over moved cells
 * with them (to another sheet as a rule there), lets the cells a move or a copy lands on lose their rules, grows a rule
 * over the copies of its cells on its sheet and makes a rule of the target sheet's, first, from another sheet's; it
 * rewrites formulas as those of the rule's first cell, made for a new first cell as a copy makes them.
 */
import type { BookInput, CellInput, ConditionalFormat, StructureOp } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'
import { valueCell as v } from '../stored.js'

/**
 * S: numbers in A1:A10 (0 to 9) and B1:G10; rules: B2:B10 above 5, D2:D10 and F2:F4 where $A is above 3, a colour scale
 * over C1:C10, a data bar over E2:E10, G1:G3 between $Z$1 and $Z$2, H2:H5 by T's A1. T: A1:A3 by S's A1.
 */
function structBook(): BookInput {
  const cells: Record<string, CellInput> = { Z1: v(2), Z2: v(6) }
  for (let r = 1; r <= 10; r++) {
    cells[`A${r}`] = v(r - 1)
    for (const c of ['B', 'C', 'D', 'E', 'F', 'G', 'H']) cells[`${c}${r}`] = v(r)
  }
  const red = { fill: { color: '#FFC7CE' } }
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['B2:B10'], operator: 'greaterThan', formulas: ['5'], style: red },
    { type: 'expression', ranges: ['D2:D10', 'F2:F4'], formula: '$A2>3', style: { font: { bold: true } } },
    { type: 'colorScale', ranges: ['C1:C10'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'dataBar', ranges: ['E2:E10'], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6' },
    { type: 'cellIs', ranges: ['G1:G3'], operator: 'between', formulas: ['$Z$1', '$Z$2'], style: { fill: { color: '#C6EFCE' } } },
    { type: 'expression', ranges: ['H2:H5'], formula: "H2>'T'!$A$1", style: { font: { italic: true } } },
  ]
  return {
    sheets: [
      { name: 'S', cells, conditionalFormats: rules },
      { name: 'T', cells: { A1: v(3), A2: v(1), A3: v(5) }, conditionalFormats: [{ type: 'expression', ranges: ['A1:A3'], formula: '$A1>S!$A$3', style: red }] },
    ],
  }
}

export function buildCondCases(): CondCases {
  const cases: CondCase[] = []
  const S = 'S'
  const read: CondCase['read'] = [['S', 'B1:H12']]
  const add = (id: string, op: StructureOp, more: CondCase['read'] = []) => cases.push({ id, book: 'struct', op, read: [...read, ...more] })
  add('insert-row-1', { op: 'insertRows', sheet: S, row: 1 })
  add('insert-row-2', { op: 'insertRows', sheet: S, row: 2 })
  add('insert-rows-5x2', { op: 'insertRows', sheet: S, row: 5, count: 2 })
  add('insert-row-11', { op: 'insertRows', sheet: S, row: 11 })
  add('insert-row-4', { op: 'insertRows', sheet: S, row: 4 })
  add('delete-row-2', { op: 'deleteRows', sheet: S, row: 2 })
  add('delete-rows-2-4', { op: 'deleteRows', sheet: S, row: 2, count: 3 })
  add('delete-row-5', { op: 'deleteRows', sheet: S, row: 5 })
  add('delete-rows-1-10', { op: 'deleteRows', sheet: S, row: 1, count: 10 })
  add('insert-col-a', { op: 'insertColumns', sheet: S, column: 'A' })
  add('insert-col-d', { op: 'insertColumns', sheet: S, column: 'D' })
  add('delete-col-a', { op: 'deleteColumns', sheet: S, column: 'A' })
  add('delete-col-d', { op: 'deleteColumns', sheet: S, column: 'D' })
  add('move-part', { op: 'moveRange', from: { sheet: S, range: 'B2:B5' }, to: { sheet: S, cell: 'J2' } }, [['S', 'J1:J6']])
  add('move-referenced', { op: 'moveRange', from: { sheet: S, range: 'A1:A10' }, to: { sheet: S, cell: 'K1' } })
  add('move-onto-rule', { op: 'moveRange', from: { sheet: S, range: 'B8:B10' }, to: { sheet: S, cell: 'D8' } })
  add('move-to-t', { op: 'moveRange', from: { sheet: S, range: 'B2:B10' }, to: { sheet: 'T', cell: 'C1' } }, [['T', 'A1:C9']])
  add('copy-within', { op: 'copyRange', from: { sheet: S, range: 'B2:B3' }, to: { sheet: S, range: 'J5' } }, [['S', 'J5:J6']])
  add('copy-to-t', { op: 'copyRange', from: { sheet: S, range: 'B2:B3' }, to: { sheet: 'T', range: 'D1' } }, [['T', 'D1:D2']])
  add('copy-plain-onto-rule', { op: 'copyRange', from: { sheet: S, range: 'A1' }, to: { sheet: S, range: 'B6' } })
  add('fill-down', { op: 'copyRange', from: { sheet: S, range: 'B10' }, to: { sheet: S, range: 'B11:B12' } })
  add('rename-t', { op: 'renameSheet', sheet: 'T', name: 'U' })
  add('delete-t', { op: 'deleteSheet', sheet: 'T' })
  add('insert-col-a-on-t', { op: 'insertColumns', sheet: 'T', column: 'A' })
  return { books: { struct: structBook() }, cases }
}
