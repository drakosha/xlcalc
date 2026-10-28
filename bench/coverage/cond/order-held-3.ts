/**
 * Held out after the set order-held-2 showed that a rule that loses its last lines stays where its formulas changed and
 * goes where a reference was lost (`lateBy`): rules cut at their end by rows and columns, beyond it, in two ranges, at
 * their first column and in their middle, a reference into the rule's own last row, a data bar, and rows deleted
 * between a rule's two ranges. Predicted before Excel ran it (`order-held-3.predicted.json`).
 */
import type { BookInput, CellInput, ConditionalFormat, StructureOp } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'
import { valueCell as v } from '../stored.js'
import { RED } from './cases.js'

/** S: numbers 1 to 10 in rows 1 to 10 of B to H and N, thresholds in M15:M17. The last rule, over N, shows which rules went after it. */
function heldBook(): BookInput {
  const cells: Record<string, CellInput> = { M15: v(4), M16: v(7), M17: v(5) }
  for (const c of ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'N']) for (let r = 1; r <= 10; r++) cells[`${c}${r}`] = v(r)
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['B1:B10'], operator: 'greaterThan', formulas: ['$M$16'], style: RED },
    { type: 'cellIs', ranges: ['C1:C10'], operator: 'lessThan', formulas: ['$C$10'], style: { font: { bold: true } } },
    { type: 'expression', ranges: ['D1:E10'], formula: 'D1>$M$15', style: { font: { italic: true } } },
    { type: 'expression', ranges: ['F1:F4', 'F7:F10'], formula: 'F1>$M$17', style: { fill: { color: '#C6EFCE' } } },
    { type: 'dataBar', ranges: ['G1:G10'], min: { type: 'num', value: '$M$15' }, max: { type: 'max' }, color: '#638EC6' },
    { type: 'cellIs', ranges: ['H1:H10'], operator: 'greaterThan', formulas: ['$M$16'], style: { font: { color: '#C00000' } } },
    { type: 'cellIs', ranges: ['N1:N10'], operator: 'greaterThan', formulas: ['5'], style: { font: { color: '#0070C0' } } },
  ]
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }] }
}

export function buildCondCases(): CondCases {
  const cases: CondCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'held', op, read: [['S', 'B1:N10']] })
  const S = 'S'
  add('delete-rows-9', { op: 'deleteRows', sheet: S, row: 9, count: 4 })
  add('delete-row-10', { op: 'deleteRows', sheet: S, row: 10 })
  add('delete-col-e', { op: 'deleteColumns', sheet: S, column: 'E' })
  add('delete-col-d', { op: 'deleteColumns', sheet: S, column: 'D' })
  add('delete-row-3', { op: 'deleteRows', sheet: S, row: 3 })
  add('delete-rows-5', { op: 'deleteRows', sheet: S, row: 5, count: 2 })
  return { books: { held: heldBook() }, cases }
}
