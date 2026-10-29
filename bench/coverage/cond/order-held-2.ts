/**
 * Held out after the set order-held fitted two more parts of the rule of which rules deleted lines put after the others
 * (`lateBy`): a rule that loses columns stays where its formulas changed, and of a rule on another sheet only a lost
 * reference counts. Its cases tell the first from the other reading of the sets (a rule cut at its last line stays,
 * rows or columns): rules cut at their last row, in their middle column and at their first; and a colour scale over
 * another sheet whose threshold went. Predicted before Excel ran it (`order-held-2.predicted.json`).
 */
import type { BookInput, CellInput, ConditionalFormat, StructureOp } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'
import { valueCell as v } from '../stored.js'
import { RED } from './cases.js'

/** S: numbers 1 to 10 in rows 1 to 10 of B to J and N, thresholds in M15:M17; T: thresholds in A5 and A8. The last rule, over N, shows which rules went after it. */
function heldBook(): BookInput {
  const cells: Record<string, CellInput> = { M15: v(4), M16: v(7), M17: v(5) }
  for (const c of ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'N']) for (let r = 1; r <= 10; r++) cells[`${c}${r}`] = v(r)
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['B1:B10'], operator: 'greaterThan', formulas: ['$M$16'], style: RED },
    { type: 'cellIs', ranges: ['C1:E10'], operator: 'greaterThan', formulas: ['$M$17'], style: { font: { bold: true } } },
    { type: 'expression', ranges: ['F1:G10'], formula: 'F1>$M$16', style: { font: { italic: true } } },
    { type: 'colorScale', ranges: ['H1:H10'], points: [{ value: { type: 'num', value: "'T'!$A$5" }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'colorScale', ranges: ['I1:I10'], points: [{ value: { type: 'num', value: '$M$15' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'expression', ranges: ['J1:J10'], formula: "J1>'T'!$A$8", style: { fill: { color: '#C6EFCE' } } },
    { type: 'cellIs', ranges: ['N1:N10'], operator: 'greaterThan', formulas: ['5'], style: { font: { color: '#0070C0' } } },
  ]
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }, { name: 'T', cells: { A5: v(4), A8: v(6) } }] }
}

export function buildCondCases(): CondCases {
  const cases: CondCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'held', op, read: [['S', 'B1:N10']] })
  const S = 'S'
  add('delete-row-10', { op: 'deleteRows', sheet: S, row: 10 })
  add('delete-col-d', { op: 'deleteColumns', sheet: S, column: 'D' })
  add('delete-col-f', { op: 'deleteColumns', sheet: S, column: 'F' })
  add('delete-col-e', { op: 'deleteColumns', sheet: S, column: 'E' })
  add('delete-row-5', { op: 'deleteRows', sheet: S, row: 5 })
  add('delete-rows-1', { op: 'deleteRows', sheet: S, row: 1, count: 2 })
  add('delete-row-5-t', { op: 'deleteRows', sheet: 'T', row: 5 })
  return { books: { held: heldBook() }, cases }
}
