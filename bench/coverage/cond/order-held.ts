/**
 * Held out after the set order fitted the rule of which rules deleted lines put after the others (`lateAfterDeletion`):
 * thresholds and formulas referring to another sheet whose lines move or go, a colour scale whose cells moved and whose
 * threshold went, a threshold over a range a deletion shrinks, a cell value rule between a lost and a moved reference,
 * a percent threshold of an icon set, a rule cut by a column deletion, a relative reference moved, and inserted lines.
 * Predicted before Excel ran it (`order-held.predicted.json`).
 */
import type { BookInput, ConditionalFormat, ConditionalValue, StructureOp } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'
import { valueCell as v } from '../stored.js'
import { RED } from './cases.js'

const num = (value: string): ConditionalValue => ({ type: 'num', value })
const max: ConditionalValue = { type: 'max' }

/** S: numbers 1 to 10 in rows 1 to 10 of B to J and N, thresholds in A12 and M15:M18; T: thresholds in A5 and A8. The last rule, over N, shows which rules went after it. */
function heldBook(): BookInput {
  const cells: Record<string, ReturnType<typeof v>> = { A12: v(3), M15: v(4), M16: v(7), M17: v(2), M18: v(9) }
  for (const c of ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'M', 'N']) for (let r = 1; r <= 10; r++) cells[`${c}${r}`] = v(r)
  const rules: ConditionalFormat[] = [
    { type: 'colorScale', ranges: ['B1:B10'], points: [{ value: num("'T'!$A$5"), color: '#F8696B' }, { value: max, color: '#63BE7B' }] },
    { type: 'expression', ranges: ['C1:C10'], formula: "C1>'T'!$A$8", style: { font: { bold: true } } },
    { type: 'colorScale', ranges: ['D1:D10'], points: [{ value: num('$A$12'), color: '#F8696B' }, { value: max, color: '#63BE7B' }] },
    { type: 'expression', ranges: ['E1:E10'], formula: 'M1>5', style: { font: { italic: true } } },
    { type: 'cellIs', ranges: ['F1:F10'], operator: 'between', formulas: ['$M$15', '$M$16'], style: RED },
    { type: 'colorScale', ranges: ['G1:G10'], points: [{ value: num('MAX($M$15:$M$18)'), color: '#F8696B' }, { value: max, color: '#63BE7B' }] },
    { type: 'iconSet', ranges: ['H1:H10'], icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '$M$15' } }, { value: { type: 'percent', value: '67' } }] },
    { type: 'cellIs', ranges: ['I1:J10'], operator: 'greaterThan', formulas: ['$M$16'], style: { fill: { color: '#C6EFCE' } } },
    { type: 'cellIs', ranges: ['N1:N10'], operator: 'greaterThan', formulas: ['5'], style: { font: { color: '#0070C0' } } },
  ]
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }, { name: 'T', cells: { A5: v(4), A8: v(6) } }] }
}

export function buildCondCases(): CondCases {
  const cases: CondCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'held', op, read: [['S', 'B1:N10']] })
  const S = 'S'
  add('delete-rows-t', { op: 'deleteRows', sheet: 'T', row: 2 })
  add('delete-row-8-t', { op: 'deleteRows', sheet: 'T', row: 8 })
  add('insert-rows-t', { op: 'insertRows', sheet: 'T', row: 2 })
  add('delete-col-a', { op: 'deleteColumns', sheet: S, column: 'A' })
  add('delete-col-k', { op: 'deleteColumns', sheet: S, column: 'K' })
  add('delete-row-15', { op: 'deleteRows', sheet: S, row: 15 })
  add('delete-row-16', { op: 'deleteRows', sheet: S, row: 16 })
  add('delete-col-j', { op: 'deleteColumns', sheet: S, column: 'J' })
  add('delete-row-1', { op: 'deleteRows', sheet: S, row: 1 })
  add('insert-col-k', { op: 'insertColumns', sheet: S, column: 'K' })
  return { books: { held: heldBook() }, cases }
}
