/**
 * Held out after the set held: which pastes grow a rule rather than add one (down its columns, as the engine takes it,
 * or with no relative reference in its formula, the other reading of the sets so far), and average and duplicate rules
 * that lose rows, taken as the other rules over statistics. Predicted before Excel ran it (`held-2.predicted.json`).
 */
import type { BookInput, CellInput, ConditionalFormat, StructureOp } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'
import { valueCell as v } from '../stored.js'
import { RED } from './cases.js'

/** S: numbers 0 to 9 in A1:A10 and 1 to 10 in B:H; a rule without references on B and H, one with a relative one on D, average and duplicates on F and G. */
function heldBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 10; r++) {
    cells[`A${r}`] = v(r - 1)
    for (const c of ['B', 'C', 'D', 'E', 'F', 'G', 'H']) cells[`${c}${r}`] = v(c === 'G' ? r % 4 : r)
  }
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['B2:B10'], operator: 'greaterThan', formulas: ['5'], style: RED },
    { type: 'expression', ranges: ['D2:D10'], formula: '$A2>4', style: { font: { italic: true } } },
    { type: 'aboveAverage', ranges: ['F2:F10'], style: { font: { bold: true } } },
    { type: 'duplicateValues', ranges: ['G2:G10'], style: RED },
    { type: 'cellIs', ranges: ['H2:H10'], operator: 'lessThan', formulas: ['3'], style: { fill: { color: '#C6EFCE' } } },
  ]
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }] }
}

export function buildCondCases(): CondCases {
  const cases: CondCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'held', op, read: [['S', 'B1:H12']] })
  const S = 'S'
  add('fill-right-constant', { op: 'copyRange', from: { sheet: S, range: 'B2:B10' }, to: { sheet: S, range: 'C2:C10' } })
  add('fill-down-relative', { op: 'copyRange', from: { sheet: S, range: 'D10' }, to: { sheet: S, range: 'D11:D12' } })
  add('fill-up-constant', { op: 'copyRange', from: { sheet: S, range: 'B2' }, to: { sheet: S, range: 'B1' } })
  add('delete-rows-average', { op: 'deleteRows', sheet: S, row: 4, count: 2 })
  return { books: { held: heldBook() }, cases }
}
