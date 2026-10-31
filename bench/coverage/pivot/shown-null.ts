/**
 * The set `shown-null`: what the empty cells of a grid show, with errors among the values: Amount by Product without
 * the text for empty cells, Price by Product with `-`, Amount by Product with `x` (to tell the text for empty cells from
 * cells that hold no value).
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const grid = (field: string): PivotOp[] => [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field } }]
  return [
    { id: 'amount-plain', ops: grid('Amount') },
    { id: 'price-dash', ops: [...grid('Price'), { table: { emptyText: '-' } }] },
    { id: 'amount-x', ops: [...grid('Amount'), { table: { emptyText: 'x' } }] },
  ]
}
