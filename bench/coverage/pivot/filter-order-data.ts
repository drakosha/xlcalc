/** Fresh numeric inputs and distinct error priorities for the record-partition holdout. */
import type { SourceColumn } from './data.js'
import { numericOrderColumns } from './numeric-order-data.js'

export function filterOrderColumns(errors = false): SourceColumn[] {
  const cols = numericOrderColumns(0x3c6ef372)
  cols[0] = { name: 'R', values: Array.from({ length: 6 }, () => ({ value: 'A' })) }
  cols[1] = { name: 'F', values: ['hide', 'keep', 'hide', 'keep', 'keep', 'keep'].map(value => ({ value })) }
  cols.push({ name: 'G', values: ['X', 'X', 'X', 'X', 'X', 'Y'].map(value => ({ value })) })
  if (errors) cols[2] = { name: 'V', values: [{ value: 1 }, { formula: 'NA()' }, { value: 2 }, { formula: '1/0' }, { formula: 'VALUE("x")' }, { formula: 'SQRT(-1)' }] }
  return cols
}
