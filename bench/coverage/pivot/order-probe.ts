/** Independent sort histories, mixed custom lists, and aggregation orders over the same source rows. */
import type { Agg, PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const cases: PivotCase[] = []
  const rows: PivotOp = { rows: ['Region'] }
  const value: PivotOp = { value: { field: 'Qty' } }
  const down: PivotOp = { sort: 'Region', order: 'descending', by: 'Sum of Qty' }
  const up: PivotOp = { sort: 'Region', order: 'ascending', by: 'Sum of Qty' }
  const read: PivotOp = { read: true }
  for (let i = 0; i < 3; i++) cases.push({ id: `ties-repeat-${i}`, data: 'shapes', ops: [rows, read, value, read, down, read, up, read, down] })
  cases.push(
    { id: 'ties-label-down-first', data: 'shapes', ops: [rows, value, { sort: 'Region', order: 'descending' }, read, down, read, up] },
    { id: 'ties-up-first', data: 'shapes', ops: [rows, value, up, read, down, read, up] },
    { id: 'ties-value-first', data: 'shapes', ops: [value, read, rows, read, down, read, up] },
    { id: 'ties-count-first', data: 'shapes', ops: [rows, { value: { field: 'Qty', fn: 'count' } }, { value: { field: 'Qty' } }, { sort: 'Region', order: 'descending', by: 'Count of Qty' }, read, down, read, up] },
    { id: 'ties-refresh', data: 'shapes', ops: [rows, value, down, read, { refresh: true }, read, { refresh: true }] },
  )
  const lists = [
    ['mon', 'December'], ['December', 'mon'], ['mon', 'Sunday'], ['Sunday', 'mon'],
    ['Jan', 'December'], ['December', 'Jan'], ['Wed', 'Jan'], ['Jan', 'Wed'],
    ['mon', 'December', 'Sunday'], ['Sunday', 'December', 'mon'],
    ['mon', 'December', 'Sunday', 'Jan', 'Feb', 'Mar', 'Apr', 'Wed'],
    ['Wed', 'Apr', 'Mar', 'Feb', 'Jan', 'Sunday', 'December', 'mon'],
  ]
  lists.forEach((labels, k) => {
    const edits: PivotOp[] = Array.from({ length: 48 }, (_, i) => ({ edit: `D!I${i + 2}`, to: labels[i % labels.length] }))
    cases.push({ id: `lists-${k}`, ops: [...edits, { cache: { missingItems: 'none' } }, { refresh: true }, { rows: ['Note'] }, { value: { field: 'Qty' } }, read, { sort: 'Note', order: 'descending' }, read, { sort: 'Note', order: 'ascending' }] })
  })
  for (const fn of ['sum', 'product', 'average', 'var', 'stdDev'] as Agg[])
    for (const columns of [false, true]) {
      const axes: PivotOp[] = [{ rows: ['Region', 'Product'] }, ...(columns ? [{ columns: ['Flag'] } as PivotOp] : [])]
      cases.push({ id: `arithmetic-${fn}-${columns ? 'columns' : 'rows'}`, ops: [...axes, { value: { field: 'Price', fn } }, read, { sort: 'Region', order: 'descending' }, { sort: 'Product', order: 'descending' }] })
    }
  return cases
}
