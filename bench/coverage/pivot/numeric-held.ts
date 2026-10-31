/** New numeric vectors, mixed aggregates, and equivalent filter definitions reached through different histories. */
import type { Agg, PivotCase, PivotOp } from './cases.js'
import { NUMERIC_ORDER, numericVectors } from './numeric-order-data.js'

export function buildPivotCases(): PivotCase[] {
  const vectors = numericVectors(0xbb67ae85)
  const variants: [string, PivotOp[]][] = (['sum', 'product', 'average', 'var', 'stdDev'] as Agg[]).map(fn => [fn, [{ value: { field: 'V', fn } }]])
  for (const fn of ['sum', 'count', 'max'] as Agg[]) variants.push([`sum-and-${fn}`, [{ value: { field: 'V' } }, { value: { field: 'V', fn, caption: 'Second V' } }]])
  variants.push(['sum-and-formula', [{ value: { field: 'V' } }, { calcField: 'Half', formula: '=V/2' }, { value: { field: 'Half' } }]])
  const cases = variants.map(([id, values]): PivotCase => {
    const ops: PivotOp[] = [{ rows: ['R'] }, { columns: ['C'] }, ...values, { read: true }]
    for (let k = 1; k < vectors.length; k++) {
      NUMERIC_ORDER.forEach((i, r) => ops.push({ edit: `D!C${r + 2}`, to: vectors[k][i] }))
      ops.push({ refresh: true })
      if (k < vectors.length - 1) ops.push({ read: true })
    }
    return { id, data: 'numeric-held', ops }
  })
  const row: PivotOp = { rows: ['Region'] }, page: PivotOp = { pages: ['Product'] }, value: PivotOp = { value: { field: 'Price' } }
  const apple: PivotOp = { page: 'Product', items: ['Apple'] }, hide: PivotOp = { hide: 'Region', items: ['North'] }
  const read: PivotOp = { read: true }, refresh: PivotOp = { refresh: true }
  const run: PivotOp = { value: { field: 'Price', show: 'runTotal', base: 'Region' } }
  cases.push(
    { id: 'page-after-value', ops: [row, page, value, read, apple, read, refresh, read, refresh] },
    { id: 'page-before-value', ops: [row, page, apple, value, read, refresh] },
    { id: 'page-via-cherry', ops: [row, page, value, { page: 'Product', items: ['Cherry'] }, read, apple, read, { page: 'Product', items: '(All)' }, read, apple] },
    { id: 'hidden-run-after', ops: [row, run, hide, read, refresh] },
    { id: 'hidden-run-before', ops: [row, hide, run, read, refresh] },
    { id: 'hidden-plain', ops: [row, value, hide, read, refresh] },
    { id: 'page-run-before', ops: [row, page, { page: 'Product', items: ['Apple', 'Cherry'] }, run, read, refresh] },
  )
  return cases
}
