/** Aggregation over the same six records and grouping for every function; values vary between refreshes. */
import type { Agg, PivotCase, PivotOp } from './cases.js'
import { NUMERIC_ORDER, numericVectors } from './numeric-order-data.js'

export function buildPivotCases(): PivotCase[] {
  return (['sum', 'product', 'average', 'var', 'stdDev'] as Agg[]).map(fn => {
    const vectors = numericVectors()
    const ops: PivotOp[] = [{ rows: ['R'] }, { columns: ['C'] }, { value: { field: 'V', fn } }, { read: true }]
    for (let k = 1; k < vectors.length; k++) {
      NUMERIC_ORDER.forEach((i, r) => ops.push({ edit: `D!C${r + 2}`, to: vectors[k][i] }))
      ops.push({ refresh: true })
      if (k < vectors.length - 1) ops.push({ read: true })
    }
    return { id: fn, data: 'numeric-order', ops }
  })
}
