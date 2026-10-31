/** Held-out filtered aggregation: the same records partitioned by one or two fields, plus distinct errors. */
import type { Agg, PivotCase, PivotOp } from './cases.js'
import { NUMERIC_ORDER, numericVectors } from './numeric-order-data.js'

export function buildPivotCases(): PivotCase[] {
  const vectors = numericVectors(0x3c6ef372).slice(0, 12)
  const forms: [string, Agg, boolean][] = (['sum', 'product', 'average', 'var', 'stdDev'] as Agg[]).map(fn => [fn, fn, false])
  forms.push(['two-filters', 'sum', true])
  const cases = forms.map(([id, fn, two]): PivotCase => {
    const ops: PivotOp[] = [{ rows: ['R'] }, { pages: two ? ['F', 'G'] : ['F'] }, { page: 'F', items: ['keep'] }]
    if (two) ops.push({ page: 'G', items: ['X'] })
    ops.push({ value: { field: 'V', fn } }, { read: true })
    for (let k = 1; k < vectors.length; k++) {
      NUMERIC_ORDER.forEach((i, r) => ops.push({ edit: `D!C${r + 2}`, to: vectors[k][i] }))
      ops.push({ refresh: true })
      if (k < vectors.length - 1) ops.push({ read: true })
    }
    return { id, data: 'filter-numeric', ops }
  })
  for (const filtered of [false, true]) cases.push({ id: filtered ? 'errors-filtered' : 'errors-all', data: 'filter-errors', ops: [
    { rows: ['R'] }, { pages: ['F'] }, ...(filtered ? [{ page: 'F', items: ['keep'] } as PivotOp] : []),
    { value: { field: 'V' } }, { calcField: 'Twice', formula: '=V*2' }, { value: { field: 'Twice' } },
  ] })
  return cases
}
