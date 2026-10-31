/** The source records Excel returns for a value, with and without manual filtering. */
import type { PivotCase } from './cases.js'
export function buildPivotCases(): PivotCase[] {
  return [
    { id: 'page-north', from: 'basic/page-one', ops: [{ drill: 'P!B5' }] },
    { id: 'hidden-west', from: 'numeric-held/hidden-plain', ops: [{ drill: 'P!B6' }] },
    { id: 'plain-west', from: 'basic/price-sum', ops: [{ drill: 'P!B7' }] },
    { id: 'plain-north', from: 'basic/price-sum', ops: [{ drill: 'P!B5' }] },
  ]
}
