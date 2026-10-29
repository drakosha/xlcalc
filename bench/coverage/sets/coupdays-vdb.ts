/**
 * COUPDAYS and VDB on the cases their rules were fitted to: the coupon schedules of `findate` on every basis, COUPDAYS
 * on basis 1 around month ends (maturities on days 28 to 31, settlements through two years), VDB over whole and
 * fractional periods, near the end of the life, without the switch, with factors and edge arguments. Excel 16.113 for
 * Mac answered them once before the functions were admitted; this run stores the answers the tests read
 * (`test/excel/cases/coupdays-vdb.json`). The rules' own check, predicted before Excel ran it, is `coupdays-vdb-held`.
 */
import type { Probe } from '../probes.js'
import { byFunction } from './arrays/common.js'

const D = (y: number, m: number, d: number) => `DATE(${y},${m},${d})`
const call = (fn: string, args: readonly (string | number | boolean)[]) => `${fn}(${args.map((a) => (typeof a === 'boolean' ? (a ? 'TRUE' : 'FALSE') : a)).join(',')})`

/** Settlement, maturity and frequency of `findate`'s coupon cases. */
const COUPONS: [string, string, number][] = [
  [D(2011, 1, 25), D(2011, 11, 15), 2], [D(2011, 1, 25), D(2011, 11, 15), 4], [D(2011, 1, 25), D(2011, 11, 15), 1], [D(2011, 3, 15), D(2015, 8, 31), 2], [D(2012, 3, 15), D(2015, 8, 31), 4],
  [D(2011, 3, 1), D(2015, 2, 28), 2], [D(2012, 2, 29), D(2015, 8, 31), 2], [D(2011, 8, 31), D(2015, 2, 28), 2], [D(2011, 5, 15), D(2011, 11, 15), 2], [D(2011, 11, 14), D(2011, 11, 15), 2],
  [D(2008, 2, 15), D(2017, 11, 15), 2], [D(2011, 9, 30), D(2012, 3, 31), 4], [D(2011, 12, 31), D(2012, 6, 30), 2], ['60', D(2011, 11, 15), 2], ['59', D(2011, 11, 15), 2], ['61', D(2011, 11, 15), 2],
]

const MATURITIES = [D(2015, 8, 31), D(2015, 8, 30), D(2015, 8, 28), D(2015, 2, 28), D(2016, 2, 29), D(2015, 6, 30), D(2015, 5, 31), D(2015, 11, 30), D(2015, 12, 31), D(2015, 9, 30), D(2015, 3, 31)]
const SETTLEMENTS: [number, number, number][] = [[2011, 1, 15], [2011, 2, 28], [2011, 3, 15], [2011, 5, 31], [2011, 6, 30], [2011, 8, 31], [2011, 9, 15], [2011, 11, 30], [2011, 12, 31], [2012, 2, 29], [2012, 3, 1], [2012, 4, 30], [2012, 8, 30]]

const VDB: (string | number | boolean)[][] = [
  [2400, 300, 3650, 0, 1], [2400, 300, 120, 0, 1], [2400, 300, 10, 0, 1], [2400, 300, 120, 6, 18], [2400, 300, 120, 6, 18, 1.5], [2400, 300, 10, 0, 0.875, 1.5],
  [2400, 300, 10, 0, 10], [2400, 300, 10, 5, 10], [2400, 300, 10, 2.5, 7.5], [2400, 300, 10, 0, 10, 2, true], [2400, 300, 10, 3.5, 9.25, 2, true], [2400, 300, 10, 8, 10, 3],
  [2400, 0, 5, 0, 5, 1], [2400, 300, 10, 0, 11], [2400, 300, 10, 4, 3], [2400, 300, 10, -1, 3], [2400, 300, 10.5, 0, 10.5], [2400, 300, 10, 9.5, 10], [10000, 1000, 5, 1, 4, 1.5, false],
  [2400, 300, 10, 0, 1, 0], [2400, 3000, 10, 0, 5], ['TRUE', 0, 10, 0, 1], [2400, 300, 10, 0, 1, 2, '"abc"'],
  // near the end of the life and in fractional periods
  ...[[9, 10], [9, 9.5], [9.25, 10], [9.75, 10], [8.5, 9], [8.5, 10], [0, 9.5], [5.5, 6], [1.5, 2], [6.5, 7], [7.5, 8], [8, 8.5], [8.9, 9.1]].map(([s, e]) => [2400, 300, 10, s, e]),
  ...[[9.5, 10], [8.5, 9]].map(([s, e]) => [2400, 300, 10, s, e, 2, true]),
  [2400, 300, 10, 2.5, 3.5, 1.5],
  [2400, 0, 10, 9.5, 10],
]

export function buildProbes(): Probe[] {
  const { out, add } = byFunction(false)
  for (const [s, m, f] of COUPONS) for (const basis of [0, 1, 2, 3, 4]) add('COUPDAYS', call('COUPDAYS', [s, m, f, basis]))
  for (const m of MATURITIES) for (const [y, mo, d] of SETTLEMENTS) for (const f of [1, 2, 4]) add('COUPDAYS', call('COUPDAYS', [D(y, mo, d), m, f, 1]))
  for (const args of VDB) add('VDB', call('VDB', args))
  return out
}
