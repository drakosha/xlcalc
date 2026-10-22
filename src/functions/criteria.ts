/** Reference handling shared by the conditional math and statistical functions. */
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { E, isErr, Rng, type CErr, type V } from '../values.js'
import type { Cx } from '../args.js'

export const asRng = (x: V | Rng): Rng | CErr => x instanceof Rng ? x : isErr(x) ? x : E.VALUE

/**
 * SUMIF/AVERAGEIF resize the value range from its anchor to the criteria range's size. Past the sheet's edge both are
 * cut alike: a position there has a blank value, which neither sums nor averages, so Excel answers as if it were not
 * there (`SUMIF(A1:A3,">0",B1048575)` adds B1048575:B1048576).
 */
function resized(values: Rng, criteria: Rng): [Rng, Rng] {
  const rows = Math.min(criteria.r2 - criteria.r1, MAX_ROWS - 1 - values.r1)
  const cols = Math.min(criteria.c2 - criteria.c1, MAX_COLS - 1 - values.c1)
  const sum = new Rng(values.s, values.r1, values.c1, values.r1 + rows, values.c1 + cols)
  const cut = rows < criteria.r2 - criteria.r1 || cols < criteria.c2 - criteria.c1
  return [sum, cut ? new Rng(criteria.s, criteria.r1, criteria.c1, criteria.r1 + rows, criteria.c1 + cols) : criteria]
}

export function criteriaPairs(rest: (V | Rng)[]): [V | Rng, V | Rng][] {
  const out: [V | Rng, V | Rng][] = []
  for (let i = 0; i < rest.length; i += 2) out.push([rest[i], rest[i + 1]])
  return out
}

type Aggregate = (values: readonly V[], selected: readonly number[]) => V

/** SUMIF/AVERAGEIF share reference validation, range resizing and selection; their reductions stay distinct. */
export const aggregateIf = (reduce: Aggregate) => (cx: Cx, range: V | Rng, crit: V | Rng, values?: V | Rng): V => {
  const cg = asRng(range)
  if (isErr(cg)) return cg
  const vg = values === undefined ? cg : asRng(values)
  if (isErr(vg)) return vg
  const [g, cut] = values === undefined ? [cg, cg] : resized(vg, cg)
  const r = cx.select([[cut, crit]], g)
  return isErr(r) ? r : reduce(r.read(g), r.sel)
}

/** SUMIFS/AVERAGEIFS/MAXIFS/MINIFS share selection over a value range of the criteria's size. */
export const aggregateIfs = (reduce: Aggregate) => (cx: Cx, values: V | Rng, rest: (V | Rng)[]): V => {
  const g = asRng(values)
  if (isErr(g)) return g
  const r = cx.select(criteriaPairs(rest), g)
  return isErr(r) ? r : reduce(r.read(g), r.sel)
}
