/** Reference handling shared by the conditional math and statistical functions. */
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { E, isErr, Rng, type CErr, type V } from '../values.js'
import type { Cx } from '../args.js'

export const asRng = (x: V | Rng): Rng | CErr => x instanceof Rng ? x : isErr(x) ? x : E.VALUE

/** SUMIF/AVERAGEIF resize the value range from its anchor, cutting it at the sheet edge. */
export const resizeRange = (values: Rng, criteria: Rng): Rng => new Rng(
  values.s, values.r1, values.c1,
  Math.min(values.r1 + criteria.r2 - criteria.r1, MAX_ROWS - 1),
  Math.min(values.c1 + criteria.c2 - criteria.c1, MAX_COLS - 1),
)

export function criteriaPairs(rest: (V | Rng)[]): [V | Rng, V | Rng][] {
  const out: [V | Rng, V | Rng][] = []
  for (let i = 0; i < rest.length; i += 2) out.push([rest[i], rest[i + 1]])
  return out
}

type Aggregate = (values: readonly V[], selected: number[]) => V

/** SUMIF/AVERAGEIF share reference validation, range resizing and selection; their reductions stay distinct. */
export const aggregateIf = (reduce: Aggregate) => (cx: Cx, range: V | Rng, crit: V | Rng, values?: V | Rng): V => {
  const cg = asRng(range)
  if (isErr(cg)) return cg
  const vg = values === undefined ? cg : asRng(values)
  if (isErr(vg)) return vg
  const g = values === undefined ? cg : resizeRange(vg, cg)
  const r = cx.select([[cg, crit]], g)
  return isErr(r) ? r : reduce(r.read(g), r.sel)
}

/** SUMIFS/AVERAGEIFS/MAXIFS/MINIFS share selection over a value range of the criteria's size. */
export const aggregateIfs = (reduce: Aggregate) => (cx: Cx, values: V | Rng, rest: (V | Rng)[]): V => {
  const g = asRng(values)
  if (isErr(g)) return g
  const r = cx.select(criteriaPairs(rest), g)
  return isErr(r) ? r : reduce(r.read(g), r.sel)
}
