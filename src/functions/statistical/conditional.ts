/** COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS and MINIFS, this engine's own, on the criteria index of the runtime. */
import type { Cx, FnDef } from '../../args.js'
import { E, isErr, type Rng, type V } from '../../values.js'
import { max, min, NEWER_MAX_ARGS } from '../common.js'
import { aggregateIf, aggregateIfs, asRng, criteriaPairs } from '../criteria.js'
import { mean, numbersOf } from './common.js'

/** A statistical function of the selected numbers, with reference errors kept in selection order. */
function over(f: (xs: number[]) => V) {
  return (values: readonly V[], selected: number[]): V => {
    const xs = numbersOf(values, selected)
    return isErr(xs) ? xs : f(xs)
  }
}

const average = over((xs) => xs.length ? mean(xs) : E.DIV0)

export const functions: Record<string, FnDef> = {
  AVERAGEIF: {
    args: 'R a R?',
    cx: true,
    criteria: true,
    fn: aggregateIf(average),
  },
  AVERAGEIFS: { args: 'R R a', rep: 2, cx: true, criteria: true, fn: aggregateIfs(average) },
  COUNTIF: {
    args: 'R a',
    cx: true,
    criteria: true,
    fn: (cx: Cx, range: V | Rng, crit: V | Rng) => {
      const g = asRng(range)
      if (isErr(g)) return g
      const r = cx.select([[g, crit]])
      return isErr(r) ? r : r.sel.length + r.rest
    },
  },
  COUNTIFS: {
    args: 'R a',
    rep: 2,
    cx: true,
    criteria: true,
    fn: (cx: Cx, rest: (V | Rng)[]) => {
      const r = cx.select(criteriaPairs(rest))
      return isErr(r) ? r : r.sel.length + r.rest
    },
  },
  MAXIFS: { args: 'R R a', rep: 2, max: NEWER_MAX_ARGS, cx: true, criteria: true, fn: aggregateIfs(over(max)) },
  MINIFS: { args: 'R R a', rep: 2, max: NEWER_MAX_ARGS, cx: true, criteria: true, fn: aggregateIfs(over(min)) },
}
