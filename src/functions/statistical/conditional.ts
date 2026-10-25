/** COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS and MINIFS, this engine's own, on the criteria index of the runtime. */
import type { Cx, FnDef } from '../../args.js'
import { isErr, type Rng, type V } from '../../values.js'
import { max, min, NEWER_MAX_ARGS } from '../common.js'
import { aggregateIf, aggregateIfs, criteriaPairs } from '../criteria.js'
import { meanOrDiv0, numbersOf } from './common.js'

/** A statistical function of the selected numbers, with reference errors kept in selection order. */
function over(f: (xs: number[]) => V) {
  return (values: readonly V[], selected: readonly number[]): V => {
    const xs = numbersOf(values, selected)
    return isErr(xs) ? xs : f(xs)
  }
}

const average = over(meanOrDiv0)

/** COUNTIF and COUNTIFS: the count of the selection, the criteria pairs' first range read by the runtime's rule. */
const count = (cx: Cx, pairs: [V | Rng, V | Rng][]) => {
  const r = cx.select(pairs)
  return isErr(r) ? r : r.sel.length + r.rest
}

export const functions: Record<string, FnDef> = {
  AVERAGEIF: {
    args: 'R a R?',
    cx: true,
    criteria: true,
    fn: aggregateIf(average),
  },
  AVERAGEIFS: { args: 'R R a', rep: 2, cx: true, criteria: true, fn: aggregateIfs(average) },
  COUNTIF: { args: 'R a', cx: true, criteria: true, fn: (cx: Cx, range: V | Rng, crit: V | Rng) => count(cx, [[range, crit]]) },
  COUNTIFS: { args: 'R a', rep: 2, cx: true, criteria: true, fn: (cx: Cx, rest: (V | Rng)[]) => count(cx, criteriaPairs(rest)) },
  MAXIFS: { args: 'R R a', rep: 2, max: NEWER_MAX_ARGS, cx: true, criteria: true, fn: aggregateIfs(over(max)) },
  MINIFS: { args: 'R R a', rep: 2, max: NEWER_MAX_ARGS, cx: true, criteria: true, fn: aggregateIfs(over(min)) },
}
