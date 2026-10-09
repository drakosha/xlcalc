/** COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS and MINIFS, this engine's own, on the criteria index of the runtime. */
import type { Cx, FnDef } from '../../args.js'
import { E, isErr, type Rng, type CErr, type V } from '../../values.js'
import { max, min, NEWER_MAX_ARGS } from '../common.js'
import { asRng, criteriaPairs, resizeRange } from '../criteria.js'
import { mean } from './common.js'

/** The numbers at the selected positions; an error there is the answer. */
function numbersAt(values: readonly V[], sel: number[]): number[] | CErr {
  const out: number[] = []
  for (const p of sel) {
    const v = values[p]
    if (typeof v === 'number') out.push(v)
    else if (isErr(v)) return v
  }
  return out
}

/** AVERAGEIFS, MAXIFS, MINIFS: `f` of the numbers of the value range where every criterion holds. */
function ifs(f: (xs: number[]) => V) {
  return (cx: Cx, values: V | Rng, rest: (V | Rng)[]) => {
    const g = asRng(values)
    if (isErr(g)) return g
    const r = cx.select(criteriaPairs(rest), g)
    if (isErr(r)) return r
    const xs = numbersAt(r.read(g), r.sel)
    return isErr(xs) ? xs : f(xs)
  }
}

export const functions: Record<string, FnDef> = {
  // the average range takes the size of the criteria range from its own top-left cell, as SUMIF's sum range
  AVERAGEIF: {
    args: 'R a R?',
    cx: true,
    criteria: true,
    fn: (cx: Cx, range: V | Rng, crit: V | Rng, avg?: V | Rng) => {
      const cg = asRng(range)
      if (isErr(cg)) return cg
      let ag = cg
      if (avg !== undefined) {
        const a = asRng(avg)
        if (isErr(a)) return a
        ag = resizeRange(a, cg)
      }
      const r = cx.select([[cg, crit]], ag)
      if (isErr(r)) return r
      const xs = numbersAt(r.read(ag), r.sel)
      return isErr(xs) ? xs : xs.length ? mean(xs) : E.DIV0
    },
  },
  AVERAGEIFS: { args: 'R R a', rep: 2, cx: true, criteria: true, fn: ifs((xs) => (xs.length ? mean(xs) : E.DIV0)) },
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
  MAXIFS: { args: 'R R a', rep: 2, max: NEWER_MAX_ARGS, cx: true, criteria: true, fn: ifs(max) },
  MINIFS: { args: 'R R a', rep: 2, max: NEWER_MAX_ARGS, cx: true, criteria: true, fn: ifs(min) },
}
