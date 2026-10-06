/** COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS and MINIFS, this engine's own, on the criteria index of the runtime. */
import { MAX_COLS, MAX_ROWS } from '../../model.js'
import type { Cx, FnDef } from '../../args.js'
import { E, isErr, Rng, type CErr, type V } from '../../values.js'
import { max, mean, min } from './common.js'

const asRng = (x: V | Rng): Rng | CErr => (x instanceof Rng ? x : isErr(x) ? x : E.VALUE)

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

/** Excel's error texts, as a criterion of COUNTIF matches the cells holding that error. */
const ERROR_CODES: Record<string, string> = { '#DIV/0!': 'DIV0', '#N/A': 'NA', '#NAME?': 'NAME', '#NUM!': 'NUM', '#REF!': 'REF', '#VALUE!': 'VALUE' }

/** An error, or its text with `=` or `<>`, as a criterion. */
function errorCriterion(crit: V): { code: string; eq: boolean } | undefined {
  if (isErr(crit)) return { code: crit.code, eq: true }
  if (typeof crit !== 'string') return undefined
  const m = /^(=|<>)?([\s\S]*)$/.exec(crit)!
  const code = ERROR_CODES[m[2].toUpperCase()]
  return code ? { code, eq: m[1] !== '<>' } : undefined
}

function criteriaPairs(rest: (V | Rng)[]): [V | Rng, V | Rng][] {
  const out: [V | Rng, V | Rng][] = []
  for (let i = 0; i < rest.length; i += 2) out.push([rest[i], rest[i + 1]])
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
    fn: (cx: Cx, range: V | Rng, crit: V | Rng, avg?: V | Rng) => {
      const cg = asRng(range)
      if (isErr(cg)) return cg
      let ag = cg
      if (avg !== undefined) {
        const a = asRng(avg)
        if (isErr(a)) return a
        ag = new Rng(a.s, a.r1, a.c1, Math.min(a.r1 + cg.r2 - cg.r1, MAX_ROWS - 1), Math.min(a.c1 + cg.c2 - cg.c1, MAX_COLS - 1))
      }
      const r = cx.select([[cg, crit]], ag)
      if (isErr(r)) return r
      const xs = numbersAt(r.read(ag), r.sel)
      return isErr(xs) ? xs : xs.length ? mean(xs) : E.DIV0
    },
  },
  AVERAGEIFS: { args: 'R R a', rep: 2, cx: true, fn: ifs((xs) => (xs.length ? mean(xs) : E.DIV0)) },
  COUNTIF: {
    args: 'R a',
    cx: true,
    fn: (cx: Cx, range: V | Rng, crit: V) => {
      const g = asRng(range)
      if (isErr(g)) return g
      const e = errorCriterion(crit)
      if (e) {
        const grid = cx.grid(g)
        const n = grid.v.filter((v) => isErr(v) && v.code === e.code).length
        return e.eq ? n : grid.rows * grid.cols - n
      }
      const r = cx.select([[g, crit]])
      return isErr(r) ? r : r.sel.length + r.rest
    },
  },
  COUNTIFS: {
    args: 'R a',
    rep: 2,
    cx: true,
    fn: (cx: Cx, rest: (V | Rng)[]) => {
      const r = cx.select(criteriaPairs(rest))
      return isErr(r) ? r : r.sel.length + r.rest
    },
  },
  MAXIFS: { args: 'R R a', rep: 2, cx: true, fn: ifs((xs) => (xs.length ? max(xs) : 0)) },
  MINIFS: { args: 'R R a', rep: 2, cx: true, fn: ifs((xs) => (xs.length ? min(xs) : 0)) },
}
