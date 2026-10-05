/**
 * Math and trigonometry functions.
 *
 * SUM and ROUND (the engine called them in formula.js before) are ported from formula.js 4.6.1
 * (https://github.com/formulajs/formulajs), src/math-trig.js: Copyright (c) 2014 Sutoiku, Inc., MIT License
 * (full text in THIRD_PARTY.md). Changed: arguments are coerced by the engine's argument layer
 * (`src/args.ts`; SUM: a range gives its numbers, a direct argument is coerced); ROUND rounds on the
 * decimal representation of any magnitude (`roundDigits`).
 * SUMIF, SUMIFS, RAND and RANDBETWEEN are this engine's own.
 */
import type { Cx, FnDef } from '../args.js'
import { E, isErr, Rng, type CErr, type V } from '../values.js'
import { roundDigits } from './common.js'

/** Excel's power: 0^0 is #NUM!, 0 to a negative power #DIV/0!. */
export function power(x: number, y: number): number | CErr {
  if (x === 0 && y === 0) return E.NUM
  if (x === 0 && y < 0) return E.DIV0
  return Math.pow(x, y)
}

/** The sum of the numbers at the selected positions of a range; an error there is the answer. */
function sumAt(cx: Cx, g: Rng, sel: number[]): V {
  const f = cx.flat(g)
  let acc = 0
  for (const p of sel) {
    const v = f[p]
    if (typeof v === 'number') acc += v
    else if (isErr(v)) return v
  }
  return acc
}

const asRng = (x: V | Rng): Rng | CErr => (x instanceof Rng ? x : isErr(x) ? x : E.VALUE)

export const functions: Record<string, FnDef> = {
  RAND: { args: '', cx: true, cell: true, volatile: true, fn: (cx: Cx, s: number, r: number, c: number) => cx.rand(s, r, c) },
  RANDBETWEEN: {
    args: 'n n',
    cx: true,
    cell: true,
    volatile: true,
    fn: (cx: Cx, s: number, r: number, c: number, lo: number, hi: number) => {
      const x = Math.ceil(lo)
      const y = Math.floor(hi)
      if (y < x) return E.NUM
      return x + Math.floor(cx.rand(s, r, c) * (y - x + 1))
    },
  },
  ROUND: { args: 'n i', fn: (x: number, d: number) => roundDigits(x, d, Math.round) },
  // a range gives its numbers (text and booleans in it are skipped), a direct argument is coerced
  SUM: { args: 'N', rep: 1, fn: (xs: number[]) => xs.reduce((s, x) => s + x, 0) },
  // the sum range takes the size of the criteria range from its own top-left cell, as in Excel
  SUMIF: {
    args: 'R a R?',
    cx: true,
    fn: (cx: Cx, range: V | Rng, crit: V | Rng, sum?: V | Rng) => {
      const cg = asRng(range)
      if (isErr(cg)) return cg
      const r = cx.select([[cg, crit]])
      if (isErr(r)) return r
      if (sum === undefined) return sumAt(cx, cg, r.sel)
      const sg = asRng(sum)
      if (isErr(sg)) return sg
      const b = cx.bounds(cg)
      return sumAt(cx, new Rng(sg.s, sg.r1, sg.c1, sg.r1 + b.h - 1, sg.c1 + b.w - 1), r.sel)
    },
  },
  SUMIFS: {
    args: 'R R a',
    rep: 2,
    cx: true,
    fn: (cx: Cx, sum: V | Rng, rest: (V | Rng)[]) => {
      const sg = asRng(sum)
      if (isErr(sg)) return sg
      const pairs: [V | Rng, V | Rng][] = []
      for (let i = 0; i < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]])
      const r = cx.select(pairs, cx.bounds(sg))
      return isErr(r) ? r : sumAt(cx, sg, r.sel)
    },
  },
}
