/**
 * Math and trigonometry functions.
 *
 * ABS … TRUNC, SUM and ROUND among them (the engine called them in formula.js before), are ported from
 * formula.js 4.6.1 (https://github.com/formulajs/formulajs), src/math-trig.js (CEILING and FLOOR also as
 * exported by src/compatibility.js): Copyright (c) 2014 Sutoiku, Inc., MIT License (full text in
 * THIRD_PARTY.md). Changed: arguments are coerced by the engine's argument layer (`src/args.ts`) instead of
 * formula.js parsing (SUM: a range gives its numbers, a direct argument is coerced); ROUND, ROUNDDOWN,
 * ROUNDUP and TRUNC share one decimal rounding (`roundDigits`; TRUNC truncated in binary before); CEILING
 * and FLOOR round a quotient that is whole to 15 digits as whole; LOG with base 1 is #DIV/0! and an explicit
 * base 0 is #NUM! (formula.js read base 0 as 10); POWER(0, negative) is #DIV/0!; PRODUCT of no numbers is 0.
 * SUMIF, SUMIFS, RAND and RANDBETWEEN are this engine's own.
 */
import { E, isErr, Rng, type CErr, type V } from '../values.js'
import type { Cx, FnDef } from '../args.js'
import { quotient, roundDigits } from './common.js'

/** CEILING.MATH and FLOOR.MATH: significance taken as positive; `mode` rounds negative numbers away from zero. */
function roundMath(n: number, sig: number | undefined, mode: number | undefined, up: boolean): number {
  const s = Math.abs(sig ?? 1)
  if (s === 0) return 0
  const toward = up ? Math.ceil : Math.floor
  const away = up ? Math.floor : Math.ceil
  return (mode && n < 0 ? away : toward)(quotient(n, s)) * s
}

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
  ABS: { args: 'n', fn: Math.abs },
  ACOS: { args: 'n', fn: Math.acos },
  ACOSH: { args: 'n', fn: (x: number) => Math.log(x + Math.sqrt(x * x - 1)) },
  ACOTH: { args: 'n', fn: (x: number) => 0.5 * Math.log((x + 1) / (x - 1)) },
  ASIN: { args: 'n', fn: Math.asin },
  ASINH: { args: 'n', fn: (x: number) => Math.log(x + Math.sqrt(x * x + 1)) },
  ATAN: { args: 'n', fn: Math.atan },
  ATANH: { args: 'n', fn: (x: number) => Math.log((1 + x) / (1 - x)) / 2 },
  CEILING: {
    args: 'n n',
    fn: (n: number, s: number) => {
      if (s === 0) return 0
      if (n > 0 && s < 0) return E.NUM
      return Math.ceil(quotient(n, s)) * s
    },
  },
  'CEILING.MATH': { args: 'n n? i?', fn: (n: number, s?: number, mode?: number) => roundMath(n, s, mode, true) },
  'CEILING.PRECISE': { args: 'n n?', fn: (n: number, s?: number) => roundMath(n, s, 0, true) },
  COSH: { args: 'n', fn: (x: number) => (Math.exp(x) + Math.exp(-x)) / 2 },
  CSCH: { args: 'n', fn: (x: number) => (x === 0 ? E.DIV0 : 2 / (Math.exp(x) - Math.exp(-x))) },
  DEGREES: { args: 'n', fn: (x: number) => (x * 180) / Math.PI },
  EVEN: { args: 'n', fn: (x: number) => roundMath(x, 2, 1, true) },
  EXP: { args: 'n', fn: Math.exp },
  FLOOR: {
    args: 'n n',
    fn: (n: number, s: number) => {
      if (s === 0) return E.DIV0
      if (n > 0 && s < 0) return E.NUM
      return Math.floor(quotient(n, s)) * s
    },
  },
  'FLOOR.MATH': { args: 'n n? i?', fn: (n: number, s?: number, mode?: number) => roundMath(n, s, mode, false) },
  'FLOOR.PRECISE': { args: 'n n?', fn: (n: number, s?: number) => roundMath(n, s, 0, false) },
  INT: { args: 'n', fn: Math.floor },
  LN: { args: 'n', fn: (x: number) => (x <= 0 ? E.NUM : Math.log(x)) },
  LOG: {
    args: 'n n?',
    fn: (x: number, base = 10) => {
      if (x <= 0 || base <= 0) return E.NUM
      return base === 1 ? E.DIV0 : Math.log(x) / Math.log(base)
    },
  },
  LOG10: { args: 'n', fn: (x: number) => (x <= 0 ? E.NUM : Math.log(x) / Math.log(10)) },
  ODD: {
    args: 'n',
    fn: (x: number) => {
      let t = Math.ceil(Math.abs(x))
      t = t & 1 ? t : t + 1
      return x >= 0 ? t : -t
    },
  },
  PI: { args: '', fn: () => Math.PI },
  POWER: { args: 'n n', fn: power },
  PRODUCT: { args: 'N', rep: 1, fn: (xs: number[]) => (xs.length ? xs.reduce((p, x) => p * x, 1) : 0) },
  RADIANS: { args: 'n', fn: (x: number) => (x * Math.PI) / 180 },
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
  ROUNDDOWN: { args: 'n i', fn: (x: number, d: number) => roundDigits(x, d, Math.floor) },
  ROUNDUP: { args: 'n i', fn: (x: number, d: number) => roundDigits(x, d, Math.ceil) },
  SECH: { args: 'n', fn: (x: number) => 2 / (Math.exp(x) + Math.exp(-x)) },
  SIGN: { args: 'n', fn: Math.sign },
  SINH: { args: 'n', fn: (x: number) => (Math.exp(x) - Math.exp(-x)) / 2 },
  SQRT: { args: 'n', fn: (x: number) => (x < 0 ? E.NUM : Math.sqrt(x)) },
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
  TRUNC: { args: 'n i?', fn: (x: number, d = 0) => roundDigits(x, d, Math.floor) },
}
