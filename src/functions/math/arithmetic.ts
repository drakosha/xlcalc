/**
 * Powers, logarithms, MOD, QUOTIENT, random numbers. ABS … SQRT ported from formula.js 4.6.1 (src/math-trig.js),
 * Copyright (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); the rest this engine's own.
 */
import { E, type CErr } from '../../values.js'
import type { Cx, FnDef } from '../../args.js'

/** The largest 1/y Excel takes as a whole number below 2^31 - 1 when it rounds it (the set `power-bound-held`). */
const MAX_ROOT = 2 ** 31 - 1.5

/**
 * Excel's power: 0^0 is #NUM!, 0 to a negative power #DIV/0!. A negative number to a fraction is its real root where
 * 1/y, to 15 significant digits, is an odd whole number below 2^31 - 1, computed as -exp(y·ln(-x)) as Excel does
 * (`(-8)^(1/3)` is -1.9999999999999998), else #NUM! (the sets `power-neg-held`, `power-edges-held`).
 */
export function power(x: number, y: number): number | CErr {
  if (x === 0 && y === 0) return E.NUM
  if (x === 0 && y < 0) return E.DIV0
  if (x >= 0 || Number.isInteger(y)) return Math.pow(x, y)
  const n = Number((1 / y).toPrecision(15))
  return Math.abs(1 / y) < MAX_ROOT && n % 2 !== 0 && Number.isInteger(n) ? -Math.exp(y * Math.log(-x)) : E.NUM
}

/** Excel: MOD is #NUM! once the quotient reaches 2^50 / 1000 (MOD(1.1258E12,1) is 0, MOD(1.1259E12,1) #NUM!). */
const MAX_MOD_QUOTIENT = 2 ** 50 / 1000

/** The remainder with the divisor's sign, exact (fmod); a remainder that rounds to the divisor is 0 (MOD(-1E-300,1)). */
function mod(n: number, d: number): number | CErr {
  if (d === 0) return E.DIV0
  if (Math.abs(n / d) >= MAX_MOD_QUOTIENT) return E.NUM
  let r = n % d
  if (r !== 0 && r < 0 !== d < 0) r += d
  return r === d ? 0 : r + 0
}

/** LOG, and LOG10 by its default base. */
function log(x: number, base = 10): number | CErr {
  if (x <= 0 || base <= 0) return E.NUM
  return base === 1 ? E.DIV0 : Math.log(x) / Math.log(base)
}

export const functions: Record<string, FnDef> = {
  ABS: { args: 'n', fn: Math.abs },
  EXP: { args: 'n', fn: Math.exp },
  LN: { args: 'n', fn: (x: number) => (x <= 0 ? E.NUM : Math.log(x)) },
  LOG: { args: 'n n?', fn: log },
  LOG10: { args: 'n', fn: log },
  MOD: { args: 'n n', fn: mod },
  POWER: { args: 'n n', fn: power },
  QUOTIENT: { args: 'x x', fn: (n: number, d: number) => (d === 0 ? E.DIV0 : Math.trunc(n / d)) },
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
  SIGN: { args: 'n', fn: Math.sign },
  SQRT: { args: 'n', fn: (x: number) => (x < 0 ? E.NUM : Math.sqrt(x)) },
  SQRTPI: { args: 'x', fn: (x: number) => (x < 0 ? E.NUM : Math.sqrt(x * Math.PI)) },
}
