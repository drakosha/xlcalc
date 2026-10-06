/**
 * BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT on whole numbers of 48 bits. This engine's own (formula.js used
 * JavaScript's 32-bit operators).
 */
import type { FnDef } from '../../args.js'
import { E, type CErr } from '../../values.js'

const MAX = 2 ** 48 - 1
const HALF = 2 ** 24

const valid = (x: number) => Number.isInteger(x) && x >= 0 && x <= MAX

/** A bitwise operation on the two 24-bit halves (JavaScript's operators take 32 bits). */
function bitwise(a: number, b: number, op: (x: number, y: number) => number): number | CErr {
  if (!valid(a) || !valid(b)) return E.NUM
  const ah = Math.floor(a / HALF)
  const bh = Math.floor(b / HALF)
  return op(ah, bh) * HALF + op(a - ah * HALF, b - bh * HALF)
}

/** Excel: zero shifts to zero; a shift past 53 bits is #NUM!, a left shift past 48 bits 0, an overflow #NUM!. */
function shift(n: number, by: number): number | CErr {
  if (!valid(n)) return E.NUM
  const s = Math.trunc(by)
  if (n === 0) return 0
  if (Math.abs(s) > 53) return E.NUM
  if (s < 0) return Math.floor(n / 2 ** -s)
  if (s > 48) return 0
  const r = n * 2 ** s
  return r > MAX ? E.NUM : r
}

export const functions: Record<string, FnDef> = {
  BITAND: { args: 'n n', fn: (a: number, b: number) => bitwise(a, b, (x, y) => x & y) },
  BITOR: { args: 'n n', fn: (a: number, b: number) => bitwise(a, b, (x, y) => x | y) },
  BITXOR: { args: 'n n', fn: (a: number, b: number) => bitwise(a, b, (x, y) => x ^ y) },
  BITLSHIFT: { args: 'n n', fn: (n: number, s: number) => shift(n, s) },
  BITRSHIFT: { args: 'n n', fn: (n: number, s: number) => shift(n, -s) },
}
