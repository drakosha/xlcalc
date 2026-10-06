/**
 * Rounding. CEILING … TRUNC ported from formula.js 4.6.1 (src/math-trig.js, src/compatibility.js), Copyright (c)
 * 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); ISO.CEILING and MROUND this engine's own.
 */
import { E, type CErr } from '../../values.js'
import type { FnDef } from '../../args.js'
import { quotient, roundDigits } from '../common.js'

/** CEILING.MATH and FLOOR.MATH: significance taken as positive; `mode` rounds negative numbers away from zero. */
function roundMath(n: number, sig: number | undefined, mode: number | undefined, up: boolean): number {
  const s = Math.abs(sig ?? 1)
  if (s === 0) return 0
  const toward = up ? Math.ceil : Math.floor
  const away = up ? Math.floor : Math.ceil
  return (mode && n < 0 ? away : toward)(quotient(n, s)) * s
}

/** Excel's MROUND rounds a quotient up from 0.5 − 5e-15 (60.5 − 2^-47 rounds down, 0.5 − 2^-48·11/8 up). */
const HALF_SLACK = 5e-15

function mround(n: number, m: number): number | CErr {
  if (m === 0 || n === 0) return 0
  if (n > 0 !== m > 0) return E.NUM
  const q = n / m
  const k = Math.floor(q)
  return (q - k >= 0.5 - HALF_SLACK ? k + 1 : k) * m
}

export const functions: Record<string, FnDef> = {
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
  EVEN: { args: 'n', fn: (x: number) => roundMath(x, 2, 1, true) },
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
  'ISO.CEILING': { args: 'n n?', fn: (n: number, s?: number) => roundMath(n, s, 0, true) },
  MROUND: { args: 'x x', fn: mround },
  ODD: {
    args: 'n',
    fn: (x: number) => {
      let t = Math.ceil(Math.abs(x))
      t = t & 1 ? t : t + 1
      return x >= 0 ? t : -t
    },
  },
  ROUND: { args: 'n i', fn: (x: number, d: number) => roundDigits(x, d, Math.round) },
  ROUNDDOWN: { args: 'n i', fn: (x: number, d: number) => roundDigits(x, d, Math.floor) },
  ROUNDUP: { args: 'n i', fn: (x: number, d: number) => roundDigits(x, d, Math.ceil) },
  TRUNC: { args: 'n i?', fn: (x: number, d = 0) => roundDigits(x, d, Math.floor) },
}
